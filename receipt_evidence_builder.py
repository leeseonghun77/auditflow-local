from __future__ import annotations

import csv
import io
import json
import re
from dataclasses import dataclass, asdict
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageOps


@dataclass
class EvidenceItem:
    number: int
    title: str
    spent_date: str
    note: str = ""
    image_paths: list[Path] | None = None


def _require_reportlab():
    try:
        import reportlab  # noqa: F401
    except ImportError as exc:
        raise RuntimeError("reportlab가 설치되어 있지 않습니다. python -m pip install -r requirements.txt 를 실행하세요.") from exc


def normalize_date_text(value: str | date | datetime) -> str:
    if isinstance(value, datetime):
        return value.strftime("%y.%m.%d")
    if isinstance(value, date):
        return value.strftime("%y.%m.%d")
    text = str(value or "").strip()
    if not text:
        return ""
    patterns = [
        r"(20\d{2})[.\-/년 ]+(\d{1,2})[.\-/월 ]+(\d{1,2})",
        r"(\d{2})[.\-/년 ]+(\d{1,2})[.\-/월 ]+(\d{1,2})",
        r"(\d{1,2})[.\-/월 ]+(\d{1,2})",
    ]
    for idx, pattern in enumerate(patterns):
        m = re.search(pattern, text)
        if not m:
            continue
        parts = m.groups()
        if idx == 0:
            y, mo, d = int(parts[0]) % 100, int(parts[1]), int(parts[2])
        elif idx == 1:
            y, mo, d = int(parts[0]), int(parts[1]), int(parts[2])
        else:
            y, mo, d = datetime.now().year % 100, int(parts[0]), int(parts[1])
        return f"{y:02d}.{mo:02d}.{d:02d}"
    return text


def safe_filename(text: str, fallback: str = "evidence") -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]+", "_", text).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or fallback


def read_items_from_csv(path: Path, image_dir: Path | None = None) -> list[EvidenceItem]:
    items: list[EvidenceItem] = []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            number = int(str(row.get("번호") or row.get("number") or "0").strip())
            title = str(row.get("항목명") or row.get("title") or "").strip()
            spent_date = normalize_date_text(str(row.get("날짜") or row.get("date") or "").strip())
            note = str(row.get("비고") or row.get("note") or "").strip()
            image_text = str(row.get("이미지") or row.get("image") or row.get("image_paths") or "").strip()
            paths: list[Path] = []
            for piece in re.split(r"[;,]", image_text):
                piece = piece.strip()
                if not piece:
                    continue
                p = Path(piece)
                if not p.is_absolute() and image_dir:
                    p = image_dir / p
                paths.append(p)
            items.append(EvidenceItem(number, title, spent_date, note, paths))
    return items


def load_items_from_json(path: Path, image_dir: Path | None = None) -> list[EvidenceItem]:
    data = json.loads(path.read_text(encoding="utf-8"))
    rows = data if isinstance(data, list) else data.get("items", [])
    items: list[EvidenceItem] = []
    for row in rows:
        paths = []
        for piece in row.get("images", row.get("image_paths", [])):
            p = Path(piece)
            if not p.is_absolute() and image_dir:
                p = image_dir / p
            paths.append(p)
        items.append(EvidenceItem(int(row["number"]), str(row["title"]), normalize_date_text(row["date"]), str(row.get("note", "")), paths))
    return items


def _register_korean_font():
    _require_reportlab()
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    candidates = [
        Path("C:/Windows/Fonts/malgun.ttf"),
        Path("C:/Windows/Fonts/malgunbd.ttf"),
        Path("C:/Windows/Fonts/NanumGothic.ttf"),
    ]
    for font_path in candidates:
        if font_path.exists():
            font_name = "KoreanBaseFont"
            if font_name not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(font_name, str(font_path)))
            return font_name
    return "Helvetica"


def _image_reader_for_reportlab(image_path: Path):
    from reportlab.lib.utils import ImageReader

    with Image.open(image_path) as image:
        image = ImageOps.exif_transpose(image)
        if image.mode not in ("RGB", "L"):
            image = image.convert("RGB")
        buffer = io.BytesIO()
        image.save(buffer, format="JPEG", quality=92)
        buffer.seek(0)
        return ImageReader(buffer), image.width, image.height


def _draw_wrapped_text(canvas, text: str, x: float, y: float, max_width: float, font_name: str, font_size: int, leading: int) -> float:
    from reportlab.pdfbase.pdfmetrics import stringWidth

    text = str(text or "")
    words = list(text) if re.search(r"[가-힣]", text) else text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        sep = "" if re.search(r"[가-힣]", text) else " "
        trial = current + (sep if current else "") + word
        if stringWidth(trial, font_name, font_size) <= max_width:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    for line in lines[:5]:
        canvas.drawString(x, y, line)
        y -= leading
    return y


def build_evidence_pdf(items: list[EvidenceItem], output_path: Path, *, title: str = "지출증빙자료") -> Path:
    _require_reportlab()
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    font_name = _register_korean_font()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(output_path), pagesize=A4)
    width, height = A4
    margin = 42

    for item in sorted(items, key=lambda x: x.number):
        image_paths = item.image_paths or []
        if not image_paths:
            image_paths = [None]
        for image_index, image_path in enumerate(image_paths, 1):
            c.setTitle(title)
            c.setStrokeColor(colors.HexColor("#D8DEE9"))
            c.setLineWidth(1)
            c.rect(margin, margin, width - margin * 2, height - margin * 2)

            c.setFillColor(colors.HexColor("#111827"))
            c.setFont(font_name, 16)
            c.drawString(margin + 12, height - margin - 30, title)

            c.setFont(font_name, 11)
            y = height - margin - 64
            c.drawString(margin + 12, y, f"번호 {item.number}")
            c.drawString(margin + 92, y, f"항목명 {item.title}")
            c.drawRightString(width - margin - 12, y, f"날짜 {normalize_date_text(item.spent_date)}")

            y -= 22
            c.setStrokeColor(colors.HexColor("#CBD5E1"))
            c.line(margin + 12, y, width - margin - 12, y)

            y -= 24
            c.setFont(font_name, 11)
            c.drawString(margin + 12, y, "비고")
            y -= 18
            c.setFont(font_name, 10)
            note = item.note or ""
            if image_path and len(image_paths) > 1:
                note = f"{note} (첨부 {image_index}/{len(image_paths)})".strip()
            y = _draw_wrapped_text(c, note, margin + 12, y, width - margin * 2 - 24, font_name, 10, 14)

            image_top = y - 18
            image_bottom = margin + 18
            image_box_h = image_top - image_bottom
            image_box_w = width - margin * 2 - 24
            image_x = margin + 12
            image_y = image_bottom

            c.setStrokeColor(colors.HexColor("#E2E8F0"))
            c.setFillColor(colors.HexColor("#F8FAFC"))
            c.roundRect(image_x, image_y, image_box_w, image_box_h, 10, fill=1, stroke=1)

            if image_path is None:
                c.setFillColor(colors.HexColor("#64748B"))
                c.setFont(font_name, 13)
                c.drawCentredString(width / 2, image_y + image_box_h / 2, "영수증 이미지 없음")
            else:
                image_path = Path(image_path)
                if not image_path.exists():
                    c.setFillColor(colors.HexColor("#B91C1C"))
                    c.setFont(font_name, 12)
                    c.drawCentredString(width / 2, image_y + image_box_h / 2, f"이미지 파일을 찾지 못했습니다: {image_path.name}")
                else:
                    reader, img_w, img_h = _image_reader_for_reportlab(image_path)
                    scale = min((image_box_w - 24) / img_w, (image_box_h - 24) / img_h)
                    draw_w = img_w * scale
                    draw_h = img_h * scale
                    draw_x = image_x + (image_box_w - draw_w) / 2
                    draw_y = image_y + (image_box_h - draw_h) / 2
                    c.drawImage(reader, draw_x, draw_y, draw_w, draw_h, preserveAspectRatio=True, mask="auto")

            c.showPage()

    c.save()
    return output_path


def make_example_csv(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    rows = [
        ["번호", "항목명", "날짜", "비고", "이미지"],
        ["1", "예시 지출 항목", "26.03.04", "영수증 사진을 같은 폴더에 넣고 이미지 파일명을 적으세요.", "receipt_001.jpg"],
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerows(rows)
    return path


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description="영수증 이미지로 지출증빙자료 PDF 생성")
    parser.add_argument("--csv", help="번호/항목명/날짜/비고/이미지 열을 가진 CSV")
    parser.add_argument("--json", help="증빙 항목 JSON")
    parser.add_argument("--image-dir", default=".", help="이미지 파일 기준 폴더")
    parser.add_argument("--output", default="outputs/지출증빙자료_자동생성.pdf", help="생성할 PDF 경로")
    parser.add_argument("--example-csv", action="store_true", help="CSV 예시 파일 생성")
    args = parser.parse_args()

    if args.example_csv:
        make_example_csv(Path("receipt_items_template.csv"))
        print("receipt_items_template.csv 생성 완료")
        return 0

    image_dir = Path(args.image_dir)
    if args.csv:
        items = read_items_from_csv(Path(args.csv), image_dir=image_dir)
    elif args.json:
        items = load_items_from_json(Path(args.json), image_dir=image_dir)
    else:
        raise SystemExit("--csv 또는 --json 중 하나를 지정하세요. 예시가 필요하면 --example-csv를 사용하세요.")

    output = build_evidence_pdf(items, Path(args.output))
    print(f"지출증빙자료 PDF 생성 완료: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())