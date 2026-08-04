from __future__ import annotations

from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st

from audit_web_core import build_workbook
from receipt_evidence_builder import EvidenceItem, build_evidence_pdf, normalize_date_text


st.set_page_config(
    page_title="중앙동아리 감사 자동화",
    page_icon="📄",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 2.2rem; padding-bottom: 3rem; }
    .hero {
        padding: 28px 30px;
        border-radius: 24px;
        background: linear-gradient(135deg, #10233f 0%, #1d4f7a 52%, #2c8a8a 100%);
        color: white;
        box-shadow: 0 18px 48px rgba(16, 35, 63, .22);
        margin-bottom: 22px;
    }
    .hero h1 { margin: 0 0 8px 0; font-size: 2.15rem; letter-spacing: -0.04em; }
    .hero p { margin: 0; opacity: .88; font-size: 1.02rem; line-height: 1.65; }
    .soft-card {
        padding: 18px 18px;
        border: 1px solid #e8edf3;
        border-radius: 18px;
        background: #fbfcfe;
    }
    .small-muted { color: #64748b; font-size: .92rem; }
    </style>
    <div class="hero">
      <h1>중앙동아리 감사 자동화</h1>
      <p>은행 PDF로 엑셀 거래내역을 채우고, 영수증 사진으로 지출증빙자료 PDF까지 만듭니다. 파일은 이 PC 안에서만 처리됩니다.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

excel_tab, receipt_tab = st.tabs(["엑셀 자동채움", "지출증빙 PDF 생성"])

with excel_tab:
    left, right = st.columns([1.15, 0.85], gap="large")

    with left:
        st.subheader("1. 파일 올리기")
        template_file = st.file_uploader(
            "엑셀 양식 파일 (.xlsx)",
            type=["xlsx"],
            help="예: 2026-01 TUBE 회계(26.05.26~26.07.17).xlsx",
            key="excel_template",
        )
        bank_files = st.file_uploader(
            "은행 거래내역서 PDF - 여러 개 가능",
            type=["pdf"],
            accept_multiple_files=True,
            help="카카오뱅크, 토스뱅크 거래내역서를 함께 올릴 수 있습니다.",
            key="bank_pdfs",
        )
        evidence_file = st.file_uploader(
            "지출증빙자료 PDF - 선택이지만 권장",
            type=["pdf"],
            help="올리면 증빙번호/항목명/감사비고까지 보완합니다.",
            key="evidence_pdf_for_excel",
        )

        with st.expander("고급 옵션", expanded=False):
            start_number = st.number_input("엑셀 A열 시작 번호", min_value=1, value=1, step=1, key="start_number")
            include_helper_sheets = st.checkbox("검산용 보조 시트 포함", value=True, key="helper_sheets")
            output_name = st.text_input("다운로드 파일명", value="감사기준보완_자동채움.xlsx", key="excel_output_name")

    with right:
        st.subheader("2. 결과에 들어가는 내용")
        st.markdown(
            """
            <div class="soft-card">
            <b>기본 입력</b><br>
            날짜, 이름, 내용, 입금, 출금, 거래 후 잔액<br><br>
            <b>감사 보완</b><br>
            감사구분, 증빙번호, 증빙항목, 감사비고<br><br>
            <b>검산 시트</b><br>
            원문 거래내역, 증빙목록, 요약<br><br>
            <span class="small-muted">※ 자동 결과는 감사자가 빠르게 확인하기 위한 초안입니다. “제출 전 확인” 메모가 있는 항목은 직접 한 번 더 확인하세요.</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    ready = bool(template_file and bank_files)
    if not ready:
        st.info("엑셀 양식과 은행 거래내역서 PDF를 올리면 자동채움 버튼이 활성화됩니다.")

    if st.button("엑셀 자동채움 생성", type="primary", disabled=not ready, use_container_width=True):
        with st.spinner("PDF를 읽고 엑셀 양식을 채우는 중입니다..."):
            try:
                with TemporaryDirectory() as tmp:
                    tmp_path = Path(tmp)
                    template_path = tmp_path / template_file.name
                    template_path.write_bytes(template_file.getbuffer())

                    bank_paths: list[Path] = []
                    for idx, uploaded in enumerate(bank_files, 1):
                        path = tmp_path / f"bank_{idx}_{uploaded.name}"
                        path.write_bytes(uploaded.getbuffer())
                        bank_paths.append(path)

                    evidence_path = None
                    if evidence_file is not None:
                        evidence_path = tmp_path / evidence_file.name
                        evidence_path.write_bytes(evidence_file.getbuffer())

                    safe_name = output_name.strip() or "감사기준보완_자동채움.xlsx"
                    if not safe_name.lower().endswith(".xlsx"):
                        safe_name += ".xlsx"
                    output_path = tmp_path / safe_name

                    result = build_workbook(
                        template_path=template_path,
                        bank_pdf_paths=bank_paths,
                        evidence_pdf_path=evidence_path,
                        output_path=output_path,
                        start_number=int(start_number),
                        include_helper_sheets=include_helper_sheets,
                    )
                    output_bytes = output_path.read_bytes()

                st.success("완성본이 생성되었습니다.")
                metric_cols = st.columns(5)
                metric_cols[0].metric("전체 거래", f"{result.transaction_count:,}건")
                metric_cols[1].metric("출금 거래", f"{result.withdrawal_count:,}건")
                metric_cols[2].metric("증빙 연결", f"{result.matched_withdrawals:,}건")
                metric_cols[3].metric("총 입금", f"{result.deposit_sum:,}원")
                metric_cols[4].metric("총 출금", f"{result.withdrawal_sum:,}원")

                if result.mismatch_candidates:
                    st.warning(f"증빙번호와 엑셀번호가 다른 후보가 {result.mismatch_candidates}건 있습니다. 결과 파일의 감사비고를 확인하세요.")
                if result.warnings:
                    with st.expander("처리 중 참고할 메시지"):
                        for warning in result.warnings:
                            st.write("- " + warning)

                st.download_button(
                    "완성 엑셀 다운로드",
                    data=output_bytes,
                    file_name=safe_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"생성 중 문제가 생겼습니다: {type(exc).__name__}: {exc}")
                st.caption("파일 형식이 다른 은행 PDF라면 먼저 샘플 구조를 확인해야 합니다.")

with receipt_tab:
    st.subheader("영수증 사진으로 지출증빙자료 PDF 만들기")
    st.write("각 항목마다 번호, 항목명, 날짜, 비고, 영수증 사진을 넣으면 A4 PDF로 정리합니다.")

    item_count = st.number_input("증빙 항목 수", min_value=1, max_value=80, value=3, step=1)
    pdf_title = st.text_input("PDF 제목", value="지출증빙자료")
    receipt_output_name = st.text_input("PDF 파일명", value="지출증빙자료_자동생성.pdf")

    items: list[dict] = []
    for idx in range(int(item_count)):
        with st.expander(f"증빙 항목 {idx + 1}", expanded=idx == 0):
            cols = st.columns([0.55, 1.4, 0.8])
            number = cols[0].number_input("번호", min_value=1, value=idx + 1, step=1, key=f"receipt_no_{idx}")
            title = cols[1].text_input("항목명", key=f"receipt_title_{idx}", placeholder="예: 입학식 부스 참여자 상품")
            spent_date = cols[2].text_input("날짜", key=f"receipt_date_{idx}", placeholder="26.03.04")
            note = st.text_area("비고", key=f"receipt_note_{idx}", placeholder="예: 부스 참여자 지급용 상품입니다.", height=80)
            images = st.file_uploader(
                "영수증 사진",
                type=["png", "jpg", "jpeg", "webp"],
                accept_multiple_files=True,
                key=f"receipt_images_{idx}",
            )
            items.append({"number": int(number), "title": title.strip(), "date": spent_date.strip(), "note": note.strip(), "images": images})

    valid_items = [item for item in items if item["title"] and item["date"] and item["images"]]
    if len(valid_items) != len(items):
        st.info("항목명, 날짜, 영수증 사진이 모두 있는 항목만 PDF에 들어갑니다.")

    if st.button("지출증빙 PDF 생성", type="primary", disabled=not valid_items, use_container_width=True):
        with st.spinner("영수증 사진을 A4 증빙자료 PDF로 정리하는 중입니다..."):
            try:
                with TemporaryDirectory() as tmp:
                    tmp_path = Path(tmp)
                    evidence_items: list[EvidenceItem] = []
                    for item_idx, item in enumerate(valid_items, 1):
                        image_paths = []
                        for image_idx, uploaded in enumerate(item["images"], 1):
                            suffix = Path(uploaded.name).suffix or ".jpg"
                            image_path = tmp_path / f"receipt_{item_idx}_{image_idx}{suffix}"
                            image_path.write_bytes(uploaded.getbuffer())
                            image_paths.append(image_path)
                        evidence_items.append(
                            EvidenceItem(
                                number=item["number"],
                                title=item["title"],
                                spent_date=normalize_date_text(item["date"]),
                                note=item["note"],
                                image_paths=image_paths,
                            )
                        )

                    safe_pdf_name = receipt_output_name.strip() or "지출증빙자료_자동생성.pdf"
                    if not safe_pdf_name.lower().endswith(".pdf"):
                        safe_pdf_name += ".pdf"
                    output_path = tmp_path / safe_pdf_name
                    build_evidence_pdf(evidence_items, output_path, title=pdf_title.strip() or "지출증빙자료")
                    pdf_bytes = output_path.read_bytes()

                st.success(f"지출증빙 PDF가 생성되었습니다. 총 {len(valid_items)}개 항목입니다.")
                st.download_button(
                    "지출증빙 PDF 다운로드",
                    data=pdf_bytes,
                    file_name=safe_pdf_name,
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"PDF 생성 중 문제가 생겼습니다: {type(exc).__name__}: {exc}")

st.divider()
st.caption(f"로컬 처리 웹앱 · {datetime.now().strftime('%Y-%m-%d')} · 업로드한 파일은 웹앱 실행 중 임시 공간에서만 사용됩니다.")