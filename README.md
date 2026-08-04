# AuditFlow - 중앙동아리 감사 자동화 도구

은행 PDF 거래내역서, 엑셀 장부 양식, 영수증 사진을 활용해 동아리 회계감사 제출물을 자동화하는 로컬 웹앱 프로젝트입니다.

이 저장소는 포트폴리오 공개를 고려해 다음 두 영역으로 구성됩니다.

## 1. 공개 소개 페이지

`public-site/` 폴더에 있는 정적 웹페이지입니다.

- 프로젝트 문제정의
- 자동화 흐름
- 주요 기능
- 보안/개인정보 설계
- 기술스택

실제 파일 업로드 기능은 공개 페이지에 포함하지 않습니다.

## 2. 비공개 로컬 실행 도구

`web_app.py`를 실행하면 내 PC에서만 작동하는 로컬 웹앱이 열립니다.

기능:

- 은행 거래내역 PDF → 엑셀 거래내역서 자동채움
- 지출증빙 PDF와 거래번호 대조
- 영수증 사진 → 지출증빙자료 PDF 생성
- 검산용 보조 시트 생성

실행:

```bash
run_web_app.bat
```

또는 직접 실행:

```bash
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python -m streamlit run web_app.py
```

## 왜 로컬 실행인가?

감사자료에는 계좌번호, 거래내역, 이름, 영수증 이미지 같은 민감정보가 포함됩니다. 그래서 실제 자동화 도구는 공개 서버가 아니라 사용자의 PC에서만 실행되도록 설계했습니다.

공개 가능한 것은 코드, 구조, 문제해결 방식이고 실제 감사자료는 공개하지 않습니다.

## 기술스택

- Python
- Streamlit
- openpyxl
- pypdf
- ReportLab
- Pillow
- HTML/CSS/JavaScript

## GitHub 공개 전 주의

실제 PDF, XLSX, 영수증 이미지는 절대 커밋하지 마세요. `.gitignore`에서 기본적으로 차단하고 있지만, 커밋 전에는 반드시 아래 문서를 확인하세요.

- `docs/GITHUB_SAFE_CHECKLIST.md`

## 포트폴리오 관점의 핵심 포인트

- 반복 업무 자동화
- 문서 처리 파이프라인 설계
- 민감정보를 고려한 로컬 우선 아키텍처
- 사용자가 직접 쓰는 업무형 웹앱 구현
- PDF/엑셀/이미지 파일 처리 자동화