# 찰칵혁신단 문서 도우미

Streamlit으로 배포하는 문서 자동화 앱입니다. 회의 사진을 사진대지로 정리하거나, 영수증 사진을 읽어 물품보고서 형태로 정리할 수 있습니다.

## 주요 기능

### 1. 회의 사진 대지

- 회의명과 회의 날짜 입력
- 여러 장의 사진 업로드
- 사진 순서 변경
- 다운로드 전 미리보기
- DOCX 및 PDF 다운로드

### 2. 영수증 물품보고서

- 영수증 사진 업로드
- OpenAI Vision 기반 영수증 OCR 분석
- OpenAI API 키가 없거나 API 요청이 실패하면 로컬 Tesseract OCR 대체 시도
- 상호명, 구매날짜, 물품명, 규격, 단위, 수량, 금액 자동 정리
- 검사항목 기본값: `물품상태 및 수량 등`
- 검사결과 기본값: `이상없음`
- 검사일자 기본값: 오늘 날짜
- 사용자가 모든 항목을 직접 수정 가능
- 다운로드 전 `이상없음을 확인했습니다` 체크 필수
- DOCX, Excel(XLSX), PDF 다운로드

> ⚠️ OCR 결과는 항상 틀릴 수 있습니다. 물품보고서 다운로드 전 상호명, 구매날짜, 물품명, 수량, 금액을 반드시 사람이 직접 확인해야 합니다.

## 프로젝트 구조

```text
.
├── app.py          # Streamlit 메인 화면 및 페이지 라우팅
├── photo.py        # 회의 사진대지 페이지
├── receipt.py      # 영수증 물품보고서 페이지
├── requirements.txt
└── README.md
```

## 로컬 실행 방법

### 1. 패키지 설치

```bash
python -m pip install -r requirements.txt
```

### 2. OpenAI API 키 설정

영수증 OCR에서 OpenAI Vision을 사용하려면 Streamlit secrets에 API 키를 설정해야 합니다.

로컬에서는 프로젝트 루트에 `.streamlit/secrets.toml` 파일을 만들고 아래처럼 입력합니다.

```toml
OPENAI_API_KEY = "sk-..."
```

OpenAI 크레딧이 없거나 API 키가 없으면 앱은 로컬 Tesseract OCR을 시도합니다. 다만 Tesseract는 Python 패키지가 아니라 서버에 별도 설치되어 있어야 합니다.

### 3. Streamlit 실행

```bash
streamlit run app.py
```

브라우저에서 안내되는 Local URL로 접속합니다.

## Streamlit Community Cloud 배포 방법

1. 이 저장소를 GitHub에 올립니다.
2. [Streamlit Community Cloud](https://streamlit.io/cloud)에서 새 앱을 생성합니다.
3. 앱 진입 파일은 `app.py`로 설정합니다.
4. `requirements.txt`가 자동으로 설치됩니다.
5. 앱 설정의 **Secrets**에 아래 값을 추가합니다.

```toml
OPENAI_API_KEY = "sk-..."
```

6. 배포 후 홈 화면에서 `회의 사진 대지` 또는 `영수증 물품보고서` 기능을 선택합니다.

## OCR 관련 주의사항

### OpenAI Vision 사용 권장

Streamlit Community Cloud 배포 환경에서는 로컬 Tesseract 설치가 제한될 수 있습니다. 따라서 영수증 OCR은 OpenAI API 키와 충분한 크레딧을 설정해 사용하는 것을 권장합니다.

### Tesseract 대체 OCR

앱에는 OpenAI API 키가 없거나 OpenAI 요청이 실패할 때 Tesseract OCR을 시도하는 코드가 포함되어 있습니다. 하지만 배포 서버에 아래 실행 파일과 언어팩이 설치되어 있어야 정상 동작합니다.

- `tesseract`
- 한국어 언어팩(`kor`)
- 영어 언어팩(`eng`)

Streamlit Community Cloud에서 시스템 패키지 설치가 필요하면 별도 패키지 설정 파일이 필요할 수 있습니다. 해당 환경에서 Tesseract 설치가 불가능하거나 정확도가 부족하면 OpenAI 크레딧을 추가 구매해 Vision OCR을 사용하는 편이 더 안정적입니다.

## 파일 다운로드 형식

| 기능 | 다운로드 형식 |
| --- | --- |
| 회의 사진 대지 | DOCX, PDF |
| 영수증 물품보고서 | DOCX, XLSX, PDF |

## 검증 명령어

개발 중에는 아래 명령어로 기본 동작을 확인할 수 있습니다.

```bash
python -m py_compile app.py photo.py receipt.py
```

```bash
streamlit run app.py
```

## 운영 시 확인할 점

- 영수증 OCR 결과는 자동 생성값이므로 다운로드 전 사람이 반드시 확인해야 합니다.
- 숫자, 금액, 수량은 임의로 보정하지 말고 영수증 원본과 대조해 수정해야 합니다.
- OpenAI API 크레딧이 부족하면 영수증 분석이 실패할 수 있습니다.
- PDF 생성은 서버에 설치된 한글 폰트 상태에 따라 글꼴 모양이 달라질 수 있습니다.

---

made by 찰칵혁신단
