import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image, ImageOps, ImageEnhance, ImageFilter
import io
import base64
import json
import requests
import re
from datetime import datetime


def fix_orientation(img):
    """EXIF 회전값을 실제 픽셀에 반영합니다."""
    try:
        return ImageOps.exif_transpose(img)
    except Exception:
        return img


def _resize_for_vision(img, min_width=1200, max_side=3500):
    """
    영수증 사진은 폭이 너무 작으면 품명/수량 열을 모델이 자주 틀립니다.
    311px 같은 작은 이미지는 먼저 키우고, 너무 큰 이미지는 API 비용/오류를 막기 위해 줄입니다.
    """
    w, h = img.size

    if w < min_width:
        ratio = min_width / w
        img = img.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)

    w, h = img.size
    if max(w, h) > max_side:
        ratio = max_side / max(w, h)
        img = img.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)

    return img


def make_receipt_vision_images(img):
    """
    모델에 2장을 보냅니다.
    1) 원본 보정본: 색/배경 맥락 유지
    2) 흑백 강조본: 작은 글씨/숫자열 판독 강화
    """
    img = fix_orientation(img).convert("RGB")
    color = _resize_for_vision(img)

    gray = ImageOps.grayscale(color)
    gray = ImageOps.autocontrast(gray, cutoff=1)
    gray = ImageEnhance.Contrast(gray).enhance(1.8)
    gray = gray.filter(ImageFilter.UnsharpMask(radius=1, percent=170, threshold=3))
    gray = gray.convert("RGB")

    return color, gray


def image_to_data_url(img, fmt="JPEG", quality=95):
    buf = io.BytesIO()
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    img.save(buf, format=fmt, quality=quality, optimize=True)
    mime = "image/jpeg" if fmt.upper() == "JPEG" else "image/png"
    return f"data:{mime};base64,{base64.b64encode(buf.getvalue()).decode('utf-8')}"


# 기존 다른 코드가 img_to_base64를 호출해도 깨지지 않게 유지
# 단, read_receipt_with_ai에서는 data URL을 직접 사용합니다.
def img_to_base64(img):
    color, _ = make_receipt_vision_images(img)
    buf = io.BytesIO()
    color.save(buf, format="JPEG", quality=95, optimize=True)
    return base64.b64encode(buf.getvalue()).decode("utf-8")


RECEIPT_SCHEMA = {
    "name": "korean_receipt_ocr",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "store": {"type": "string"},
            "date": {"type": "string"},
            "lines": {
                "type": "array",
                "items": {"type": "string"}
            },
            "items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "raw_line": {"type": "string"},
                        "name": {"type": "string"},
                        "spec": {"type": "string"},
                        "quantity": {"type": "number"},
                        "unit_price": {"type": "number"},
                        "amount": {"type": "number"}
                    },
                    "required": ["raw_line", "name", "spec", "quantity", "unit_price", "amount"],
                    "additionalProperties": False
                }
            },
            "warnings": {
                "type": "array",
                "items": {"type": "string"}
            }
        },
        "required": ["store", "date", "lines", "items", "warnings"],
        "additionalProperties": False
    }
}


RECEIPT_PROMPT = """너는 한국 영수증 OCR 엔진이다. 최종 판단보다 원문 보존이 더 중요하다.

해야 할 일:
1. 영수증에서 보이는 텍스트 줄을 위에서 아래 순서대로 lines에 최대한 그대로 넣어라.
2. 상품 구매로 보이는 줄은 items에도 넣어라.
3. 상품행은 raw_line에 한 줄 원문을 반드시 넣어라.
4. 상품명 안의 숫자, 괄호, 용량, 규격, 세트표기는 절대 떼지 마라.
   예: 25P, 21*21cm, 500ml, 2입, A4, 80g, 3개입은 상품명 일부다.
5. 단가/수량/금액이 보이면 그대로 넣되, 확실하지 않으면 raw_line 중심으로 넣어라.
6. 날짜는 YYYY-MM-DD로 가능하면 넣고, 아니면 빈 문자열.
7. JSON 외 텍스트를 절대 쓰지 마라.

주의:
- 이마트/다이소/편의점/식당/온라인몰/카드전표 등 형식이 다를 수 있다.
- 오른쪽 숫자열이 [단가 수량 금액]일 수도 있고, [수량 단가 금액]일 수도 있고, 금액만 있을 수도 있다.
- Python 후처리가 최종 검증을 하므로 lines와 raw_line을 최대한 정확히 보존하는 것이 최우선이다.
"""


_AMOUNT_TOKEN = r"(?:\d{1,3}(?:,\d{3})+|\d+)"
_NUM_RE = re.compile(_AMOUNT_TOKEN)

# 진짜 합계/결제/세금 줄만 제외. '봉투'는 실제 구매품일 수 있어서 제외하지 않는다.
_EXCLUDE_NAME_RE = re.compile(
    r"(과세\s*합계|부가\s*세|판매\s*합계|총\s*합계|합\s*계|받은\s*금액|거스름|"
    r"신용\s*카드|체크\s*카드|현금\s*영수증|승인|카드번호|포인트|멤버십|"
    r"사업자|대표|전화|주소|POS|교환|환불|영수증|바코드)",
    re.IGNORECASE,
)
_CODE_ONLY_RE = re.compile(r"^\s*\[?\s*\d{4,}\s*\]?\s*$")


def _clean_space(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _to_int(value):
    if isinstance(value, bool):
        return 0
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(round(value))
    s = re.sub(r"[^0-9]", "", str(value or ""))
    return int(s) if s else 0


def _normalize_date(value):
    s = str(value or "")
    m = re.search(r"(20\d{2})\D+(\d{1,2})\D+(\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
        return f"{y:04d}-{mo:02d}-{d:02d}"
    m = re.search(r"(\d{2})[./-](\d{1,2})[./-](\d{1,2})", s)
    if m:
        yy, mo, d = map(int, m.groups())
        return f"20{yy:02d}-{mo:02d}-{d:02d}"
    return ""


def _strip_code_lines(text):
    # 상품코드 [1002504] 같은 토큰만 제거. 상품명 내부 숫자는 보존.
    text = re.sub(r"\s*\[\s*\d{4,}\s*\]\s*", " ", str(text or ""))
    return _clean_space(text.replace("원", ""))


def _looks_like_non_item_name(name):
    name = _clean_space(name)
    if not name:
        return True
    if _CODE_ONLY_RE.match(name):
        return True
    if _EXCLUDE_NAME_RE.search(name):
        return True
    if not re.search(r"[가-힣A-Za-z]", name):
        return True
    return False


def _parse_item_raw_line(raw_line):
    """
    영수증 종류별 숫자 패턴을 Python이 최종 판정한다.
    지원:
    - 다이소형: 품명 단가 수량 금액
    - 마트형: 품명 수량 단가 금액
    - 일반형: 품명 금액  -> 수량 1, 단가=금액
    - x 표기형: 품명 2 x 1,000 2,000
    """
    raw = _strip_code_lines(raw_line)
    if not raw or _EXCLUDE_NAME_RE.search(raw):
        return None

    # 할인/취소/마이너스 줄은 일단 제외
    if re.search(r"(할인|에누리|쿠폰|취소|반품)", raw):
        return None

    # 1) x 표기: 품명 2 x 1,000 2,000 또는 품명 1,000 x 2 2,000
    m = re.match(rf"^(.+?)\s+({_AMOUNT_TOKEN})\s*[xX×]\s*({_AMOUNT_TOKEN})\s+({_AMOUNT_TOKEN})\s*$", raw)
    if m:
        name = _clean_space(m.group(1))
        a, b, total = _to_int(m.group(2)), _to_int(m.group(3)), _to_int(m.group(4))
        candidates = []
        if a * b == total:
            # 보통 작은 숫자는 수량, 큰 숫자는 단가
            if a <= 99 and b > 99:
                candidates.append((b, a, total))
            elif b <= 99 and a > 99:
                candidates.append((a, b, total))
            else:
                candidates.append((b, a, total))
        if candidates and not _looks_like_non_item_name(name):
            unit_price, quantity, amount = candidates[0]
            return {"raw_line": raw, "name": name, "spec": "", "quantity": quantity, "unit_price": unit_price, "amount": amount}

    nums = list(_NUM_RE.finditer(raw))
    if not nums:
        return None

    # 2) 오른쪽 끝 숫자 3개에서 산술검증. 품명 내부 숫자는 앞쪽에 남긴다.
    if len(nums) >= 3:
        best = None
        for i in range(len(nums) - 3, -1, -1):
            n1 = _to_int(nums[i].group())
            n2 = _to_int(nums[i + 1].group())
            n3 = _to_int(nums[i + 2].group())
            if n1 <= 0 or n2 <= 0 or n3 <= 0:
                continue

            name = _clean_space(raw[:nums[i].start()])
            if _looks_like_non_item_name(name):
                continue

            # 단가 수량 금액
            if n1 * n2 == n3:
                best = (name, n2, n1, n3)
                break
            # 수량 단가 금액
            if n2 * n1 == n3:
                best = (name, n1, n2, n3)
                break

        if best:
            name, quantity, unit_price, amount = best
            return {"raw_line": raw, "name": name, "spec": "", "quantity": quantity, "unit_price": unit_price, "amount": amount}

    # 3) 금액만 있는 상품행: 품명 금액 -> 수량 1
    last = nums[-1]
    amount = _to_int(last.group())
    name = _clean_space(raw[:last.start()])
    if amount > 0 and not _looks_like_non_item_name(name):
        # 너무 작은 숫자만 있으면 코드/수량일 가능성이 높음
        if amount >= 100:
            return {"raw_line": raw, "name": name, "spec": "", "quantity": 1, "unit_price": amount, "amount": amount}

    return None


def _merge_wrapped_lines(lines):
    """온라인몰/마트처럼 상품명이 다음 줄로 이어지는 경우를 조금 더 버티게 합친다."""
    cleaned = [_strip_code_lines(x) for x in lines or []]
    cleaned = [x for x in cleaned if x and not _CODE_ONLY_RE.match(x)]

    merged = []
    pending = ""
    for line in cleaned:
        if _EXCLUDE_NAME_RE.search(line):
            if pending:
                merged.append(pending)
                pending = ""
            continue

        parsed = _parse_item_raw_line(line)
        if parsed:
            if pending and not _parse_item_raw_line(pending):
                combined = _clean_space(pending + " " + line)
                if _parse_item_raw_line(combined):
                    merged.append(combined)
                    pending = ""
                    continue
                merged.append(pending)
                pending = ""
            merged.append(line)
            continue

        # 숫자가 거의 없고 글자가 있으면 다음 줄 상품명 후보로 보관
        if re.search(r"[가-힣A-Za-z]", line):
            if pending:
                merged.append(pending)
            pending = line

    if pending:
        merged.append(pending)
    return merged


def _normalize_receipt_result(data):
    warnings = list(data.get("warnings") or [])
    normalized = {
        "store": _clean_space(data.get("store", "")),
        "date": _normalize_date(data.get("date", "")),
        "items": [],
        "warnings": warnings,
    }

    candidates = []
    for line in data.get("lines", []) or []:
        candidates.append(line)
    for item in data.get("items", []) or []:
        candidates.append(item.get("raw_line") or "")

    for line in _merge_wrapped_lines(candidates):
        parsed = _parse_item_raw_line(line)
        if parsed:
            normalized["items"].append(parsed)

    # raw_line에서 못 잡힌 것만 AI 필드를 보조로 사용
    for item in data.get("items", []) or []:
        raw_line = item.get("raw_line", "")
        name = _clean_space(item.get("name", ""))
        unit_price = _to_int(item.get("unit_price"))
        quantity = _to_int(item.get("quantity"))
        amount = _to_int(item.get("amount"))

        if _looks_like_non_item_name(name):
            continue
        if amount <= 0:
            continue
        if quantity <= 0:
            quantity = 1
        if unit_price <= 0:
            unit_price = amount // quantity if quantity else amount
        if unit_price * quantity != amount:
            # OCR이 금액만 정확한 경우는 1개 구매로 살림
            if quantity == 1:
                unit_price = amount
            else:
                normalized["warnings"].append(f"금액 검증 실패로 제외: {raw_line or name}")
                continue

        normalized["items"].append({
            "raw_line": _clean_space(raw_line),
            "name": name,
            "spec": "",
            "quantity": quantity,
            "unit_price": unit_price,
            "amount": amount,
        })

    # 중복 제거
    deduped = []
    seen = set()
    for it in normalized["items"]:
        key = (it["name"], it["quantity"], it["unit_price"], it["amount"])
        if key in seen:
            continue
        seen.add(key)
        deduped.append(it)
    normalized["items"] = deduped

    return normalized


def read_receipt_with_ai(img):
    color_img, contrast_img = make_receipt_vision_images(img)
    color_url = image_to_data_url(color_img)
    contrast_url = image_to_data_url(contrast_img)
    api_key = st.secrets["OPENAI_API_KEY"]

    payload = {
        "model": "gpt-4o-2024-08-06",
        "temperature": 0,
        "max_tokens": 3000,
        "response_format": {
            "type": "json_schema",
            "json_schema": RECEIPT_SCHEMA,
        },
        "messages": [
            {
                "role": "system",
                "content": "너는 한국 영수증 OCR 엔진이다. 원문 줄 보존을 최우선으로 하고, JSON만 반환한다.",
            },
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": RECEIPT_PROMPT + "\n첫 번째 이미지는 원본 보정본, 두 번째 이미지는 흑백 대비 강화본이다. 둘을 대조해서 lines를 최대한 정확히 작성해라."},
                    {"type": "image_url", "image_url": {"url": color_url, "detail": "high"}},
                    {"type": "image_url", "image_url": {"url": contrast_url, "detail": "high"}},
                ],
            },
        ],
    }

    try:
        res = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
            },
            json=payload,
            timeout=60,
        )
        res.raise_for_status()
        res_json = res.json()

        choice = (res_json.get("choices") or [{}])[0]
        message = choice.get("message", {})
        if message.get("refusal"):
            return None, f"모델 거절: {message['refusal']}"

        text = (message.get("content") or "").strip()
        if not text:
            return None, f"빈 응답: {str(res_json)[:300]}"

        parsed = json.loads(text)
        parsed = _normalize_receipt_result(parsed)
        if not parsed["items"]:
            return None, "상품행을 검증하지 못했습니다. 사진을 더 가까이/선명하게 찍거나 수동 입력해 주세요."
        return parsed, None

    except requests.exceptions.HTTPError as e:
        try:
            detail = res.json()
        except Exception:
            detail = res.text
        return None, f"API HTTP 오류: {e} / {str(detail)[:500]}"
    except Exception as e:
        return None, str(e)

def show():
    if st.button("← 홈으로"):
        st.session_state.page = "home"
        st.session_state.receipt_items = []
        st.session_state.last_uploaded = None
        st.session_state.analyzed = False
        st.rerun()

    st.title("🧾 영수증 물품보고서 생성기")
    st.caption("made by 찰칵혁신단")
    st.markdown("영수증 사진을 올리고 분석하기를 누르면 AI가 항목을 읽어드려요.")
    st.divider()

    report_title = st.text_input("📝 보고서 제목", placeholder="예: 2025년 5월 사무용품 구매내역")
    report_date = st.date_input("📅 보고서 날짜", value=datetime.today())
    st.divider()

    uploaded_file = st.file_uploader(
        "📷 영수증 사진 업로드 (1장)",
        type=["jpg", "jpeg", "png", "heic"],
        accept_multiple_files=False
    )

    if "receipt_items" not in st.session_state:
        st.session_state.receipt_items = []
    if "receipt_meta" not in st.session_state:
        st.session_state.receipt_meta = {"store": "", "date": ""}
    if "last_uploaded" not in st.session_state:
        st.session_state.last_uploaded = None
    if "analyzed" not in st.session_state:
        st.session_state.analyzed = False
    if "confirmed" not in st.session_state:
        st.session_state.confirmed = False
    if "receipt_warnings" not in st.session_state:
        st.session_state.receipt_warnings = []

    current_name = uploaded_file.name if uploaded_file else None
    if current_name != st.session_state.last_uploaded:
        st.session_state.receipt_items = []
        st.session_state.receipt_meta = {"store": "", "date": ""}
        st.session_state.analyzed = False
        st.session_state.confirmed = False
        st.session_state.receipt_warnings = []
        st.session_state.img_rotation = 0
        st.session_state.last_uploaded = current_name

    if "img_rotation" not in st.session_state:
        st.session_state.img_rotation = 0

    if uploaded_file:
        uploaded_file.seek(0)
        img_orig = Image.open(io.BytesIO(uploaded_file.read()))

        # 회전 적용
        rotation = st.session_state.img_rotation
        if rotation != 0:
            img = img_orig.rotate(-rotation, expand=True)
        else:
            img = img_orig

        col_img, col_btn = st.columns([2, 1])
        with col_img:
            st.image(img, caption=uploaded_file.name, width="stretch")
        with col_btn:
            st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
            if st.button("🔄 90도 회전", width="stretch"):
                st.session_state.img_rotation = (st.session_state.img_rotation + 90) % 360
                st.rerun()
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
            if st.button("🤖 AI 분석하기", type="primary", width="stretch"):
                with st.spinner("AI가 영수증을 읽는 중..."):
                    result, err = read_receipt_with_ai(img)
                    if result and result.get("items"):
                        st.session_state.receipt_meta = {
                            "store": result.get("store", ""),
                            "date": result.get("date", ""),
                        }
                        st.session_state.receipt_items = [
                            {
                                "물품명": item.get("name", ""),
                                "규격": item.get("spec", "") or "-",
                                "단위": "개",
                                "수량": item.get("quantity", 1),
                                "단가": item.get("unit_price", 0),
                                "금액": item.get("amount", 0),
                                "통화": "KRW",
                                "검사항목": "물품수량 및 상태 등",
                                "검사결과": "이상없음",
                                "비고": "",
                            }
                            for item in result["items"]
                        ]
                        st.session_state.analyzed = True
                        st.session_state.confirmed = False
                        st.session_state.receipt_warnings = result.get("warnings", [])
                        if st.session_state.receipt_warnings:
                            st.warning("⚠️ 일부 줄은 금액 검증 실패/불확실로 제외했어요. 아래 경고를 확인하세요.")
                            with st.expander("AI 검증 경고 보기"):
                                st.write(st.session_state.receipt_warnings)
                        st.success(f"✅ {len(st.session_state.receipt_items)}개 항목을 읽었어요!")
                    else:
                        st.error("❌ 항목을 읽지 못했어요.")
                        if err:
                            st.caption(f"오류 상세: {err}")

    # ── 헤더 정보 (구매날짜, 거래처, 구매자명) ──
    if st.session_state.analyzed and st.session_state.receipt_items:
        st.divider()
        st.subheader("📋 기본 정보")

        meta = st.session_state.receipt_meta
        try:
            meta_date = datetime.strptime(meta["date"], "%Y-%m-%d").date()
        except:
            meta_date = datetime.today().date()

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            purchase_date = st.date_input("구매날짜", value=meta_date, key="meta_date", disabled=st.session_state.confirmed)
        with col_b:
            store_name = st.text_input("거래처명", value=meta["store"], key="meta_store", disabled=st.session_state.confirmed)
        with col_c:
            buyer_name = st.text_input("구매자명", value="K-water", key="meta_buyer", disabled=st.session_state.confirmed)

        # ── 항목 표 ──
        st.divider()

        locked = st.session_state.confirmed

        # 잠금 상태 배너
        if locked:
            st.success("✅ 내용이 확인되었습니다. 수정하려면 '잠금 해제'를 눌러주세요.")
        else:
            st.subheader("📝 항목 확인 및 수정")
            st.caption("내용을 모두 확인했으면 맨 아래 체크박스를 체크해주세요.")

        # 헤더
        # 물품명, 규격, 단위, 수량, 단가, 금액, 통화, 검사항목, 검사결과, 비고, 삭제
        cols_w = [2.2, 1.5, 0.8, 0.7, 1.2, 1.2, 1.0, 2.2, 1.5, 1.5, 0.6]
        h_labels = ["물품명", "규격", "단위", "수량", "단가", "금액", "통화", "검사항목", "검사결과", "비고", ""]
        hcols = st.columns(cols_w)
        for hc, hl in zip(hcols, h_labels):
            hc.markdown(f"**{hl}**")

        items = st.session_state.receipt_items
        updated_items = []
        to_delete = []

        for i, item in enumerate(items):
            c = st.columns(cols_w)
            with c[0]:
                name = st.text_input("물품명", value=item["물품명"], key=f"name_{i}", label_visibility="collapsed", disabled=locked)
            with c[1]:
                spec = st.text_input("규격", value=item["규격"], key=f"spec_{i}", label_visibility="collapsed", disabled=locked)
            with c[2]:
                unit = st.text_input("단위", value=item["단위"], key=f"unit_{i}", label_visibility="collapsed", disabled=locked)
            with c[3]:
                qty = st.number_input("수량", value=int(item["수량"]), min_value=1, key=f"qty_{i}", label_visibility="collapsed", disabled=locked)
            with c[4]:
                unit_price = st.number_input("단가", value=int(item["단가"]), min_value=0, key=f"uprice_{i}", label_visibility="collapsed", disabled=locked)
            with c[5]:
                amount = unit_price * qty
                st.markdown(f"<div style='padding-top:8px'>{amount:,}</div>", unsafe_allow_html=True)
            with c[6]:
                currency = st.text_input("통화", value=item["통화"], key=f"curr_{i}", label_visibility="collapsed", disabled=locked)
            with c[7]:
                insp_item = st.text_input("검사항목", value=item["검사항목"], key=f"insp_{i}", label_visibility="collapsed", disabled=locked)
            with c[8]:
                insp_result = st.text_input("검사결과", value=item["검사결과"], key=f"iresult_{i}", label_visibility="collapsed", disabled=locked)
            with c[9]:
                note = st.text_input("비고", value=item["비고"], key=f"note_{i}", label_visibility="collapsed", disabled=locked)
            with c[10]:
                if not locked:
                    if st.button("🗑️", key=f"del_{i}"):
                        to_delete.append(i)

            updated_items.append({
                "물품명": name,
                "규격": spec if spec.strip() else "-",
                "단위": unit,
                "수량": qty,
                "단가": unit_price,
                "금액": amount,
                "통화": currency,
                "검사항목": insp_item,
                "검사결과": insp_result,
                "검사일자": purchase_date.strftime("%Y-%m-%d"),
                "비고": note,
            })

        if to_delete:
            st.session_state.receipt_items = [it for idx, it in enumerate(updated_items) if idx not in to_delete]
            st.rerun()
        else:
            st.session_state.receipt_items = updated_items

        if not locked:
            if st.button("➕ 항목 추가"):
                st.session_state.receipt_items.append({
                    "물품명": "",
                    "규격": "-",
                    "단위": "개",
                    "수량": 1,
                    "단가": 0,
                    "금액": 0,
                    "통화": "KRW",
                    "검사항목": "물품수량 및 상태 등",
                    "검사결과": "이상없음",
                    "검사일자": purchase_date.strftime("%Y-%m-%d"),
                    "비고": "",
                })
                st.rerun()

        # ── 확인 체크박스 / 잠금 해제 버튼 ──
        st.divider()
        if not locked:
            confirmed = st.checkbox("✅ 위 내용이 모두 정확함을 확인했습니다. (체크 시 수정 불가)")
            if confirmed:
                st.session_state.confirmed = True
                st.rerun()
        else:
            if st.button("🔓 잠금 해제 (내용 수정하기)"):
                st.session_state.confirmed = False
                st.rerun()

        st.divider()

        if st.button("📄 Word 문서 생성", type="primary", width="stretch", disabled=not st.session_state.confirmed):
            if not st.session_state.confirmed:
                st.warning("⚠️ 내용 확인 체크박스를 먼저 체크해주세요.")
            elif not report_title.strip():
                st.error("⚠️ 보고서 제목을 입력해주세요!")
            else:
                with st.spinner("문서를 생성하는 중..."):
                    try:
                        doc = Document()
                        section = doc.sections[0]
                        section.top_margin = Inches(1)
                        section.bottom_margin = Inches(1)
                        section.left_margin = Inches(0.8)
                        section.right_margin = Inches(0.8)

                        # ── 제목 ──
                        title_p = doc.add_paragraph()
                        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        t_run = title_p.add_run(report_title.strip())
                        t_run.bold = True
                        t_run.font.size = Pt(18)
                        t_run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)

                        date_p = doc.add_paragraph()
                        date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        d_run = date_p.add_run(report_date.strftime("%Y년 %m월 %d일"))
                        d_run.font.size = Pt(11)
                        d_run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
                        doc.add_paragraph()

                        # ── 기본 정보 (구매날짜 / 거래처 / 구매자명) ──
                        def add_info_row(label, value):
                            p = doc.add_paragraph()
                            label_run = p.add_run(f"{label}: ")
                            label_run.bold = True
                            label_run.font.size = Pt(11)
                            val_run = p.add_run(value)
                            val_run.font.size = Pt(11)

                        add_info_row("구매날짜", purchase_date.strftime("%Y년 %m월 %d일"))
                        add_info_row("거래처명", store_name)
                        add_info_row("구매자명", buyer_name)
                        doc.add_paragraph()

                        # ── 표 ──
                        headers = ["번호", "품명", "규격", "단위", "수량", "단가", "금액", "통화", "검사항목", "검사결과", "검사일자", "비고"]
                        aligns = [
                            WD_ALIGN_PARAGRAPH.CENTER,  # 번호
                            WD_ALIGN_PARAGRAPH.LEFT,    # 품명
                            WD_ALIGN_PARAGRAPH.CENTER,  # 규격
                            WD_ALIGN_PARAGRAPH.CENTER,  # 단위
                            WD_ALIGN_PARAGRAPH.CENTER,  # 수량
                            WD_ALIGN_PARAGRAPH.RIGHT,   # 단가
                            WD_ALIGN_PARAGRAPH.RIGHT,   # 금액
                            WD_ALIGN_PARAGRAPH.CENTER,  # 통화
                            WD_ALIGN_PARAGRAPH.LEFT,    # 검사항목
                            WD_ALIGN_PARAGRAPH.CENTER,  # 검사결과
                            WD_ALIGN_PARAGRAPH.CENTER,  # 검사일자
                            WD_ALIGN_PARAGRAPH.LEFT,    # 비고
                        ]

                        table = doc.add_table(rows=1, cols=len(headers))
                        table.style = "Table Grid"

                        # 헤더 행
                        hdr_row = table.rows[0]
                        for j, h in enumerate(headers):
                            cell = hdr_row.cells[j]
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run = cell.paragraphs[0].add_run(h)
                            run.bold = True
                            run.font.size = Pt(9)
                            tcPr = cell._tc.get_or_add_tcPr()
                            shd = OxmlElement('w:shd')
                            shd.set(qn('w:val'), 'clear')
                            shd.set(qn('w:color'), 'auto')
                            shd.set(qn('w:fill'), 'E8EDF5')
                            tcPr.append(shd)

                        # 데이터 행
                        for i, item in enumerate(st.session_state.receipt_items):
                            row = table.add_row()
                            try:
                                insp_date = datetime.strptime(item["검사일자"], "%Y-%m-%d").strftime("%Y.%m.%d")
                            except:
                                insp_date = item.get("검사일자", "")

                            vals = [
                                str(i + 1),
                                item["물품명"],
                                item["규격"],
                                item["단위"],
                                str(item["수량"]),
                                f"{int(item['단가']):,}",
                                f"{int(item['금액']):,}",
                                item["통화"],
                                item["검사항목"].replace("물품수량 및 상태 등", "물품수량 및 상태 등"),
                                item["검사결과"],
                                insp_date,
                                item["비고"],
                            ]
                            for j, (v, a) in enumerate(zip(vals, aligns)):
                                cell = row.cells[j]
                                cell.paragraphs[0].alignment = a
                                run = cell.paragraphs[0].add_run(v)
                                run.font.size = Pt(9)

                        doc_buffer = io.BytesIO()
                        doc.save(doc_buffer)
                        doc_buffer.seek(0)

                        date_filename = report_date.strftime("%Y%m%d")
                        safe_name = report_title.strip().replace(" ", "_").replace("/", "-")
                        filename = f"{safe_name}_{date_filename}_물품보고서.docx"

                        st.success("✅ 문서가 생성되었어요!")
                        st.download_button(
                            label="⬇️ Word 문서 다운로드",
                            data=doc_buffer,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            width="stretch"
                        )
                    except Exception as e:
                        st.error(f"오류: {str(e)}")
                        st.exception(e)
