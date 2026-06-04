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
    "name": "korean_receipt_extraction",
    "strict": True,
    "schema": {
        "type": "object",
        "properties": {
            "store": {"type": "string"},
            "date": {"type": "string"},
            "items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "raw_line": {"type": "string"},
                        "name": {"type": "string"},
                        "spec": {"type": "string"},
                        "quantity": {"type": "integer"},
                        "unit_price": {"type": "integer"},
                        "amount": {"type": "integer"},
                    },
                    "required": ["raw_line", "name", "spec", "quantity", "unit_price", "amount"],
                    "additionalProperties": False,
                },
            },
            "warnings": {
                "type": "array",
                "items": {"type": "string"},
            },
        },
        "required": ["store", "date", "items", "warnings"],
        "additionalProperties": False,
    },
}


RECEIPT_PROMPT = """너는 한국 영수증 상품행 추출기다. 반드시 JSON만 반환한다.

핵심 목표:
- 영수증의 상품 구매 행만 추출한다.
- 상품행의 오른쪽 끝 숫자 3개 열은 보통 [단가] [수량] [금액] 순서다.
- 단가 × 수량 = 금액 검증이 되는 행만 items에 넣는다.

매우 중요한 규칙:
1. 상품명은 오른쪽 끝의 숫자 3개 열(단가/수량/금액)을 제거한 왼쪽 전체 문자열이다.
2. 상품명 안의 숫자, 괄호, 크기, 수량 표기처럼 보이는 문자는 절대 수량/규격으로 떼지 마라.
   예: "도트냅킨(21*21cm)20 1,000 1 1,000"이면
       name="도트냅킨(21*21cm)20", unit_price=1000, quantity=1, amount=1000, spec=""
   예: "신지카드종이컵25P(21 1,000 12 12,000"이면
       name="신지카드종이컵25P(21", unit_price=1000, quantity=12, amount=12000, spec=""
3. 다이소 영수증처럼 별도 '규격' 열이 없는 경우 spec은 빈 문자열로 둔다.
   상품명 안의 20, 25P, 21*21cm, 괄호 안 숫자는 spec으로 빼지 말고 name에 그대로 둔다.
4. raw_line에는 해당 상품행의 원문 한 줄을 오른쪽 숫자 3개까지 포함해 적는다.
   다음 줄의 [ 1002504 ] 같은 바코드번호는 raw_line에도 넣지 않는다.
5. [ 숫자 ] 형태의 바코드 행, 과세합계, 부가세, 판매합계, 봉투, 포인트, 신용카드, 승인번호, 영수증 안내문은 items에 넣지 않는다.
6. 숫자는 콤마와 원 표시 없이 정수로 반환한다.
7. 날짜는 가능하면 YYYY-MM-DD 형식으로 반환한다. 날짜가 불확실하면 빈 문자열.
8. 애매하거나 단가×수량=금액이 맞지 않는 행은 items에 넣지 말고 warnings에 이유를 남긴다.

반환 JSON 구조:
{
  "store": "거래처명",
  "date": "YYYY-MM-DD",
  "items": [
    {
      "raw_line": "상품행 원문",
      "name": "상품명",
      "spec": "",
      "quantity": 1,
      "unit_price": 1000,
      "amount": 1000
    }
  ],
  "warnings": []
}
"""


_AMOUNT_TOKEN = r"(?:\d{1,3}(?:,\d{3})+|\d+)"
_ITEM_LINE_RE = re.compile(
    rf"^\s*(?P<name>.+?)\s+(?P<unit>{_AMOUNT_TOKEN})\s+(?P<qty>\d+)\s+(?P<amount>{_AMOUNT_TOKEN})\s*$"
)
_EXCLUDE_NAME_RE = re.compile(
    r"(과세\s*합계|부가\s*세|판매\s*합계|총\s*합계|합\s*계|봉투|포인트|신용\s*카드|체크\s*카드|현금|거스름|승인|영수증|교환|환불|멤버십|사업자|대표|전화|POS)",
    re.IGNORECASE,
)


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


def _parse_item_raw_line(raw_line):
    """
    모델이 raw_line만 제대로 읽으면, 품명/단가/수량/금액은 코드가 확정한다.
    이게 '단가 1,000 / 수량 7 / 금액 7,000'을 '단가 7,000 / 수량 1'로 뒤집는 문제를 막는다.
    """
    raw = _clean_space(raw_line)
    if not raw:
        return None

    raw = raw.replace("원", "")
    m = _ITEM_LINE_RE.match(raw)
    if not m:
        return None

    name = _clean_space(m.group("name"))
    unit_price = _to_int(m.group("unit"))
    quantity = _to_int(m.group("qty"))
    amount = _to_int(m.group("amount"))

    if not name or _EXCLUDE_NAME_RE.search(name):
        return None
    if unit_price <= 0 or quantity <= 0 or amount <= 0:
        return None
    if unit_price * quantity != amount:
        return None

    return {
        "raw_line": raw,
        "name": name,
        "spec": "",
        "quantity": quantity,
        "unit_price": unit_price,
        "amount": amount,
    }


def _normalize_receipt_result(data):
    warnings = list(data.get("warnings") or [])
    normalized = {
        "store": _clean_space(data.get("store", "")),
        "date": _normalize_date(data.get("date", "")),
        "items": [],
        "warnings": warnings,
    }

    seen = set()
    for item in data.get("items", []) or []:
        raw_line = item.get("raw_line", "")
        parsed = _parse_item_raw_line(raw_line)

        # raw_line 파싱이 안 될 때만 모델이 나눈 필드를 사용하되, 검증 실패하면 버린다.
        if parsed is None:
            name = _clean_space(item.get("name", ""))
            spec = _clean_space(item.get("spec", ""))
            unit_price = _to_int(item.get("unit_price"))
            quantity = _to_int(item.get("quantity"))
            amount = _to_int(item.get("amount"))

            if not name or _EXCLUDE_NAME_RE.search(name):
                continue
            if unit_price <= 0 or quantity <= 0 or amount <= 0:
                normalized["warnings"].append(f"숫자 부족으로 제외: {raw_line or name}")
                continue
            if unit_price * quantity != amount:
                normalized["warnings"].append(
                    f"금액 검증 실패로 제외: {raw_line or name} / {unit_price}*{quantity}!={amount}"
                )
                continue

            # 규격이 단순 숫자/괄호/크기만 있으면 대부분 상품명에서 잘못 뜯어낸 값이라 비운다.
            if re.fullmatch(r"[0-9A-Za-z가-힣()*/.*\-\s]{1,20}", spec) and not re.search(r"규격|SIZE|호|형", spec, re.IGNORECASE):
                spec = ""

            parsed = {
                "raw_line": _clean_space(raw_line),
                "name": name,
                "spec": spec,
                "quantity": quantity,
                "unit_price": unit_price,
                "amount": amount,
            }

        key = (parsed["name"], parsed["quantity"], parsed["unit_price"], parsed["amount"])
        if key in seen:
            continue
        seen.add(key)
        normalized["items"].append(parsed)

    return normalized


def read_receipt_with_ai(img):
    color_img, contrast_img = make_receipt_vision_images(img)
    color_url = image_to_data_url(color_img)
    contrast_url = image_to_data_url(contrast_img)
    api_key = st.secrets["OPENAI_API_KEY"]

    payload = {
        "model": "gpt-4o-2024-08-06",
        "temperature": 0,
        "max_tokens": 1800,
        "response_format": {
            "type": "json_schema",
            "json_schema": RECEIPT_SCHEMA,
        },
        "messages": [
            {
                "role": "system",
                "content": "너는 한국 영수증에서 상품행만 정확히 추출하는 데이터 추출기다. 추측하지 말고 보이는 내용과 산술 검증만 사용한다.",
            },
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": RECEIPT_PROMPT + "\n첫 번째 이미지는 원본 보정본, 두 번째 이미지는 흑백 대비 강화본이다. 둘을 대조해서 읽어라."},
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
