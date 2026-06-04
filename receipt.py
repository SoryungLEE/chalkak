import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image
import io
import base64
import json
import requests
from datetime import datetime


def fix_orientation(img):
    try:
        exif_data = img._getexif()
        if exif_data and 274 in exif_data:
            orientation = exif_data[274]
            rotations = {3: 180, 6: 270, 8: 90}
            if orientation in rotations:
                img = img.rotate(rotations[orientation], expand=True)
    except:
        pass
    return img


def img_to_base64(img):
    img = fix_orientation(img)

    # 너무 크면 리사이즈 (긴 쪽 2000px로 제한, 화질 유지)
    max_side = 2000
    w, h = img.size
    if max(w, h) > max_side:
        ratio = max_side / max(w, h)
        img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)

    buf = io.BytesIO()
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    img.save(buf, format="JPEG", quality=90)
    return base64.b64encode(buf.getvalue()).decode()


def read_receipt_with_ai(img):
    b64 = img_to_base64(img)

    # Streamlit secrets에서 OpenAI API 키 가져오기
    api_key = st.secrets["OPENAI_API_KEY"]

    payload = {
        "model": "gpt-4o",
        "max_tokens": 1000,
        "messages": [{
            "role": "user",
            "content": [
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/jpeg;base64,{b64}"
                    }
                },
                {
                    "type": "text",
                    "text": """이 이미지는 한국어 구매영수증입니다. 이미지가 회전되어 있거나 기울어져 있어도 모든 방향의 텍스트를 읽어주세요. 이미지에서 보이는 모든 텍스트를 먼저 읽고, 구매 항목 정보를 추출하세요.

반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트는 절대 포함하지 마세요.

{
  "store": "거래처명(회사명/판매자명/가게명 중 보이는 것)",
  "date": "구매날짜 YYYY-MM-DD 형식 (주문날짜/결제일/구매일 중 보이는 것, 없으면 빈 문자열)",
  "items": [
    {"name": "물품명", "spec": "규격(모델명/사이즈 등, 없으면 빈 문자열)", "quantity": 수량숫자, "unit_price": 단가숫자, "amount": 금액숫자}
  ]
}

중요 규칙:
- 상품명/품목/아이템 어떤 표현이든 물품명으로 추출할 것
- 상품명 끝에 수량이 붙어있으면(예: '칫솔 2개') 수량 분리해서 추출
- 단가/금액을 알 수 없으면 합계금액을 amount에 넣고 unit_price도 동일하게
- 수량 모르면 1
- 금액은 숫자만 (원, 콤마 없이)
- 합계/배송비는 items에 포함 금지
- 이미지에 구매 정보가 하나라도 있으면 반드시 items에 1개 이상 넣을 것
- items가 빈 배열이면 절대 안됨"""
                }
            ]
        }]
    }

    try:
        res = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            },
            json=payload,
            timeout=30
        )
        res_json = res.json()
        if "choices" not in res_json:
            return None, f"API 오류: {str(res_json)[:200]}"
        text = res_json["choices"][0]["message"]["content"].strip()
        text = text.replace("```json", "").replace("```", "").strip()
        parsed = json.loads(text)
        return parsed, None
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
    if "last_uploaded" not in st.session_state:
        st.session_state.last_uploaded = None
    if "analyzed" not in st.session_state:
        st.session_state.analyzed = False

    current_name = uploaded_file.name if uploaded_file else None
    if current_name != st.session_state.last_uploaded:
        st.session_state.receipt_items = []
        st.session_state.analyzed = False
        st.session_state.last_uploaded = current_name

    if uploaded_file:
        uploaded_file.seek(0)
        img = Image.open(io.BytesIO(uploaded_file.read()))

        col_img, col_btn = st.columns([2, 1])
        with col_img:
            st.image(img, caption=uploaded_file.name, width="stretch")
        with col_btn:
            st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
            if st.button("🤖 AI 분석하기", type="primary", width="stretch"):
                with st.spinner("AI가 영수증을 읽는 중..."):
                    result, err = read_receipt_with_ai(img)
                    if result and result.get("items"):
                        st.session_state.receipt_items = [
                            {
                                "구매날짜": result.get("date", ""),
                                "물품명": item.get("name", ""),
                                "규격": item.get("spec", "") or "-",
                                "수량": item.get("quantity", 1),
                                "거래처명": result.get("store", ""),
                                "구매자명": "",
                            }
                            for item in result["items"]
                        ]
                        st.session_state.analyzed = True
                        st.success(f"✅ {len(st.session_state.receipt_items)}개 항목을 읽었어요!")
                    else:
                        st.error(f"❌ 항목을 읽지 못했어요.")
                        if err:
                            st.caption(f"오류 상세: {err}")

    # 항목 수정 테이블
    if st.session_state.analyzed and st.session_state.receipt_items:
        st.divider()
        st.subheader("📝 항목 확인 및 수정")
        st.caption("내용을 확인하고 필요하면 수정하세요. 구매자명을 꼭 입력해주세요.")

        # 헤더
        h1, h2, h3, h4, h5, h6, h7 = st.columns([1.8, 2.5, 1.5, 1, 1.8, 1.8, 0.6])
        h1.markdown("**구매날짜**")
        h2.markdown("**물품명**")
        h3.markdown("**규격**")
        h4.markdown("**수량**")
        h5.markdown("**거래처명**")
        h6.markdown("**구매자명**")

        items = st.session_state.receipt_items
        updated_items = []
        to_delete = []

        for i, item in enumerate(items):
            c1, c2, c3, c4, c5, c6, c7 = st.columns([1.8, 2.5, 1.5, 1, 1.8, 1.8, 0.6])
            with c1:
                date_val = item["구매날짜"]
                try:
                    date_obj = datetime.strptime(date_val, "%Y-%m-%d").date()
                except:
                    date_obj = datetime.today().date()
                purchase_date = st.date_input("구매날짜", value=date_obj, key=f"date_{i}", label_visibility="collapsed")
            with c2:
                name = st.text_input("물품명", value=item["물품명"], key=f"name_{i}", label_visibility="collapsed")
            with c3:
                spec = st.text_input("규격", value=item["규격"], key=f"spec_{i}", label_visibility="collapsed")
            with c4:
                qty = st.number_input("수량", value=int(item["수량"]), min_value=1, key=f"qty_{i}", label_visibility="collapsed")
            with c5:
                store = st.text_input("거래처명", value=item["거래처명"], key=f"store_{i}", label_visibility="collapsed")
            with c6:
                buyer = st.text_input("구매자명", value=item["구매자명"], placeholder="입력", key=f"buyer_{i}", label_visibility="collapsed")
            with c7:
                if st.button("🗑️", key=f"del_{i}"):
                    to_delete.append(i)

            updated_items.append({
                "구매날짜": purchase_date.strftime("%Y-%m-%d"),
                "물품명": name,
                "규격": spec if spec.strip() else "-",
                "수량": qty,
                "거래처명": store,
                "구매자명": buyer,
            })

        if to_delete:
            st.session_state.receipt_items = [item for i, item in enumerate(updated_items) if i not in to_delete]
            st.rerun()
        else:
            st.session_state.receipt_items = updated_items

        if st.button("➕ 항목 추가"):
            st.session_state.receipt_items.append({
                "구매날짜": datetime.today().strftime("%Y-%m-%d"),
                "물품명": "",
                "규격": "-",
                "수량": 1,
                "거래처명": "",
                "구매자명": "",
            })
            st.rerun()

        st.divider()

        if st.button("📄 Word 문서 생성", type="primary", width="stretch"):
            if not report_title.strip():
                st.error("⚠️ 보고서 제목을 입력해주세요!")
            else:
                with st.spinner("문서를 생성하는 중..."):
                    try:
                        doc = Document()
                        section = doc.sections[0]
                        section.top_margin = Inches(1)
                        section.bottom_margin = Inches(1)
                        section.left_margin = Inches(1)
                        section.right_margin = Inches(1)

                        # 제목
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

                        # 표
                        headers = ["번호", "구매날짜", "물품명", "규격", "수량", "거래처명", "구매자명"]
                        aligns = [
                            WD_ALIGN_PARAGRAPH.CENTER,
                            WD_ALIGN_PARAGRAPH.CENTER,
                            WD_ALIGN_PARAGRAPH.LEFT,
                            WD_ALIGN_PARAGRAPH.CENTER,
                            WD_ALIGN_PARAGRAPH.CENTER,
                            WD_ALIGN_PARAGRAPH.CENTER,
                            WD_ALIGN_PARAGRAPH.CENTER,
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
                            run.font.size = Pt(10)
                            tcPr = cell._tc.get_or_add_tcPr()
                            shd = OxmlElement('w:shd')
                            shd.set(qn('w:val'), 'clear')
                            shd.set(qn('w:color'), 'auto')
                            shd.set(qn('w:fill'), 'E8EDF5')
                            tcPr.append(shd)

                        # 데이터 행
                        items = st.session_state.receipt_items
                        for i, item in enumerate(items):
                            row = table.add_row()
                            try:
                                d = datetime.strptime(item["구매날짜"], "%Y-%m-%d")
                                date_str = d.strftime("%Y.%m.%d")
                            except:
                                date_str = item["구매날짜"]

                            vals = [
                                str(i + 1),
                                date_str,
                                item["물품명"],
                                item["규격"],
                                str(item["수량"]),
                                item["거래처명"],
                                item["구매자명"],
                            ]
                            for j, (v, a) in enumerate(zip(vals, aligns)):
                                cell = row.cells[j]
                                cell.paragraphs[0].alignment = a
                                run = cell.paragraphs[0].add_run(v)
                                run.font.size = Pt(10)

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
