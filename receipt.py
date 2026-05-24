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


def img_to_base64(img):
    buf = io.BytesIO()
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    img.save(buf, format="JPEG", quality=85)
    return base64.b64encode(buf.getvalue()).decode()


def read_receipt_with_ai(img):
    b64 = img_to_base64(img)
    payload = {
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 1000,
        "messages": [{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}
                },
                {
                    "type": "text",
                    "text": """이 영수증 이미지에서 구매 항목을 추출해주세요.
반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트는 절대 포함하지 마세요.

{
  "store": "거래처명(가게명)",
  "date": "구매날짜 (YYYY-MM-DD 형식, 모르면 빈 문자열)",
  "items": [
    {"name": "물품명", "spec": "규격(없으면 빈 문자열)", "quantity": 수량(숫자), "unit_price": 단가(숫자), "amount": 금액(숫자)}
  ]
}

- 수량 불명확하면 1로 설정
- 금액은 숫자만 (원, 콤마 제외)
- 규격은 용량/사이즈/모델명 등, 없으면 빈 문자열"""
                }
            ]
        }]
    }
    try:
        res = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"Content-Type": "application/json"},
            json=payload,
            timeout=30
        )
        text = res.json()["content"][0]["text"].strip()
        text = text.replace("```json", "").replace("```", "").strip()
        return json.loads(text)
    except:
        return None


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
            st.image(img, caption=uploaded_file.name, use_container_width=True)
        with col_btn:
            st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)
            if st.button("🤖 AI 분석하기", type="primary", use_container_width=True):
                with st.spinner("AI가 영수증을 읽는 중..."):
                    result = read_receipt_with_ai(img)
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
                        st.error("❌ 항목을 읽지 못했어요. 사진을 다시 확인해주세요.")

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

        if st.button("📄 Word 문서 생성", type="primary", use_container_width=True):
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
                            # 날짜 포맷 변환
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
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"오류: {str(e)}")
                        st.exception(e)
