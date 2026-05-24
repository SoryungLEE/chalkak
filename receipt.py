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
    """Claude API로 영수증 분석"""
    b64 = img_to_base64(img)
    payload = {
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 1000,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/jpeg",
                            "data": b64
                        }
                    },
                    {
                        "type": "text",
                        "text": """이 영수증 이미지에서 구매 항목을 추출해주세요.
반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트는 절대 포함하지 마세요.

{
  "store": "가게명",
  "date": "구매날짜 (YYYY-MM-DD 형식, 모르면 빈 문자열)",
  "items": [
    {"name": "품목명", "quantity": 수량(숫자), "unit_price": 단가(숫자), "amount": 금액(숫자)},
    ...
  ],
  "total": 합계금액(숫자)
}

- 수량이 불명확하면 1로 설정
- 금액은 숫자만 (원, 콤마 제외)
- 품목이 없으면 items를 빈 배열로"""
                    }
                ]
            }
        ]
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
    except Exception as e:
        return None


def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{edge}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:color'), 'CCCCCC')
        tcBorders.append(border)
    tcPr.append(tcBorders)


def show():
    if st.button("← 홈으로"):
        st.session_state.page = "home"
        st.rerun()

    st.title("🧾 영수증 물품보고서 생성기")
    st.caption("made by 찰칵혁신단")
    st.markdown("영수증 사진을 올리면 AI가 항목을 읽어드려요. 수정 후 Word 문서로 저장하세요.")
    st.divider()

    # 보고서 기본 정보
    report_title = st.text_input("📝 보고서 제목", placeholder="예: 2025년 5월 사무용품 구매내역")
    report_date = st.date_input("📅 보고서 날짜", value=datetime.today())

    st.divider()

    uploaded_files = st.file_uploader(
        "📷 영수증 사진 업로드 (여러 장 가능)",
        type=["jpg", "jpeg", "png", "heic"],
        accept_multiple_files=True
    )

    if "receipt_items" not in st.session_state:
        st.session_state.receipt_items = []
    if "last_uploaded" not in st.session_state:
        st.session_state.last_uploaded = []

    # 새 파일 업로드 시 AI 분석
    if uploaded_files:
        current_names = [f.name for f in uploaded_files]
        if current_names != st.session_state.last_uploaded:
            st.session_state.receipt_items = []
            st.session_state.last_uploaded = current_names

            for file in uploaded_files:
                with st.spinner(f"🤖 AI가 {file.name} 분석 중..."):
                    file.seek(0)
                    img = Image.open(io.BytesIO(file.read()))
                    result = read_receipt_with_ai(img)

                    if result and result.get("items"):
                        for item in result["items"]:
                            st.session_state.receipt_items.append({
                                "출처": result.get("store", file.name),
                                "구매일": result.get("date", ""),
                                "품목명": item.get("name", ""),
                                "수량": item.get("quantity", 1),
                                "단가": item.get("unit_price", 0),
                                "금액": item.get("amount", 0),
                            })
                    else:
                        st.warning(f"⚠️ {file.name}: 항목을 읽지 못했어요. 아래에서 직접 추가해주세요.")

    # 항목 편집 테이블
    if st.session_state.receipt_items or uploaded_files:
        st.divider()
        st.subheader("📝 항목 확인 및 수정")
        st.caption("AI가 읽은 내용을 확인하고 필요하면 수정하세요.")

        items = st.session_state.receipt_items
        updated_items = []
        to_delete = []

        for i, item in enumerate(items):
            with st.container():
                col_del, col1, col2, col3, col4 = st.columns([0.5, 3, 1.5, 1.5, 1.5])
                with col_del:
                    if st.button("🗑️", key=f"del_{i}"):
                        to_delete.append(i)
                with col1:
                    name = st.text_input("품목명", value=item["품목명"], key=f"name_{i}", label_visibility="collapsed")
                with col2:
                    qty = st.number_input("수량", value=int(item["수량"]), min_value=1, key=f"qty_{i}", label_visibility="collapsed")
                with col3:
                    price = st.number_input("단가", value=int(item["단가"]), min_value=0, key=f"price_{i}", label_visibility="collapsed")
                with col4:
                    amount = qty * price if price > 0 else int(item["금액"])
                    st.text_input("금액", value=f"{amount:,}원", key=f"amt_{i}", disabled=True, label_visibility="collapsed")

                updated_items.append({
                    "출처": item["출처"],
                    "구매일": item["구매일"],
                    "품목명": name,
                    "수량": qty,
                    "단가": price,
                    "금액": amount,
                })

        # 삭제 처리
        if to_delete:
            st.session_state.receipt_items = [item for i, item in enumerate(updated_items) if i not in to_delete]
            st.rerun()
        else:
            st.session_state.receipt_items = updated_items

        # 항목 추가
        st.divider()
        if st.button("➕ 항목 직접 추가"):
            st.session_state.receipt_items.append({
                "출처": "", "구매일": "", "품목명": "새 항목", "수량": 1, "단가": 0, "금액": 0
            })
            st.rerun()

        # 합계
        if st.session_state.receipt_items:
            total = sum(item["금액"] for item in st.session_state.receipt_items)
            st.markdown(f"### 합계: **{total:,}원**")

        st.divider()

        # 문서 생성
        if st.button("📄 Word 문서 생성", type="primary", use_container_width=True):
            if not report_title.strip():
                st.error("⚠️ 보고서 제목을 입력해주세요!")
            elif not st.session_state.receipt_items:
                st.error("⚠️ 항목이 없습니다!")
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

                        # 표 헤더
                        headers = ["번호", "품목명", "수량", "단가", "금액", "구매처"]
                        col_widths = [700, 3000, 700, 1500, 1500, 1960]
                        total_w = sum(col_widths)

                        table = doc.add_table(rows=1, cols=len(headers))
                        table.style = "Table Grid"

                        from docx.shared import Pt as DPt
                        from docx.shared import RGBColor as DRGB

                        # 헤더 행
                        hdr_row = table.rows[0]
                        for j, (h, w) in enumerate(zip(headers, col_widths)):
                            cell = hdr_row.cells[j]
                            cell.width = w * 635  # EMU 근사
                            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run = cell.paragraphs[0].add_run(h)
                            run.bold = True
                            run.font.size = Pt(10)
                            # 헤더 배경색
                            from docx.oxml.ns import qn as _qn
                            from docx.oxml import OxmlElement as _OE
                            tcPr = cell._tc.get_or_add_tcPr()
                            shd = _OE('w:shd')
                            shd.set(_qn('w:val'), 'clear')
                            shd.set(_qn('w:color'), 'auto')
                            shd.set(_qn('w:fill'), 'E8EDF5')
                            tcPr.append(shd)

                        # 데이터 행
                        items = st.session_state.receipt_items
                        for i, item in enumerate(items):
                            row = table.add_row()
                            vals = [
                                str(i + 1),
                                item["품목명"],
                                str(item["수량"]),
                                f"{int(item['단가']):,}",
                                f"{int(item['금액']):,}",
                                item["출처"],
                            ]
                            aligns = [
                                WD_ALIGN_PARAGRAPH.CENTER,
                                WD_ALIGN_PARAGRAPH.LEFT,
                                WD_ALIGN_PARAGRAPH.CENTER,
                                WD_ALIGN_PARAGRAPH.RIGHT,
                                WD_ALIGN_PARAGRAPH.RIGHT,
                                WD_ALIGN_PARAGRAPH.CENTER,
                            ]
                            for j, (v, a) in enumerate(zip(vals, aligns)):
                                cell = row.cells[j]
                                cell.paragraphs[0].alignment = a
                                run = cell.paragraphs[0].add_run(v)
                                run.font.size = Pt(10)

                        # 합계 행
                        total = sum(item["금액"] for item in items)
                        total_row = table.add_row()
                        total_row.cells[0].merge(total_row.cells[4])
                        total_row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        tr = total_row.cells[0].paragraphs[0].add_run(f"합계: {total:,}원")
                        tr.bold = True
                        tr.font.size = Pt(11)

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
