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
    api_key = st.secrets["OPENAI_API_KEY"]

    payload = {
        "model": "gpt-4o",
        "max_tokens": 1500,
        "messages": [{
            "role": "user",
            "content": [
                {
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{b64}"}
                },
                {
                    "type": "text",
                    "text": """이 이미지는 구매영수증입니다. 이미지가 회전되어 있을 수 있으니 모든 방향에서 텍스트를 읽어주세요.

[1단계] 이미지에 보이는 모든 텍스트를 빠짐없이 그대로 읽어내세요. 숫자, 한글, 영문, 기호 전부 포함합니다.

[2단계] 읽어낸 텍스트에서 아래 정보를 파악하세요:
- 어떤 숫자가 단가인지, 수량인지, 합계금액인지 — 단가×수량=금액 관계로 검증하세요.
- 품명은 숫자가 아닌 텍스트 부분입니다. 바코드번호([ 숫자 ] 형태)는 품명이 아닙니다.
- 거래처명, 날짜가 어디 있는지 찾으세요.

[3단계] 반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트는 절대 포함하지 마세요.

{
  "store": "거래처명",
  "date": "YYYY-MM-DD (없으면 빈 문자열)",
  "items": [
    {"name": "품명", "spec": "규격(없으면 빈 문자열)", "quantity": 수량숫자, "unit_price": 단가숫자, "amount": 금액숫자}
  ]
}

규칙:
- 단가×수량=금액이 맞는지 반드시 검증 후 넣을 것
- 합계, 배송비, 부가세, 바코드번호는 items에 넣지 말 것
- 금액은 숫자만 (콤마, 원 없이)
- items는 반드시 1개 이상"""
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
    if "receipt_meta" not in st.session_state:
        st.session_state.receipt_meta = {"store": "", "date": ""}
    if "last_uploaded" not in st.session_state:
        st.session_state.last_uploaded = None
    if "analyzed" not in st.session_state:
        st.session_state.analyzed = False
    if "confirmed" not in st.session_state:
        st.session_state.confirmed = False

    current_name = uploaded_file.name if uploaded_file else None
    if current_name != st.session_state.last_uploaded:
        st.session_state.receipt_items = []
        st.session_state.receipt_meta = {"store": "", "date": ""}
        st.session_state.analyzed = False
        st.session_state.confirmed = False
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
