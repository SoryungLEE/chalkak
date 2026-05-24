import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import io

st.set_page_config(
    page_title="회의록 사진 문서 생성기",
    page_icon="📋",
    layout="centered"
)

st.title("📋 회의록 사진 문서 생성기")
st.markdown("회의명을 입력하고 사진을 업로드하면 Word 문서를 자동으로 만들어드려요.")

st.divider()

meeting_name = st.text_input("📝 회의명", placeholder="예: 2024년 1월 기획팀 월례회의")

uploaded_files = st.file_uploader(
    "📷 사진 업로드 (여러 장 선택 가능)",
    type=["jpg", "jpeg", "png", "heic"],
    accept_multiple_files=True
)

if uploaded_files:
    st.divider()
    st.subheader("🖼️ 업로드된 사진 미리보기")
    cols = st.columns(3)
    for i, file in enumerate(uploaded_files):
        with cols[i % 3]:
            try:
                img = Image.open(file)
                if hasattr(img, '_getexif') and img._getexif():
                    exif_data = img._getexif()
                    if exif_data and 274 in exif_data:
                        orientation = exif_data[274]
                        rotations = {3: 180, 6: 270, 8: 90}
                        if orientation in rotations:
                            img = img.rotate(rotations[orientation], expand=True)
                st.image(img, caption=f"{i+1}. {file.name}", use_container_width=True)
            except:
                st.warning(f"미리보기 불가: {file.name}")

st.divider()

if st.button("📄 Word 문서 생성", type="primary", use_container_width=True):
    if not meeting_name.strip():
        st.error("⚠️ 회의명을 입력해주세요!")
    elif not uploaded_files:
        st.error("⚠️ 사진을 최소 1장 이상 업로드해주세요!")
    else:
        with st.spinner("문서를 생성하는 중..."):
            try:
                doc = Document()
                section = doc.sections[0]
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)

                # 회의명 제목
                title = doc.add_paragraph()
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = title.add_run(meeting_name.strip())
                run.bold = True
                run.font.size = Pt(20)
                run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)
                doc.add_paragraph()

                PHOTOS_PER_PAGE = 3
                page_width = Inches(6.0)

                for i, file in enumerate(uploaded_files):
                    file.seek(0)
                    try:
                        img = Image.open(io.BytesIO(file.read()))
                        if hasattr(img, '_getexif') and img._getexif():
                            exif_data = img._getexif()
                            if exif_data and 274 in exif_data:
                                orientation = exif_data[274]
                                rotations = {3: 180, 6: 270, 8: 90}
                                if orientation in rotations:
                                    img = img.rotate(rotations[orientation], expand=True)
                        if img.mode in ("RGBA", "P"):
                            img = img.convert("RGB")

                        buf = io.BytesIO()
                        img.save(buf, format="JPEG", quality=85)
                        buf.seek(0)

                        w, h = img.size
                        aspect = h / w
                        img_width = page_width
                        img_height = page_width * aspect
                        if img_height > Inches(3.5):
                            img_height = Inches(3.5)
                            img_width = img_height / aspect

                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = p.add_run()
                        run.add_picture(buf, width=img_width)

                        caption = doc.add_paragraph(f"사진 {i+1}  |  {file.name}")
                        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        caption.runs[0].font.size = Pt(9)
                        caption.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)
                        doc.add_paragraph()

                        if (i + 1) % PHOTOS_PER_PAGE == 0 and (i + 1) < len(uploaded_files):
                            doc.add_page_break()
                            title2 = doc.add_paragraph()
                            title2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run2 = title2.add_run(meeting_name.strip())
                            run2.bold = True
                            run2.font.size = Pt(20)
                            run2.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)
                            doc.add_paragraph()
                    except Exception as e:
                        doc.add_paragraph(f"[이미지 로드 실패: {file.name}]")

                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)

                safe_name = meeting_name.strip().replace(" ", "_").replace("/", "-")
                filename = f"{safe_name}_회의록.docx"

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
