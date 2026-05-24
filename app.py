import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
from PIL.ExifTags import TAGS
import io
from datetime import datetime

st.set_page_config(
    page_title="회의 사진 대지 생성기",
    page_icon="📋",
    layout="centered"
)

st.title("📋 회의 사진 대지 생성기")
st.markdown("회의명을 입력하고 사진을 업로드하면 Word 문서를 자동으로 만들어드려요.")

st.divider()

meeting_name = st.text_input("📝 회의명", placeholder="예: 2024년 1월 기획팀 월례회의")

uploaded_files = st.file_uploader(
    "📷 사진 업로드 (여러 장 선택 가능)",
    type=["jpg", "jpeg", "png", "heic"],
    accept_multiple_files=True
)


def get_exif_datetime(img):
    """EXIF에서 촬영 날짜/시간 추출"""
    try:
        exif_data = img._getexif()
        if not exif_data:
            return None
        for tag_id, value in exif_data.items():
            tag = TAGS.get(tag_id, tag_id)
            if tag == "DateTimeOriginal":
                dt = datetime.strptime(value, "%Y:%m:%d %H:%M:%S")
                return dt.strftime("%Y년 %m월 %d일 %H:%M")
    except:
        pass
    return None


def fix_orientation(img):
    """EXIF 회전 보정"""
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


# 미리보기
if uploaded_files:
    st.divider()
    st.subheader("🖼️ 업로드된 사진 미리보기")
    cols = st.columns(3)
    for i, file in enumerate(uploaded_files):
        with cols[i % 3]:
            try:
                img = Image.open(file)
                img = fix_orientation(img)
                shot_time = get_exif_datetime(img)
                caption_text = f"{i+1}. {file.name}"
                if shot_time:
                    caption_text += f"\n📅 {shot_time}"
                st.image(img, caption=caption_text, use_container_width=True)
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

                def add_title(doc, text):
                    title = doc.add_paragraph()
                    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = title.add_run(text)
                    run.bold = True
                    run.font.size = Pt(20)
                    run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)
                    doc.add_paragraph()

                add_title(doc, meeting_name.strip())

                PHOTOS_PER_PAGE = 3
                page_width = Inches(6.0)

                for i, file in enumerate(uploaded_files):
                    file.seek(0)
                    try:
                        img = Image.open(io.BytesIO(file.read()))
                        shot_time = get_exif_datetime(img)
                        img = fix_orientation(img)

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

                        # 캡션: 촬영 날짜/시간 있으면 표시, 없으면 파일명
                        if shot_time:
                            caption_str = f"사진 {i+1}  |  {shot_time}"
                        else:
                            caption_str = f"사진 {i+1}  |  {file.name}"

                        caption = doc.add_paragraph(caption_str)
                        caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        caption.runs[0].font.size = Pt(9)
                        caption.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)
                        doc.add_paragraph()

                        # 3장마다 페이지 나누기
                        if (i + 1) % PHOTOS_PER_PAGE == 0 and (i + 1) < len(uploaded_files):
                            doc.add_page_break()
                            add_title(doc, meeting_name.strip())

                    except Exception as e:
                        doc.add_paragraph(f"[이미지 로드 실패: {file.name}]")

                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)

                # 파일명: 회의명_날짜_사진대지.docx
                today = datetime.now().strftime("%Y%m%d")
                safe_name = meeting_name.strip().replace(" ", "_").replace("/", "-")
                filename = f"{safe_name}_{today}_사진대지.docx"

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
