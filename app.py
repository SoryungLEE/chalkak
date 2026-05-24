import streamlit as st

st.set_page_config(
    page_title="찰칵혁신단 문서 도우미",
    page_icon="⚡",
    layout="centered"
)

if "page" not in st.session_state:
    st.session_state.page = "home"

if st.session_state.page == "photo":
    import photo
    photo.show()

elif st.session_state.page == "receipt":
    import receipt
    receipt.show()

else:
    st.title("⚡ 찰칵혁신단 문서 도우미")
    st.caption("필요한 기능을 선택하세요")
    st.divider()

    # 카드처럼 보이는 버튼 CSS
    st.markdown("""
    <style>
    [data-testid="stButton"] > button {
        height: 220px !important;
        border-radius: 16px !important;
        font-size: 15px !important;
        font-weight: 600 !important;
        white-space: pre-line !important;
        line-height: 1.8 !important;
        transition: transform 0.15s ease, box-shadow 0.15s ease !important;
    }
    [data-testid="stButton"] > button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 6px 20px rgba(0,0,0,0.12) !important;
    }
    [data-testid="stButton"]:nth-of-type(1) > button {
        background: #f0f4ff !important;
        border: 2px solid #c7d7f9 !important;
        color: #1a1a2e !important;
    }
    [data-testid="stButton"]:nth-of-type(1) > button:hover {
        background: #e0eaff !important;
        border-color: #6699ff !important;
    }
    [data-testid="stButton"]:nth-of-type(2) > button {
        background: #fff4f0 !important;
        border: 2px solid #f9c7b0 !important;
        color: #1a1a2e !important;
    }
    [data-testid="stButton"]:nth-of-type(2) > button:hover {
        background: #ffe8df !important;
        border-color: #ff8855 !important;
    }
    </style>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        if st.button(
            "📋\n\n회의 사진 대지\n\n회의 사진을 업로드하면\nWord 문서로 자동 정리",
            use_container_width=True,
            key="btn_photo"
        ):
            st.session_state.page = "photo"
            st.rerun()

    with col2:
        if st.button(
            "🧾\n\n영수증 물품보고서\n\n영수증 사진을 올리면\nAI가 항목을 읽어 표로 정리",
            use_container_width=True,
            key="btn_receipt"
        ):
            st.session_state.page = "receipt"
            st.rerun()

    st.divider()
    st.caption("made by 찰칵혁신단")
