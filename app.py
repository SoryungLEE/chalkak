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

    st.markdown("""
    <style>
    /* 카드 스타일 */
    .card {
        border-radius: 16px;
        padding: 36px 20px 28px 20px;
        text-align: center;
        margin-bottom: 12px;
    }
    .card-photo { background: #f0f4ff; border: 2px solid #c7d7f9; }
    .card-receipt { background: #fff4f0; border: 2px solid #f9c7b0; }
    .card-icon { font-size: 52px; margin-bottom: 14px; }
    .card-title { font-size: 19px; font-weight: 700; color: #1a1a2e; margin-bottom: 8px; }
    .card-desc { font-size: 13px; color: #777; line-height: 1.65; }

    /* 버튼을 카드 바로 아래 딱 붙이기 & 스타일 */
    div[data-testid="stButton"] button {
        border-radius: 12px !important;
        font-size: 15px !important;
        font-weight: 600 !important;
        height: 52px !important;
        transition: all 0.15s ease !important;
        width: 100% !important;
    }
    /* 첫 번째 버튼 (사진 대지) */
    div[data-testid="column"]:nth-of-type(1) div[data-testid="stButton"] button {
        background: #4a7cf7 !important;
        color: white !important;
        border: none !important;
    }
    div[data-testid="column"]:nth-of-type(1) div[data-testid="stButton"] button:hover {
        background: #3366e6 !important;
        box-shadow: 0 4px 14px rgba(74,124,247,0.35) !important;
        transform: translateY(-1px) !important;
    }
    /* 두 번째 버튼 (물품보고서) */
    div[data-testid="column"]:nth-of-type(2) div[data-testid="stButton"] button {
        background: #ff7043 !important;
        color: white !important;
        border: none !important;
    }
    div[data-testid="column"]:nth-of-type(2) div[data-testid="stButton"] button:hover {
        background: #e85d2f !important;
        box-shadow: 0 4px 14px rgba(255,112,67,0.35) !important;
        transform: translateY(-1px) !important;
    }
    </style>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="card card-photo">
            <div class="card-icon">📋</div>
            <div class="card-title">회의 사진 대지</div>
            <div class="card-desc">회의 사진을 업로드하면<br>Word 문서로 자동 정리</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("사진 대지 만들기 →", use_container_width=True, key="btn_photo"):
            st.session_state.page = "photo"
            st.rerun()

    with col2:
        st.markdown("""
        <div class="card card-receipt">
            <div class="card-icon">🧾</div>
            <div class="card-title">영수증 물품보고서</div>
            <div class="card-desc">영수증 사진을 올리면<br>AI가 항목을 읽어 표로 정리</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("물품보고서 만들기 →", use_container_width=True, key="btn_receipt"):
            st.session_state.page = "receipt"
            st.rerun()

    st.divider()
    st.caption("made by 찰칵혁신단")
