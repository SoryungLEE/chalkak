import streamlit as st

st.set_page_config(
    page_title="찰칵혁신단 문서 도우미",
    page_icon="⚡",
    layout="centered"
)

# 페이지 상태 초기화
if "page" not in st.session_state:
    st.session_state.page = "home"

# 페이지 라우팅
if st.session_state.page == "photo":
    import photo
    photo.show()

elif st.session_state.page == "receipt":
    import receipt
    receipt.show()

else:
    # 홈 화면
    st.title("⚡ 찰칵혁신단 문서 도우미")
    st.caption("필요한 기능을 선택하세요")
    st.divider()

    # 카드 버튼 스타일 주입
    st.markdown("""
    <style>
    div[data-testid="stButton"] button {
        height: auto !important;
        padding: 0 !important;
        border: none !important;
        background: transparent !important;
        box-shadow: none !important;
    }
    div[data-testid="stButton"] button:hover {
        background: transparent !important;
        box-shadow: none !important;
    }
    .card-photo {
        background: #f0f4ff;
        border-radius: 16px;
        padding: 36px 24px 28px 24px;
        text-align: center;
        border: 2px solid #c7d7f9;
        min-height: 220px;
        cursor: pointer;
        transition: all 0.18s ease;
        width: 100%;
    }
    .card-photo:hover {
        background: #e0eaff;
        border-color: #6699ff;
        box-shadow: 0 4px 16px rgba(100,140,255,0.18);
        transform: translateY(-2px);
    }
    .card-receipt {
        background: #fff4f0;
        border-radius: 16px;
        padding: 36px 24px 28px 24px;
        text-align: center;
        border: 2px solid #f9c7b0;
        min-height: 220px;
        cursor: pointer;
        transition: all 0.18s ease;
        width: 100%;
    }
    .card-receipt:hover {
        background: #ffe8df;
        border-color: #ff8855;
        box-shadow: 0 4px 16px rgba(255,140,100,0.18);
        transform: translateY(-2px);
    }
    .card-title {
        font-size: 20px;
        font-weight: 700;
        color: #1a1a2e;
        margin-bottom: 10px;
        font-family: 'Malgun Gothic', sans-serif;
    }
    .card-desc {
        font-size: 13px;
        color: #666;
        line-height: 1.6;
        font-family: 'Malgun Gothic', sans-serif;
    }
    </style>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        if st.button(
            "📋\n\n**회의 사진 대지**\n\n회의 사진을 업로드하면 Word 문서로 자동 정리",
            use_container_width=True,
            key="btn_photo"
        ):
            st.session_state.page = "photo"
            st.rerun()
        # 카드 스타일 덮어쓰기
        st.markdown("""
        <style>
        #btn_photo { display: none; }
        </style>
        """, unsafe_allow_html=True)

    with col2:
        if st.button(
            "🧾\n\n**영수증 물품보고서**\n\n영수증 사진을 올리면 AI가 항목을 읽어 표로 정리",
            use_container_width=True,
            key="btn_receipt"
        ):
            st.session_state.page = "receipt"
            st.rerun()

    # 실제 클릭 가능한 카드 (JavaScript로 버튼 트리거)
    st.markdown("""
    <style>
    div[data-testid="stButton"]:has(button[kind="secondary"]),
    div[data-testid="stButton"]:has(button[kind="primary"]) {
        display: none;
    }
    </style>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        <div class="card-photo" onclick="
            const btns = window.parent.document.querySelectorAll('button');
            for(const b of btns){ if(b.innerText.includes('회의 사진 대지')){ b.click(); break; } }
        ">
            <div style="font-size:48px; margin-bottom:14px;">📋</div>
            <div class="card-title">회의 사진 대지</div>
            <div class="card-desc">회의 사진을 업로드하면<br>Word 문서로 자동 정리</div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="card-receipt" onclick="
            const btns = window.parent.document.querySelectorAll('button');
            for(const b of btns){ if(b.innerText.includes('영수증 물품보고서')){ b.click(); break; } }
        ">
            <div style="font-size:48px; margin-bottom:14px;">🧾</div>
            <div class="card-title">영수증 물품보고서</div>
            <div class="card-desc">영수증 사진을 올리면<br>AI가 항목을 읽어 표로 정리</div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()
    st.caption("made by 찰칵혁신단")
