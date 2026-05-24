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

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div style="
            background: #f0f4ff;
            border-radius: 16px;
            padding: 32px 24px;
            text-align: center;
            border: 2px solid #c7d7f9;
            min-height: 200px;
        ">
            <div style="font-size: 48px; margin-bottom: 12px;">📋</div>
            <div style="font-size: 18px; font-weight: 700; color: #1a1a2e; margin-bottom: 8px;">회의 사진 대지</div>
            <div style="font-size: 13px; color: #666; line-height: 1.5;">회의 사진을 업로드하면<br>Word 문서로 자동 정리</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
        if st.button("📋 사진 대지 만들기", use_container_width=True, type="primary"):
            st.session_state.page = "photo"
            st.rerun()

    with col2:
        st.markdown("""
        <div style="
            background: #fff4f0;
            border-radius: 16px;
            padding: 32px 24px;
            text-align: center;
            border: 2px solid #f9c7b0;
            min-height: 200px;
        ">
            <div style="font-size: 48px; margin-bottom: 12px;">🧾</div>
            <div style="font-size: 18px; font-weight: 700; color: #1a1a2e; margin-bottom: 8px;">영수증 물품보고서</div>
            <div style="font-size: 13px; color: #666; line-height: 1.5;">영수증 사진을 올리면<br>AI가 항목을 읽어 표로 정리</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
        if st.button("🧾 물품보고서 만들기", use_container_width=True):
            st.session_state.page = "receipt"
            st.rerun()

    st.divider()
    st.caption("made by 찰칵혁신단")
