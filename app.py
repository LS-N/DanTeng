import streamlit as st

# 页面配置
st.set_page_config(
    page_title="CSS Test",
    page_icon="😈",
    layout="wide",
    initial_sidebar_state="expanded"
)

def inject_css():
    st.markdown(
        """
        <style>
            :root {
                --bg-color: #0d1117;
                --card-bg: #161b22;
                --accent: #238636;
                --text: #c9d1d9;
                --border-color: #30363d;
            }

            .stApp {
                background-color: var(--bg-color);
                color: var(--text);
            }

            /* 表头居中 */
            [data-testid="stDataFrame"] th {
                text-align: center !important;
                vertical-align: middle !important;
            }

            /* 单元格居中 */
            [data-testid="stDataFrame"] td {
                text-align: center !important;
                vertical-align: middle !important;
            }

            /* 下拉框居中 */
            [data-testid="stDataFrame"] select {
                text-align: center !important;
                margin: 0 auto !important;
            }
        </style>
        """,
        unsafe_allow_html=True
    )

inject_css()

st.title("✅ CSS 注入成功")
st.dataframe(
    {
        "列A": [1, 2, 3],
        "列B": ["a", "b", "c"]
    }
)
