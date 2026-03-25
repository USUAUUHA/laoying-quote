"""
多客户钣金报价系统首页（Streamlit 多页面主入口）

作用：
- 作为统一导航首页，提供 4 个客户报价入口。
- 使用 st.page_link 跳转，兼容 Streamlit 原生侧边栏导航。
- 不改动任何现有报价脚本，属于纯导航层。
"""

from __future__ import annotations

import streamlit as st


st.set_page_config(
    page_title="青岛宏泰铭润机械 · 钣金智能报价系统（多客户专属版）",
    page_icon="🏭",
    layout="wide",
)


# 轻量样式：保持工业软件风格，强调已上线与待更新状态差异
st.markdown(
    """
    <style>
    .main-title {
        font-size: 40px;
        font-weight: 800;
        margin-bottom: 6px;
        letter-spacing: 0.5px;
    }
    .sub-title {
        font-size: 19px;
        color: #4b5563;
        margin-bottom: 20px;
    }
    .card-online, .card-pending {
        border-radius: 14px;
        padding: 16px 14px;
        border: 1px solid #e5e7eb;
        min-height: 120px;
    }
    .card-online {
        background: linear-gradient(180deg, #fff1f2 0%, #ffffff 100%);
        border-color: #fecdd3;
    }
    .card-pending {
        background: linear-gradient(180deg, #f8fafc 0%, #ffffff 100%);
        border-color: #e5e7eb;
    }
    .card-title {
        font-size: 21px;
        font-weight: 800;
        margin-bottom: 6px;
    }
    .online-text { color: #b91c1c; }
    .pending-text { color: #6b7280; }
    .card-desc {
        font-size: 14px;
        color: #374151;
        line-height: 1.5;
        margin-bottom: 10px;
    }
    .footer-note {
        margin-top: 26px;
        padding: 12px 14px;
        border-radius: 10px;
        background: #f9fafb;
        border: 1px dashed #d1d5db;
        color: #374151;
        font-size: 14px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="main-title">青岛宏泰铭润机械 · 钣金智能报价系统（多客户专属版）</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">适配不同客户专属报价格式，一键切换报价模板，批量报价更高效</div>',
    unsafe_allow_html=True,
)


col1, col2, col3, col4 = st.columns(4, gap="large")

with col1:
    st.markdown(
        """
        <div class="card-pending">
            <div class="card-title pending-text">① 常规报价</div>
            <div class="card-desc">通用钣金报价模板（待更新）</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.page_link("pages/1_常规报价.py", label="进入常规报价", icon="⚙️", use_container_width=True)

with col2:
    st.markdown(
        """
        <div class="card-online">
            <div class="card-title online-text">② 崂应报价</div>
            <div class="card-desc">已上线 · 专属报价格式</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.page_link("pages/2_崂应报价.py", label="进入崂应报价", icon="🚀", use_container_width=True)

with col3:
    st.markdown(
        """
        <div class="card-pending">
            <div class="card-title pending-text">③ 新星报价</div>
            <div class="card-desc">待更新 · 专属报价格式</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.page_link("pages/3_新星报价.py", label="进入新星报价", icon="🧩", use_container_width=True)

with col4:
    st.markdown(
        """
        <div class="card-pending">
            <div class="card-title pending-text">④ 鼎信报价</div>
            <div class="card-desc">待更新 · 专属报价格式</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.page_link("pages/4_鼎信报价.py", label="进入鼎信报价", icon="🏗️", use_container_width=True)


st.markdown(
    """
    <div class="footer-note">
    后续新增客户报价模板，仅需在 <b>pages</b> 文件夹新增对应脚本即可，无需修改现有功能。
    </div>
    """,
    unsafe_allow_html=True,
)
