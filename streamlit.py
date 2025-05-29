import streamlit as st
from PIL import Image

# Page configuration
st.set_page_config(page_title="SlideCopilot - Slide Editor", layout="wide")

# Header
st.markdown("# SlideCopilot: スライド解析・編集")

# Sidebar: Slide Thumbnails
st.sidebar.header("Slides")
num_slides = 5  # example slide count
selected_index = st.sidebar.radio(
    "Select Slide",
    list(range(1, num_slides + 1)),
    index=0,
    format_func=lambda x: f"Slide {x}",
)

# Main: Slide Preview
st.subheader(f"Slide {selected_index} Preview")
placeholder = Image.new("RGB", (800, 450), color=(240, 240, 240))
# use_container_width に置き換え
st.image(placeholder, use_container_width=True)

# 想定質問セクション
st.markdown("---")
st.markdown("### 想定質問")
questions = ["このアルゴリズムの計算量は？", "この手法の適用条件は？"]
for i, q in enumerate(questions, start=1):
    cols = st.columns([8, 1, 1])
    cols[0].write(f"{i}. {q}")
    if cols[1].button("👍", key=f"q_up_{i}"):
        st.write(f"Feedback: Question {i} 👍")
    if cols[2].button("👎", key=f"q_down_{i}"):
        st.write(f"Feedback: Question {i} 👎")

# 補足情報セクション
st.markdown("---")
st.markdown("### 補足情報")
infos = ["アルゴリズムの背景に関する論文紹介", "関連するベンチマークの結果概要"]
for i, info in enumerate(infos, start=1):
    cols = st.columns([8, 1, 1])
    cols[0].write(f"{i}. {info}")
    if cols[1].button("👍", key=f"info_up_{i}"):
        st.write(f"Feedback: Info {i} 👍")
    if cols[2].button("👎", key=f"info_down_{i}"):
        st.write(f"Feedback: Info {i} 👎")
