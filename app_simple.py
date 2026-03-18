import streamlit as st

st.set_page_config(
    page_title="智能CAD报价生成器",
    page_icon="📐",
    layout="wide"
)

st.title("📐 智能CAD报价生成器")
st.markdown("---")

st.success("✅ 应用已成功部署！")

st.subheader("使用说明")
st.markdown("""
1. 上传价格模板Excel文件
2. 上传样式模板Excel文件
3. 上传CAD图纸（DXF格式）
4. 点击生成报价
""")

st.info("🔧 正在完善功能，敬请期待...")
