import streamlit as st
import pandas as pd
import time
import io
from cnstats.stats import stats

# 页面基本设置
st.set_page_config(page_title="中国宏观经济数据获取器", page_icon="🇨🇳")
st.title("📊 中国宏观经济核心数据获取器")
st.markdown("数据来源：中国国家统计局 (NBS) | 基于 `cn-stats` 引擎构建")

indicators = [
    {"name": "居民消费价格指数 (CPI)", "code": "A010101"},
    {"name": "工业生产者出厂价格指数 (PPI)", "code": "A010801"},
    {"name": "制造业采购经理指数 (PMI)", "code": "A0B0101"},
    {"name": "M2货币供应量", "code": "A0D0101"},
    {"name": "社会消费品零售总额", "code": "A070101"},
    {"name": "固定资产投资", "code": "A050101"},
    {"name": "出口总值", "code": "A080103"},
    {"name": "进口总值", "code": "A080104"},
    {"name": "城镇调查失业率", "code": "A0C0101"},
    {"name": "外汇储备", "code": "A0E0101"}
]

# 当用户点击按钮时触发
if st.button("🚀 开始向国家统计局获取最新数据"):
    all_data = []
    
    # 创建一个进度条和状态提示
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, ind in enumerate(indicators):
        status_text.text(f"📡 正在拉取: {ind['name']}...")
        try:
            df = stats(zbcode=ind['code'], as_df=True)
            if df is not None and not df.empty:
                for index, row in df.iterrows():
                    date_val = str(row.get('查询日期', row.get('时间', '')))
                    value = row.get('数值', row.get('数据', ''))
                    all_data.append([ind['name'], ind['code'], date_val, value])
        except Exception as e:
            st.warning(f"❌ 获取 {ind['name']} 时报错: {e}")
            
        time.sleep(1.5) # 防止请求过快被拦截
        progress_bar.progress((i + 1) / len(indicators))
        
    status_text.text("✅ 数据获取完成！正在生成 Excel 文件...")
    
    if all_data:
        # 将数据转换为 DataFrame
        result_df = pd.DataFrame(all_data, columns=['指标名称', '官方指标代码', '年月/周期', '数值'])
        
        # 在内存中生成 Excel 文件（不保存在硬盘，方便网页直接下载）
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='宏观数据')
        excel_data = output.getvalue()
        
        st.success(f"🎉 成功抓取了 {len(all_data)} 条数据！请点击下方按钮下载。")
        
        # 显示下载按钮
        st.download_button(
            label="📥 下载 Excel 数据文件",
            data=excel_data,
            file_name="中国核心宏观经济数据.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("⚠️ 未能获取到任何数据，可能是国家统计局暂时拦截了云端服务器的 IP。")
