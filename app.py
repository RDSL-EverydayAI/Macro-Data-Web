import streamlit as st
import pandas as pd
import time
import io
import datetime
from cnstats.stats import stats

# 页面基本设置
st.set_page_config(page_title="中国宏观经济数据获取器", page_icon="🇨🇳")
st.title("📊 中国宏观经济核心数据获取器")
st.markdown("数据来源：中国国家统计局 (NBS) | 基于 `cn-stats` 引擎构建")

indicators = [
    {"name": "居民消费价格指数 (CPI)", "code": "A010101"},
    {"name": "工业生产者出厂价格指数 (PPI)", "code": "A010801"},
    {"name": "制造业采购经理指数 (PMI)", "code": "A0B0101"},
    {"name": "货币和准货币供应量_期末值 (M2)", "code": "A0D0101"},
    {"name": "社会消费品零售总额", "code": "A070101"},
    {"name": "固定资产投资(不含农户)", "code": "A050101"},
    {"name": "出口总值", "code": "A080103"},
    {"name": "进口总值", "code": "A080104"},
    {"name": "城镇调查失业率", "code": "A0C0101"},
    {"name": "外汇储备", "code": "A0E0101"}
]

# 当用户点击按钮时触发
if st.button("🚀 开始向国家统计局获取最新数据"):
    all_data = []
    
    # 动态计算时间区间：过往3年 + 当年
    current_year = datetime.datetime.now().year # 2026
    start_year = current_year - 3 # 2023
    
    # 构造符合 cn-stats 规范的日期字符串，例如 "2023-2026"
    # 这样国家统计局接口就会精确返回这四年内的所有月度历史数据
    target_date_str = f"{start_year}-{current_year}" 
    
    # 创建一个进度条和状态提示
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, ind in enumerate(indicators):
        status_text.text(f"📡 正在拉取: {ind['name']} ({target_date_str})...")
        try:
            # 💡 核心修复：传入明确的 datestr 参数
            df = stats(zbcode=ind['code'], datestr=target_date_str, as_df=True)
            
            if df is not None and not df.empty:
                for index, row in df.iterrows():
                    # 兼容不同版本 cn-stats 返回的列名
                    date_val = str(row.get('查询日期', row.get('时间', '')))
                    value = row.get('数值', row.get('数据', ''))
                    all_data.append([ind['name'], ind['code'], date_val, value])
                st.write(f"  ✅ {ind['name']} 成功获取 {len(df)} 条记录")
            else:
                st.warning(f"  ⚠️ {ind['name']} 返回了空数据")
        except Exception as e:
            st.error(f"❌ 获取 {ind['name']} 时报错: {e}")
            
        time.sleep(2.0) # 维持 2 秒休眠，这是在云端高成功率越过统计局防火墙的关键
        progress_bar.progress((i + 1) / len(indicators))
        
    status_text.text("✅ 数据获取完成！正在打包 Excel...")
    
    if all_data:
        # 将数据转换为 DataFrame 并按照指标名和时间倒序排序
        result_df = pd.DataFrame(all_data, columns=['指标名称', '官方指标代码', '年月/周期', '数值'])
        result_df = result_df.sort_values(by=['指标名称', '年月/周期'], ascending=[True, False])
        
        # 在内存中生成 Excel 文件
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='宏观数据一览')
        excel_data = output.getvalue()
        
        st.success(f"🎉 全部核心指标拉取成功！共筛选出 {len(all_data)} 条月度历史记录。")
        
        # 显示下载按钮
        st.download_button(
            label="📥 点击这里下载完整的 Excel 数据表",
            data=excel_data,
            file_name=f"中国宏观经济核心数据_{target_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("⚠️ 未能获取到任何有效数据，请检查运行日志。")
