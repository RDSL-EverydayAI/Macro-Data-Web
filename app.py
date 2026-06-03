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
    
    current_year = datetime.datetime.now().year # 2026
    start_year = current_year - 3 # 2023
    
    # 创建一个进度条和状态提示
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, ind in enumerate(indicators):
        status_text.text(f"📡 正在拉取: {ind['name']}...")
        
        # 💡 策略优化：分年度循环抓取，进一步降低单次被拦截的概率
        for year in range(start_year, current_year + 1):
            target_date_str = str(year)
            try:
                # 核心防线：用 try-except 包裹请求，防止单次解析失败导致整个程序退出
                df = stats(zbcode=ind['code'], datestr=target_date_str, as_df=True)
                
                if df is not None and not df.empty:
                    for index, row in df.iterrows():
                        date_val = str(row.get('查询日期', row.get('时间', '')))
                        value = row.get('数值', row.get('数据', ''))
                        all_data.append([ind['name'], ind['code'], date_val, value])
            except Exception as e:
                # 如果被拦截报错，不弹红，而是以黄色警告框提示，程序继续运行
                st.warning(f"⚠️ {ind['name']} ({target_date_str}年) 被统计局防火墙阻截，已自动跳过。错误原因: {e}")
            
            time.sleep(2.5) # 每个小请求休眠 2.5 秒，温柔对待统计局服务器
            
        # 汇报当前指标的整体完成情况
        ind_count = len([x for x in all_data if x[0] == ind['name']])
        if ind_count > 0:
            st.write(f"  ✅ {ind['name']} 获取成功，共 {ind_count} 条月度数据")
            
        progress_bar.progress((i + 1) / len(indicators))
        
    status_text.text("✅ 数据拉取流程结束，正在打包 Excel...")
    
    if all_data:
        # 将数据转换为 DataFrame 并按照指标名和时间倒序排序
        result_df = pd.DataFrame(all_data, columns=['指标名称', '官方指标代码', '年月/周期', '数值'])
        result_df = result_df.sort_values(by=['指标名称', '年月/周期'], ascending=[True, False])
        
        # 在内存中生成 Excel 文件
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='宏观数据一览')
        excel_data = output.getvalue()
        
        st.success(f"🎉 任务完成！成功为您抢救回 {len(all_data)} 条有效数据记录。")
        
        # 显示下载按钮
        st.download_button(
            label="📥 点击这里下载已获取到的 Excel 数据表",
            data=excel_data,
            file_name=f"中国宏观经济核心数据_{start_year}-{current_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("❌ 非常遗憾，本次所有尝试均被统计局防火墙全面拦截（返回了空页面）。请几分钟后再次点击按钮重试，或尝试刷新网页更换云端节点 IP。")
