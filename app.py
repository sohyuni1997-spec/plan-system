import streamlit as st
import pandas as pd
import io
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="생산계획 배분 시스템", layout="wide")

st.title("📊 생산계획 배분 시스템")
st.markdown("---")

uploaded_file = st.file_uploader("📁 0차계획.xlsx 파일을 업로드하세요", type=['xlsx'])

if uploaded_file:
    
    df = pd.read_excel(uploaded_file, header=None, skiprows=11, nrows=6)
    df_filtered = df[df[0].astype(str).str.contains('FAN|FLANGE', case=True, na=False)].copy()
    
    numbers_original = df_filtered.iloc[:, 6:34].copy()
    
    numbers = df_filtered.iloc[:, 6:34].copy()
    units = df_filtered[2]
    result = pd.DataFrame(0, index=numbers.index, columns=numbers.columns)
    
    for row_idx in numbers.index:
        unit = units.loc[row_idx] if pd.notna(units.loc[row_idx]) else 1
        
        for col_idx, col in enumerate(numbers.columns):
            value = numbers.loc[row_idx, col]
            
            if pd.isna(value) or value == 0:
                continue
            
            for i in range(min(4, col_idx + 1)):
                target_col = numbers.columns[col_idx - i]
                current_sum = result[target_col].sum()
                
                if current_sum < 3300:
                    add = min(unit, 3300 - current_sum)
                    result.loc[row_idx, target_col] += add
                    value -= add
                    
                    if value <= 0:
                        break
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📋 원본 데이터")
        st.dataframe(numbers_original, use_container_width=True)
        st.metric("원본 총 생산량", f"{numbers_original.sum().sum():,.0f}")
    
    with col2:
        st.subheader("✅ 배분 결과")
        st.dataframe(result, use_container_width=True)
        st.metric("배분 후 총 생산량", f"{result.sum().sum():,.0f}")
    
    st.markdown("---")
    
    st.subheader("📈 데이터 시각화")
    
    tab1, tab2, tab3, tab4 = st.tabs(["📊 제품별 비교", "📉 일별 생산량", "🎯 CAPA 활용률", "📦 제품별 분포"])
    
    with tab1:
        product_names = [df_filtered.loc[idx, 0] for idx in result.index]
        
        comparison_data = pd.DataFrame({
            '제품': product_names,
            '원본': numbers_original.sum(axis=1).values,
            '배분 후': result.sum(axis=1).values
        })
        
        fig1 = go.Figure(data=[
            go.Bar(name='원본', x=comparison_data['제품'], y=comparison_data['원본']),
            go.Bar(name='배분 후', x=comparison_data['제품'], y=comparison_data['배분 후'])
        ])
        fig1.update_layout(barmode='group', title='제품별 생산량 비교')
        st.plotly_chart(fig1, use_container_width=True)
    
    with tab2:
        daily_sum = result.sum(axis=0)
        
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=list(range(len(daily_sum))),
            y=daily_sum.values,
            mode='lines+markers',
            name='일별 생산량',
            line=dict(color='#2196F3', width=3)
        ))
        fig2.add_hline(y=3300, line_dash="dash", line_color="red", 
                      annotation_text="CAPA 3300")
        fig2.update_layout(title='일별 생산량 추이', 
                          xaxis_title='일자', 
                          yaxis_title='생산량')
        st.plotly_chart(fig2, use_container_width=True)
        
        col_a, col_b, col_c, col_d = st.columns(4)
        col_a.metric("평균", f"{daily_sum.mean():,.0f}")
        col_b.metric("최대", f"{daily_sum.max():,.0f}")
        col_c.metric("최소", f"{daily_sum.min():,.0f}")
        col_d.metric("표준편차", f"{daily_sum.std():,.0f}")
    
    with tab3:
        daily_sum = result.sum(axis=0)
        utilization = (daily_sum / 3300 * 100).round(1)
        
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(
            x=list(range(len(utilization))),
            y=utilization.values,
            marker_color=['green' if x <= 100 else 'red' for x in utilization.values]
        ))
        fig3.add_hline(y=100, line_dash="dash", line_color="red", 
                      annotation_text="100% CAPA")
        fig3.update_layout(title='일별 CAPA 활용률 (%)', 
                          xaxis_title='일자', 
                          yaxis_title='활용률 (%)')
        st.plotly_chart(fig3, use_container_width=True)
        
        over_capa = utilization[utilization > 100]
        if len(over_capa) > 0:
            st.error(f"⚠️ CAPA 초과 일자: {len(over_capa)}일")
        else:
            st.success("✅ 모든 일자가 CAPA 이내입니다!")
    
    with tab4:
        product_totals = result.sum(axis=1)
        
        fig4 = go.Figure(data=[go.Pie(
            labels=product_names,
            values=product_totals.values,
            hole=.3
        )])
        fig4.update_layout(title='제품별 생산량 분포')
        st.plotly_chart(fig4, use_container_width=True)
    
    st.markdown("---")
    
    with st.expander("📋 상세 데이터 보기"):
        st.subheader("각 열(일자)별 합계")
        daily_detail = pd.DataFrame({
            '일자': list(range(len(result.columns))),
            '생산량': result.sum(axis=0).values,
            'CAPA 활용률(%)': (result.sum(axis=0) / 3300 * 100).round(1).values,
            'CAPA 여유': (3300 - result.sum(axis=0)).values
        })
        st.dataframe(daily_detail, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📥 결과 다운로드")
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result_with_product = result.copy()
        result_with_product.insert(0, '제품명', product_names)
        result_with_product.to_excel(writer, sheet_name='배분결과', index=False)
        
        original_with_product = numbers_original.copy()
        original_with_product.insert(0, '제품명', product_names)
        original_with_product.to_excel(writer, sheet_name='원본데이터', index=False)
        
        daily_detail.to_excel(writer, sheet_name='일별합계', index=False)
        
        comparison_data.to_excel(writer, sheet_name='제품별합계', index=False)
    
    col_download1, col_download2, col_download3 = st.columns([1, 1, 2])
    
    with col_download1:
        st.download_button(
            label="📥 전체 결과 다운로드",
            data=output.getvalue(),
            file_name="배분결과_전체.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col_download2:
        csv = result.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📄 CSV 다운로드",
            data=csv,
            file_name="배분결과.csv",
            mime="text/csv",
            use_container_width=True
        )

else:
    st.info("👆 엑셀 파일을 업로드하면 자동으로 처리됩니다.")
    
    st.markdown('''
    ### 📌 사용 방법
    1. **0차계획.xlsx** 파일을 업로드하세요
    2. 자동으로 FAN, FLANGE 제품이 추출됩니다
    3. CAPA 3300 기준으로 생산량이 배분됩니다
    4. 그래프로 한눈에 확인하세요
    5. 수정된 엑셀 파일을 다운로드하세요
    
    ### ✅ 기능
    - 📊 제품별 생산량 비교
    - 📉 일별 생산량 추이
    - 🎯 CAPA 활용률 분석
    - 📦 제품별 분포 차트
    - 📥 엑셀/CSV 다운로드
    ''')

st.markdown("---")
st.caption("생산계획 배분 시스템 v1.0")
