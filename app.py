                "[타계정대체] 수량": etc_kg, "[타계정대체] 금액": etc_amt,
                "[당월재고] 수량": rem_kg, "[당월재고] 금액": rem_amt
            })
            
        report_df = pd.DataFrame(report_data)
        
        # 가독성 개선: 스타일링 적용하여 매입/불출/대체/재고 칸 배경색 분리
        def highlight_columns(s):
            if '매입' in s.name: return ['background-color: #e6f2ff'] * len(s) # 연한 파랑
            elif '불출' in s.name: return ['background-color: #fff0f0'] * len(s) # 연한 빨강
            elif '대체' in s.name: return ['background-color: #fff9e6'] * len(s) # 연한 노랑
            elif '재고' in s.name: return ['background-color: #e6ffe6; font-weight: bold'] * len(s) # 연한 초록
            return [''] * len(s)

        styled_df = report_df.style.apply(highlight_columns).format(
            {col: "{:,.0f}" if '금액' in col else "{:,.1f}" for col in report_df.columns if '수량' in col or '단가' in col or '금액' in col}
        )
        
        st.dataframe(styled_df, hide_index=True, use_container_width=True)

elif menu.startswith("4"):
    st.markdown('<div class="section-title">■ 경영/총괄 관리 > 전사 재고 자산 현황</div>', unsafe_allow_html=True)
    if not st.session_state.po_master.empty:
        total_in_kg = st.session_state.po_master['수입량(Kg)'].sum(); total_out_kg = st.session_state.transactions['출고수량(Kg)'].sum()
        total_in_amt = st.session_state.po_master['총원가(원)'].sum(); total_out_amt = st.session_state.transactions['출고금액(원)'].sum()
        c1, c2, c3 = st.columns(3)
        c1.metric("📦 창고 실재고 총합", f"{total_in_kg - total_out_kg:,.1f} Kg")
        c2.metric("💰 현재고 자산총액", f"{total_in_amt - total_out_amt:,.0f} 원")
        c3.metric("📈 전체 출고(사용)량", f"{total_out_kg:,.1f} Kg")
        
        summary = st.session_state.transactions.groupby('PO번호').agg(출고합_Kg=('출고수량(Kg)', 'sum'), 출고합_원=('출고금액(원)', 'sum')).reset_index()
        dash_df = pd.merge(st.session_state.po_master[['PO번호', '원산지', '품목명(CL)', '수입량(Kg)', '총원가(원)']], summary, on='PO번호', how='left').fillna(0)
        dash_df['잔여재고(Kg)'] = dash_df['수입량(Kg)'] - dash_df['출고합_Kg']
        dash_df['잔여자산(원)'] = dash_df['총원가(원)'] - dash_df['출고합_원']
        dash_df = dash_df[dash_df['잔여재고(Kg)'] > 0.1].drop(columns=['수입량(Kg)', '총원가(원)', '출고합_Kg', '출고합_원'])
        
        st.markdown("##### 📌 원산지/품목별 상세 잔액")
        st.dataframe(dash_df.style.format({"잔여재고(Kg)":"{:,.1f}", "잔여자산(원)":"{:,.0f}"}), hide_index=True, use_container_width=True)
