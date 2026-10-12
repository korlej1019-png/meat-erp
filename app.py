import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="금화 수입육 통합 ERP (가독성 개선)", page_icon="🏢", layout="wide", initial_sidebar_state="expanded")

st.markdown('''
<style>
body { font-family: 'Malgun Gothic', sans-serif; background-color: #f3f4f6; }
.erp-header { background-color: #1c3d5a; color: white; padding: 15px 20px; font-size: 1.5rem; font-weight: bold; border-radius: 5px; margin-bottom: 20px; }
.section-title { font-size: 1.1rem; font-weight: 600; color: #333; border-bottom: 2px solid #1c3d5a; padding-bottom: 8px; margin-bottom: 15px; margin-top: 20px; }
.stMetric { background-color: #fff; padding: 15px; border-radius: 4px; border: 1px solid #ddd; border-top: 3px solid #1c3d5a; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
div[data-testid="stForm"] { background-color: #fff; border: 1px solid #ddd; padding: 20px; border-radius: 4px; }
</style>
''', unsafe_allow_html=True)

st.markdown('<div class="erp-header">🏢 금화 수입육 관리 시스템 (가독성 개선 V8)</div>', unsafe_allow_html=True)

def fix_dtypes():
    num_cols_po = ['계약환율', '외화금액(USD)', '수입량(Kg)', '원료육대금(원)', '적하보험료(원)', '포워딩비용(원)', '관세비용(원)', '보세창고비용(원)', '제비용(원)', '총원가(원)', 'Kg당단가(원)']
    if not st.session_state.po_master.empty:
        for c in num_cols_po:
            if c in st.session_state.po_master.columns:
                st.session_state.po_master[c] = pd.to_numeric(st.session_state.po_master[c], errors='coerce').fillna(0)
    
    num_cols_tx = ['출고수량(Kg)', '적용단가', '출고금액(원)']
    if not st.session_state.transactions.empty:
        for c in num_cols_tx:
            if c in st.session_state.transactions.columns:
                st.session_state.transactions[c] = pd.to_numeric(st.session_state.transactions[c], errors='coerce').fillna(0)
        try:
            st.session_state.transactions['출고일자'] = pd.to_datetime(st.session_state.transactions['출고일자']).dt.strftime('%Y-%m-%d')
            st.session_state.transactions.sort_values(by='출고일자', inplace=True)
            st.session_state.transactions.reset_index(drop=True, inplace=True)
        except: pass

def auto_migrate_raw_excel(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file)
        po_master_list = []
        if '수입비용' in xls.sheet_names:
            df_cost = pd.read_excel(xls, sheet_name='수입비용', header=None)
            for col_idx in range(1, df_cost.shape[1]):
                po_val = df_cost.iloc[2, col_idx]
                if pd.isna(po_val): continue
                po_no = str(po_val).strip()
                if po_no == '' or po_no == 'nan': continue
                if 'PO' not in po_no.upper(): po_no = f"PO-{po_no}"
                    
                import_kg = df_cost.iloc[15, col_idx]
                import_kg_num = pd.to_numeric(import_kg, errors='coerce')
                
                meat_cost = pd.to_numeric(df_cost.iloc[29:46, col_idx], errors='coerce').fillna(0).sum()
                insurance = pd.to_numeric(df_cost.iloc[47, col_idx], errors='coerce')
                insurance = insurance if pd.notna(insurance) else 0
                forwarding = pd.to_numeric(df_cost.iloc[49:67, col_idx], errors='coerce').fillna(0).sum()
                customs = pd.to_numeric(df_cost.iloc[68:76, col_idx], errors='coerce').fillna(0).sum()
                storage = pd.to_numeric(df_cost.iloc[76:85, col_idx], errors='coerce').fillna(0).sum()
                
                total_expenses = insurance + forwarding + customs + storage
                total_cost = meat_cost + total_expenses
                
                unit_price = total_cost / import_kg_num if pd.notna(import_kg_num) and import_kg_num > 0 else 0
                
                ct_date = df_cost.iloc[4, col_idx]
                cc_date = df_cost.iloc[87, col_idx]
                
                po_master_list.append({
                    '계약일자': str(ct_date)[:10] if pd.notna(ct_date) else '',
                    'PO번호': po_no,
                    '원산지': str(df_cost.iloc[11, col_idx]) if pd.notna(df_cost.iloc[11, col_idx]) else '',
                    '품목명(CL)': str(df_cost.iloc[3, col_idx]) if pd.notna(df_cost.iloc[3, col_idx]) else '',
                    '계약환율': pd.to_numeric(df_cost.iloc[5, col_idx], errors='coerce'),
                    '외화금액(USD)': pd.to_numeric(df_cost.iloc[24, col_idx], errors='coerce'),
                    '수입량(Kg)': import_kg_num,
                    '통관일자': str(cc_date)[:10] if pd.notna(cc_date) else '',
                    '원료육대금(원)': meat_cost, '적하보험료(원)': insurance, '포워딩비용(원)': forwarding,
                    '관세비용(원)': customs, '보세창고비용(원)': storage, '제비용(원)': total_expenses,
                    '총원가(원)': total_cost, 'Kg당단가(원)': unit_price
                })
        
        tx_list = []
        cl_sheets = [s for s in xls.sheet_names if s.endswith('CL')]
        for sheet in cl_sheets:
            df_cl = pd.read_excel(xls, sheet_name=sheet, header=None)
            current_po = None; current_origin = None
            for row_idx in range(3, df_cl.shape[0]):
                po_val = df_cl.iloc[row_idx, 2]
                if pd.notna(po_val) and str(po_val).strip() not in ['', 'nan']:
                    po_str = str(po_val).strip()
                    if 'PO' not in po_str.upper(): po_str = f"PO-{po_str}"
                    current_po = po_str
                    origin_val = df_cl.iloc[row_idx, 7]
                    current_origin = str(origin_val) if pd.notna(origin_val) else ''
                
                tx_date = df_cl.iloc[row_idx, 11]
                if pd.notna(tx_date) and current_po:
                    out_kg = pd.to_numeric(df_cl.iloc[row_idx, 12], errors='coerce')
                    out_price = pd.to_numeric(df_cl.iloc[row_idx, 13], errors='coerce')
                    if pd.notna(out_kg) and out_kg > 0:
                        tx_list.append({
                            '출고일자': str(tx_date)[:10], '구분': '당월불출(생산투입)', '원산지': current_origin,
                            'PO번호': current_po, '품목명(CL)': sheet, '출고수량(Kg)': out_kg,
                            '적용단가': out_price if pd.notna(out_price) else 0,
                            '출고금액(원)': (out_kg * out_price) if pd.notna(out_price) else 0, '비고': '자동 이관'
                        })
                        
                    etc_kg = pd.to_numeric(df_cl.iloc[row_idx, 15], errors='coerce')
                    etc_price = pd.to_numeric(df_cl.iloc[row_idx, 16], errors='coerce')
                    if pd.notna(etc_kg) and etc_kg > 0:
                        tx_list.append({
                            '출고일자': str(tx_date)[:10], '구분': '타계정대체(폐기/판매등)', '원산지': current_origin,
                            'PO번호': current_po, '품목명(CL)': sheet, '출고수량(Kg)': etc_kg,
                            '적용단가': etc_price if pd.notna(etc_price) else 0,
                            '출고금액(원)': (etc_kg * etc_price) if pd.notna(etc_price) else 0, '비고': '자동 이관'
                        })
                        
        if po_master_list: st.session_state.po_master = pd.DataFrame(po_master_list)
        if tx_list: st.session_state.transactions = pd.DataFrame(tx_list)
        fix_dtypes()
        return True, len(po_master_list), len(tx_list)
    except Exception as e:
        return False, str(e), 0

if 'po_master' not in st.session_state:
    st.session_state.po_master = pd.DataFrame(columns=[
        '계약일자', 'PO번호', '원산지', '품목명(CL)', '계약환율', '외화금액(USD)', 
        '수입량(Kg)', '통관일자', '원료육대금(원)', '적하보험료(원)', '포워딩비용(원)', '관세비용(원)', '보세창고비용(원)', 
        '제비용(원)', '총원가(원)', 'Kg당단가(원)'
    ])
if 'transactions' not in st.session_state:
    st.session_state.transactions = pd.DataFrame(columns=[
        '출고일자', '구분', '원산지', 'PO번호', '품목명(CL)', '출고수량(Kg)', '적용단가', '출고금액(원)', '비고'
    ])

with st.sidebar:
    st.markdown("### 📋 ERP 메뉴")
    menu = st.radio("작업 영역 선택", [
        "0. [초기세팅] 원본 엑셀 자동 마이그레이션",
        "1. [수입/구매] PO 등록 및 단가 확정", 
        "2. [생산/재고] 공장 생산일보 수불 등록", 
        "3. [회계/결산] 품목별 수불명세서 조회",
        "4. [경영/총괄] 재고 자산 대시보드"
    ])
    
    st.divider()
    st.markdown("### 💾 데이터 마감 및 복구")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.po_master.to_excel(writer, index=False, sheet_name='PO마스터')
        st.session_state.transactions.to_excel(writer, index=False, sheet_name='수불부')
    st.download_button("📥 전체 장부 엑셀 백업", data=output.getvalue(), file_name=f"금화_수입육장부_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    
    uploaded_file = st.file_uploader("📂 백업 파일 복구", type=['xlsx'], label_visibility="collapsed")
    if uploaded_file is not None:
        if st.button("장부 복구", use_container_width=True):
            st.session_state.po_master = pd.read_excel(uploaded_file, sheet_name='PO마스터')
            st.session_state.transactions = pd.read_excel(uploaded_file, sheet_name='수불부')
            fix_dtypes()
            st.success("복구 완료!")
            st.rerun()

if menu.startswith("0"):
    st.markdown('<div class="section-title">■ 시스템 초기화 > 원본 엑셀(생산정리.xlsx) 자동 이관기</div>', unsafe_allow_html=True)
    st.info("과장님께서 기존에 사용하시던 복잡한 '생산(BKR Beef Patty 수율)정리.xlsx' 파일을 업로드하시면, 프로그램이 수십 개의 탭과 행/열을 0.1초 만에 자동 분석하여 ERP 데이터베이스로 100% 깔끔하게 마이그레이션합니다.")
    raw_file = st.file_uploader("원본 '생산정리.xlsx' 파일 업로드", type=['xlsx'])
    if raw_file is not None:
        if st.button("🚀 원본 데이터 자동 마이그레이션 실행", type="primary", use_container_width=True):
            with st.spinner('복잡한 원본 데이터를 ERP 양식으로 변환 중입니다...'):
                success, po_cnt, tx_cnt = auto_migrate_raw_excel(raw_file)
                if success: st.success(f"🎉 데이터 마이그레이션 대성공! (총 {po_cnt}개 PO, {tx_cnt}건의 수불 내역 자동 변환 완료)")
                else: st.error(f"변환 중 오류 발생: {po_cnt}")

elif menu.startswith("1"):
    st.markdown('<div class="section-title">■ 수입/구매 관리 > 수입육 PO 원가 및 제비용 등록</div>', unsafe_allow_html=True)
    
    existing_pos = ["(+) 신규 PO 등록"] + st.session_state.po_master['PO번호'].astype(str).tolist()
    selected_action = st.selectbox("진행 상태 선택 (신규 등록 또는 기존 PO 비용 업데이트)", existing_pos)
    
    d_ct_date = datetime.today(); d_po = ""; d_org = "호주"; d_item = "95CL"; d_rate = 1350.0; d_usd = 0.0
    d_kg = 0.0; d_cc_date = datetime.today()
    d_meat = 0; d_ins = 0; d_fwd = 0; d_cust = 0; d_stor = 0
    
    is_update = False
    if selected_action != "(+) 신규 PO 등록":
        is_update = True
        row = st.session_state.po_master[st.session_state.po_master['PO번호'] == selected_action].iloc[0]
        try: d_ct_date = pd.to_datetime(row['계약일자'])
        except: pass
        d_po = row['PO번호']; d_org = row['원산지']; d_item = row['품목명(CL)']
        d_rate = float(row['계약환율']); d_usd = float(row['외화금액(USD)']); d_kg = float(row['수입량(Kg)'])
        try: d_cc_date = pd.to_datetime(row['통관일자'])
        except: pass
        d_meat = int(row['원료육대금(원)']); d_ins = int(row['적하보험료(원)'])
        d_fwd = int(row['포워딩비용(원)']); d_cust = int(row['관세비용(원)']); d_stor = int(row['보세창고비용(원)'])

    with st.form("po_form"):
        st.markdown("**[외환 및 계약 정보]**")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: ct_date = st.date_input("계약일자", value=d_ct_date)
        with c2: po_no = st.text_input("PO 번호", value=d_po, disabled=is_update)
        with c3: origin = st.selectbox("원산지", ["호주", "뉴질랜드", "미국", "기타"], index=["호주", "뉴질랜드", "미국", "기타"].index(d_org) if d_org in ["호주", "뉴질랜드", "미국", "기타"] else 0)
        with c4: item_name = st.selectbox("품목명(CL)", ["95CL", "90CL", "85CL", "80CL", "75CL", "70CL", "65CL"], index=["95CL", "90CL", "85CL", "80CL", "75CL", "70CL", "65CL"].index(d_item) if d_item in ["95CL", "90CL", "85CL", "80CL", "75CL", "70CL", "65CL"] else 0)
        with c5: cc_date = st.date_input("통관일자", value=d_cc_date)
        
        c6, c7, c8 = st.columns(3)
        with c6: ex_rate = st.number_input("계약환율 (원/USD)", value=d_rate, step=10.0, format="%.2f")
        with c7: usd_amt = st.number_input("외화금액 (USD)", value=d_usd, step=1.0)
        with c8: import_kg = st.number_input("수입 수량 (Kg)", value=d_kg, step=100.0)
        
        st.markdown("**[원가 및 제비용 청구 내역]**")
        c9, c10, c11, c12, c13 = st.columns(5)
        with c9: meat_cost = st.number_input("원료육 대금 (원)", value=d_meat, step=100000)
        with c10: insurance = st.number_input("적하보험료 (원)", value=d_ins, step=10000)
        with c11: forwarding = st.number_input("포워딩 비용 (원)", value=d_fwd, step=10000)
        with c12: customs = st.number_input("관세비용 (원)", value=d_cust, step=10000)
        with c13: storage = st.number_input("보세창고 비용 (원)", value=d_stor, step=10000)
        
        submit_po = st.form_submit_button("저장 및 원가 단가 확정(소급)", type="primary")
        if submit_po and po_no and import_kg > 0:
            total_expenses = insurance + forwarding + customs + storage
            total_cost = meat_cost + total_expenses
            unit_price = total_cost / import_kg
            if is_update:
                idx = st.session_state.po_master[st.session_state.po_master['PO번호'] == po_no].index
                st.session_state.po_master.loc[idx, ['계약일자', '원산지', '품목명(CL)', '계약환율', '외화금액(USD)', '수입량(Kg)', '통관일자', '원료육대금(원)', '적하보험료(원)', '포워딩비용(원)', '관세비용(원)', '보세창고비용(원)', '제비용(원)', '총원가(원)', 'Kg당단가(원)']] = [ct_date.strftime('%Y-%m-%d'), origin, item_name, ex_rate, usd_amt, import_kg, cc_date.strftime('%Y-%m-%d'), meat_cost, insurance, forwarding, customs, storage, total_expenses, total_cost, unit_price]
                tx_mask = st.session_state.transactions['PO번호'] == po_no
                st.session_state.transactions.loc[tx_mask, '적용단가'] = unit_price
                st.session_state.transactions.loc[tx_mask, '출고금액(원)'] = st.session_state.transactions.loc[tx_mask, '출고수량(Kg)'] * unit_price
                fix_dtypes()
                st.success(f"[{po_no}] 단가 업데이트 완료 (1Kg당 {unit_price:,.2f}원)")
            else:
                if po_no in st.session_state.po_master['PO번호'].values: st.error("이미 등록된 PO입니다.")
                else:
                    st.session_state.po_master.loc[len(st.session_state.po_master)] = [ct_date.strftime('%Y-%m-%d'), po_no, origin, item_name, ex_rate, usd_amt, import_kg, cc_date.strftime('%Y-%m-%d'), meat_cost, insurance, forwarding, customs, storage, total_expenses, total_cost, unit_price]
                    fix_dtypes()
                    st.success(f"[{po_no}] 신규 등록 완료. 가단가: 1Kg당 {unit_price:,.2f}원")
            st.rerun()

    st.markdown("##### 📌 수입 원가 마스터 (입고 내역)")
    st.dataframe(st.session_state.po_master.style.format({col: "{:,.0f}" for col in st.session_state.po_master.columns if '(원)' in col or '환율' in col or '(Kg)' in col or 'USD' in col}), hide_index=True, use_container_width=True)

elif menu.startswith("2"):
    st.markdown('<div class="section-title">■ 생산/재고 관리 > 현장 수불 일보 등록</div>', unsafe_allow_html=True)
    if st.session_state.po_master.empty:
        st.warning("수입/구매 관리에서 PO를 먼저 등록해주세요.")
    else:
        with st.form("tx_form"):
            c1, c2, c3, c4 = st.columns(4)
            with c1: tx_date = st.date_input("출고일자")
            with c2: tx_type = st.selectbox("출고 구분", ["당월불출(생산투입)", "타계정대체(폐기/판매등)"])
            with c3: 
                po_choices = st.session_state.po_master['PO번호'].astype(str) + " (" + st.session_state.po_master['품목명(CL)'] + ")"
                selected_po_display = st.selectbox("대상 PO 선택", po_choices)
            with c4: tx_kg = st.number_input("출고 수량 (Kg)", min_value=0.0, step=10.0)
            tx_note = st.text_input("적요/비고 (예: 버거패티 BKR 투입분)")
            if st.form_submit_button("전표 발행 (수량 차감)", type="primary") and tx_kg > 0:
                tx_po = selected_po_display.split(" ")[0]
                po_info = st.session_state.po_master[st.session_state.po_master['PO번호'] == tx_po].iloc[0]
                unit_price = po_info['Kg당단가(원)']
                st.session_state.transactions.loc[len(st.session_state.transactions)] = [tx_date.strftime('%Y-%m-%d'), tx_type, po_info['원산지'], tx_po, po_info['품목명(CL)'], tx_kg, unit_price, tx_kg * unit_price, tx_note]
                fix_dtypes()
                st.success("정상적으로 출고 처리되었습니다.")
                st.rerun()

        st.markdown("##### 📌 당월 수불 전표 내역")
        # 가독성 개선: 컬럼명 한글화 및 포맷팅, 인덱스 숨기기
        st.dataframe(
            st.session_state.transactions,
            hide_index=True,
            use_container_width=True,
            column_config={
                "출고일자": st.column_config.TextColumn("📅 출고일자", width="small"),
                "구분": st.column_config.TextColumn("🏷️ 구분", width="medium"),
                "원산지": st.column_config.TextColumn("🌍 원산지", width="small"),
                "PO번호": st.column_config.TextColumn("📑 PO번호", width="small"),
                "품목명(CL)": st.column_config.TextColumn("🥩 품목", width="small"),
                "출고수량(Kg)": st.column_config.NumberColumn("⚖️ 출고량(Kg)", format="%,.1f"),
                "적용단가": st.column_config.NumberColumn("💲 단가(원)", format="%,.2f"),
                "출고금액(원)": st.column_config.NumberColumn("💵 출고금액", format="%,.0f"),
                "비고": st.column_config.TextColumn("📝 비고")
            }
        )

elif menu.startswith("3"):
    st.markdown('<div class="section-title">■ 회계/결산 관리 > 품목별 수불명세서 (더존 양식)</div>', unsafe_allow_html=True)
    if st.session_state.po_master.empty:
        st.info("데이터가 없습니다.")
    else:
        cl_list = sorted(list(set(st.session_state.po_master['품목명(CL)'].tolist())))
        selected_cl = st.selectbox("조회할 품목(CL) 선택", cl_list)
        po_filtered = st.session_state.po_master[st.session_state.po_master['품목명(CL)'] == selected_cl]
        tx_filtered = st.session_state.transactions[st.session_state.transactions['품목명(CL)'] == selected_cl]
        
        st.markdown(f"#### 📄 소고기트리밍 {selected_cl} 수불명세서")
        report_data = []
        for index, po in po_filtered.iterrows():
            po_no = po['PO번호']; in_kg = po['수입량(Kg)']; in_price = po['Kg당단가(원)']; in_amt = po['총원가(원)']
            out_prod = tx_filtered[(tx_filtered['PO번호'] == po_no) & (tx_filtered['구분'] == '당월불출(생산투입)')]
            prod_kg = out_prod['출고수량(Kg)'].sum(); prod_amt = out_prod['출고금액(원)'].sum()
            out_etc = tx_filtered[(tx_filtered['PO번호'] == po_no) & (tx_filtered['구분'] == '타계정대체(폐기/판매등)')]
            etc_kg = out_etc['출고수량(Kg)'].sum(); etc_amt = out_etc['출고금액(원)'].sum()
            rem_kg = in_kg - prod_kg - etc_kg; rem_amt = in_amt - prod_amt - etc_amt
            report_data.append({
                "PO No": po_no, "원산지": po['원산지'],
                "[매입] 수량": in_kg, "[매입] 단가": in_price, "[매입] 금액": in_amt,
                "[당월불출] 수량": prod_kg, "[당월불출] 금액": prod_amt,
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
