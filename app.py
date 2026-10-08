import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="금화 수입육 통합 ERP", page_icon="🥩", layout="wide", initial_sidebar_state="expanded")

# --- 디자인 요소 (CSS 적용) ---
st.markdown('''
<style>
.main-header { font-size: 2.2rem; font-weight: 800; color: #1E3A8A; margin-bottom: 0.2rem; }
.sub-header { font-size: 1.05rem; color: #6B7280; margin-bottom: 2rem; }
.stMetric { background-color: #F8FAFC; padding: 1rem; border-radius: 0.5rem; border-left: 5px solid #1E3A8A; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
div[data-testid="stForm"] { background-color: #FFFFFF; border-radius: 0.5rem; padding: 1.5rem; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
</style>
''', unsafe_allow_html=True)

st.markdown('<div class="main-header">🥩 금화 수입육 통합 ERP</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">PO 단위 원가 정산 및 실시간 단가 소급 시스템</div>', unsafe_allow_html=True)

# --- 데이터베이스 초기화 ---
if 'po_master' not in st.session_state:
    st.session_state.po_master = pd.DataFrame(columns=[
        'PO번호', '원산지', '품목명(CL)', '수입량(Kg)', 
        '원료육대금(원)', '관세및제세금(원)', '포워딩제비용(원)', '보세창고하역비(원)', '내륙운송비(원)', 
        '총원가(원)', 'Kg당단가(원)'
    ])
if 'transactions' not in st.session_state:
    st.session_state.transactions = pd.DataFrame(columns=['일자', '구분', '원산지', 'PO번호', '변동중량(Kg)', '적용단가', '원가반영액', '비고'])

# --- 사이드바 (데이터 관리) ---
with st.sidebar:
    st.header("⚙️ 시스템 관리")
    st.info("클라우드 초기화 방지를 위해 퇴근 전 반드시 엑셀 백업을 진행하세요.")
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.po_master.to_excel(writer, index=False, sheet_name='PO마스터')
        st.session_state.transactions.to_excel(writer, index=False, sheet_name='수불부')
    
    st.download_button(
        label="💾 일일 장부 백업 (엑셀)", 
        data=output.getvalue(), 
        file_name=f"금화_수입육장부_{datetime.now().strftime('%Y%m%d')}.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.divider()
    st.markdown("**📂 이전 장부 복구**")
    uploaded_file = st.file_uploader("어제 다운로드한 엑셀 업로드", type=['xlsx'], label_visibility="collapsed")
    if uploaded_file is not None:
        if st.button("장부 복구 실행", use_container_width=True, type="primary"):
            st.session_state.po_master = pd.read_excel(uploaded_file, sheet_name='PO마스터')
            st.session_state.transactions = pd.read_excel(uploaded_file, sheet_name='수불부')
            st.success("복구 완료!")
            st.rerun()

# --- 메인 화면 탭 ---
tab1, tab2, tab3 = st.tabs(["📝 PO 원가 관리 (단가 소급)", "🏭 생산일보 (수불 차감)", "📊 실시간 재고 대시보드"])

with tab1:
    st.markdown("### 📝 PO 원가 세팅 및 비용 업데이트")
    st.caption("초기 통관 비용만 입력하여 가단가를 생성하고, 추후 청구서 도착 시 비용을 추가하면 과거 수불부 원가가 자동 소급 재계산됩니다.")
    
    existing_pos = ["(+) 신규 PO 등록"] + st.session_state.po_master['PO번호'].tolist()
    selected_action = st.selectbox("작업 선택", existing_pos)
    
    def_po = ""; def_org = "호주"; def_item = "95CL"; def_kg = 0.0
    def_meat = 0; def_cust = 0; def_fwd = 0; def_stor = 0; def_trans = 0
    
    is_update = False
    if selected_action != "(+) 신규 PO 등록":
        is_update = True
        row = st.session_state.po_master[st.session_state.po_master['PO번호'] == selected_action].iloc[0]
        def_po = row['PO번호']
        def_org = row['원산지']
        def_item = row['품목명(CL)']
        def_kg = float(row['수입량(Kg)'])
        def_meat = int(row['원료육대금(원)'])
        def_cust = int(row['관세및제세금(원)'])
        def_fwd = int(row['포워딩제비용(원)'])
        def_stor = int(row['보세창고하역비(원)'])
        def_trans = int(row['내륙운송비(원)'])

    with st.form("po_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**1. 기본 정보**")
            po_no = st.text_input("PO 번호", value=def_po, disabled=is_update)
            origin = st.selectbox("원산지", ["호주", "뉴질랜드", "미국", "기타"], index=["호주", "뉴질랜드", "미국", "기타"].index(def_org) if def_org in ["호주", "뉴질랜드", "미국", "기타"] else 0)
            item_name = st.selectbox("품목명 (CL)", ["95CL", "90CL", "85CL", "80CL", "75CL", "70CL", "65CL"], index=["95CL", "90CL", "85CL", "80CL", "75CL", "70CL", "65CL"].index(def_item) if def_item in ["95CL", "90CL", "85CL", "80CL", "75CL", "70CL", "65CL"] else 0)
            import_kg = st.number_input("수입 실중량(Kg)", min_value=0.0, step=100.0, value=def_kg)
        with c2:
            st.markdown("**2. 선발생 비용 (가단가 기준)**")
            meat_cost = st.number_input("원료육 대금 (LC 포함)", min_value=0, step=100000, value=def_meat)
            customs = st.number_input("관세 및 제세금", min_value=0, step=100000, value=def_cust)
            forwarding = st.number_input("포워딩 제비용", min_value=0, step=100000, value=def_fwd)
        with c3:
            st.markdown("**3. 후발생 비용 (청구서 수취 시 업데이트)**")
            storage = st.number_input("보세창고 / 하역비", min_value=0, step=100000, value=def_stor)
            transport = st.number_input("내륙 운송비", min_value=0, step=10000, value=def_trans)
            
        submit_label = "PO 원가 업데이트 및 소급 적용" if is_update else "신규 PO 원가 (가단가) 확정"
        submit_po = st.form_submit_button(submit_label, type="primary")
        
        if submit_po and po_no and import_kg > 0:
            total_cost = meat_cost + customs + forwarding + storage + transport
            unit_price = total_cost / import_kg
            
            if is_update:
                idx = st.session_state.po_master[st.session_state.po_master['PO번호'] == po_no].index
                st.session_state.po_master.loc[idx, ['원산지', '품목명(CL)', '수입량(Kg)', '원료육대금(원)', '관세및제세금(원)', '포워딩제비용(원)', '보세창고하역비(원)', '내륙운송비(원)', '총원가(원)', 'Kg당단가(원)']] = [origin, item_name, import_kg, meat_cost, customs, forwarding, storage, transport, total_cost, unit_price]
                
                tx_mask = st.session_state.transactions['PO번호'] == po_no
                st.session_state.transactions.loc[tx_mask & (st.session_state.transactions['구분'] == '최초입고'), '변동중량(Kg)'] = import_kg
                st.session_state.transactions.loc[tx_mask, '적용단가'] = unit_price
                st.session_state.transactions.loc[tx_mask, '원가반영액'] = st.session_state.transactions.loc[tx_mask, '변동중량(Kg)'] * unit_price
                
                st.success(f"✅ [{po_no}] 단가 업데이트 완료 (1Kg당 {unit_price:,.2f}원). 과거 생산일보 사용 내역이 모두 새 단가로 소급 정산되었습니다.")
            else:
                if po_no in st.session_state.po_master['PO번호'].values:
                    st.error("이미 존재하는 PO 번호입니다. 업데이트하려면 상단의 '작업 선택'에서 해당 PO를 고르세요.")
                else:
                    new_row = [po_no, origin, item_name, import_kg, meat_cost, customs, forwarding, storage, transport, total_cost, unit_price]
                    st.session_state.po_master.loc[len(st.session_state.po_master)] = new_row
                    new_tx = [datetime.today().strftime('%Y-%m-%d'), '최초입고', origin, po_no, import_kg, unit_price, total_cost, '신규 입고 (가단가)']
                    st.session_state.transactions.loc[len(st.session_state.transactions)] = new_tx
                    st.success(f"✅ [{po_no}] 등록 완료! 현재 임시 가단가: 1Kg당 {unit_price:,.2f}원")
            st.rerun()

    st.dataframe(st.session_state.po_master.style.format({col: "{:,.0f}" for col in st.session_state.po_master.columns if '(원)' in col or '(Kg)' in col}), use_container_width=True)

with tab2:
    st.markdown("### 🏭 생산일보 수불 등록")
    if st.session_state.po_master.empty:
        st.warning("먼저 1번 탭에서 PO를 등록하세요.")
    else:
        with st.form("tx_form"):
            t1, t2 = st.columns(2)
            with t1:
                tx_date = st.date_input("사용 일자")
                po_choices = st.session_state.po_master['PO번호'] + " (" + st.session_state.po_master['원산지'] + " " + st.session_state.po_master['품목명(CL)'] + ")"
                selected_po_display = st.selectbox("차감 대상 PO 선택", po_choices)
                tx_type = st.radio("구분", ["생산투입", "폐기", "판매(출고)"], horizontal=True)
            with t2:
                tx_kg = st.number_input("사용 중량(Kg) - 생산일보 수량 입력", min_value=0.0, step=10.0)
                tx_note = st.text_input("비고 (예: 버거패티 1251.2)")
                
            submit_tx = st.form_submit_button("사용 수량 차감", type="primary")
            
            if submit_tx and tx_kg > 0:
                tx_po = selected_po_display.split(" ")[0]
                origin = st.session_state.po_master.loc[st.session_state.po_master['PO번호'] == tx_po, '원산지'].values[0]
                unit_price = st.session_state.po_master.loc[st.session_state.po_master['PO번호'] == tx_po, 'Kg당단가(원)'].values[0]
                total_tx_cost = -tx_kg * unit_price
                
                new_tx = [tx_date.strftime('%Y-%m-%d'), tx_type, origin, tx_po, -tx_kg, unit_price, total_tx_cost, tx_note]
                st.session_state.transactions.loc[len(st.session_state.transactions)] = new_tx
                
                st.success(f"✅ {tx_kg}Kg 차감 완료! ({origin}산 {tx_po} / 반영 원가: {abs(total_tx_cost):,.0f}원)")
                st.rerun()

    st.dataframe(st.session_state.transactions.style.format({"변동중량(Kg)":"{:,.1f}", "적용단가":"{:,.2f}", "원가반영액":"{:,.0f}"}), use_container_width=True)

with tab3:
    st.markdown("### 📊 실시간 재고 및 자산 대시보드")
    if not st.session_state.po_master.empty:
        summary = st.session_state.transactions.groupby(['원산지', 'PO번호']).agg(
            잔여재고_Kg=('변동중량(Kg)', 'sum'),
            잔여자산_원=('원가반영액', 'sum')
        ).reset_index()
        
        dash_df = pd.merge(st.session_state.po_master[['PO번호', '품목명(CL)', 'Kg당단가(원)']], summary, on='PO번호', how='right')
        dash_df = dash_df[dash_df['잔여재고_Kg'] > 0.1] # 미미한 소수점 오차 제외
        
        col1, col2, col3 = st.columns(3)
        col1.metric("📦 창고 총 잔여 재고", f"{dash_df['잔여재고_Kg'].sum():,.1f} Kg")
        col2.metric("💰 총 재고 자산액", f"{dash_df['잔여자산_원'].sum():,.0f} 원")
        col3.metric("🧾 현재 활성 PO 개수", f"{len(dash_df)} 건")
        
        st.markdown("#### PO별 세부 현황 (잔여 재고 존재 건)")
        st.dataframe(dash_df.style.format({
            "Kg당단가(원)":"{:,.2f}",
            "잔여재고_Kg":"{:,.1f}",
            "잔여자산_원":"{:,.0f}"
        }), use_container_width=True)
    else:
        st.info("등록된 PO 데이터가 없습니다.")
