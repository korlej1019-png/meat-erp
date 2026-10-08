import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="금화 수입육 원가/수불 ERP", layout="wide")
st.title("🥩 금화 수입육 원가/수불 관리 시스템 (클라우드 웹 버전)")

# 1. 메모리(세션 상태)에 데이터베이스(데이터프레임) 초기화
if 'po_master' not in st.session_state:
    st.session_state.po_master = pd.DataFrame(columns=['PO번호', '품목명', '수입량(Kg)', '물대', 'LC수수료', '포워딩', '관세', '초기보관비', '운송비', '총원가', 'Kg당단가'])
if 'transactions' not in st.session_state:
    st.session_state.transactions = pd.DataFrame(columns=['일자', '구분', 'PO번호', '변동중량(Kg)', '적용단가', '원가반영액', '비고'])

# 2. 사이드바: 엑셀 백업 및 복구 기능 (DB 서버 완벽 대체)
with st.sidebar:
    st.header("💾 일일 마감 (엑셀 백업)")
    st.info("무료 클라우드는 가끔 서버가 초기화됩니다. 퇴근 전 아래 버튼을 눌러 오늘 장부를 엑셀로 저장하세요.")
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.po_master.to_excel(writer, index=False, sheet_name='PO마스터')
        st.session_state.transactions.to_excel(writer, index=False, sheet_name='수불부')
    
    st.download_button(
        label="📥 현재 장부 엑셀로 다운로드", 
        data=output.getvalue(), 
        file_name=f"금화_수입육장부_{datetime.now().strftime('%Y%m%d')}.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    
    st.divider()
    st.header("📂 데이터 불러오기")
    uploaded_file = st.file_uploader("어제 저장한 엑셀 파일 업로드", type=['xlsx'])
    if uploaded_file is not None:
        if st.button("장부 복구하기"):
            st.session_state.po_master = pd.read_excel(uploaded_file, sheet_name='PO마스터')
            st.session_state.transactions = pd.read_excel(uploaded_file, sheet_name='수불부')
            st.success("데이터 복구 완료!")
            st.rerun()

# 3. 메인 화면 탭 구성
tab1, tab2, tab3 = st.tabs(["1️⃣ 신규 PO 원가 세팅", "2️⃣ 수불부 (공장 투입/출고)", "3️⃣ 실시간 재고 대시보드"])

with tab1:
    st.subheader("새로운 수입육 단가 세팅")
    with st.form("po_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            po_no = st.text_input("PO 번호 (예: PO-12)")
            item_name = st.text_input("품목명 (예: 65CL)")
            import_kg = st.number_input("수입 실중량(Kg)", min_value=0.0, step=100.0)
        with col2:
            meat_cost = st.number_input("원료육 대금(원)", min_value=0.0, step=10000.0)
            lc_fee = st.number_input("LC개설수수료(원)", min_value=0.0, step=1000.0)
            forwarding = st.number_input("포워딩비용(원)", min_value=0.0, step=10000.0)
        with col3:
            customs = st.number_input("관세비용(원)", min_value=0.0, step=10000.0)
            storage = st.number_input("초기 보관비용(원)", min_value=0.0, step=10000.0)
            transport = st.number_input("운송비용(원)", min_value=0.0, step=10000.0)
            
        submit_po = st.form_submit_button("PO 원가 확정")
        
        if submit_po and po_no and import_kg > 0:
            if po_no in st.session_state.po_master['PO번호'].values:
                st.error("이미 존재하는 PO 번호입니다.")
            else:
                total_cost = meat_cost + lc_fee + forwarding + customs + storage + transport
                unit_price = total_cost / import_kg
                
                new_po = pd.DataFrame([[po_no, item_name, import_kg, meat_cost, lc_fee, forwarding, customs, storage, transport, total_cost, unit_price]], 
                                      columns=st.session_state.po_master.columns)
                st.session_state.po_master = pd.concat([st.session_state.po_master, new_po], ignore_index=True)
                
                new_tx = pd.DataFrame([[datetime.today().strftime('%Y-%m-%d'), '최초입고', po_no, import_kg, unit_price, total_cost, '신규 PO 등록']],
                                      columns=st.session_state.transactions.columns)
                st.session_state.transactions = pd.concat([st.session_state.transactions, new_tx], ignore_index=True)
                
                st.success(f"[{po_no}] 등록 완료! 1Kg당 매입단가: {unit_price:,.2f}원")
                st.rerun()

    st.dataframe(st.session_state.po_master.style.format({"수입량(Kg)":"{:,.1f}", "물대":"{:,.0f}", "총원가":"{:,.0f}", "Kg당단가":"{:,.2f}"}), use_container_width=True)

with tab2:
    st.subheader("창고 고기 사용 및 판매 등록")
    po_list = st.session_state.po_master['PO번호'].tolist()
    
    if not po_list:
        st.warning("먼저 1번 탭에서 PO를 등록해주세요.")
    else:
        with st.form("tx_form"):
            t_col1, t_col2 = st.columns(2)
            with t_col1:
                tx_date = st.date_input("일자")
                tx_po = st.selectbox("대상 PO 번호", po_list)
                tx_type = st.radio("구분", ["공장투입(사용)", "외부판매(도매)"])
            with t_col2:
                tx_kg = st.number_input("사용 중량(Kg) - 양수로 입력", min_value=0.0, step=10.0)
                tx_note = st.text_input("비고 (예: 버거킹 패티 라인 투입)")
                
            submit_tx = st.form_submit_button("수불부 기록")
            
            if submit_tx and tx_kg > 0:
                unit_price = st.session_state.po_master.loc[st.session_state.po_master['PO번호'] == tx_po, 'Kg당단가'].values[0]
                total_tx_cost = tx_kg * unit_price
                
                new_tx = pd.DataFrame([[tx_date.strftime('%Y-%m-%d'), tx_type, tx_po, -tx_kg, unit_price, -total_tx_cost, tx_note]],
                                      columns=st.session_state.transactions.columns)
                st.session_state.transactions = pd.concat([st.session_state.transactions, new_tx], ignore_index=True)
                
                st.success(f"{tx_kg}Kg 출고 완료! (원가 {total_tx_cost:,.0f}원 반영)")
                st.rerun()

    st.dataframe(st.session_state.transactions.style.format({"변동중량(Kg)":"{:,.1f}", "적용단가":"{:,.2f}", "원가반영액":"{:,.0f}"}), use_container_width=True)

with tab3:
    st.subheader("현재 잔여 재고 및 자산 평가액")
    if not st.session_state.po_master.empty and not st.session_state.transactions.empty:
        summary = st.session_state.transactions.groupby('PO번호').agg(
            잔여재고_Kg=('변동중량(Kg)', 'sum'),
            잔여자산_원=('원가반영액', 'sum')
        ).reset_index()
        
        dash_df = pd.merge(st.session_state.po_master[['PO번호', '품목명', '수입량(Kg)', 'Kg당단가']], summary, on='PO번호', how='left')
        
        total_asset = dash_df['잔여자산_원'].sum()
        st.metric(label="총 창고 잔여 재고자산액", value=f"{total_asset:,.0f} 원")
        
        st.dataframe(dash_df.style.format({
            "수입량(Kg)":"{:,.1f}", 
            "Kg당단가":"{:,.2f}",
            "잔여재고_Kg":"{:,.1f}",
            "잔여자산_원":"{:,.0f}"
        }), use_container_width=True)
    else:
        st.info("데이터가 없습니다.")
