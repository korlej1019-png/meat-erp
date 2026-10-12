import streamlit as st
import pandas as pd
from datetime import datetime
import io
import json

st.set_page_config(page_title="금화 수입육 통합 ERP", page_icon="🏢", layout="wide", initial_sidebar_state="expanded")

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

st.markdown('''
<style>
body { font-family: 'Malgun Gothic', sans-serif; background-color: #f3f4f6; }
.erp-header { background-color: #1c3d5a; color: white; padding: 15px 20px; font-size: 1.5rem; font-weight: bold; border-radius: 5px; margin-bottom: 20px; }
.section-title { font-size: 1.1rem; font-weight: 600; color: #333; border-bottom: 2px solid #1c3d5a; padding-bottom: 8px; margin-bottom: 15px; margin-top: 20px; }
.stMetric { background-color: #fff; padding: 15px; border-radius: 4px; border: 1px solid #ddd; border-top: 3px solid #1c3d5a; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
div[data-testid="stForm"] { background-color: #fff; border: 1px solid #ddd; padding: 20px; border-radius: 4px; }
</style>
''', unsafe_allow_html=True)

st.markdown('<div class="erp-header">🏢 금화 수입육 관리 시스템 (표준 인수인계 버전)</div>', unsafe_allow_html=True)

if not GSPREAD_AVAILABLE:
    st.error("🚨 필수 라이브러리가 누락되었습니다! GitHub의 `requirements.txt` 파일에 `gspread` 와 `google-auth` 를 추가해 주세요.")
    st.stop()

def init_session_state():
    if 'po_master' not in st.session_state or not isinstance(st.session_state.po_master, pd.DataFrame):
        st.session_state.po_master = pd.DataFrame(columns=[
            '계약일자', 'PO번호', '원산지', '품목명(CL)', '계약환율', '외화금액(USD)', 
            '수입량(Kg)', '통관일자', '원료육대금(원)', '적하보험료(원)', '포워딩비용(원)', '관세비용(원)', '보세창고비용(원)', 
            '제비용(원)', '총원가(원)', 'Kg당단가(원)'
        ])
    if 'transactions' not in st.session_state or not isinstance(st.session_state.transactions, pd.DataFrame):
        st.session_state.transactions = pd.DataFrame(columns=[
            '출고일자', '구분', '원산지', 'PO번호', '품목명(CL)', '출고수량(Kg)', '적용단가', '출고금액(원)', '비고'
        ])

init_session_state()

def enforce_columns():
    init_session_state()
    expected_po = ['계약일자', 'PO번호', '원산지', '품목명(CL)', '계약환율', '외화금액(USD)', '수입량(Kg)', '통관일자', '원료육대금(원)', '적하보험료(원)', '포워딩비용(원)', '관세비용(원)', '보세창고비용(원)', '제비용(원)', '총원가(원)', 'Kg당단가(원)']
    for col in expected_po:
        if col not in st.session_state.po_master.columns:
            st.session_state.po_master[col] = None
            
    expected_tx = ['출고일자', '구분', '원산지', 'PO번호', '품목명(CL)', '출고수량(Kg)', '적용단가', '출고금액(원)', '비고']
    for col in expected_tx:
        if col not in st.session_state.transactions.columns:
            st.session_state.transactions[col] = None

@st.cache_resource
def get_gs_client():
    try:
        creds_json = st.secrets["google_credentials"]
        creds_dict = json.loads(creds_json)
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        return None

def sync_from_gs():
    client = get_gs_client()
    if not client: return
    try:
        sheet = client.open("금화_수입육장부")
        try:
            ws_po = sheet.worksheet("PO마스터")
            data_po = ws_po.get_all_records()
            if data_po:
                st.session_state.po_master = pd.DataFrame(data_po)
        except gspread.exceptions.WorksheetNotFound:
            sheet.add_worksheet(title="PO마스터", rows="1000", cols="20")
            
        try:
            ws_tx = sheet.worksheet("수불부")
            data_tx = ws_tx.get_all_records()
            if data_tx:
                st.session_state.transactions = pd.DataFrame(data_tx)
        except gspread.exceptions.WorksheetNotFound:
            sheet.add_worksheet(title="수불부", rows="10000", cols="20")
            
        enforce_columns()
        fix_dtypes()
    except Exception as e:
        pass

def sync_to_gs():
    client = get_gs_client()
    if not client: return
    try:
        sheet = client.open("금화_수입육장부")
        enforce_columns()
        
        try:
            ws_po = sheet.worksheet("PO마스터")
        except gspread.exceptions.WorksheetNotFound:
            ws_po = sheet.add_worksheet(title="PO마스터", rows="1000", cols="20")
        ws_po.clear()
        po_data = [st.session_state.po_master.columns.values.tolist()] + st.session_state.po_master.fillna("").values.tolist()
        ws_po.update(values=po_data, range_name='A1')
        
        try:
            ws_tx = sheet.worksheet("수불부")
        except gspread.exceptions.WorksheetNotFound:
            ws_tx = sheet.add_worksheet(title="수불부", rows="10000", cols="20")
        ws_tx.clear()
        tx_data = [st.session_state.transactions.columns.values.tolist()] + st.session_state.transactions.fillna("").values.tolist()
        ws_tx.update(values=tx_data, range_name='A1')
    except Exception as e:
        st.error(f"구글 시트 저장 실패: {e}")

def fix_dtypes():
    init_session_state()
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

if 'db_loaded' not in st.session_state:
    st.session_state.db_loaded = True
    sync_from_gs()

with st.sidebar:
    st.markdown("### 📋 업무 메뉴얼 (인수인계용)")
    menu = st.radio("작업 영역 선택", [
        "0. [초기세팅] 시스템 데이터 일괄 마이그레이션",
        "1. [1단계] 수입 고기 및 제비용 등록", 
        "2. [2단계] 공장 생산 출고일보 입력", 
        "3. [3단계] 월말 결산 수불명세서 출력",
        "4. [경영보고] 전사 재고 현황 요약",
        "5. [검증] 실무진 엑셀 오차 검사하기"
    ])
    
    st.divider()
    st.markdown("### ☁️ 클라우드 동기화")
    if st.button("🔄 최신 장부 불러오기", use_container_width=True):
        sync_from_gs()
        st.success("최신 데이터 동기화 완료!")
        st.rerun()

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
        enforce_columns()
        fix_dtypes()
        sync_to_gs()
        return True, len(po_master_list), len(tx_list)
    except Exception as e:
        return False, str(e), 0

if menu.startswith("0."):
    st.markdown('<div class="section-title">■ 시스템 초기화 > 원본 엑셀(생산정리.xlsx) 데이터 일괄 마이그레이션</div>', unsafe_allow_html=True)
    st.info("💡 본 메뉴는 시스템 최초 도입 시 기존 엑셀 장부 데이터를 클라우드 데이터베이스로 일괄 등록하는 전용 페이지입니다. 일반적인 일일 업무 시에는 사용하지 않습니다.")
    
    raw_file = st.file_uploader("기준 '생산정리.xlsx' 파일 업로드", type=['xlsx'])
    if raw_file is not None:
        if st.button("🚀 데이터 일괄 변환 및 클라우드 업로드 실행", type="primary", use_container_width=True):
            with st.spinner('기존 엑셀 데이터를 표준 ERP 양식으로 변환 중입니다...'):
                success, po_cnt, tx_cnt = auto_migrate_raw_excel(raw_file)
                if success: st.success(f"🎉 데이터 일괄 마이그레이션 완료! (총 {po_cnt}개 PO, {tx_cnt}건의 수불 내역 등록 완료)")
                else: st.error(f"변환 중 오류 발생: {po_cnt}")

elif menu.startswith("1."):
    st.markdown('<div class="section-title">■ [1단계] 수입 고기 및 제비용 등록</div>', unsafe_allow_html=True)
    st.info("💡 해외에서 수입한 고기의 계약 정보와 통관 시 발생한 제비용(원료육, 보험, 포워딩, 관세, 보세창고)을 입력하여 1Kg당 원가 단가를 확정하는 단계입니다.")
    
    enforce_columns()
    po_list = []
    if not st.session_state.po_master.empty and 'PO번호' in st.session_state.po_master.columns:
        po_list = st.session_state.po_master['PO번호'].dropna().astype(str).tolist()
    existing_pos = ["(+) 신규 PO 등록"] + [p for p in po_list if p.strip() != ""]
        
    selected_action = st.selectbox("작업 구분 선택 (신규 등록 또는 기존 내역 수정)", existing_pos)
    
    d_ct_date = datetime.today(); d_po = ""; d_org = "호주"; d_item = "95CL"; d_rate = 1350.0; d_usd = 0.0
    d_kg = 0.0; d_cc_date = datetime.today(); d_meat = 0; d_ins = 0; d_fwd = 0; d_cust = 0; d_stor = 0
    
    is_update = False
    if selected_action != "(+) 신규 PO 등록":
        is_update = True
        row = st.session_state.po_master[st.session_state.po_master['PO번호'] == selected_action].iloc[0]
        try: d_ct_date = pd.to_datetime(row['계약일자'])
        except: pass
        d_po = row['PO번호']; d_org = row['원산지']; d_item = row['품목명(CL)']; d_rate = float(row['계약환율']); d_usd = float(row['외화금액(USD)']); d_kg = float(row['수입량(Kg)'])
        try: d_cc_date = pd.to_datetime(row['통관일자'])
        except: pass
        d_meat = int(row['원료육대금(원)']); d_ins = int(row['적하보험료(원)']); d_fwd = int(row['포워딩비용(원)']); d_cust = int(row['관세비용(원)']); d_stor = int(row['보세창고비용(원)'])

    with st.form("po_form"):
        st.markdown("**[1] 계약 및 외환 정보**")
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
        
        st.markdown("**[2] 원가 및 제비용 청구 내역**")
        c9, c10, c11, c12, c13 = st.columns(5)
        with c9: meat_cost = st.number_input("원료육 대금 (원)", value=d_meat, step=100000)
        with c10: insurance = st.number_input("적하보험료 (원)", value=d_ins, step=10000)
        with c11: forwarding = st.number_input("포워딩 비용 (원)", value=d_fwd, step=10000)
        with c12: customs = st.number_input("관세비용 (원)", value=d_cust, step=10000)
        with c13: storage = st.number_input("보세창고 비용 (원)", value=d_stor, step=10000)
        
        submit_po = st.form_submit_button("저장 및 원가 단가 확정", type="primary")
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
                sync_to_gs()
                st.success(f"[{po_no}] 단가 업데이트 완료 및 클라우드 저장 완료.")
            else:
                if not st.session_state.po_master.empty and po_no in st.session_state.po_master['PO번호'].values: st.error("이미 등록된 PO 번호입니다.")
                else:
                    new_row = pd.DataFrame([[ct_date.strftime('%Y-%m-%d'), po_no, origin, item_name, ex_rate, usd_amt, import_kg, cc_date.strftime('%Y-%m-%d'), meat_cost, insurance, forwarding, customs, storage, total_expenses, total_cost, unit_price]], columns=st.session_state.po_master.columns)
                    st.session_state.po_master = pd.concat([st.session_state.po_master, new_row], ignore_index=True)
                    fix_dtypes()
                    sync_to_gs()
                    st.success(f"[{po_no}] 신규 등록 완료 및 클라우드 저장 완료.")
            
    display_cols = [c for c in st.session_state.po_master.columns if '(원)' in c or '환율' in c or '(Kg)' in c or 'USD' in c]
    if not st.session_state.po_master.empty and all(c in st.session_state.po_master.columns for c in display_cols):
        st.markdown("##### 📌 등록된 수입 원가 마스터 내역")
        st.dataframe(st.session_state.po_master.style.format({col: "{:,.0f}" for col in display_cols}), hide_index=True, use_container_width=True)

elif menu.startswith("2."):
    st.markdown('<div class="section-title">■ [2단계] 공장 생산 출고일보 입력</div>', unsafe_allow_html=True)
    st.info("💡 공장에서 생산 투입을 위해 고기를 불출하거나 타계정으로 대체(폐기/판매 등)한 내역을 일자별로 등록하여 재고를 차감하는 단계입니다.")
    
    enforce_columns()
    valid_pos = []
    if 'PO번호' in st.session_state.po_master.columns and '품목명(CL)' in st.session_state.po_master.columns and not st.session_state.po_master.empty:
        valid_pos = (st.session_state.po_master['PO번호'].astype(str) + " (" + st.session_state.po_master['품목명(CL)'].astype(str) + ")").tolist()
        
    if not valid_pos:
        st.warning("등록된 수입 PO가 없습니다. 먼저 [1단계] 메뉴에서 PO를 등록해 주세요.")
    else:
        with st.form("tx_form"):
            c1, c2, c3, c4 = st.columns(4)
            with c1: tx_date = st.date_input("출고일자")
            with c2: tx_type = st.selectbox("출고 구분", ["당월불출(생산투입)", "타계정대체(폐기/판매등)"])
            with c3: selected_po_display = st.selectbox("대상 PO 선택", valid_pos)
            with c4: tx_kg = st.number_input("출고 수량 (Kg)", min_value=0.0, step=10.0)
            tx_note = st.text_input("적요/비고 (예: 버거패티 투입분)")
            if st.form_submit_button("출고 전표 발행 및 저장", type="primary") and tx_kg > 0:
                tx_po = selected_po_display.split(" ")[0]
                po_info = st.session_state.po_master[st.session_state.po_master['PO번호'] == tx_po].iloc[0]
                unit_price = po_info['Kg당단가(원)']
                new_tx = pd.DataFrame([[tx_date.strftime('%Y-%m-%d'), tx_type, po_info['원산지'], tx_po, po_info['품목명(CL)'], tx_kg, unit_price, tx_kg * unit_price, tx_note]], columns=st.session_state.transactions.columns)
                st.session_state.transactions = pd.concat([st.session_state.transactions, new_tx], ignore_index=True)
                fix_dtypes()
                sync_to_gs()
                st.success("출고 전표가 정상 발행되었으며, 클라우드에 실시간 반영되었습니다.")
                
        if not st.session_state.transactions.empty:
            st.markdown("##### 📌 당월 등록된 출고 전표 내역")
            st.dataframe(st.session_state.transactions, hide_index=True, use_container_width=True, column_config={"출고수량(Kg)": st.column_config.NumberColumn(format="%,.1f"), "적용단가": st.column_config.NumberColumn(format="%,.2f"), "출고금액(원)": st.column_config.NumberColumn(format="%,.0f")})

elif menu.startswith("3."):
    st.markdown('<div class="section-title">■ [3단계] 월말 결산 수불명세서 출력</div>', unsafe_allow_html=True)
    st.info("💡 품목(CL)별로 매입, 당월불출, 타계정대체, 월말재고 수량 및 금액을 더존 회계 양식으로 조회하고 결산에 활용하는 단계입니다.")
    
    enforce_columns()
    if st.session_state.po_master.empty or '품목명(CL)' not in st.session_state.po_master.columns: 
        st.info("조회할 데이터가 없습니다.")
    else:
        cl_list = sorted(list(set(st.session_state.po_master['품목명(CL)'].dropna().tolist())))
        selected_cl = st.selectbox("조회할 품목(CL) 선택", cl_list)
        po_filtered = st.session_state.po_master[st.session_state.po_master['품목명(CL)'] == selected_cl]
        tx_filtered = st.session_state.transactions[st.session_state.transactions['품목명(CL)'] == selected_cl] if not st.session_state.transactions.empty else pd.DataFrame(columns=st.session_state.transactions.columns)
        
        report_data = []
        for index, po in po_filtered.iterrows():
            po_no = po['PO번호']; in_kg = po['수입량(Kg)']; in_price = po['Kg당단가(원)']; in_amt = po['총원가(원)']
            out_prod = tx_filtered[(tx_filtered['PO번호'] == po_no) & (tx_filtered['구분'] == '당월불출(생산투입)')] if not tx_filtered.empty else pd.DataFrame()
            prod_kg = out_prod['출고수량(Kg)'].sum() if not out_prod.empty else 0
            prod_amt = out_prod['출고금액(원)'].sum() if not out_prod.empty else 0
            
            out_etc = tx_filtered[(tx_filtered['PO번호'] == po_no) & (tx_filtered['구분'] == '타계정대체(폐기/판매등)')] if not tx_filtered.empty else pd.DataFrame()
            etc_kg = out_etc['출고수량(Kg)'].sum() if not out_etc.empty else 0
            etc_amt = out_etc['출고금액(원)'].sum() if not out_etc.empty else 0
            
            rem_kg = in_kg - prod_kg - etc_kg; rem_amt = in_amt - prod_amt - etc_amt
            report_data.append({
                "PO No": po_no, "원산지": po['원산지'],
                "[매입] 수량": in_kg, "[매입] 단가": in_price, "[매입] 금액": in_amt,
                "[당월불출] 수량": prod_kg, "[당월불출] 금액": prod_amt,
                "[타계정대체] 수량": etc_kg, "[타계정대체] 금액": etc_amt,
                "[당월재고] 수량": rem_kg, "[당월재고] 금액": rem_amt
            })
            
        def highlight_columns(s):
            if '매입' in s.name: return ['background-color: #e6f2ff'] * len(s)
            elif '불출' in s.name: return ['background-color: #fff0f0'] * len(s)
            elif '대체' in s.name: return ['background-color: #fff9e6'] * len(s)
            elif '재고' in s.name: return ['background-color: #e6ffe6; font-weight: bold'] * len(s)
            return [''] * len(s)
        
        if report_data:
            st.markdown(f"#### 📄 소고기트리밍 {selected_cl} 수불명세서")
            st.dataframe(pd.DataFrame(report_data).style.apply(highlight_columns).format({col: "{:,.0f}" if '금액' in col else "{:,.1f}" for col in pd.DataFrame(report_data).columns if '수량' in col or '단가' in col or '금액' in col}), hide_index=True, use_container_width=True)

elif menu.startswith("4."):
    st.markdown('<div class="section-title">■ 경영보고 > 전사 재고 현황 요약</div>', unsafe_allow_html=True)
    st.info("💡 원산지 및 품목별 잔여 재고 현황을 한눈에 파악할 수 있는 경영진 보고용 화면입니다.")
    enforce_columns()
    if not st.session_state.po_master.empty:
        summary = st.session_state.transactions.groupby('PO번호').agg(출고합_Kg=('출고수량(Kg)', 'sum'), 출고합_원=('출고금액(원)', 'sum')).reset_index() if not st.session_state.transactions.empty else pd.DataFrame(columns=['PO번호', '출고합_Kg', '출고합_원'])
        dash_df = pd.merge(st.session_state.po_master[['PO번호', '원산지', '품목명(CL)', '수입량(Kg)', '총원가(원)']], summary, on='PO번호', how='left').fillna(0)
        dash_df['잔여재고(Kg)'] = dash_df['수입량(Kg)'] - dash_df['출고합_Kg']
        dash_df['잔여자산(원)'] = dash_df['총원가(원)'] - dash_df['출고합_원']
        dash_df = dash_df[dash_df['잔여재고(Kg)'] > 0.1].drop(columns=['수입량(Kg)', '총원가(원)', '출고합_Kg', '출고합_원'])
        
        st.markdown("##### 📌 원산지/품목별 상세 잔여 현황")
        st.dataframe(dash_df.style.format({"잔여재고(Kg)":"{:,.1f}", "잔여자산(원)":"{:,.0f}"}), hide_index=True, use_container_width=True)

elif menu.startswith("5."):
    st.markdown('<div class="section-title">■ [검증] 실무진 엑셀 오차 검사하기</div>', unsafe_allow_html=True)
    st.info("💡 실무진이 작성한 '생산일보 엑셀' 파일을 업로드하면, 시스템 장부와 1:1로 자동 대조하여 오차가 있는 항목을 빨간색으로 즉시 찾아냅니다.")
    
    check_file = st.file_uploader("검증할 현장 실무진 엑셀 업로드", type=['xlsx'])
    if check_file is not None:
        if st.button("🔍 크로스체크(대사) 실행", type="primary"):
            with st.spinner("엑셀 데이터와 클라우드 장부를 대조 중입니다..."):
                try:
                    xls = pd.ExcelFile(check_file)
                    field_data = []
                    for sheet in [s for s in xls.sheet_names if s.endswith('CL')]:
                        df_field = pd.read_excel(xls, sheet_name=sheet, header=None)
                        for row_idx in range(3, df_field.shape[0]):
                            po_val = df_field.iloc[row_idx, 2]
                            if pd.notna(po_val) and str(po_val).strip() not in ['', 'nan']:
                                po_str = str(po_val).strip()
                                if 'PO' not in po_str.upper(): po_str = f"PO-{po_str}"
                                out_kg = pd.to_numeric(df_field.iloc[row_idx, 12], errors='coerce')
                                if pd.notna(out_kg) and out_kg > 0:
                                    field_data.append({'품목명(CL)': sheet, 'PO번호': po_str, '현장보고_출고량(Kg)': out_kg})
                    
                    df_field_summary = pd.DataFrame(field_data).groupby(['품목명(CL)', 'PO번호'])['현장보고_출고량(Kg)'].sum().reset_index()
                    tx = st.session_state.transactions
                    erp_prod = tx[tx['구분'] == '당월불출(생산투입)'] if not tx.empty else pd.DataFrame(columns=['품목명(CL)', 'PO번호', '출고수량(Kg)'])
                    df_erp_summary = erp_prod.groupby(['품목명(CL)', 'PO번호'])['출고수량(Kg)'].sum().reset_index().rename(columns={'출고수량(Kg)': 'ERP장부_출고량(Kg)'})
                    
                    compare_df = pd.merge(df_erp_summary, df_field_summary, on=['품목명(CL)', 'PO번호'], how='outer').fillna(0)
                    compare_df['오차(Kg)'] = compare_df['ERP장부_출고량(Kg)'] - compare_df['현장보고_출고량(Kg)']
                    
                    def highlight_diff(val): return 'background-color: #ffcccc' if abs(val) > 0.1 else ''
                    
                    st.markdown("#### ⚖️ 크로스체크 결과 대사표")
                    if compare_df.empty:
                        st.warning("비교할 데이터가 없습니다.")
                    else:
                        st.dataframe(compare_df.style.applymap(highlight_diff, subset=['오차(Kg)']).format({"ERP장부_출고량(Kg)":"{:,.1f}", "현장보고_출고량(Kg)":"{:,.1f}", "오차(Kg)":"{:,.1f}"}), hide_index=True, use_container_width=True)
                        errors = compare_df[abs(compare_df['오차(Kg)']) > 0.1]
                        if not errors.empty: st.error(f"🚨 총 {len(errors)}개의 PO에서 장부와 현장 보고 수량 간의 불일치가 발견되었습니다! 빨간색으로 표시된 오차를 확인하세요.")
                        else: st.success("✅ 완벽합니다! 시스템 장부의 모든 출고량 데이터가 현장 실무진 엑셀 데이터와 100% 일치합니다.")
                except Exception as e:
                    st.error(f"엑셀 분석 중 오류 발생: {e}")
