import streamlit as st
import pandas as pd
from datetime import datetime
import io
import json

st.set_page_config(page_title="금화 수입육 통합 ERP (구글 DB 연동)", page_icon="☁️", layout="wide", initial_sidebar_state="expanded")

# --- 에러 방지용 라이브러리 체크 ---
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

st.markdown('<div class="erp-header">☁️ 금화 수입육 관리 시스템 (구글 클라우드 DB V9)</div>', unsafe_allow_html=True)

if not GSPREAD_AVAILABLE:
    st.error("🚨 필수 라이브러리가 누락되었습니다! GitHub의 `requirements.txt` 파일에 `gspread` 와 `google-auth` 를 추가해 주세요.")
    st.stop()

# --- 구글 시트 연동 함수 ---
@st.cache_resource
def get_gs_client():
    try:
        creds_json = st.secrets["google_credentials"]
        creds_dict = json.loads(creds_json)
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"구글 인증 실패: {e}")
        return None

def sync_from_gs():
    client = get_gs_client()
    if not client: return
    try:
        sheet = client.open("금화_수입육장부")
        try:
            ws_po = sheet.worksheet("PO마스터")
            st.session_state.po_master = pd.DataFrame(ws_po.get_all_records())
        except gspread.exceptions.WorksheetNotFound:
            sheet.add_worksheet(title="PO마스터", rows="1000", cols="20")
            st.session_state.po_master = pd.DataFrame(columns=[
                '계약일자', 'PO번호', '원산지', '품목명(CL)', '계약환율', '외화금액(USD)', 
                '수입량(Kg)', '통관일자', '원료육대금(원)', '적하보험료(원)', '포워딩비용(원)', '관세비용(원)', '보세창고비용(원)', 
                '제비용(원)', '총원가(원)', 'Kg당단가(원)'
            ])
            
        try:
            ws_tx = sheet.worksheet("수불부")
            st.session_state.transactions = pd.DataFrame(ws_tx.get_all_records())
        except gspread.exceptions.WorksheetNotFound:
            sheet.add_worksheet(title="수불부", rows="10000", cols="20")
            st.session_state.transactions = pd.DataFrame(columns=[
                '출고일자', '구분', '원산지', 'PO번호', '품목명(CL)', '출고수량(Kg)', '적용단가', '출고금액(원)', '비고'
            ])
        fix_dtypes()
    except Exception as e:
        st.error(f"데이터베이스 로드 중 오류: {e}")

def sync_to_gs():
    client = get_gs_client()
    if not client: return
    try:
        sheet = client.open("금화_수입육장부")
        
        ws_po = sheet.worksheet("PO마스터")
        ws_po.clear()
        po_data = [st.session_state.po_master.columns.values.tolist()] + st.session_state.po_master.fillna("").values.tolist()
        ws_po.update(values=po_data, range_name='A1')
        
        ws_tx = sheet.worksheet("수불부")
        ws_tx.clear()
        tx_data = [st.session_state.transactions.columns.values.tolist()] + st.session_state.transactions.fillna("").values.tolist()
        ws_tx.update(values=tx_data, range_name='A1')
    except Exception as e:
        st.error(f"구글 시트 저장 실패 (API 한도 초과 등): {e}")

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

# --- 앱 실행 시 무조건 구글 시트에서 최신 데이터 로드 ---
if 'db_loaded' not in st.session_state:
    st.session_state.db_loaded = True
    sync_from_gs()

with st.sidebar:
    st.markdown("### 📋 ERP 메뉴")
    menu = st.radio("작업 영역 선택", [
        "1. [수입/구매] PO 등록 및 단가 확정", 
        "2. [생산/재고] 공장 생산일보 수불 등록", 
        "3. [회계/결산] 품목별 수불명세서 조회",
        "4. [경영/총괄] 재고 자산 대시보드",
        "5. [검증/대사] 현장 엑셀 vs ERP 크로스체크"
    ])
    
    st.divider()
    st.markdown("### ☁️ 클라우드 동기화")
    if st.button("🔄 최신 데이터 불러오기", use_container_width=True):
        sync_from_gs()
        st.success("구글 시트와 동기화 완료!")
        st.rerun()

if menu.startswith("1"):
    st.markdown('<div class="section-title">■ 수입/구매 관리 > 수입육 PO 원가 및 제비용 등록</div>', unsafe_allow_html=True)
    
    existing_pos = ["(+) 신규 PO 등록"] + st.session_state.po_master['PO번호'].astype(str).tolist() if not st.session_state.po_master.empty else ["(+) 신규 PO 등록"]
    selected_action = st.selectbox("진행 상태 선택", existing_pos)
    
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
        
        submit_po = st.form_submit_button("클라우드 저장 및 원가 단가 확정", type="primary")
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
                sync_to_gs() # DB에 즉시 쓰기
                st.success(f"[{po_no}] 단가 업데이트 완료 (1Kg당 {unit_price:,.2f}원). 구글 시트에 안전하게 기록되었습니다.")
            else:
                if not st.session_state.po_master.empty and po_no in st.session_state.po_master['PO번호'].values: st.error("이미 등록된 PO입니다.")
                else:
                    st.session_state.po_master.loc[len(st.session_state.po_master)] = [ct_date.strftime('%Y-%m-%d'), po_no, origin, item_name, ex_rate, usd_amt, import_kg, cc_date.strftime('%Y-%m-%d'), meat_cost, insurance, forwarding, customs, storage, total_expenses, total_cost, unit_price]
                    fix_dtypes()
                    sync_to_gs() # DB에 즉시 쓰기
                    st.success(f"[{po_no}] 신규 등록 완료. 구글 시트에 안전하게 기록되었습니다.")
            
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
            if st.form_submit_button("클라우드 전표 발행", type="primary") and tx_kg > 0:
                tx_po = selected_po_display.split(" ")[0]
                po_info = st.session_state.po_master[st.session_state.po_master['PO번호'] == tx_po].iloc[0]
                unit_price = po_info['Kg당단가(원)']
                st.session_state.transactions.loc[len(st.session_state.transactions)] = [tx_date.strftime('%Y-%m-%d'), tx_type, po_info['원산지'], tx_po, po_info['품목명(CL)'], tx_kg, unit_price, tx_kg * unit_price, tx_note]
                fix_dtypes()
                sync_to_gs() # DB에 즉시 쓰기
                st.success("정상적으로 출고 처리되었으며, 구글 시트에 안전하게 기록되었습니다.")
                
        st.dataframe(st.session_state.transactions, hide_index=True, use_container_width=True, column_config={"출고수량(Kg)": st.column_config.NumberColumn(format="%,.1f"), "적용단가": st.column_config.NumberColumn(format="%,.2f"), "출고금액(원)": st.column_config.NumberColumn(format="%,.0f")})

elif menu.startswith("3"):
    st.markdown('<div class="section-title">■ 회계/결산 관리 > 품목별 수불명세서 (더존 양식)</div>', unsafe_allow_html=True)
    if st.session_state.po_master.empty: st.info("데이터가 없습니다.")
    else:
        cl_list = sorted(list(set(st.session_state.po_master['품목명(CL)'].tolist())))
        selected_cl = st.selectbox("조회할 품목(CL) 선택", cl_list)
        po_filtered = st.session_state.po_master[st.session_state.po_master['품목명(CL)'] == selected_cl]
        tx_filtered = st.session_state.transactions[st.session_state.transactions['품목명(CL)'] == selected_cl]
        
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
            
        def highlight_columns(s):
            if '매입' in s.name: return ['background-color: #e6f2ff'] * len(s)
            elif '불출' in s.name: return ['background-color: #fff0f0'] * len(s)
            elif '대체' in s.name: return ['background-color: #fff9e6'] * len(s)
            elif '재고' in s.name: return ['background-color: #e6ffe6; font-weight: bold'] * len(s)
            return [''] * len(s)
        
        st.dataframe(pd.DataFrame(report_data).style.apply(highlight_columns).format({col: "{:,.0f}" if '금액' in col else "{:,.1f}" for col in pd.DataFrame(report_data).columns if '수량' in col or '단가' in col or '금액' in col}), hide_index=True, use_container_width=True)

elif menu.startswith("4"):
    st.markdown('<div class="section-title">■ 경영/총괄 관리 > 전사 재고 자산 현황</div>', unsafe_allow_html=True)
    if not st.session_state.po_master.empty:
        total_in_kg = st.session_state.po_master['수입량(Kg)'].sum(); total_out_kg = st.session_state.transactions['출고수량(Kg)'].sum()
        total_in_amt = st.session_state.po_master['총원가(원)'].sum(); total_out_amt = st.session_state.transactions['출고금액(원)'].sum()
        c1, c2, c3 = st.columns(3)
        c1.metric("📦 창고 실재고 총합", f"{total_in_kg - total_out_kg:,.1f} Kg")
        c2.metric("💰 현재고 자산총액", f"{total_in_amt - total_out_amt:,.0f} 원")
        c3.metric("📈 전체 출고(사용)량", f"{total_out_kg:,.1f} Kg")

elif menu.startswith("5"):
    st.markdown('<div class="section-title">■ 검증/대사 > 현장 엑셀 vs ERP 수불부 크로스체크</div>', unsafe_allow_html=True)
    st.info("💡 실무진이 보내온 '생산일보 엑셀'을 업로드하면, 현재 ERP 클라우드에 기록된 사용량과 1:1로 비교하여 오차가 발생한 PO를 즉시 찾아냅니다.")
    
    check_file = st.file_uploader("현장 실무진 엑셀 업로드", type=['xlsx'])
    if check_file is not None:
        if st.button("🔍 데이터 대사(Cross-check) 실행", type="primary"):
            with st.spinner("엑셀 데이터와 클라우드 DB를 대조 중입니다..."):
                try:
                    # 현장 엑셀 읽기 (가정: 현장 엑셀에 'PO번호'와 '사용중량' 컬럼이 있거나, 탭별로 관리)
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
                    
                    # ERP DB 요약
                    tx = st.session_state.transactions
                    erp_prod = tx[tx['구분'] == '당월불출(생산투입)']
                    df_erp_summary = erp_prod.groupby(['품목명(CL)', 'PO번호'])['출고수량(Kg)'].sum().reset_index().rename(columns={'출고수량(Kg)': 'ERP장부_출고량(Kg)'})
                    
                    # 두 데이터 병합 및 차이 계산
                    compare_df = pd.merge(df_erp_summary, df_field_summary, on=['품목명(CL)', 'PO번호'], how='outer').fillna(0)
                    compare_df['오차(Kg)'] = compare_df['ERP장부_출고량(Kg)'] - compare_df['현장보고_출고량(Kg)']
                    
                    def highlight_diff(val):
                        color = 'red' if abs(val) > 0.1 else 'transparent'
                        return f'background-color: {color}'
                    
                    st.markdown("#### ⚖️ 크로스체크 결과")
                    if compare_df.empty:
                        st.warning("비교할 데이터가 없습니다.")
                    else:
                        st.dataframe(compare_df.style.applymap(highlight_diff, subset=['오차(Kg)']).format({"ERP장부_출고량(Kg)":"{:,.1f}", "현장보고_출고량(Kg)":"{:,.1f}", "오차(Kg)":"{:,.1f}"}), hide_index=True, use_container_width=True)
                        
                        errors = compare_df[abs(compare_df['오차(Kg)']) > 0.1]
                        if not errors.empty:
                            st.error(f"🚨 총 {len(errors)}개의 PO에서 장부와 현장 보고 수량 간의 불일치가 발견되었습니다! 빨간색으로 표시된 오차를 확인하세요.")
                        else:
                            st.success("✅ 완벽합니다! ERP 장부의 모든 출고량 데이터가 현장 실무진 엑셀 데이터와 100% 일치합니다.")
                except Exception as e:
                    st.error(f"엑셀 분석 중 오류 발생. 양식이 맞는지 확인해주세요: {e}")
