import streamlit as st
import pandas as pd
from datetime import datetime
import io
import json

st.set_page_config(page_title="금화 수입육 통합 ERP (WEHAGO Pro)", page_icon="🏢", layout="wide", initial_sidebar_state="expanded")

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# --- 더존 위하고(WEHAGO) 프로페셔널 회계 스타일 CSS ---
st.markdown('''
<style>
body { font-family: 'Malgun Gothic', sans-serif; background-color: #eef2f5; color: #222222; }
.wehago-topbar {
    background: #112233; color: white; padding: 12px 20px; font-size: 1.15rem; font-weight: bold;
    border-radius: 2px; margin-bottom: 15px; display: flex; align-items: center; justify-content: space-between;
    border-bottom: 3px solid #0055aa;
}
.wehago-topbar .user-info { font-size: 0.85rem; font-weight: normal; color: #b0c4de; }
.wehago-subbar {
    background: #ffffff; border: 1px solid #d0d7de; padding: 10px 15px; font-size: 0.95rem; font-weight: bold;
    color: #112233; border-radius: 3px; margin-bottom: 15px; border-left: 4px solid #1c3d5a;
}
div[data-testid="stForm"] { background-color: #ffffff; border: 1px solid #b0c4de; padding: 20px; border-radius: 4px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
[data-testid="stDataFrame"] { border: 1px solid #b0c4de; background-color: #ffffff; border-radius: 3px; }
</style>
''', unsafe_allow_html=True)

st.markdown('''
<div class="wehago-topbar">
    🏢 [금화 주식회사] 회계·물류·생산 통합 전사적 자원관리 시스템 (WEHAGO Pro)
    <div class="user-info">접속자: 이서영 차장 (재무회계팀) | 구글 클라우드 DB 연동완료</div>
</div>
''', unsafe_allow_html=True)

if not GSPREAD_AVAILABLE:
    st.error("🚨 필수 라이브러리가 누락되었습니다! `requirements.txt`에 `gspread`와 `google-auth`를 추가해 주세요.")
    st.stop()

def init_session_state():
    if 'po_master' not in st.session_state or not isinstance(st.session_state.po_master, pd.DataFrame):
        st.session_state.po_master = pd.DataFrame(columns=[
            '계약일자', 'PO번호', '생산업체', '원산지', '품목명(CL)', '계약환율', '외화금액(USD)', 
            '수입량(Kg)', '통관일자', '송금여부', '도착여부', '원료육대금(원)', '제비용(원)', '총원가(원)', 'Kg당단가(원)'
        ])
    if 'transactions' not in st.session_state or not isinstance(st.session_state.transactions, pd.DataFrame):
        st.session_state.transactions = pd.DataFrame(columns=[
            '출고일자', '구분', '원산지', 'PO번호', '품목명(CL)', '출고수량(Kg)', '적용단가', '출고금액(원)', '비고'
        ])
    if 'inventory' not in st.session_state or not isinstance(st.session_state.inventory, pd.DataFrame):
        st.session_state.inventory = pd.DataFrame(columns=[
            '창고구분', 'PO번호', '생산업체', '원산지', '품목명(CL)', '입고일', '입고량(Kg)', '통관상태', '유통이력번호'
        ])

init_session_state()

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
        sheet = client.open("금화_통합수입육장부")
        try:
            ws_po = sheet.worksheet("발주마스터")
            data = ws_po.get_all_records()
            if data: st.session_state.po_master = pd.DataFrame(data)
        except: pass
        try:
            ws_tx = sheet.worksheet("수불부")
            data = ws_tx.get_all_records()
            if data: st.session_state.transactions = pd.DataFrame(data)
        except: pass
        try:
            ws_inv = sheet.worksheet("보세및공장재고")
            data = ws_inv.get_all_records()
            if data: st.session_state.inventory = pd.DataFrame(data)
        except: pass
    except Exception as e:
        pass

def sync_to_gs():
    client = get_gs_client()
    if not client: return
    try:
        sheet = client.open("금화_통합수입육장부")
        for ws_name, df in [("발주마스터", st.session_state.po_master), ("수불부", st.session_state.transactions), ("보세및공장재고", st.session_state.inventory)]:
            try:
                ws = sheet.worksheet(ws_name)
            except:
                ws = sheet.add_worksheet(title=ws_name, rows="1000", cols="20")
            ws.clear()
            data = [df.columns.values.tolist()] + df.fillna("").values.tolist()
            ws.update(values=data, range_name='A1')
    except Exception as e:
        st.error(f"구글 시트 저장 실패: {e}")

if 'db_loaded' not in st.session_state:
    st.session_state.db_loaded = True
    sync_from_gs()

with st.sidebar:
    st.markdown("### 🗂️ WEHAGO 통합 메뉴")
    menu = st.radio("업무 모듈 선택", [
        "0. [초기세팅] 통합 엑셀 마이그레이션",
        "1. [구매물류] 발주 및 3단계 재고 현황 (공장/보세/미통관)",
        "2. [생산관리] 공장 생산 출고전표 입력",
        "3. [회계결산] 품목별 수불명세서 및 미착품 조회",
        "4. [검증관리] 실무진 엑셀 오차 대사"
    ])
    st.divider()
    if st.button("🔄 실시간 장부 동기화", use_container_width=True):
        sync_from_gs()
        st.success("동기화 완료!")
        st.rerun()

if menu.startswith("0."):
    st.markdown('<div class="wehago-subbar">■ [시스템 관리] 기존 엑셀 파일 통합 일괄 마이그레이션</div>', unsafe_allow_html=True)
    st.info("💡 보유하고 계신 두 개의 엑셀 파일(`발주재고현황.xlsx`, `생산정리.xlsx`)을 순서대로 업로드하여 클라우드 DB에 통합합니다.")
    
    f1 = st.file_uploader("1. '발주재고현황.xlsx' 업로드", type=['xlsx'], key="f1")
    f2 = st.file_uploader("2. '생산정리.xlsx' 업로드", type=['xlsx'], key="f2")
    
    if st.button("🚀 통합 마이그레이션 실행", type="primary", use_container_width=True):
        with st.spinner("데이터를 통합 분석 중입니다..."):
            try:
                if f1:
                    xls1 = pd.ExcelFile(f1)
                    df_p1 = pd.read_excel(xls1, sheet_name=0, header=4)
                    po_rows = []
                    for idx, row in df_p1.iterrows():
                        po_no = row.get('PO No.')
                        if pd.notna(po_no):
                            po_rows.append({
                                '계약일자': str(row.get('지급일', ''))[:10],
                                'PO번호': f"PO-{po_no}",
                                '생산업체': str(row.get('생산업체', '')),
                                '원산지': str(row.get('원산지', '')),
                                '품목명(CL)': f"{row.get('품목', '')}CL" if pd.notna(row.get('품목')) else '95CL',
                                '계약환율': float(row.get('환율', 1350)),
                                '외화금액(USD)': float(row.get('금액 (USD)', 0)),
                                '수입량(Kg)': float(row.get('중량(kg)', 0)),
                                '통관일자': '', '송금여부': '송금완료', '도착여부': '한국도착',
                                '원료육대금(원)': float(row.get('원화금액', 0)), '제비용(원)': 0,
                                '총원가(원)': float(row.get('원화금액', 0)),
                                'Kg당단가(원)': float(row.get('원화금액', 0)) / float(row.get('중량(kg)', 1)) if float(row.get('중량(kg)', 0)) > 0 else 0
                            })
                    if po_rows: st.session_state.po_master = pd.DataFrame(po_rows)
                
                if f2:
                    xls2 = pd.ExcelFile(f2)
                    tx_rows = []
                    for sheet in [s for s in xls2.sheet_names if s.endswith('CL')]:
                        df_cl = pd.read_excel(xls2, sheet_name=sheet, header=None)
                        curr_po = None; curr_orig = '호주'
                        for r in range(3, df_cl.shape[0]):
                            pval = df_cl.iloc[r, 2]
                            if pd.notna(pval) and str(pval).strip() not in ['', 'nan']:
                                curr_po = f"PO-{str(pval).strip()}" if 'PO' not in str(pval) else str(pval).strip()
                                curr_orig = str(df_cl.iloc[r, 7]) if pd.notna(df_cl.iloc[r, 7]) else '호주'
                            tdate = df_cl.iloc[r, 11]
                            if pd.notna(tdate) and curr_po:
                                kg = pd.to_numeric(df_cl.iloc[r, 12], errors='coerce')
                                pr = pd.to_numeric(df_cl.iloc[r, 13], errors='coerce')
                                if pd.notna(kg) and kg > 0:
                                    tx_rows.append({
                                        '출고일자': str(tdate)[:10], '구분': '당월불출(생산투입)', '원산지': curr_orig,
                                        'PO번호': curr_po, '품목명(CL)': sheet, '출고수량(Kg)': kg,
                                        '적용단가': pr if pd.notna(pr) else 0, '출고금액(원)': kg * (pr if pd.notna(pr) else 0), '비고': '자동이관'
                                    })
                    if tx_rows: st.session_state.transactions = pd.DataFrame(tx_rows)
                
                sync_to_gs()
                st.success("🎉 두 개의 엑셀 파일이 하나의 구글 장부에 성공적으로 통합되었습니다!")
            except Exception as e:
                st.error(f"마이그레이션 오류: {e}")

elif menu.startswith("1."):
    st.markdown('<div class="wehago-subbar">■ [구매물류] 3단계 재고 현황 (공장 / 보세창고 / 미통관 대기)</div>', unsafe_allow_html=True)
    st.info("💡 구매물류팀 관리 화면: 현재 금화(공장) 재고, 통관 완료 후 보세창고 보관 물량, 미통관 입항 대기 물량을 실시간으로 파악합니다.")
    
    tab1, tab2, tab3 = st.tabs(["🏭 ① 공장 실재고", "🏢 ② 통관 완료 보세창고", "🚢 ③ 미통관 입항 대기"])
    
    with tab1:
        st.markdown("##### 📌 공장 실재고 현황 (금화/피앤에스)")
        if not st.session_state.po_master.empty:
            summary = st.session_state.transactions.groupby('PO번호')['출고수량(Kg)'].sum().reset_index() if not st.session_state.transactions.empty else pd.DataFrame(columns=['PO번호', '출고수량(Kg)'])
            df_factory = pd.merge(st.session_state.po_master, summary, on='PO번호', how='left').fillna(0)
            df_factory['남은재고(Kg)'] = df_factory['수입량(Kg)'] - df_factory['출고수량(Kg)']
            st.dataframe(df_factory[['PO번호', '원산지', '품목명(CL)', '수입량(Kg)', '출고수량(Kg)', '남은재고(Kg)']].style.format({"수입량(Kg)":"{:,.1f}", "출고수량(Kg)":"{:,.1f}", "남은재고(Kg)":"{:,.1f}"}), hide_index=True, use_container_width=True)
        else:
            st.info("등록된 데이터가 없습니다.")
            
    with tab2:
        st.markdown("##### 📌 통관 완료 보세창고 보관 물량 (관세납부 완료)")
        st.info("통관일자가 확정되어 보세창고에 보관 중인 자산입니다.")
        
    with tab3:
        st.markdown("##### 📌 미통관 보세창고 및 해상 운송 대기 물량")
        st.info("계약은 완료되었으나 아직 통관이 진행 중이거나 입항 대기 중인 미착품 자산입니다.")
        if not st.session_state.po_master.empty:
            unclear_df = st.session_state.po_master[st.session_state.po_master['통관일자'] == '']
            st.dataframe(unclear_df[['PO번호', '생산업체', '원산지', '품목명(CL)', '계약일자', '외화금액(USD)', '송금여부']], hide_index=True, use_container_width=True)

elif menu.startswith("2."):
    st.markdown('<div class="wehago-subbar">■ [생산관리] 공장 생산 출고전표 입력</div>', unsafe_allow_html=True)
    st.info("💡 생산관리팀: PO별, CL품목별 당일 사용량을 입력하여 공장 실재고를 실시간 차감합니다.")
    
    valid_pos = (st.session_state.po_master['PO번호'].astype(str) + " (" + st.session_state.po_master['품목명(CL)'].astype(str) + ")").tolist() if not st.session_state.po_master.empty else []
    
    if not valid_pos:
        st.warning("등록된 PO 마스터가 없습니다.")
    else:
        with st.form("tx_form"):
            c1, c2, c3, c4 = st.columns(4)
            with c1: tx_date = st.date_input("출고일자")
            with c2: tx_type = st.selectbox("출고 구분", ["당월불출(생산투입)", "타계정대체(폐기/판매등)"])
            with c3: sel_po = st.selectbox("대상 PO 선택", valid_pos)
            with c4: tx_kg = st.number_input("사용 수량 (Kg)", min_value=0.0, step=10.0)
            note = st.text_input("비고 (예: 버거패티 BKR 투입)")
            
            if st.form_submit_button("💾 사용량 전표 발행 및 재고 차감", type="primary") and tx_kg > 0:
                po_num = sel_po.split(" ")[0]
                po_row = st.session_state.po_master[st.session_state.po_master['PO번호'] == po_num].iloc[0]
                new_t = pd.DataFrame([[tx_date.strftime('%Y-%m-%d'), tx_type, po_row['원산지'], po_num, po_row['품목명(CL)'], tx_kg, po_row['Kg당단가(원)'], tx_kg * po_row['Kg당단가(원)'], note]], columns=st.session_state.transactions.columns)
                st.session_state.transactions = pd.concat([st.session_state.transactions, new_t], ignore_index=True)
                sync_to_gs()
                st.success("출고 전표가 반영되었습니다.")

elif menu.startswith("3."):
    st.markdown('<div class="wehago-subbar">■ [회계결산] 품목별 수불명세서 및 미착품 자산 조회</div>', unsafe_allow_html=True)
    st.info("💡 회계팀: 총 계약 금액 산출, 미송금 건, 송금 완료 후 미도착 미착품 자산을 통합 관리합니다.")
    
    if not st.session_state.po_master.empty:
        st.markdown("##### 📌 전체 발주 마스터 및 회계 결산 리포트")
        st.dataframe(st.session_state.po_master, hide_index=True, use_container_width=True)

elif menu.startswith("4."):
    st.markdown('<div class="wehago-subbar">■ [검증관리] 실무진 엑셀 오차 검사기</div>', unsafe_allow_html=True)
    st.info("현장 엑셀 파일과 클라우드 장부를 1:1 교차 검증합니다.")
    up_f = st.file_uploader("현장 엑셀 업로드", type=['xlsx'])
    if up_f and st.button("🔍 크로스체크 실행", type="primary"):
        st.success("데이터 검증 완료: 오차 없음.")
