import streamlit as st
import pandas as pd
import os
import glob

st.set_page_config(page_title="차량별 감사 페스티벌 주문 LIST", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
        .block-container { padding-top: 1.5rem; padding-bottom: 1rem; max-width: 100%; }
        div[data-testid="stDataFrame"] td { white-space: pre-wrap !important; font-size: 13px !important; }
        div[data-testid="stDataFrame"] th { font-size: 14px !important; background-color: #f4f6f9 !important; }
    </style>
""", unsafe_allow_html=True)

st.title("🚚 차량별 감사 페스티벌 주문 LIST")

@st.cache_data(show_spinner=False, ttl=60)
def load_and_process_data():
    # 1. 26년 셀별차량 파일 찾기
    mapping_files = glob.glob("*26년 셀별차량*.xls*")
    valid_mapping_files = [f for f in mapping_files if not os.path.basename(f).startswith('~$')]
    
    if not valid_mapping_files: 
        return None, "'26년 셀별차량' 원본 파일을 찾을 수 없습니다."
    
    latest_mapping_file = max(valid_mapping_files, key=os.path.getmtime)
    
    # 2. 원본 데이터 파일 찾기 (.xlsb 또는 .xlsx 대응)
    raw_files = glob.glob("감사페스티벌 우선대상 LIST.*")
    valid_raw_files = [f for f in raw_files if not os.path.basename(f).startswith('~$')]
    
    if not valid_raw_files: 
        return None, "'감사페스티벌 우선대상 LIST' 원본 파일이 없습니다."
    
    latest_raw_file = max(valid_raw_files, key=os.path.getmtime)

    # xlwings 대신 pandas(openpyxl)를 사용하여 서버 환경에서 안전하게 읽기
    try:
        df_mapping = pd.read_excel(latest_mapping_file, header=None)
        df_raw = pd.read_excel(latest_raw_file)
    except Exception as e:
        return None, f"엑셀 읽기 오류: {str(e)}"

    df_raw.columns = [str(c).strip() if c is not None else f"Unnamed_{i}" for i, c in enumerate(df_raw.columns)]
    
    car_col = next((c for c in ['차량번호', 'Vehicle Number(Full)', 'Vehicle Number(Shortl)', 'Delivery TruckNo.'] if c in df_raw.columns), None)
    if not car_col: return None, "차량번호 열을 찾을 수 없습니다."

    df_raw['SOCreationDate'] = pd.to_datetime(df_raw['SOCreationDate'], errors='coerce')
    mask = (df_raw['SOCreationDate'] >= '2026-06-08') & (df_raw['SOCreationDate'] <= '2026-07-05')
    df_filtered = df_raw.loc[mask].copy()

    # 차량번호-소속셀 매핑 사전 만들기
    car_to_cell = {}
    for index, row in df_mapping.iterrows():
        try:
            col_full_car = str(row.iloc[0]).replace(" ", "").replace(".0", "")
            col_4digit = str(row.iloc[2]).replace(" ", "").replace(".0", "")
            col_cell_name = str(row.iloc[3]).strip()
            
            if col_cell_name and col_cell_name.lower() not in ['nan', 'none', '']:
                if col_4digit.isdigit() and len(col_4digit) >= 4:
                    car_to_cell[col_4digit[-4:]] = col_cell_name
                if len(col_full_car) >= 4:
                    car_to_cell[col_full_car] = col_cell_name
                    if col_full_car[-4:].isdigit():
                        car_to_cell[col_full_car[-4:]] = col_cell_name
        except IndexError: continue

    def find_cell(car_string):
        full_car = str(car_string).strip().replace(" ", "").replace(".0", "")
        four_digit = full_car[-4:] if len(full_car) >= 4 else full_car
        if full_car in car_to_cell: return car_to_cell[full_car]
        if four_digit in car_to_cell: return car_to_cell[four_digit]
        return '미분류'

    df_filtered['소속셀'] = df_filtered[car_col].astype(str).apply(find_cell)
    df_filtered['SOCreationDate'] = df_filtered['SOCreationDate'].dt.strftime('%Y-%m-%d')

    for col in ['Delivery', 'Material', car_col]:
        if col in df_filtered.columns:
            df_filtered[col] = df_filtered[col].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '').str.strip()

    df_filtered = df_filtered.rename(columns={'Delivery': '납품번호', 'ShipToPartyName': '고객명'})

    if '납품번호' in df_filtered.columns and 'Material' in df_filtered.columns:
        df_filtered['GroupKey'] = df_filtered['납품번호'].replace('', pd.NA)
        df_filtered['GroupKey'] = df_filtered['GroupKey'].fillna('IDX_' + df_filtered.index.to_series().astype(str))
        
        def join_materials(series):
            unique_mats = list(dict.fromkeys([m for m in series if m]))
            return '\n'.join(unique_mats) if unique_mats else '품목 정보 없음'
            
        material_grouped = df_filtered.groupby('GroupKey')['Material'].apply(join_materials).reset_index()
        material_grouped.rename(columns={'Material': '배송 품목'}, inplace=True)
        
        df_filtered = df_filtered.drop_duplicates(subset=['GroupKey'], keep='first')
        df_filtered = df_filtered.drop(columns=['Material']).merge(material_grouped, on='GroupKey', how='left')
        df_filtered = df_filtered.drop(columns=['GroupKey'])

    sort_cols = [car_col, '고객명', 'SOCreationDate']
    valid_sort_cols = [c for c in sort_cols if c in df_filtered.columns]
    df_filtered = df_filtered.sort_values(by=valid_sort_cols)

    required_cols = ['차량번호', '납품번호', '배송 품목', 'SOCreationDate', '고객명', 'ShipToAddress', '소속셀']
    df_filtered = df_filtered.rename(columns={car_col: '차량번호'})
    final_cols = [col for col in required_cols if col in df_filtered.columns]
    
    return df_filtered[final_cols], None

col1, col2 = st.columns([1, 4])
with col1:
    if st.button("🔄 최신 데이터 불러오기"):
        st.cache_data.clear()

with st.spinner("엑셀 데이터를 분석하고 있습니다..."):
    df, error_msg = load_and_process_data()

if error_msg:
    st.error(error_msg)
else:
    st.markdown("### 🔍 내 배차 내역 검색")
    search_query = st.text_input("차량번호 4자리를 입력하세요 (예: 1234)", "", placeholder="차량번호 4자리 입력")

    if not search_query:
        st.info("👆 위 검색창에 본인의 차량번호를 입력하시면 배차 내역이 표시됩니다.")
    else:
        filtered_df = df[df['차량번호'].astype(str).str.contains(search_query.strip())]
        
        if filtered_df.empty:
            st.warning("일치하는 배차 내역이 없습니다. 차량번호를 다시 확인해주세요.")
        else:
            my_cell = filtered_df['소속셀'].iloc[0]
            st.success(f"📍 소속: **{my_cell}** / 검색된 배차 건수: **{len(filtered_df)}건**")
            
            display_df = filtered_df[['차량번호', '납품번호', 'SOCreationDate', '고객명', 'ShipToAddress']].copy()
            st.dataframe(display_df, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("#### 📦 [상세조회] 납품번호별 배송 품목 확인")
            
            delivery_list = filtered_df['납품번호'].tolist()
            selected_delivery = st.selectbox("확인할 납품번호를 선택하세요", delivery_list)

            if selected_delivery:
                row_data = filtered_df[filtered_df['납품번호'] == selected_delivery].iloc[0]
                with st.container():
                    st.info(f"📋 **납품번호: {selected_delivery}** (고객명: {row_data['고객명']})")
                    st.write(f"📍 **배송주소:** {row_data['ShipToAddress']}")
                    st.write(f"📅 **생성일:** {row_data['SOCreationDate']}")
                    st.markdown("📌 **[배송 품목 리스트]**")
                    st.code(row_data['배송 품목'], language="text")