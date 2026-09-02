import streamlit as st
import pandas as pd
import os
import glob

st.set_page_config(page_title="차량별 감사 페스티벌 주문 LIST", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
        .block-container { padding-top: 1rem; padding-bottom: 1rem; padding-left: 0.8rem; padding-right: 0.8rem; max-width: 100%; }
        div[data-testid="stDataFrame"] td { white-space: pre-wrap !important; font-size: 12px !important; word-break: break-all !important; }
        div[data-testid="stDataFrame"] th { font-size: 12px !important; background-color: #f4f6f9 !important; text-align: center !important; }
        input[type="text"] { font-size: 16px !important; }
        .custom-card { background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 12px; margin-bottom: 10px; }
    </style>
""", unsafe_allow_html=True)

st.title("🚚 감사 페스티벌 배차 조회")

@st.cache_data(show_spinner=False, ttl=60)
def load_and_process_data():
    mapping_files = glob.glob("*26년 셀별차량*.xls*")
    valid_mapping_files = [f for f in mapping_files if not os.path.basename(f).startswith('~$')]
    if not valid_mapping_files: return None, "'26년 셀별차량' 원본 파일을 찾을 수 없습니다."
    latest_mapping_file = max(valid_mapping_files, key=os.path.getmtime)
    
    raw_files = glob.glob("감사페스티벌 우선대상 LIST.*")
    valid_raw_files = [f for f in raw_files if not os.path.basename(f).startswith('~$')]
    if not valid_raw_files: return None, "'감사페스티벌 우선대상 LIST' 원본 파일이 없습니다."
    latest_raw_file = max(valid_raw_files, key=os.path.getmtime)

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

    # =================================================================
    # [지정 범위 기반 확실한 소속셀 매핑 로직]
    # 이미지의 블록 구조 특성을 반영하여 행을 내려가며 셀 이름을 누적 기억하고
    # 차량번호와 매칭시킵니다.
    # =================================================================
    car_to_cell = {}
    current_cell = "미분류"

    # 엑셀 전체 행을 탐색하며 차량번호와 소속셀 추출
    for idx, row in df_mapping.iterrows():
        row_values = [str(val).strip() for val in row.values if pd.notna(val)]
        
        # 행 안에서 소속셀 이름(예: 하누리, 스마일 등 짧은 한글 단어)이 포함되어 있는지 확인
        for val in row_values:
            clean_val = val.replace(" ", "").replace(".0", "")
            # 지역명이거나 차번 형식이 아니면서 글자 수가 짧은 경우 소속셀 이름으로 간주
            if len(clean_val) >= 2 and len(clean_val) <= 5 and not any(char.isdigit() for char in clean_val):
                if not any(k in clean_val for k in ['경북', '서울', '경기', '부산', '대구', '인천', '광주', '대전', '울산', '충청', '전라', '경상', '강원', '제주', '차량', '기사', '성명']):
                    current_cell = val

        # 행 안에서 차량번호 패턴(예: 경북86...)이 있는지 확인하여 현재 소속셀과 묶어줌
        for val in row_values:
            clean_val = val.replace(" ", "").replace(".0", "")
            if len(clean_val) >= 4 and ('경북' in clean_val or any(char.isdigit() for char in clean_val)):
                four_digit = clean_val[-4:] if clean_val[-4:].isdigit() else ""
                if current_cell != "미분류":
                    car_to_cell[clean_val] = current_cell
                    if four_digit:
                        car_to_cell[four_digit] = current_cell

    def find_cell(car_string):
        full_car = str(car_string).strip().replace(" ", "").replace(".0", "")
        four_digit = full_car[-4:] if len(full_car) >= 4 else full_car
        if full_car in car_to_cell: return car_to_cell[full_car]
        if four_digit in car_to_cell: return car_to_cell[four_digit]
        return '미분류'

    df_filtered['소속셀'] = df_filtered[car_col].astype(str).apply(find_cell)
    df_filtered['SOCreationDate'] = df_filtered['SOCreationDate'].dt.strftime('%m-%d')

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

with st.spinner("데이터 로딩 중..."):
    df, error_msg = load_and_process_data()

if error_msg:
    st.error(error_msg)
else:
    search_query = st.text_input("🔍 차량번호 4자리 입력", "", placeholder="예: 6791")

    if not search_query:
        st.info("👆 본인의 차량번호 4자리를 입력해주세요.")
    else:
        filtered_df = df[df['차량번호'].astype(str).str.contains(search_query.strip())]
        
        if filtered_df.empty:
            st.warning("일치하는 내역이 없습니다.")
        else:
            my_cell = filtered_df['소속셀'].iloc[0]
            st.success(f"📍 소속: **{my_cell}** (총 {len(filtered_df)}건)")
            
            display_df = filtered_df[['차량번호', '납품번호', 'SOCreationDate', '고객명']].copy()
            st.dataframe(display_df, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("📦 **상세 품목 및 주소 확인**")
            
            delivery_list = filtered_df['납품번호'].tolist()
            selected_delivery = st.selectbox("확인할 납품번호 선택", delivery_list)

            if selected_delivery:
                row_data = filtered_df[filtered_df['납품번호'] == selected_delivery].iloc[0]
                
                st.markdown(f"""
                <div class="custom-card">
                    <b>납품번호:</b> {selected_delivery}<br>
                    <b>고객명:</b> {row_data['고객명']}<br>
                    <b>배송주소:</b> {row_data['ShipToAddress']}<br>
                    <b>날짜:</b> {row_data['SOCreationDate']}
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("📌 **배송 품목:**")
                st.code(row_data['배송 품목'], language="text")