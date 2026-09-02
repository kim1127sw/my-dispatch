import streamlit as st
import pandas as pd
import os
import glob

# 페이지 설정
st.set_page_config(page_title="감사 페스티벌 배차 조회", layout="wide", initial_sidebar_state="collapsed")

# 모바일 최적화 여백 및 글자 크기
st.markdown("""
    <style>
        .block-container { padding-top: 1.5rem; padding-bottom: 1rem; padding-left: 0.8rem; padding-right: 0.8rem; max-width: 100%; }
        div[data-testid="stDataFrame"] td { white-space: pre-wrap !important; font-size: 13px !important; word-break: break-all !important; }
        div[data-testid="stDataFrame"] th { font-size: 13px !important; background-color: #f1f5f9 !important; text-align: center !important; }
    </style>
""", unsafe_allow_html=True)

# 1. 폰에서 안리지 않는 적절한 크기의 제목 복구
st.subheader("🚚 감사 페스티벌 배차 조회")

@st.cache_data(show_spinner=False, ttl=60)
def load_and_process_data():
    mapping_files = glob.glob("*26년 셀별차량*.xls*")
    valid_mapping_files = [f for f in mapping_files if not os.path.basename(f).startswith('~$')]
    if not valid_mapping_files: 
        return None, "'26년 셀별차량' 원본 파일을 찾을 수 없습니다.", "", ""
    latest_mapping_file = max(valid_mapping_files, key=os.path.getmtime)
    
    raw_files = glob.glob("감사페스티벌 우선대상 LIST.*")
    valid_raw_files = [f for f in raw_files if not os.path.basename(f).startswith('~$')]
    if not valid_raw_files: 
        return None, "'감사페스티벌 우선대상 LIST' 원본 파일이 없습니다.", "", ""
    latest_raw_file = max(valid_raw_files, key=os.path.getmtime)

    try:
        xls_map = pd.ExcelFile(latest_mapping_file)
        target_sheet = next((s for s in xls_map.sheet_names if '셀별차량' in s), xls_map.sheet_names[0])
        df_mapping_raw = pd.read_excel(xls_map, sheet_name=target_sheet, header=None)
        
        df_raw = pd.read_excel(latest_raw_file)
    except Exception as e:
        return None, f"엑셀 읽기 오류: {str(e)}", "", ""

    # A1:C132 범위 지정
    df_mapping = df_mapping_raw.iloc[0:132, 0:3]
    df_raw.columns = [str(c).strip() if c is not None else f"Unnamed_{i}" for i, c in enumerate(df_raw.columns)]
    
    car_col = next((c for c in ['차량번호', 'Vehicle Number(Full)', 'Vehicle Number(Shortl)', 'Delivery TruckNo.'] if c in df_raw.columns), None)
    if not car_col: return None, "차량번호 열을 찾을 수 없습니다.", "", ""

    df_raw['SOCreationDate'] = pd.to_datetime(df_raw['SOCreationDate'], errors='coerce')
    mask = (df_raw['SOCreationDate'] >= '2026-06-08') & (df_raw['SOCreationDate'] <= '2026-07-05')
    df_filtered = df_raw.loc[mask].copy()

    # 모든 소속 분류군을 완벽히 포함하여 매핑 탐색
    valid_cells = ['하누리', '스마일', '엘리트', '글로벌', '임팩트', '어울림', '올포원', '신세계']
    car_to_cell = {}
    current_cell = "미분류"
    
    for idx, row in df_mapping.iterrows():
        row_vals = [str(v).strip().replace(" ", "").replace(".0", "") for v in row.values if pd.notna(v) and str(v).strip().lower() not in ('nan', 'none', '')]
        if not row_vals: continue
            
        row_cell = None
        for v in row_vals:
            for vc in valid_cells:
                if vc in v:
                    row_cell = vc
                    break
            if row_cell: break
                
        if not row_cell:
            for v in row_vals:
                if len(v) >= 2 and len(v) <= 5 and not any(c.isdigit() for c in v):
                    if not any(k in v for k in ['경북', '서울', '경기', '부산', '대구', '인천', '광주', '대전', '울산', '충청', '전라', '경상', '강원', '제주', '차량', '번호', '성명', '기사']):
                        row_cell = v
                        break

        if row_cell: current_cell = row_cell

        for v in row_vals:
            if len(v) >= 4 and any(c.isdigit() for c in v):
                if not any(k in v for k in ['차량', '번호', '성명', '기사']):
                    target_cell = row_cell if row_cell else current_cell
                    if target_cell != "미분류":
                        car_to_cell[v] = target_cell
                        if v[-4:].isdigit(): car_to_cell[v[-4:]] = target_cell

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
    
    return df_filtered[final_cols], None, f"{os.path.basename(latest_mapping_file)} [{target_sheet}]", os.path.basename(latest_raw_file)

with st.spinner("데이터 로딩 중..."):
    df, error_msg, map_file_name, raw_file_name = load_and_process_data()

if error_msg:
    st.error(error_msg)
else:
    # 2. 검색 반응성 개선 (조회 버튼 추가)
    col1, col2 = st.columns([3, 1])
    with col1:
        search_query = st.text_input("차량번호 4자리 입력", "", placeholder="번호 입력 후 조회 터치", label_visibility="collapsed", max_chars=4)
    with col2:
        search_btn = st.button("🔍 조회", use_container_width=True)

    if not search_query.strip():
        st.info("👆 빈칸에 차량번호 4자리를 입력하고 돋보기 버튼을 눌러주세요.")
    else:
        filtered_df = df[df['차량번호'].astype(str).str.contains(search_query.strip())].reset_index(drop=True)
        
        if filtered_df.empty:
            st.warning("일치하는 배차 내역이 없습니다.")
        else:
            my_cell = filtered_df['소속셀'].iloc[0]
            st.success(f"📍 소속: **{my_cell}** (총 {len(filtered_df)}건)")
            
            display_df = filtered_df[['차량번호', '납품번호', 'SOCreationDate', '고객명', 'ShipToAddress']].copy()
            display_df = display_df.rename(columns={'SOCreationDate': '날짜', 'ShipToAddress': '배송주소'})
            
            st.markdown("👇 **목록을 터치하시면 하단에 배송 품목이 표시됩니다.**")
            
            # 표 클릭 시 즉시 반응
            event = st.dataframe(
                display_df, 
                use_container_width=True, 
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row"
            )

            selected_rows = event.selection.get("rows", [])
            st.markdown("---")
            
            delivery_list = filtered_df['납품번호'].tolist()
            default_idx = selected_rows[0] if selected_rows else 0
            
            # 라디오 버튼 추가로 터치 편의성 극대화
            selected_delivery = st.radio("📦 **납품번호 선택 (터치):**", delivery_list, index=default_idx, horizontal=True)

            if selected_delivery:
                row_data = filtered_df[filtered_df['납품번호'] == selected_delivery].iloc[0]
                
                # 3. 다크모드/라이트모드 모두 글씨가 완벽하게 보이는 전용 컨테이너
                with st.container(border=True):
                    st.markdown(f"**📋 납품번호:** {selected_delivery}")
                    st.markdown(f"**👤 고객명:** {row_data['고객명']}")
                    st.markdown(f"**📍 배송주소:** {row_data['ShipToAddress']}")
                    st.markdown(f"**📅 생성일자:** {row_data['SOCreationDate']}")
                
                st.markdown("📌 **[배송 품목 상세 리스트]**")
                st.code(row_data['배송 품목'], language="text")