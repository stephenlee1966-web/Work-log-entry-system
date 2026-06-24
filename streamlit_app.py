import streamlit as st
import pandas as pd
import os
from datetime import datetime

# =====================================================================
# [設定] 網頁版面自適應設定
# =====================================================================
st.set_page_config(page_title="伺服器-工作日誌填寫系統", page_icon="📝", layout="wide")

st.markdown("### 📝 工作日誌填寫系統")

# =====================================================================
# 0. 設定伺服器存放路徑與初始化暫存記憶體
# =====================================================================
TARGET_FOLDER = "excel_files"
EMPLOYEE_FILE = "公司人員名單.xlsx"
OUTPUT_FOLDER = "reports"  # 報表儲存的目標資料夾

# 自動建立伺服器需要的資料夾
for folder in [TARGET_FOLDER, OUTPUT_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# 初始化全域暫存記憶體
if "export_buffer" not in st.session_state:
    st.session_state["export_buffer"] = []

# =====================================================================
# 1. 讀取伺服器資料夾內所有 Excel 檔名的函式
# =====================================================================
def get_server_excel_files(folder, employee_list):
    emp_files = [f"{name}.xlsx" for name in employee_list]
    return [f for f in os.listdir(folder) if (f.endswith('.xlsx') or f.endswith('.xls')) and f != EMPLOYEE_FILE and f not in emp_files]

# =====================================================================
# 1.5 獨立讀取「公司人員名單.xlsx」的函式
# =====================================================================
@st.cache_data
def load_company_employees():
    file_path = os.path.join(TARGET_FOLDER, EMPLOYEE_FILE)
    if not os.path.exists(file_path):
        st.error(f"❌ 找不到人員名單檔案！請將 `{EMPLOYEE_FILE}` 放入 `{TARGET_FOLDER}` 資料夾中。")
        return []
    try:
        emp_df = pd.read_excel(file_path)
        emp_df.columns = emp_df.columns.astype(str).str.strip()
        if '員工姓名' in emp_df.columns:
            names = emp_df['員工姓名'].dropna().astype(str).str.strip().unique().tolist()
            return sorted(names)
        else:
            st.error(f"❌ 在 `{EMPLOYEE_FILE}` 內找不到「員工姓名」這一個欄位！")
            return []
    except Exception as e:
        st.error(f"❌ 讀取人員名單檔案失敗: {e}")
        return []

# =====================================================================
# 1.6 讀取員工個人選單檔的函式（工作內容來源）
# =====================================================================
def load_employee_job_contents(employee_name):
    file_path = os.path.join(TARGET_FOLDER, f"{employee_name}.xlsx")
    if not os.path.exists(file_path):
        return [f"⚠️ 尚未建立您的專屬工作項目清單，請通知管理員建立 {employee_name}.xlsx"]
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None, header=None)
        all_text_items = []
        for sheet_name, df in all_sheets.items():
            flat_list = df.stack().dropna().astype(str).str.strip().tolist()
            valid_items = [item for item in flat_list if item != "" and not item.isnumeric()]
            all_text_items.extend(valid_items)
        unique_items = sorted(list(set(all_text_items)))
        return unique_items if unique_items else [f"💡 {employee_name}.xlsx 內容是空的。"]
    except Exception as e:
        return [f"❌ 讀取 {employee_name}.xlsx 失敗: {str(e)}"]

# =====================================================================
# 2. 核心功能：讀取單一 Excel 內「所有」工作表並自動合併清洗工程案號
# =====================================================================
@st.cache_data
def load_single_file_data(file_name):
    try:
        file_path = os.path.join(TARGET_FOLDER, file_name)
        all_sheets = pd.read_excel(file_path, sheet_name=None, header=None)
        
        combined_list = []
        final_id_column_name = None
        
        for sheet_name, df in all_sheets.items():
            if df.shape[0] < 2:
                continue
                
            row1 = df.iloc[0].fillna('').astype(str).str.strip().tolist()
            row2 = df.iloc[1].fillna('').astype(str).str.strip().tolist()
            combined_headers = [f"{r1} | {r2}" for r1, r2 in zip(row1, row2)]
            
            has_project_name = any('工程名稱' in col for col in combined_headers)
            has_id = any('工程案號' in col or '報價案號' in col for col in combined_headers)
            
            if has_project_name and has_id:
                name_idx = [i for i, col in enumerate(combined_headers) if '工程名稱' in col][0]
                id_idx = [i for i, col in enumerate(combined_headers) if '工程案號' in col or '報價案號' in col][0]
                
                if final_id_column_name is None:
                    final_id_column_name = '工程案號' if any('工程案號' in col for col in combined_headers) else '報價案號'
                
                sheet_df = pd.DataFrame()
                sheet_df['案號'] = df[id_idx]
                sheet_df['工程名稱'] = df[name_idx]
                sheet_df = sheet_df.iloc[2:].reset_index(drop=True)
                combined_list.append(sheet_df)
        
        if combined_list:
            final_df = pd.concat(combined_list, ignore_index=True)
            final_df = final_df.dropna(subset=['案號'])
            final_df['案號'] = final_df['案號'].astype(str).str.strip()
            final_df['案號'] = final_df['案號'].apply(lambda x: x.split('.')[0] if x.endswith('.0') else x)
            final_df['工程名稱'] = final_df['工程名稱'].fillna('未命名項目').astype(str).str.strip()
            
            final_df = final_df.rename(columns={'案號': final_id_column_name})
            return final_df, final_id_column_name
            
        return None, None
    except Exception as e:
        st.error(f"❌ 讀取檔案【{file_name}】時發生錯誤，已自動跳過。")
        return None, None

# =====================================================================
# 3. 主要網頁渲染邏輯
# =====================================================================
company_employees = load_company_employees()
available_files = get_server_excel_files(TARGET_FOLDER, company_employees)

if not company_employees:
    st.warning(f"⚠️ 員工名單載入失敗，請確認 `{EMPLOYEE_FILE}` 欄位是否包含「員工姓名」。")
elif not available_files:
    st.info(f"💡 目前伺服器的 `{TARGET_FOLDER}/` 資料夾內沒有任何工程/報價資料庫 Excel 檔案。")
else:
    st.markdown("#### 📋 填表基礎設定")
    
    row1_col1, row1_col2 = st.columns([1, 1], gap="small")
    with row1_col1:
        selected_date = st.date_input("📅 1. 日期：", value=datetime.today().date())
    with row1_col2:
        selected_employee = st.selectbox("👤 2. 姓名：", company_employees)
        
    selected_file = available_files[0] if available_files else ""
    
    row2_col1, row2_col2 = st.columns([4, 6], gap="small")
    with row2_col1:
        selected_file = st.selectbox("📁 3. 資料庫：", available_files)
        
    file_df, id_column = load_single_file_data(selected_file)
    id_label = id_column if id_column else '案號'
    
    project_list = []
    project_display_dict = {}
    if file_df is not None and not file_df.empty:
        project_list = file_df[id_label].unique().tolist()
        for index, row in file_df.drop_duplicates(subset=[id_label]).iterrows():
            project_display_dict[row[id_label]] = f"{row[id_label]} | {row['工程名稱']}"
            
    with row2_col2:
        selected_project = st.selectbox(
            f"📋 4. {id_label}：", 
            project_list,
            format_func=lambda x: project_display_dict.get(x, x),
            disabled=len(project_list) == 0
        )

    has_valid_selection = False
    current_selected_info = {}

    if file_df is not None and not file_df.empty and project_list:
        final_match_df = file_df[file_df[id_label] == selected_project]
        if not final_match_df.empty:
            project_info = final_match_df.iloc[0]
            has_valid_selection = True
            
            current_selected_info = {
                "填表日期": str(selected_date),
                "員工姓名": selected_employee,
                "工程/報價案號": selected_project,
                "工程名稱": project_info.get('工程名稱', '無資料')
            }

# =====================================================================
# 4. 中央主畫面：🛠️ 工作日誌內容填寫面板
# =====================================================================
    st.write("---")
    st.markdown("#### 🛠️ 工作日誌內容填寫")
    
    if has_valid_selection:
        st.info(f"選定：**{current_selected_info['工程/報價案號']} | {current_selected_info['工程名稱']}**")
        
        input_col1, input_col2, input_col3 = st.columns(
