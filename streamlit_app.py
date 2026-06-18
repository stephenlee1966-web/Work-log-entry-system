import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime
# 引進 Google 官方核心 API 套件
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

# =====================================================================
# [設定] 網頁版面自適應設定
# =====================================================================
st.set_page_config(page_title="伺服器-工作日誌填寫系統", page_icon="📝", layout="wide")

st.markdown("### 📝 工作日誌填寫系統")

# =====================================================================
# 0. 初始化儲存設定與暫存記憶體
# =====================================================================
TARGET_FOLDER = "excel_files"
EMPLOYEE_FILE = "公司人員名單.xlsx"

# ⚠️ 【修改點 1】請確認為您雲端硬碟建立的資料夾 ID (不含網址)
GOOGLE_DRIVE_FOLDER_ID = "11z2FrCaJhspliWlZ96gKFNjQYJjHCZrh"

if not os.path.exists(TARGET_FOLDER):
    os.makedirs(TARGET_FOLDER)

if "export_buffer" not in st.session_state:
    st.session_state["export_buffer"] = []

# =====================================================================
# 🛠️ 核心功能：使用官方 API 連線至 Google Drive
# =====================================================================
def get_google_drive_service():
    """利用 Streamlit Cloud Secrets 中的 GCP 憑證初始化 Google Drive 官方服務物件"""
    try:
        gcp_info = {
            "type": st.secrets["connections"]["gsheets"]["type"],
            "project_id": st.secrets["connections"]["gsheets"]["project_id"],
            "private_key_id": st.secrets["connections"]["gsheets"]["private_key_id"],
            "private_key": st.secrets["connections"]["gsheets"]["private_key"],
            "client_email": st.secrets["connections"]["gsheets"]["client_email"],
            "auth_uri": st.secrets["connections"]["gsheets"]["auth_uri"],
            "token_uri": st.secrets["connections"]["gsheets"]["token_uri"],
            "auth_provider_x509_cert_url": st.secrets["connections"]["gsheets"]["auth_provider_x509_cert_url"],
            "client_x509_cert_url": st.secrets["connections"]["gsheets"]["client_x509_cert_url"]
        }
        
        scope = ['https://www.googleapis.com/auth/drive']
        credentials = service_account.Credentials.from_service_account_info(gcp_info, scopes=scope)
        
        # 建立官方 Drive API 服務物件
        service = build('drive', 'v3', credentials=credentials)
        return service
    except Exception as e:
        st.error(f"❌ Google 雲端硬碟連線初始化失敗，請檢查 Secrets 設定。錯誤: {e}")
        return None

def upload_excel_to_drive(file_name, dataframe):
    """將 Pandas DataFrame 轉成 Excel 並上傳/追加至 Google Drive (修正空間配額問題)"""
    service = get_google_drive_service()
    if service is None:
        return False
        
    try:
        # 1. 檢查雲端硬碟資料夾內是否已有同名檔案
        query = f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and name = '{file_name}' and trashed = false"
        results = service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])
        
        final_df = dataframe
        file_id = None
        
        if files:
            # 找到既有檔案的 ID
            file_id = files[0]['id']
            try:
                # 下載舊檔案進行資料合併追加
                request = service.files().get_media(fileId=file_id)
                downloaded_bytes = io.BytesIO()
                downloader = MediaIoBaseDownload(downloaded_bytes, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                
                downloaded_bytes.seek(0)
                existing_df = pd.read_excel(downloaded_bytes)
                
                if "備註" not in existing_df.columns:
                    existing_df["備註"] = ""
                    
                # 合併 舊資料 + 新資料
                final_df = pd.concat([existing_df, dataframe], ignore_index=True)
            except Exception:
                pass # 如果下載或讀取失敗，則退回直接覆蓋
                
        # 2. 將最終的 DataFrame 寫入記憶體二進位流
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, sheet_name="工作日誌報表", index=False)
            # 自動調整欄寬
            worksheet = writer.sheets["工作日誌報表"]
            for i, col in enumerate(final_df.columns):
                column_len = max(final_df[col].astype(str).str.len().max(), len(col)) + 4
                worksheet.set_column(i, i, column_len)
        
        excel_buffer.seek(0)
        media = MediaIoBaseUpload(excel_buffer, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
        
        # 3. 執行儲存
        if file_id:
            # 修改既有檔案
            service.files().update(fileId=file_id, media_body=media).execute()
        else:
            # 建立全新檔案
            file_metadata = {
                'name': file_name,
                'parents': [GOOGLE_DRIVE_FOLDER_ID]
            }
            new_file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            new_file_id = new_file.get('id')
            
            # ⚠️ 【修改點 2】請將下方改成您個人用來管理該雲端資料夾的真實 Google 帳號 Email
            USER_GMAIL_ACCOUNT = "your-email@gmail.com" 
            
            # 🌟 核心修正：將檔案權限建立給您自己，藉此調用您的硬碟空間配額，解決服務帳戶 0GB 的限制
            try:
                user_permission = {
                    'type': 'user',
                    'role': 'writer', 
                    'emailAddress': USER_GMAIL_ACCOUNT
                }
                service.permissions().create(fileId=new_file_id, body=user_permission).execute()
            except Exception:
                pass # 防止權限覆蓋失敗導致整個程式中斷
            
        return True
    except Exception as e:
        st.error(f"❌ 上傳至雲端硬碟失敗: {e}")
        return False

# =====================================================================
# 1. 讀取伺服器資料夾內所有 Excel 檔名的函式
# =====================================================================
def get_server_excel_files(folder, employee_list):
    emp_files = [f"{name}.xlsx" for name in employee_list]
    return [f_name for f_name in os.listdir(folder) if (f_name.endswith('.xlsx') or f_name.endswith('.xls')) and f_name != EMPLOYEE_FILE and f_name not in emp_files]

# =====================================================================
# 1.5 獨立讀取「公司人員名單.xlsx」的函式
# =====================================================================
@st.cache_data
def load_company_employees():
    file_path = os.path.join(TARGET_FOLDER, EMPLOYEE_FILE)
    if not os.path.exists(file_path):
        st.error(f"❌ 找不到人員名單檔案！請確認 GitHub 倉庫的 `{TARGET_FOLDER}` 資料夾內包含 `{EMPLOYEE_FILE}`。")
        return []
    try:
        emp_df = pd.read_excel(file_path)
        emp_df.columns = emp_df.columns.astype(str).str.strip()
        if '員工姓名' in emp_df.columns:
            names = emp_df['員工姓名'].dropna().astype(str).str.strip().unique().tolist()
            return sorted(names)
        else:
            st.error(f"❌ 內找不到「員工姓名」這一個欄位！")
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
        return None, None

# =====================================================================
# 3. 主要網頁渲染邏輯
# =====================================================================
company_employees = load_company_employees()
available_files = get_server_excel_files(TARGET_FOLDER, company_employees)

if company_employees and available_files:
    st.markdown("#### 📋 填表基礎設定")
    
    row1_col1, row1_col2 = st.columns([1, 1], gap="small")
    with row1_col1:
        selected_date = st.date_input("📅 1. 日期：", value=datetime.today().date())
    with row1_col2:
        selected_employee = st.selectbox("👤 2. 姓名：", company_employees)
        
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
        
        input_col1, input_col2, input_col3 = st.columns([44, 44, 12], gap="small")
        
        with input_col1:
            job_options = load_employee_job_contents(selected_employee)
            user_job_content = st.selectbox(
                f"📝 5. 工作內容：", options=job_options, key="employee_job_select_box"
            )
            current_selected_info["工作內容"] = user_job_content.strip()
            
        with input_col2:
            user_memo = st.text_input("✏️ 6. 備註：", value="", key="employee_job_memo_input")
            current_selected_info["備註"] = user_memo.strip()
            
        with input_col3:
            user_hours = st.number_input("⏱️ 7. 時數：", min_value=0.0, max_value=24.0, value=1.0, step=0.5)
            current_selected_info["填寫時數"] = user_hours
    else:
        st.warning("💡 請先完成上方 1~4 步驟項目。")

    # ─── 💾 暫存與清空按鈕 ───
    st.write("---")
    btn_col1, btn_col2 = st.columns(2, gap="small")
    
    with btn_col1:
        is_btn_disabled = (not has_valid_selection or 
                           "⚠️" in current_selected_info.get("工作內容", "") or 
                           "💡" in current_selected_info.get("工作內容", "") or 
                           "❌" in current_selected_info.get("工作內容", ""))
        
        if st.button("➕ 加入暫存區", width="stretch", disabled=is_btn_disabled, type="secondary", key="add_to_buffer_btn"):
            is_duplicate = any(
                item["填表日期"] == current_selected_info["填表日期"] and 
                item["員工姓名"] == current_selected_info["員工姓名"] and
                item["工程/報價案號"] == current_selected_info["工程/報價案號"] and
                item["工作內容"] == current_selected_info["工作內容"] and
                item["備註"] == current_selected_info["備註"]
                for item in st.session_state["export_buffer"]
            )
            if is_duplicate:
                st.warning("⚠️ 已在暫存區。")
            else:
                st.session_state["export_buffer"].append(current_selected_info.copy())
                st.rerun()

    with btn_col2:
        if st.button("🗑️ 清空暫存", width="stretch", key="clear_buffer_btn"):
            st.session_state["export_buffer"] = []
            st.rerun()

    # =====================================================================
    # 5. 渲染暫存區表格與儲存功能 (上傳至 Google Drive 雲端硬碟)
    # =====================================================================
    st.write("---")
    output_container = st.container(key="stable_output_container")
    
    with output_container:
        if st.session_state["export_buffer"]:
            st.markdown("##### 📝 待匯出暫存清單")
            buffer_df = pd.DataFrame(st.session_state["export_buffer"])
            display_df = buffer_df[["填表日期", "員工姓名", "工程/報價案號", "工程名稱", "工作內容", "備註", "填寫時數"]]
            
            st.dataframe(display_df, width="stretch", hide_index=True, key="main_data_table")
            
            # 點擊按鈕上傳至 Google Drive
            if st.button("💾 儲存至雲端硬碟 Google Drive", width="stretch", type="primary", key="save_report_btn"):
                report_owner = st.session_state["export_buffer"][0]["員工姓名"]
                # 每位員工每月獨立存成一個 Excel 檔案
                file_name = f"{report_owner}_工作日誌時數報表_{datetime.now().strftime('%Y%m')}.xlsx"
                
                with st.spinner("正在上傳檔案至 Google 雲端硬碟..."):
                    success = upload_excel_to_drive(file_name, display_df)
                    
                if success:
                    st.success(f"🎉 儲存成功！檔案 `{file_name}` 已透過官方 API 安全同步至您的 Google Drive。")
                    st.session_state["export_buffer"] = []
                    st.rerun()
        else:
            st.info("💡 暫存區無資料。請先選擇內容與時數後點擊「加入暫存區」。")
