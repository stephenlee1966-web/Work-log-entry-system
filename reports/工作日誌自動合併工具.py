import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
from datetime import datetime
import re

def get_resource_path(relative_path):
    """ 取得資源絕對路徑，相容於 PyInstaller 打包後的環境 """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class ExcelMergerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("工作日誌自動合併工具 (分頁細化版)")
        self.root.geometry("550x340")
        self.root.resizable(False, False)
        
        # === 視窗與工作列 ICON 設定 ===
        icon_path = get_resource_path("logo.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)
        
        # 變量儲存
        self.folder_path = tk.StringVar(value=os.getcwd()) # 預設當前目錄
        self.merge_mode = tk.StringVar(value="month")      # 預設模式為按月合併
        
        current_year = datetime.now().year
        current_month = datetime.now().month
        
        # --- 介面佈局 ---
        
        # 1. 資料夾選擇
        frame_dir = tk.LabelFrame(root, text=" 1. 選擇來源資料夾 ", padx=10, pady=10)
        frame_dir.pack(fill="x", padx=15, pady=10)
        
        entry_dir = tk.Entry(frame_dir, textvariable=self.folder_path, width=45)
        entry_dir.pack(side="left", padx=5)
        
        btn_browse = tk.Button(frame_dir, text="瀏覽...", command=self.browse_folder)
        btn_browse.pack(side="left", padx=5)
        
        # 2. 合併模式選擇
        frame_mode = tk.LabelFrame(root, text=" 2. 選擇合併模式 ", padx=10, pady=5)
        frame_mode.pack(fill="x", padx=15, pady=5)
        
        rb_month = tk.Radiobutton(frame_mode, text="按月份合併 (例如: 工作日誌時數報表_202606)", 
                                  variable=self.merge_mode, value="month", command=self.toggle_mode)
        rb_month.pack(anchor="w", padx=5, pady=2)
        
        rb_year = tk.Radiobutton(frame_mode, text="按整年份合併 (不同月份各自獨立分頁 Sheet)", 
                                 variable=self.merge_mode, value="year", command=self.toggle_mode)
        rb_year.pack(anchor="w", padx=5, pady=2)
        
        # 3. 時間範圍選擇
        frame_date = tk.LabelFrame(root, text=" 3. 選擇時間範圍 ", padx=10, pady=10)
        frame_date.pack(fill="x", padx=15, pady=5)
        
        # 年份下拉選單
        years_list = [str(y) for y in range(current_year - 5, current_year + 3)]
        self.combo_year = ttk.Combobox(frame_date, values=years_list, width=8, state="readonly")
        self.combo_year.set(str(current_year))
        self.combo_year.pack(side="left", padx=5)
        
        self.lbl_year = tk.Label(frame_date, text="年")
        self.lbl_year.pack(side="left", padx=2)
        
        # 月份下拉選單
        months_list = [f"{m:02d}" for m in range(1, 13)]
        self.combo_month = ttk.Combobox(frame_date, values=months_list, width=6, state="readonly")
        self.combo_month.set(f"{current_month:02d}")
        self.combo_month.pack(side="left", padx=5)
        
        self.lbl_month = tk.Label(frame_date, text="月")
        self.lbl_month.pack(side="left", padx=2)
        
        # 4. 執行按鈕
        btn_merge = tk.Button(root, text="開始合併檔案", bg="#4CAF50", fg="white", 
                              font=("Microsoft JhengHei", 12, "bold"), width=20, height=2,
                              command=self.start_merge)
        btn_merge.pack(pady=15)

    def browse_folder(self):
        """開啟資料夾選擇視窗"""
        selected_dir = filedialog.askdirectory(initialdir=self.folder_path.get())
        if selected_dir:
            self.folder_path.set(selected_dir)

    def toggle_mode(self):
        """當切換模式時，自動啟用或停用月份下拉選單"""
        if self.merge_mode.get() == "year":
            self.combo_month.configure(state="disabled")
        else:
            self.combo_month.configure(state="readonly")

    def start_merge(self):
        """執行 Excel 合併邏輯"""
        folder = self.folder_path.get()
        
        if not folder or not os.path.exists(folder):
            messagebox.showerror("錯誤", "請選擇正確的來源資料夾！")
            return
        
        year_str = self.combo_year.get()
        mode = self.merge_mode.get()
        
        try:
            files = os.listdir(folder)
        except Exception as e:
            messagebox.showerror("錯誤", f"無法讀取資料夾內容：\n{e}")
            return

        # ==================== 模式一：按月份合併 ====================
        if mode == "month":
            month_str = self.combo_month.get()
            keyword = f"工作日誌時數報表_{year_str}{month_str}"
            output_filename = f"{keyword}.xlsx"
            output_path = os.path.join(folder, output_filename)
            
            all_data = []
            for file in files:
                if keyword in file and file.endswith(('.xlsx', '.xls')) and file != output_filename:
                    file_path = os.path.join(folder, file)
                    try:
                        df = pd.read_excel(file_path)
                        if not df.empty:
                            all_data.append(df)
                    except Exception as e:
                        print(f"讀取檔案 {file} 時發生錯誤: {e}")
            
            if all_data:
                try:
                    merged_df = pd.concat(all_data, ignore_index=True)
                    merged_df.to_excel(output_path, index=False)
                    messagebox.showinfo("成功", f"【月份合併完成】\n已產生新檔案：\n{output_filename}")
                except Exception as e:
                    messagebox.showerror("失敗", f"儲存檔案時發生錯誤：\n{e}")
            else:
                messagebox.showwarning("提示", f"找不到包含「{keyword}」的員工原始檔案！")

        # ==================== 模式二：按整年份合併 (分頁 Sheet) ====================
        else:
            keyword = f"工作日誌時數報表_{year_str}"
            output_filename = f"{keyword}_全年.xlsx"
            output_path = os.path.join(folder, output_filename)
            
            # 使用字典來分類存放各個月份的資料 DataFrame 清單。範例：{"06": [df1, df2]}
            monthly_buckets = {}
            
            for file in files:
                if keyword in file and file.endswith(('.xlsx', '.xls')) and file != output_filename:
                    # 排除機制
                    if "_全年" in file:
                        continue
                    if file.startswith("工作日誌時數報表_"):
                        continue
                    
                    # 使用正則表達式精準擷取檔名中的 6 位數年月（例如從 202606 擷取出 06 月）
                    match = re.search(rf"工作日誌時數報表_{year_str}(\d{{2}})", file)
                    if match:
                        month_part = match.group(1) # 得到 "01" ~ "12" 之間的字串
                        
                        file_path = os.path.join(folder, file)
                        try:
                            df = pd.read_excel(file_path)
                            if not df.empty:
                                if month_part not in monthly_buckets:
                                    monthly_buckets[month_part] = []
                                monthly_buckets[month_part].append(df)
                        except Exception as e:
                            print(f"讀取檔案 {file} 時發生錯誤: {e}")
            
            # 開始寫入多工作表 Excel
            if monthly_buckets:
                try:
                    # 透過 ExcelWriter 進行多 Sheet 寫入
                    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                        # 依照月份大小排序寫入 (由 01月 到 12月)
                        for month_key in sorted(monthly_buckets.keys()):
                            # 將該月份所有員工的 DataFrame 進行上下合併
                            month_merged_df = pd.concat(monthly_buckets[month_key], ignore_index=True)
                            # 寫入對應的工作表分頁，名稱格式定為 "06月"
                            sheet_name = f"{month_key}月"
                            month_merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
                            
                    messagebox.showinfo("成功", f"【年度分頁合併完成】\n各月份已分開放置於不同工作表中！\n已產生新檔案：\n{output_filename}")
                except Exception as e:
                    messagebox.showerror("失敗", f"儲存年度分頁檔案時發生錯誤：\n{e}")
            else:
                messagebox.showwarning("提示", f"找不到該年度「{year_str}」任何符合條件的員工原始檔案！")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerGUI(root)
    root.mainloop()