import os
import glob
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import pandas as pd

class SpecDataApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Spec Data 轉換工具")
        self.root.geometry("700x550")

        # 變數設定
        self.folder_path = tk.StringVar()

        # 建立 UI 介面
        self.create_widgets()

    def create_widgets(self):
        # 1. 路徑選取區塊
        frame_top = tk.Frame(self.root, padx=10, pady=10)
        frame_top.pack(fill=tk.X)

        lbl_path = tk.Label(frame_top, text="讀取資料夾:")
        lbl_path.pack(side=tk.LEFT, padx=(0, 5))

        entry_path = tk.Entry(frame_top, textvariable=self.folder_path, width=50)
        entry_path.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        btn_browse = tk.Button(frame_top, text="瀏覽...", command=self.browse_folder)
        btn_browse.pack(side=tk.LEFT)

        # 2. 功能按鈕區塊
        frame_btns = tk.Frame(self.root, padx=10, pady=5)
        frame_btns.pack(fill=tk.X)

        self.btn_step1 = tk.Button(frame_btns, text="spec Data", bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), command=self.run_step_1)
        self.btn_step1.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.btn_step2_wi = tk.Button(frame_btns, text="SAVE Data (wi vtgm)", bg="#2196F3", fg="white", font=("Arial", 10, "bold"), command=self.run_step_2_wi)
        self.btn_step2_wi.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.btn_step2_wo = tk.Button(frame_btns, text="SAVE Data (wo vtgm)", bg="#FF9800", fg="white", font=("Arial", 10, "bold"), command=self.run_step_2_wo)
        self.btn_step2_wo.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        # 3. Log 視窗區塊
        frame_log = tk.Frame(self.root, padx=10, pady=10)
        frame_log.pack(expand=True, fill=tk.BOTH)

        lbl_log = tk.Label(frame_log, text="操作記錄 (Log):")
        lbl_log.pack(anchor=tk.W)

        self.log_text = scrolledtext.ScrolledText(frame_log, wrap=tk.WORD, height=20)
        self.log_text.pack(expand=True, fill=tk.BOTH, pady=(5, 0))
        
        # 記憶體中的暫存資料結構 (用來對應原本活頁簿中多個工作表的資料)
        self.processed_data = {} 

    def log(self, message):
        """寫入 Log 訊息到介面文字框"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def browse_folder(self):
        """選擇資料夾並列出檔案資訊"""
        dir_ = filedialog.askdirectory()
        if dir_:
            self.folder_path.set(dir_)
            self.log(f"已選擇資料夾: {dir_}")
            
            # 列出資料夾有哪些檔案 (符合需求 2)
            try:
                files = os.listdir(dir_)
                self.log(f"--- 該資料夾中的檔案清單 ({len(files)} 個項目) ---")
                for f in files:
                    self.log(f"  • {f}")
                self.log("---------------------------------------------")
            except Exception as e:
                self.log(f"讀取資料夾檔案失敗: {e}")

    def get_target_dir(self):
        """取得並自動建立 after_trans 資料夾路徑"""
        base_dir = self.folder_path.get()
        if not base_dir:
            messagebox.showerror("錯誤", "請先選擇讀取資料夾！")
            return None
        
        after_trans_dir = os.path.join(base_dir, "after_trans")
        if not os.path.exists(after_trans_dir):
            os.makedirs(after_trans_dir)
            self.log(f"建立存放資料夾: {after_trans_dir}")
        return after_trans_dir

    def run_step_1(self):
        """對應 spec_data_step_1"""
        base_dir = self.folder_path.get()
        if not base_dir:
            messagebox.showerror("錯誤", "請先選擇資料夾位置！")
            return

        self.processed_data.clear()
        self.log("=== 開始執行 spec Data (Step 1) ===")

        # 搜尋指定資料夾下的 .xls 或 .xlsx 檔案
        search_pattern = os.path.join(base_dir, "*.xls*")
        files = glob.glob(search_pattern)

        if not files:
            messagebox.showwarning("提示", "指定資料夾中找不到任何 Excel 檔案 (.xls / .xlsx)！")
            self.log("錯誤: 找不到原始資料檔案。")
            return

        self.log(f"找到 {len(files)} 個原始檔案進行處理。")

        for file_path in files:
            filename = os.path.basename(file_path)
            self.log(f"正在讀取檔案: {filename}")
            try:
                # 讀取 Excel 中的 "Model team Fitting data" 工作表
                # 這裡使用 pandas 讀取
                df_fitting = pd.read_excel(file_path, sheet_name="Model team Fitting data", header=None)
                
                # 檔名分割命名規則： sp1(0) & "_" & sp1(1)
                name_parts = filename.split("_")
                if len(name_parts) >= 2:
                    newname = f"{name_parts[0]}_{name_parts[1]}"
                else:
                    newname = os.path.splitext(filename)[0]

                # 暫存進記憶體字典中 (模擬多個工作表)
                self.processed_data[newname] = df_fitting
                self.log(f"  -> 成功萃取並建立工作表資料: {newname}")

            except Exception as e:
                self.log(f"  -> 讀取檔案 {filename} 失敗: {e}")

        self.log("Step 1 執行完成！")
        messagebox.println = messagebox.showinfo("完成", "Step 1 執行完畢！\n不需要 vtgm 請執行 SAVE Data(wo vtgm)\n需要 vtgm 請執行 SAVE Data(wi vtgm)")

    def _save_dat_files(self):
        """對應 step_3_save_dat"""
        after_dir = self.get_target_dir()
        if not after_dir:
            return

        if not self.processed_data:
            messagebox.showwarning("提示", "目前沒有處理好的資料可供儲存，請先執行 'spec Data'！")
            return

        self.log("=== 開始儲存 .dat 檔案 ===")
        for sheet_name, df in self.processed_data.items():
            out_path = os.path.join(after_dir, f"{sheet_name}.dat")
            try:
                # 轉成定位字元分隔 (TextPrinter 格式在 Python 中可用 sep='\t' 模擬)
                df.to_csv(out_path, sep='\t', index=False, header=False, na_rep='')
                self.log(f"已儲存: {out_path}")
            except Exception as e:
                self.log(f"儲存 {sheet_name}.dat 失敗: {e}")

        self.log("全部檔案儲存完畢！")
        messagebox.showinfo("成功", f"檔案已全數儲存至:\n{after_dir}")

    def run_step_2_wi(self):
        """對應 spec_data_step_2_save_and_remove_space 與 step_3_save_dat"""
        if not self.processed_data:
            messagebox.showwarning("提示", "無資料可處理，請先執行 'spec Data'！")
            return

        self.log("=== 執行 SAVE Data (wi vtgm): 清理空白列並存檔 ===")
        for sheet_name, df in self.processed_data.items():
            # 移除第一欄 (Col 1) 為空值的列
            # 原 VBA 是從底往上掃描若 Cells(i, 1) == "" 則刪除整列
            if not df.empty:
                # 保留第一欄有填寫資料的列 (dropna 或過濾)
                df_cleaned = df.dropna(subset=[df.columns[0]])
                # 也可以選擇清除全空白的列，此處嚴格對應原本 VBA 邏輯：檢查第 1 欄是否為空
                # 由於 pandas 處理方式較快，過濾掉第 1 欄是 NaN 或空字串的列
                df_cleaned = df[df.iloc[:, 0].notna() & (df.iloc[:, 0].astype(str).str.strip() != "")]
                self.processed_data[sheet_name] = df_cleaned
                self.log(f"工作表 [{sheet_name}] 已清除空白列。")

        # 執行儲存
        self._save_dat_files()

    def run_step_2_wo(self):
        """對應 spec_data_step2_remove_vtgm_option"""
        if not self.processed_data:
            messagebox.showwarning("提示", "無資料可處理，請先執行 'spec Data'！")
            return

        self.log("=== 執行 SAVE Data (wo vtgm): 移除 {vtgm_Time} 及其下方資料 ===")
        search_vtgm = "{vtgm_Time}"

        for sheet_name, df in self.processed_data.items():
            found = False
            # 搜尋包含 '{vtgm_Time}' 的儲存格
            for r_idx, row in df.iterrows():
                for c_idx, val in enumerate(row):
                    if val and search_vtgm in str(val):
                        # 找到後，刪除該列以及之後的所有列 (相當於 VBA 中的 Row 到 G3)
                        # 在 pandas 中，我們直接切片保留該列以前的資料
                        df = df.iloc[:r_idx]
                        found = True
                        break
                if found:
                    break
            
            if found:
                self.log(f"工作表 [{sheet_name}] 成功找到並切除自 {{vtgm_Time}} 起的下方資料。")
            else:
                self.log(f"工作表 [{sheet_name}] 未找到關鍵字 {{vtgm_Time}}。")

            # 接著進行清除空白列動作 (對應原 VBA 會接著呼叫 step_2_save_and_remove_space)
            if not df.empty:
                df = df[df.iloc[:, 0].notna() & (df.iloc[:, 0].astype(str).str.strip() != "")]

            self.processed_data[sheet_name] = df

        # 執行儲存
        self._save_dat_files()

if __name__ == "__main__":
    root = tk.Tk()
    app = SpecDataApp(root)
    root.mainloop()
