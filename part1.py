import os
import glob
import re
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import pandas as pd

class SpecDataApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Spec Data 轉換工具 (文字檔剖析修正版)")
        self.root.geometry("700x550")

        # 變數設定
        self.folder_path = tk.StringVar()
        self.processed_data = {} # 記憶體中的工作表資料 (儲存 DataFrame)

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
        """對應 spec_data_step_1 (以文字檔方式解析，對應 OpenText 與 ConsecutiveDelimiter)"""
        base_dir = self.folder_path.get()
        if not base_dir:
            messagebox.showerror("錯誤", "請先選擇資料夾位置！")
            return

        self.processed_data.clear()
        self.log("=== 開始執行 spec Data (Step 1) ===")

        # 搜尋指定資料夾下的檔案
        search_pattern = os.path.join(base_dir, "*")
        all_paths = glob.glob(search_pattern)
        # 過濾掉資料夾，只抓檔案
        files = [p for p in all_paths if os.path.isfile(p) and not p.endswith(".dat")]

        if not files:
            messagebox.showwarning("提示", "指定資料夾中找不到任何檔案！")
            self.log("錯誤: 找不到原始資料檔案。")
            return

        self.log(f"找到 {len(files)} 個原始檔案進行處理。")

        for file_path in files:
            filename = os.path.basename(file_path)
            # 略過 after_trans 資料夾本身若被掃到
            if filename == "after_trans":
                continue
                
            self.log(f"正在以文字模式解析檔案: {filename}")
            try:
                # 模擬 VBA 的 Workbooks.OpenText 與 ConsecutiveDelimiter:=True
                # 逐行讀取純文字檔案，並以連續空白或 Tab 切割欄位
                rows = []
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    for line in f:
                        # 移除結尾換行並依空白/Tab切割 (ConsecutiveDelimiter 效果)
                        cols = re.split(r'\s+', line.strip())
                        if cols and cols[0] != '':
                            rows.append(cols)

                if not rows:
                    self.log(f"  -> 警告: 檔案 {filename} 內容為空或無法解析")
                    continue

                # 轉換為 DataFrame 格式
                df_fitting = pd.DataFrame(rows)

                # 檔名分割命名規則： sp1(0) & "_" & sp1(1)
                name_parts = filename.split("_")
                if len(name_parts) >= 2:
                    newname = f"{name_parts[0]}_{name_parts[1]}"
                else:
                    newname = os.path.splitext(filename)[0]

                # 暫存進記憶體字典中
                self.processed_data[newname] = df_fitting
                self.log(f"  -> 成功解析並建立工作表資料: {newname} (共 {len(rows)} 行)")

            except Exception as e:
                self.log(f"  -> 解析檔案 {filename} 失敗: {e}")

        if not self.processed_data:
            messagebox.showwarning("提示", "沒有成功解析任何檔案！")
            return

        self.log("Step 1 執行完成！")
        messagebox.showinfo("完成", "Step 1 執行完畢！\n不需要 vtgm 請執行 SAVE Data(wo vtgm)\n需要 vtgm 請執行 SAVE Data(wi vtgm)")

    def _save_dat_files(self):
        """對應 step_3_save_dat (以類似 xlTextPrinter 格式輸出)"""
        after_dir = self.get_target_dir()
        if not after_dir:
            return

        if not self.processed_data:
            messagebox.showwarning("提示", "目前沒有處理好的資料可供儲存！")
            return

        self.log("=== 開始儲存 .dat 檔案 ===")
        for sheet_name, df in self.processed_data.items():
            out_path = os.path.join(after_dir, f"{sheet_name}.dat")
            try:
                # 使用 tab 分隔輸出，不包含表頭與索引，模擬 xlTextPrinter 輸出
                df.to_csv(out_path, sep='\t', index=False, header=False, na_rep='')
                self.log(f"已儲存: {out_path}")
            except Exception as e:
                self.log(f"儲存 {sheet_name}.dat 失敗: {e}")

        self.log("全部檔案儲存完畢！")
        messagebox.showinfo("成功", f"檔案已全數儲存至:\n{after_dir}")

    def run_step_2_wi(self):
        """對應 spec_data_step_2_save_and_remove_space"""
        if not self.processed_data:
            messagebox.showwarning("提示", "無資料可處理，請先執行 'spec Data'！")
            return

        self.log("=== 執行 SAVE Data (wi vtgm): 清理空白列並存檔 ===")
        for sheet_name, df in self.processed_data.items():
            if not df.empty:
                # 嚴格對應原 VBA: 檢查第 1 欄若為空值則刪除該列
                df_cleaned = df[df.iloc[:, 0].notna() & (df.iloc[:, 0].astype(str).str.strip() != "")]
                self.processed_data[sheet_name] = df_cleaned
                self.log(f"工作表 [{sheet_name}] 已清除空白列。")

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
                        # 找到後，保留該列以前的資料（刪除該列及之後所有列）
                        df = df.iloc[:r_idx]
                        found = True
                        break
                if found:
                    break
            
            if found:
                self.log(f"工作表 [{sheet_name}] 成功找到並切除自 {{vtgm_Time}} 起的下方資料。")
            else:
                self.log(f"工作表 [{sheet_name}] 未找到關鍵字 {{vtgm_Time}}。")

            # 清除空白列
            if not df.empty:
                df = df[df.iloc[:, 0].notna() & (df.iloc[:, 0].astype(str).str.strip() != "")]

            self.processed_data[sheet_name] = df

        self._save_dat_files()

if __name__ == "__main__":
    root = tk.Tk()
    app = SpecDataApp(root)
    root.mainloop()
