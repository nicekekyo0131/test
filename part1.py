import os
import glob
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext

class SpecDataConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Spec Data 轉換工具")
        self.root.geometry("750x550")
        
        # 變數
        self.folder_path = tk.StringVar()
        self.processed_data_dict = {}
        
        # 建立 UI 介面
        self.create_widgets()
        
    def create_widgets(self):
        # 1. 路徑選取區塊
        frame_top = tk.LabelFrame(self.root, text=" 資料夾設定 ", padx=10, pady=10)
        frame_top.pack(fill="x", padx=10, pady=10)
        
        lbl_path = tk.Label(frame_top, text="指定資料夾:")
        lbl_path.pack(side="left", padx=5)
        
        entry_path = tk.Entry(frame_top, textvariable=self.folder_path, width=55)
        entry_path.pack(side="left", padx=5, expand=True, fill="x")
        
        btn_browse = tk.Button(frame_top, text="瀏覽...", command=self.browse_folder)
        btn_browse.pack(side="left", padx=5)
        
        # 2. 功能按鈕區塊
        frame_btns = tk.LabelFrame(self.root, text=" 執行操作 ", padx=10, pady=10)
        frame_btns.pack(fill="x", padx=10, pady=5)
        
        btn_step1 = tk.Button(frame_btns, text="spec Data", bg="#e1f5fe", width=18, command=self.run_spec_data)
        btn_step1.pack(side="left", padx=10, expand=True)
        
        btn_step2_wi = tk.Button(frame_btns, text="SAVE Data (wi vtgm)", bg="#e8f5e9", width=20, command=self.run_save_wi_vtgm)
        btn_step2_wi.pack(side="left", padx=10, expand=True)
        
        btn_step2_wo = tk.Button(frame_btns, text="SAVE Data (wo vtgm)", bg="#fff3e0", width=20, command=self.run_save_wo_vtgm)
        btn_step2_wo.pack(side="left", padx=10, expand=True)
        
        # 3. Log 訊息顯示區塊
        frame_log = tk.LabelFrame(self.root, text=" 執行紀錄 (Log) ", padx=10, pady=10)
        frame_log.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(frame_log, wrap=tk.WORD, height=15, bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)
        
    def log(self, message):
        """寫入 Log 訊息到介面"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        
    def browse_folder(self):
        """選擇資料夾並掃描檔案"""
        dir_selected = filedialog.askdirectory()
        if dir_selected:
            self.folder_path.set(dir_selected)
            self.log(f"已選擇資料夾: {dir_selected}")
            self.scan_folder_files(dir_selected)
            
    def scan_folder_files(self, path):
        """掃描資料夾內的檔案並列在 Log 中"""
        try:
            files = os.listdir(path)
            self.log("-" * 40)
            self.log(f"資料夾內檔案列表 ({len(files)} 個項目):")
            for f in files:
                self.log(f"  - {f}")
            self.log("-" * 40)
        except Exception as e:
            self.log(f"讀取資料夾檔案失敗: {str(e)}")

    def run_spec_data(self):
        """對應 spec_data_step_1"""
        folder = self.folder_path.get()
        if not folder or not os.path.exists(folder):
            self.log("錯誤: 請先選擇有效的資料夾！")
            return
        
        self.log("開始執行 spec Data 轉檔...")
        search_pattern = os.path.join(folder, "*.xls*")
        file_list = glob.glob(search_pattern)
        
        if not file_list:
            self.log("警告: 指定資料夾中找不到任何 .xls 或 .xlsx 檔案！")
            return
            
        results_folder = os.path.join(folder, "after_trans")
        if not os.path.exists(results_folder):
            os.makedirs(results_folder)
            self.log(f"已建立存放資料夾: {results_folder}")
            
        self.processed_data_dict = {} 
        
        for file_path in file_list:
            filename = os.path.basename(file_path)
            self.log(f"正在處理檔案: {filename}")
            try:
                sp1 = filename.split("_")
                if len(sp1) >= 2:
                    newname = f"{sp1[0]}_{sp1[1]}"
                else:
                    newname = os.path.splitext(filename)[0]
                
                xls = pd.ExcelFile(file_path)
                if "Model team Fitting data" not in xls.sheet_names:
                    self.log(f"  -> 錯誤: 檔案 {filename} 中找不到 'Model team Fitting data' 工作表")
                    continue
                    
                ws1_fitting = pd.read_excel(file_path, sheet_name="Model team Fitting data", header=None)
                
                W = ws1_fitting.iloc[8, 1]  # Row 9 Index 8
                L = ws1_fitting.iloc[9, 1]  # Row 10 Index 9
                T = ws1_fitting.iloc[7, 1]  # Row 8 Index 7
                
                cell_3_2 = ws1_fitting.iloc[2, 1] # Row 3 Index 2
                mos_type = "NMOS" if pd.notna(cell_3_2) and float(cell_3_2) > 0 else "PMOS"
                
                a_str = f"Instance{{W={W},L={L},T={T}}}"
                b_str = f"DataType{{{mos_type}}}"
                
                col_g = ws1_fitting.iloc[11:, 6].dropna()
                g1_last_row = col_g.index[-1] + 1
                num = g1_last_row - 10
                
                target_rows = [
                    "//Copyright (C) 2010-2020 Primarius Technologies Co., Ltd",
                    "// All rights reserved.",
                    "",  # 這裡的空字串代表第三行為空白
                    "ObjInfo{}",
                    "ProbeInfo{}",
                    "Version{4.1}",
                    "ModelType{DC}",
                    b_str,
                    "Delimiter{,}",
                    "WorkingMode{Forward}",
                    a_str,
                    "Node{D|7|0.000000,G|8|0.000000,S|6|0.000000,B|5|0.000000}",
                    "MeasurementSetup{PDA}",
                    "Input{Time}",
                    "Output{vtlin,vtgm,vtsat,idlin,idsat}"
                ]
                
                df_target = pd.DataFrame({0: target_rows})
                
                extra_data = []
                for j in range(5):
                    col_idx = 6 + j * 3
                    subset = ws1_fitting.iloc[11 : 11 + num, col_idx].reset_index(drop=True)
                    extra_data.append(subset)
                    
                combined_extra = pd.concat(extra_data, ignore_index=True)
                df_target = pd.concat([df_target, combined_extra.to_frame(name=0)], ignore_index=True)
                
                self.processed_data_dict[newname] = df_target
                self.log(f"  -> 成功處理並暫存: {newname}")
                
            except Exception as e:
                self.log(f"  -> 處理 {filename} 時發生錯誤: {str(e)}")
                
        self.log("spec Data 執行完畢！")
        self.log("提示: 不需要 vtgm 請執行 SAVE Data (wo vtgm)，否則請執行 SAVE Data (wi vtgm)")

    def run_save_wi_vtgm(self):
        """對應 spec_data_step_2_save_and_remove_space + step_3_save_dat"""
        if not self.processed_data_dict:
            self.log("錯誤: 沒有可儲存的資料，請先執行 'spec Data'！")
            return
            
        folder = self.folder_path.get()
        results_folder = os.path.join(folder, "after_trans")
        
        self.log("開始清理空白列並儲存 .dat 檔案...")
        for name, df in self.processed_data_dict.items():
            df_cleaned = df.dropna(how='all').reset_index(drop=True)
            out_path = os.path.join(results_folder, f"{name}.dat")
            
            # 改用純文字逐行寫入，確保空字串不會變成 ""，而是真正的空白行
            with open(out_path, 'w', encoding='utf-8') as f:
                for val in df_cleaned[0]:
                    if pd.isna(val) or str(val).strip() == "":
                        f.write("\n")
                    else:
                        f.write(f"{val}\n")
                        
            self.log(f"已儲存: {out_path}")
            
        self.log("全部檔案儲存完成！")

    def run_save_wo_vtgm(self):
        """對應 spec_data_step2_remove_vtgm_option"""
        if not self.processed_data_dict:
            self.log("錯誤: 沒有可處理的資料，請先執行 'spec Data'！")
            return
            
        self.log("開始移除 vtgm 選項相關資料...")
        folder = self.folder_path.get()
        
        for name, df in self.processed_data_dict.items():
            match_idx = df[df[0].astype(str).str.contains(r"\{vtgm_Time\}", case=False, na=False)].index
            
            if not match_idx.empty:
                row_idx = match_idx[0]
                df_trimmed = df.iloc[:row_idx].reset_index(drop=True)
                self.processed_data_dict[name] = df_trimmed
                self.log(f"[{name}] 成功移除 {row_idx} 行之後包含 vtgm 的資料")
            else:
                self.log(f"[{name}] 未找到 {{vtgm_Time}} 關鍵字")
                
        self.run_save_wi_vtgm()

if __name__ == "__main__":
    root = tk.Tk()
    app = SpecDataConverterApp(root)
    root.mainloop()
