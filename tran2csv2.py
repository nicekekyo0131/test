import os
import glob
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext

class SpecDataConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("IC 數據自動轉檔整合工具 (Part 1, 2, 3)")
        self.root.geometry("820x720")
        
        # 變數
        self.folder_path = tk.StringVar()
        self.processed_data_dict = {}
        
        # 建立 Grid 網格狀 UI 介面
        self.create_widgets()
        
    def create_widgets(self):
        # 讓主視窗的欄位可以隨著視窗大小延展
        self.root.columnconfigure(0, weight=1)
        
        # ==================== 1. 資料夾設定區 (Row 0) ====================
        frame_top = tk.LabelFrame(self.root, text=" 資料夾設定 ", padx=10, pady=10)
        frame_top.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        frame_top.columnconfigure(1, weight=1)
        
        lbl_path = tk.Label(frame_top, text="指定資料夾:")
        lbl_path.grid(row=0, column=0, sticky="w", padx=5)
        
        entry_path = tk.Entry(frame_top, textvariable=self.folder_path, width=50)
        entry_path.grid(row=0, column=1, sticky="ew", padx=5)
        
        btn_browse = tk.Button(frame_top, text="瀏覽...", command=self.browse_folder, width=10)
        btn_browse.grid(row=0, column=2, padx=5)
        
        # ==================== 2. 功能按鈕操作區 (Row 1 ~ 3) ====================
        # Part 1: Spec Data
        frame_part1 = tk.LabelFrame(self.root, text=" Part 1: Spec Data 操作 ", padx=10, pady=10)
        frame_part1.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        frame_part1.columnconfigure((0, 1, 2), weight=1)
        
        btn_step1 = tk.Button(frame_part1, text="spec Data", bg="#e1f5fe", command=self.run_spec_data)
        btn_step1.grid(row=0, column=0, sticky="ew", padx=5, ipady=4)
        
        btn_step2_wi = tk.Button(frame_part1, text="SAVE Data (wi vtgm)", bg="#e8f5e9", command=self.run_save_wi_vtgm)
        btn_step2_wi.grid(row=0, column=1, sticky="ew", padx=5, ipady=4)
        
        btn_step2_wo = tk.Button(frame_part1, text="SAVE Data (wo vtgm)", bg="#fff3e0", command=self.run_save_wo_vtgm)
        btn_step2_wo.grid(row=0, column=2, sticky="ew", padx=5, ipady=4)
        
        # Part 2: HCI IV Data
        frame_part2 = tk.LabelFrame(self.root, text=" Part 2: HCI IV Data 操作 ", padx=10, pady=10)
        frame_part2.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
        frame_part2.columnconfigure((0, 1), weight=1)
        
        btn_hci_trans = tk.Button(frame_part2, text="HCI IV Data", bg="#e1f5fe", command=self.run_hci_iv_trans)
        btn_hci_trans.grid(row=0, column=0, sticky="ew", padx=5, ipady=4)
        
        btn_hci_save = tk.Button(frame_part2, text="SAVE IV Data", bg="#e8f5e9", command=self.run_save_iv_dat)
        btn_hci_save.grid(row=0, column=1, sticky="ew", padx=5, ipady=4)

        # Part 3: BTI IV Data
        frame_part3 = tk.LabelFrame(self.root, text=" Part 3: BTI IV Data 操作 ", padx=10, pady=10)
        frame_part3.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
        frame_part3.columnconfigure((0, 1), weight=1)
        
        btn_bti_trans = tk.Button(frame_part3, text="BTI IV Data", bg="#e1f5fe", command=self.run_bti_iv_trans)
        btn_bti_trans.grid(row=0, column=0, sticky="ew", padx=5, ipady=4)
        
        btn_bti_save = tk.Button(frame_part3, text="SAVE IV Data", bg="#e8f5e9", command=self.run_save_iv_dat)
        btn_bti_save.grid(row=0, column=1, sticky="ew", padx=5, ipady=4)

        # ==================== 3. Log 訊息顯示區 (Row 4) ====================
        frame_log = tk.LabelFrame(self.root, text=" 執行紀錄 (Log) ", padx=10, pady=10)
        frame_log.grid(row=4, column=0, sticky="nsew", padx=10, pady=5)
        frame_log.rowconfigure(0, weight=1)
        frame_log.columnconfigure(0, weight=1)
        self.root.rowconfigure(4, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(frame_log, wrap=tk.WORD, bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 10))
        self.log_text.grid(row=0, column=0, sticky="nsew")
        
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

    def check_folder(self):
        folder = self.folder_path.get()
        if not folder or not os.path.exists(folder):
            self.log("錯誤: 請先選擇有效的資料夾！")
            return None
        results_folder = os.path.join(folder, "after_trans")
        if not os.path.exists(results_folder):
            os.makedirs(results_folder)
        return folder

    # ==================== 共用純文字寫入 (第三行空白、無引號) ====================
    def save_to_dat_files(self):
        if not self.processed_data_dict:
            self.log("錯誤: 沒有可儲存的資料！")
            return
            
        folder = self.folder_path.get()
        results_folder = os.path.join(folder, "after_trans")
        
        self.log("開始清理空白列並儲存 .dat 檔案...")
        for name, df in self.processed_data_dict.items():
            df_cleaned = df.dropna(how='all').reset_index(drop=True)
            out_path = os.path.join(results_folder, f"{name}.dat")
            
            with open(out_path, 'w', encoding='utf-8') as f:
                for val in df_cleaned[0]:
                    if pd.isna(val) or str(val).strip() == "":
                        f.write("\n")
                    else:
                        f.write(f"{val}\n")
            self.log(f"已儲存: {out_path}")
        self.log("全部檔案儲存完成！")

    # ==================== Part 1 邏輯 ====================
    def run_spec_data(self):
        folder = self.check_folder()
        if not folder: return
        
        self.log("開始執行 spec Data 轉檔...")
        file_list = glob.glob(os.path.join(folder, "*.xls*"))
        if not file_list:
            self.log("警告: 指定資料夾中找不到任何檔案！")
            return
            
        self.processed_data_dict = {} 
        for file_path in file_list:
            filename = os.path.basename(file_path)
            self.log(f"正在處理檔案: {filename}")
            try:
                sp1 = filename.split("_")
                newname = f"{sp1[0]}_{sp1[1]}" if len(sp1) >= 2 else os.path.splitext(filename)[0]
                
                xls = pd.ExcelFile(file_path)
                if "Model team Fitting data" not in xls.sheet_names:
                    continue
                ws1_fitting = pd.read_excel(file_path, sheet_name="Model team Fitting data", header=None)
                
                W = ws1_fitting.iloc[8, 1]
                L = ws1_fitting.iloc[9, 1]
                T = ws1_fitting.iloc[7, 1]
                cell_3_2 = ws1_fitting.iloc[2, 1]
                mos_type = "NMOS" if pd.notna(cell_3_2) and float(cell_3_2) > 0 else "PMOS"
                
                a_str = f"Instance{{W={W},L={L},T={T}}}"
                b_str = f"DataType{{{mos_type}}}"
                
                col_g = ws1_fitting.iloc[11:, 6].dropna()
                g1_last_row = col_g.index[-1] + 1
                num = g1_last_row - 10
                
                target_rows = [
                    "//Copyright (C) 2010-2020 Primarius Technologies Co., Ltd",
                    "// All rights reserved.", "",
                    "ObjInfo{}", "ProbeInfo{}", "Version{4.1}", "ModelType{DC}",
                    b_str, "Delimiter{,}", "WorkingMode{Forward}", a_str,
                    "Node{D|7|0.000000,G|8|0.000000,S|6|0.000000,B|5|0.000000}",
                    "MeasurementSetup{PDA}", "Input{Time}", "Output{vtlin,vtgm,vtsat,idlin,idsat}"
                ]
                df_target = pd.DataFrame({0: target_rows})
                
                extra_data = [ws1_fitting.iloc[11:11+num, 6 + j * 3].reset_index(drop=True) for j in range(5)]
                combined_extra = pd.concat(extra_data, ignore_index=True)
                df_target = pd.concat([df_target, combined_extra.to_frame(name=0)], ignore_index=True)
                
                self.processed_data_dict[newname] = df_target
                self.log(f"  -> 成功處理並暫存: {newname}")
            except Exception as e:
                self.log(f"  -> 處理 {filename} 錯誤: {str(e)}")
        self.log("spec Data 執行完畢！")

    def run_save_wi_vtgm(self):
        self.save_to_dat_files()

    def run_save_wo_vtgm(self):
        if not self.processed_data_dict: return
        for name, df in self.processed_data_dict.items():
            match_idx = df[df[0].astype(str).str.contains(r"\{vtgm_Time\}", case=False, na=False)].index
            if not match_idx.empty:
                row_idx = match_idx[0]
                self.processed_data_dict[name] = df.iloc[:row_idx].reset_index(drop=True)
        self.save_to_dat_files()

    # ==================== Part 2: HCI IV Data 邏輯 ====================
    def run_hci_iv_trans(self):
        folder = self.check_folder()
        if not folder: return
        
        self.log("開始執行 HCI IV Data 轉檔...")
        file_list = glob.glob(os.path.join(folder, "*.xls*"))
        if not file_list:
            self.log("警告: 找不到檔案！")
            return
            
        self.processed_data_dict = {}
        for file_path in file_list:
            filename = os.path.basename(file_path)
            self.log(f"正在處理 HCI 檔案: {filename}")
            try:
                sp1 = filename.split("_")
                prefix = f"{sp1[0]}_{sp1[1]}" if len(sp1) >= 2 else os.path.splitext(filename)[0]
                
                xls = pd.ExcelFile(file_path)
                ws1_fitting = pd.read_excel(file_path, sheet_name="Model team Fitting data", header=None)
                ws_idvg = pd.read_excel(file_path, sheet_name="IdVg_raw", header=None)
                ws_idvd = pd.read_excel(file_path, sheet_name="IdVd_raw", header=None)
                ws_param = pd.read_excel(file_path, sheet_name="Parameter", header=None) if "Parameter" in xls.sheet_names else None
                
                W, L, T = ws1_fitting.iloc[8, 1], ws1_fitting.iloc[9, 1], ws1_fitting.iloc[7, 1]
                vdlin, vdd = ws1_fitting.iloc[13, 1], ws1_fitting.iloc[12, 1]
                vdstr, vgstr, vbstr = ws1_fitting.iloc[3, 1], ws1_fitting.iloc[2, 1], ws1_fitting.iloc[5, 1]
                mos_type = "NMOS" if ws1_fitting.iloc[2, 1] > 0 else "PMOS"
                
                a_str = f"Instance{{W={W},L={L},T={T}}}"
                b_str = f"DataType{{{mos_type}}}"
                
                # 1. hci_IV_t0
                self.processed_data_dict[f"{prefix}_IV_t0"] = self._build_hci_t0(
                    ws_idvg, ws_idvd, a_str, b_str, vdlin, vdd, vdstr, vgstr, vbstr
                )
                
                # 2. hci_iv_final_stress
                if ws_param is not None:
                    self.processed_data_dict[f"{prefix}_IV_filnal_stress"] = self._build_hci_final(
                        ws_idvg, ws_idvd, ws_param, a_str, b_str, vdlin, vdd, vdstr, vgstr, vbstr
                    )
                    
                self.log(f"  -> HCI 處理成功: {prefix}")
            except Exception as e:
                self.log(f"  -> HCI 處理 {filename} 錯誤: {str(e)}")
        self.log("HCI IV Data 轉檔完畢！可點擊 SAVE IV Data 輸出。")

    def _build_hci_t0(self, ws_idvg, ws_idvd, a, b, vdlin, vdd, vdstr, vgstr, vbstr):
        time_str = "0"
        c = f"[vgs,Id(vbs=0.00000),vds={vdlin},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time={time_str}]"
        d = f"[vgs,Id(vbs=0.00000),vds={vdd},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time={time_str}]"
        e = f"[vds,Id(vgs={vdd}),vbs=0.00000,ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr=0,Time={time_str}]"
        
        num = ws_idvg.iloc[14:, 1].dropna().index[-1] + 1 - 14
        num2 = ws_idvd.iloc[17:, 1].dropna().index[-1] + 1 - 17
        
        headers = [
            "//Copyright (C) 2010-2020 Primarius Technologies Co., Ltd", "// All rights reserved.", "",
            "ObjInfo{}", "ProbeInfo{}", "Version{4.1}", "ModelType{DC}", b, "Delimiter{,}",
            "WorkingMode{Forward}", a, "Node{D|7|0.000000,G|8|0.000000,S|6|0.000000,B|5|0.000000}",
            "MeasurementSetup{PDA}", "Input{vgs,vbs,vds}", "Output{Id,Isub}",
            "{MOS_Id_Vgs_Vbs_DC}", c
        ]
        rows = [""] * (26 + num * 2 + num2 + 5)
        for idx, h in enumerate(headers): rows[idx] = h
        
        rows[18 + num - 1] = "{MOS_Id_Vgs_Vbs_DC}"
        rows[19 + num - 1] = d
        rows[23 + num * 2 - 1] = "{MOS_Id_Vds_Vgs_DC}"
        rows[24 + num * 2 - 1] = e
        
        for i in range(num):
            rows[17 + i] = f"{ws_idvg.iloc[14 + i, 0]},{ws_idvg.iloc[14 + i, 1]}"
            rows[20 + num + i] = f"{ws_idvg.iloc[14 + i, 0]},{ws_idvg.iloc[14 + i, 3]}"
        for j in range(num2):
            rows[25 + num * 2 + j] = f"{ws_idvd.iloc[17 + j, 0]},{ws_idvd.iloc[17 + j, 1]}"
        return pd.DataFrame({0: rows})

    def _build_hci_final(self, ws_idvg, ws_idvd, ws_param, a, b, vdlin, vdd, vdstr, vgstr, vbstr):
        tstress_row = ws_param[ws_param.apply(lambda row: row.astype(str).str.contains("Tstress").any(), axis=1)]
        if tstress_row.empty:
            time_str = "0"
        else:
            r_idx = tstress_row.index[0]
            g3 = ws_param.iloc[r_idx:, 0].dropna().index[-1]
            time_str = str(float(ws_param.iloc[g3, 0]) / 60)
            
        c = f"[vgs,Id(vbs=0.00000),vds={vdlin},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time={time_str}]"
        d = f"[vgs,Id(vbs=0.00000),vds={vdd},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time={time_str}]"
        e = f"[vds,Id(vgs={vdd}),vbs=0.00000,ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr=0,Time={time_str}]"
        
        num = ws_idvg.iloc[14:, 1].dropna().index[-1] + 1 - 14
        num2 = ws_idvd.iloc[17:, 1].dropna().index[-1] + 1 - 17
        num3 = ws_idvg.iloc[:, 0].dropna().index[-1] + 1 - num
        num4 = ws_idvd.iloc[:, 0].dropna().index[-1] + 1 - num2
        
        headers = [
            "//Copyright (C) 2010-2020 Primarius Technologies Co., Ltd", "// All rights reserved.", "",
            "ObjInfo{}", "ProbeInfo{}", "Version{4.1}", "ModelType{DC}", b, "Delimiter{,}",
            "WorkingMode{Forward}", a, "Node{D|7|0.000000,G|8|0.000000,S|6|0.000000,B|5|0.000000}",
            "MeasurementSetup{PDA}", "Input{vgs,vbs,vds}", "Output{Id,Isub}",
            "{MOS_Id_Vgs_Vbs_DC}", c
        ]
        rows = [""] * (26 + num * 2 + num2 + 5)
        for idx, h in enumerate(headers): rows[idx] = h
        
        rows[18 + num - 1] = "{MOS_Id_Vgs_Vbs_DC}"
        rows[19 + num - 1] = d
        rows[23 + num * 2 - 1] = "{MOS_Id_Vds_Vgs_DC}"
        rows[24 + num * 2 - 1] = e
        
        for i in range(num):
            rows[17 + i] = f"{ws_idvg.iloc[num3 + i, 0]},{ws_idvg.iloc[num3 + i, 1]}"
            rows[20 + num + i] = f"{ws_idvg.iloc[num3 + i, 0]},{ws_idvg.iloc[num3 + i, 3]}"
        for j in range(num2):
            rows[25 + num * 2 + j] = f"{ws_idvd.iloc[num4 + j, 0]},{ws_idvd.iloc[num4 + j, 1]}"
        return pd.DataFrame({0: rows})

    # ==================== Part 3: BTI IV Data 邏輯 ====================
    def run_bti_iv_trans(self):
        folder = self.check_folder()
        if not folder: return
        
        self.log("開始執行 BTI IV Data 轉檔...")
        file_list = glob.glob(os.path.join(folder, "*.xls*"))
        if not file_list:
            self.log("警告: 找不到檔案！")
            return
            
        self.processed_data_dict = {}
        for file_path in file_list:
            filename = os.path.basename(file_path)
            self.log(f"正在處理 BTI 檔案: {filename}")
            try:
                sp1 = filename.split("_")
                prefix = f"{sp1[0]}_{sp1[1]}" if len(sp1) >= 2 else os.path.splitext(filename)[0]
                
                xls = pd.ExcelFile(file_path)
                ws1_fitting = pd.read_excel(file_path, sheet_name="Model team Fitting data", header=None)
                ws_vtlin = pd.read_excel(file_path, sheet_name="IdVg_vtlin_raw", header=None)
                ws_vtsat = pd.read_excel(file_path, sheet_name="IdVg_vtsat_raw", header=None)
                ws_single = pd.read_excel(file_path, sheet_name="single Id", header=None) if "single Id" in xls.sheet_names else None
                
                W, L, T = ws1_fitting.iloc[8, 1], ws1_fitting.iloc[9, 1], ws1_fitting.iloc[7, 1]
                vdlin, vdd = ws1_fitting.iloc[13, 1], ws1_fitting.iloc[12, 1]
                vdstr, vgstr, vbstr = ws1_fitting.iloc[3, 1], ws1_fitting.iloc[2, 1], ws1_fitting.iloc[5, 1]
                mos_type = "NMOS" if ws1_fitting.iloc[2, 1] > 0 else "PMOS"
                
                a_str = f"Instance{{W={W},L={L},T={T}}}"
                b_str = f"DataType{{{mos_type}}}"
                
                # 1. BTI_IV_t0
                self.processed_data_dict[f"{prefix}_IV_t0"] = self._build_bti_t0(
                    ws_vtlin, ws_vtsat, a_str, b_str, vdlin, vdd, vdstr, vgstr, vbstr
                )
                
                # 2. BTI_IV_final_stress
                if ws_single is not None:
                    self.processed_data_dict[f"{prefix}_IV_filnal_stress"] = self._build_bti_final(
                        ws_vtlin, ws_vtsat, ws_single, a_str, b_str, vdlin, vdd, vdstr, vgstr, vbstr
                    )
                    
                self.log(f"  -> BTI 處理成功: {prefix}")
            except Exception as e:
                self.log(f"  -> BTI 處理 {filename} 錯誤: {str(e)}")
        self.log("BTI IV Data 轉檔完畢！可點擊 SAVE IV Data 輸出。")

    def _build_bti_t0(self, ws_vtlin, ws_vtsat, a, b, vdlin, vdd, vdstr, vgstr, vbstr):
        time_str = "0"
        c = f"[vgs,Id(vbs=0.00000),vds={vdlin},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time={time_str}]"
        d = f"[vgs,Id(vbs=0.00000),vds={vdd},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time={time_str}]"
        
        num = ws_vtlin.iloc[13:, 0].dropna().index[-1] + 1 - 13
        
        headers = [
            "//Copyright (C) 2010-2020 Primarius Technologies Co., Ltd", "// All rights reserved.", "",
            "ObjInfo{}", "ProbeInfo{}", "Version{4.1}", "ModelType{DC}", b, "Delimiter{,}",
            "WorkingMode{Forward}", a, "Node{D|7|0.000000,G|8|0.000000,S|6|0.000000,B|5|0.000000}",
            "MeasurementSetup{PDA}", "Input{vgs,vbs,vds}", "Output{Id,Isub}",
            "{MOS_Id_Vgs_Vbs_DC}", c
        ]
        rows = [""] * (20 + num * 2)
        for idx, h in enumerate(headers): rows[idx] = h
        
        rows[18 + num - 1] = "{MOS_Id_Vgs_Vbs_DC}"
        rows[19 + num - 1] = d
        
        for i in range(num):
            rows[17 + i] = f"{ws_vtlin.iloc[13 + i, 0]},{ws_vtlin.iloc[13 + i, 1]}"
            rows[20 + num + i] = f"{ws_vtsat.iloc[13 + i, 0]},{ws_vtsat.iloc[13 + i, 1]}"
        return pd.DataFrame({0: rows})

    def _build_bti_final(self, ws_vtlin, ws_vtsat, ws_single, a, b, vdlin, vdd, vdstr, vgstr, vbstr):
        g3 = ws_single.iloc[:, 0].dropna().index[-1]
        time_str = str(float(ws_single.iloc[g3, 0]) / 60)
        
        c = f"[vgs,Id(vbs=0.00000),vds={vdlin},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time={time_str}]"
        d = f"[vgs,Id(vbs=0.00000),vds={vdd},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time={time_str}]"
        
        num = ws_vtlin.iloc[13:, 0].dropna().index[-1] + 1 - 13
        
        match_lin = ws_vtlin[ws_vtlin.apply(lambda row: row.astype(str).str.contains("Vg\(V\)").any(), axis=1)]
        g4 = match_lin.index[0] if not match_lin.empty else 0
        
        match_sat = ws_vtsat[ws_vtsat.apply(lambda row: row.astype(str).str.contains("Vg\(V\)").any(), axis=1)]
        g5 = match_sat.index[0] if not match_sat.empty else 0
        
        headers = [
            "//Copyright (C) 2010-2020 Primarius Technologies Co., Ltd", "// All rights reserved.", "",
            "ObjInfo{}", "ProbeInfo{}", "Version{4.1}", "ModelType{DC}", b, "Delimiter{,}",
            "WorkingMode{Forward}", a, "Node{D|7|0.000000,G|8|0.000000,S|6|0.000000,B|5|0.000000}",
            "MeasurementSetup{PDA}", "Input{vgs,vbs,vds}", "Output{Id,Isub}",
            "{MOS_Id_Vgs_Vbs_DC}", c
        ]
        rows = [""] * (20 + num * 2)
        for idx, h in enumerate(headers): rows[idx] = h
        
        rows[18 + num - 1] = "{MOS_Id_Vgs_Vbs_DC}"
        rows[19 + num - 1] = d
        
        for i in range(num):
            rows[17 + i] = f"{ws_vtlin.iloc[g4 + 1 + i, 0]},{ws_vtlin.iloc[g4 + 1 + i, 1]}"
            rows[20 + num + i] = f"{ws_vtsat.iloc[g5 + 1 + i, 0]},{ws_vtsat.iloc[g5 + 1 + i, 1]}"
        return pd.DataFrame({0: rows})

    # ==================== 共用 SAVE_IV_dat ====================
    def run_save_iv_dat(self):
        self.save_to_dat_files()

if __name__ == "__main__":
    root = tk.Tk()
    app = SpecDataConverterApp(root)
    root.mainloop()
