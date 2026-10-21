import os
import glob
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

class IVTransApp:
    def __init__(self, root):
        self.root = root
        self.root.title("HCI IV Data 轉檔工具")
        self.root.geometry("800x650")

        # 儲存轉換後的資料結構（格式：{檔名: 資料內容清單}）
        self.trans_results = {}

        self.create_widgets()

    def create_widgets(self):
        # 1. 路徑選取區塊
        path_frame = tk.LabelFrame(self.root, text=" 資料夾設定 ", font=("Arial", 12, "bold"))
        path_frame.pack(fill="x", padx=15, pady=10)

        self.path_entry = tk.Entry(path_frame, font=("Arial", 11))
        self.path_entry.pack(side="left", fill="x", expand=True, padx=10, pady=10)

        browse_btn = tk.Button(path_frame, text="瀏覽資料夾", font=("Arial", 10), command=self.browse_folder)
        browse_btn.pack(side="right", padx=10, pady=10)

        # 2. 功能按鈕區塊
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill="x", padx=15, pady=5)

        self.btn_hci = tk.Button(btn_frame, text="HCI IV DATA", font=("Arial", 11, "bold"), 
                                 bg="#4CAF50", fg="white", width=20, height=2, command=self.run_hci_trans)
        self.btn_hci.pack(side="left", padx=10)

        self.btn_save = tk.Button(btn_frame, text="SAVE IV DATA", font=("Arial", 11, "bold"), 
                                  bg="#2196F3", fg="white", width=20, height=2, command=self.run_save_dat)
        self.btn_save.pack(side="right", padx=10)

        # 3. Log 記錄視窗區塊
        log_frame = tk.LabelFrame(self.root, text=" 操作記錄 (Log) ", font=("Arial", 12, "bold"))
        log_frame.pack(fill="both", expand=True, padx=15, pady=10)

        self.log_text = scrolledtext.ScrolledText(log_frame, font=("Consolas", 10), bg="#1e1e1e", fg="#00ff00")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)

        self.log("系統就緒，請選擇要讀取的資料夾。\n")

    def log(self, message):
        """寫入 Log 訊息到介面"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def browse_folder(self):
        """選擇資料夾按鈕事件"""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, folder_selected)
            self.log(f"已選擇資料夾: {folder_selected}")
            
            # 列出資料夾內有哪些檔案
            files = os.listdir(folder_selected)
            self.log(f"--- 該資料夾內容清單 ({len(files)} 個項目) ---")
            for f in files:
                self.log(f"  • {f}")
            self.log("-" * 50)

    def format_val(self, val):
        """將數值轉換為字串，強制避免科學記號 e"""
        if pd.isna(val):
            return "0.000000"
        try:
            f_val = float(val)
            # 若為整數或浮點數，轉為一般小數點表示法，避免科學記號
            if f_val.is_integer():
                return str(int(f_val))
            else:
                # 預留足夠小數點位數，可依需求調整
                return f"{f_val:.6f}".rstrip('0').rstrip('.') if '.' in f"{f_val:.6f}" else f"{f_val:.6f}"
        except ValueError:
            return str(val)

    def run_hci_trans(self):
        """執行 HCI IV DATA 轉檔邏輯"""
        target_dir = self.path_entry.get().strip()
        if not target_dir or not os.path.isdir(target_dir):
            messagebox.showerror("錯誤", "請先選擇有效的資料夾路徑！")
            return

        self.log("=== 開始執行 HCI IV DATA 轉檔 ===")
        self.trans_results.clear()

        # 搜尋指定資料夾底下的所有 excel 檔案
        search_pattern = os.path.join(target_dir, "*.[xlX][lsL]*")
        raw_files = [f for f in glob.glob(search_pattern) if not os.path.basename(f).startswith("~$")]

        if not raw_files:
            self.log("警告: 指定資料夾中找不到任何 Excel 檔案！")
            messagebox.showwarning("提示", "指定資料夾中找不到 Excel 檔案！")
            return

        self.log(f"找到 {len(raw_files)} 個原始 Excel 檔案進行處理。")

        try:
            # 建立 after_trans 資料夾
            results_folder = os.path.join(target_dir, "after_trans")
            os.makedirs(results_folder, exist_ok=True)

            con1_file = None

            for file_path in raw_files:
                source_file = os.path.basename(file_path)
                self.log(f"正在處理檔案: {source_file}")

                # 讀取 Excel 分頁並建立不分大小寫的對應字典
                xls = pd.ExcelFile(file_path)
                sheet_lower_map = {s.lower(): s for s in xls.sheet_names}

                required_sheets = ["Model team Fitting data", "IdVg_raw", "IdVd_raw", "Parameter"]
                real_sheets = {}
                for req in required_sheets:
                    req_lower = req.lower()
                    if req_lower in sheet_lower_map:
                        real_sheets[req] = sheet_lower_map[req_lower]
                    else:
                        raise ValueError(f"檔案 {source_file} 缺少必要的工作表: {req} (找不到對應的分頁)")

                df_fitting = pd.read_excel(file_path, sheet_name=real_sheets["Model team Fitting data"], header=None)
                df_idvg = pd.read_excel(file_path, sheet_name=real_sheets["IdVg_raw"], header=None)
                df_idvd = pd.read_excel(file_path, sheet_name=real_sheets["IdVd_raw"], header=None)
                df_param = pd.read_excel(file_path, sheet_name=real_sheets["Parameter"], header=None)

                # 剖析檔名以建立新工作表名稱
                sp1 = source_file.split("_")
                if len(sp1) >= 2:
                    newname = f"{sp1[0]}_{sp1[1]}_IV_t0"
                    newname2 = f"{sp1[0]}_{sp1[1]}_IV_filnal_stress"
                else:
                    newname = f"{sp1[0]}_IV_t0"
                    newname2 = f"{sp1[0]}_IV_filnal_stress"

                # 執行轉檔運算
                lines_t0 = self.hci_IV_t0(df_fitting, df_idvg, df_idvd)
                lines_final = self.hci_iv_final_stress(df_fitting, df_idvg, df_idvd, df_param)

                self.trans_results[newname] = lines_t0
                self.trans_results[newname2] = lines_final

                if "con1_" in source_file and not con1_file:
                    con1_file = file_path

            # 處理 isub_igate (若找到 con1_ 檔案)
            if con1_file:
                self.log(f"發現包含 con1_ 的檔案: {os.path.basename(con1_file)}，正在產生 isub_igate...")
                xls_con1 = pd.ExcelFile(con1_file)
                con1_lower_map = {s.lower(): s for s in xls_con1.sheet_names}
                
                if "ibig_vg" in con1_lower_map and "model team fitting data" in con1_lower_map:
                    df_fitting_c1 = pd.read_excel(con1_file, sheet_name=con1_lower_map["model team fitting data"], header=None)
                    df_ibig = pd.read_excel(con1_file, sheet_name=con1_lower_map["ibig_vg"], header=None)
                    lines_isub = self.isub_igate_iv(df_fitting_c1, df_ibig)
                    self.trans_results["isub_igate"] = lines_isub

            self.log("=== HCI IV DATA 轉檔完成！ ===")
            messagebox.showinfo("成功", "HCI IV 轉檔程序已全部完成！")

        except Exception as e:
            self.log(f"轉檔過程發生錯誤: {str(e)}")
            messagebox.showerror("錯誤", f"轉檔失敗:\n{str(e)}")

    def hci_IV_t0(self, df_fitting, df_idvg, df_idvd):
        # 參數讀取 (對應 VBA Excel row - 1)
        w = self.format_val(df_fitting.iloc[8, 1])
        l_val = self.format_val(df_fitting.iloc[9, 1])
        t_val = self.format_val(df_fitting.iloc[7, 1])
        vdlin = self.format_val(df_fitting.iloc[13, 1])
        vdd = self.format_val(df_fitting.iloc[12, 1])
        vdstr = self.format_val(df_fitting.iloc[3, 1])
        vgstr = self.format_val(df_fitting.iloc[2, 1])
        vbstr = self.format_val(df_fitting.iloc[5, 1])
        time_str = "0"

        mos_type = "NMOS" if float(df_fitting.iloc[2, 1]) > 0 else "PMOS"

        a = f"Instance{{W={w},L={l_val},T={t_val}}}"
        b = f"DataType{{{mos_type}}}"
        c = f"[vgs,Id(vbs=0.00000),vds={vdlin},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time={time_str}]"
        d = f"[vgs,Id(vbs=0.00000),vds={vdd},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time={time_str}]"
        e = f"[vds,Id(vgs={vdd}),vbs=0.00000,ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr=0,Time={time_str}]"

        # 對應 VBA: G1 = Cells(15,2).End(xlDown).Row -> num = G1 - 15 (Python index 14 為 Excel 第 15 列)
        # 尋找 IdVg_raw 第 2 欄有資料的最末列
        g1 = 14
        while g1 + 1 < len(df_idvg) and pd.notna(df_idvg.iloc[g1 + 1, 1]):
            g1 += 1
        num = g1 - 14

        # 對應 VBA: G2 = Cells(18,2).End(xlDown).Row -> num2 = G2 - 18 (Python index 17 為 Excel 第 18 列)
        g2 = 17
        while g2 + 1 < len(df_idvd) and pd.notna(df_idvd.iloc[g2 + 1, 1]):
            g2 += 1
        num2 = g2 - 17

        lines = [""] * (26 + 2 * num + num2)
        lines[0] = "//Copyright (C) 2010-2020 Primarius Technologies Co., Ltd"
        lines[1] = "// All rights reserved."
        lines[3] = "ObjInfo{}"
        lines[4] = "ProbeInfo{}"
        lines[5] = "Version{4.1}"
        lines[6] = "ModelType{DC}"
        lines[7] = b
        lines[8] = "Delimiter{,}"
        lines[9] = "WorkingMode{Forward}"
        lines[10] = a
        lines[11] = "Node{D|7|0.000000,G|8|0.000000,S|6|0.000000,B|5|0.000000}"
        lines[12] = "MeasurementSetup{PDA}"
        lines[13] = "Input{vgs,vbs,vds}"
        lines[14] = "Output{Id,Isub}"
        lines[15] = "{MOS_Id_Vgs_Vbs_DC}"
        lines[16] = c

        lines[18 + num] = "{MOS_Id_Vgs_Vbs_DC}"
        lines[19 + num] = d
        lines[23 + 2 * num] = "{MOS_Id_Vds_Vgs_DC}"
        lines[24 + 2 * num] = e

        for i in range(num + 1):
            row_idx = 14 + i
            if row_idx < len(df_idvg):
                v1 = self.format_val(df_idvg.iloc[row_idx, 0])
                v2 = self.format_val(df_idvg.iloc[row_idx, 1])
                v4 = self.format_val(df_idvg.iloc[row_idx, 3])
                lines[17 + i] = f"{v1},{v2}"
                lines[20 + num + i] = f"{v1},{v4}"

        for j in range(num2 + 1):
            row_idx = 17 + j
            if row_idx < len(df_idvd):
                v1 = self.format_val(df_idvd.iloc[row_idx, 0])
                v2 = self.format_val(df_idvd.iloc[row_idx, 1])
                lines[25 + 2 * num + j] = f"{v1},{v2}"

        return [line for line in lines if line is not None]

    def hci_iv_final_stress(self, df_fitting, df_idvg, df_idvd, df_param):
        w = self.format_val(df_fitting.iloc[8, 1])
        l_val = self.format_val(df_fitting.iloc[9, 1])
        t_val = self.format_val(df_fitting.iloc[7, 1])
        vdlin = self.format_val(df_fitting.iloc[13, 1])
        vdd = self.format_val(df_fitting.iloc[12, 1])
        vdstr = self.format_val(df_fitting.iloc[3, 1])
        vgstr = self.format_val(df_fitting.iloc[2, 1])
        vbstr = self.format_val(df_fitting.iloc[5, 1])

        # 尋找 Parameter sheet 中的 Tstress
        tstress_row = None
        for r_idx, row in df_param.iterrows():
            for c_idx, val in enumerate(row):
                if val and str(val).strip().lower() == "tstress":
                    tstress_row = r_idx
                    break
            if tstress_row is not None:
                break

        if tstress_row is None:
            raise ValueError("Parameter sheet 找不到 'Tstress'")

        g3 = tstress_row
        while g3 + 1 < len(df_param) and pd.notna(df_param.iloc[g3 + 1, 0]):
            g3 += 1
        
        raw_time = float(df_param.iloc[g3, 0])
        time_str = self.format_val(raw_time / 60)

        mos_type = "NMOS" if float(df_fitting.iloc[2, 1]) > 0 else "PMOS"

        a = f"Instance{{W={w},L={l_val},T={t_val}}}"
        b = f"DataType{{{mos_type}}}"
        c = f"[vgs,Id(vbs=0.00000),vds={vdlin},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time={time_str}]"
        d = f"[vgs,Id(vbs=0.00000),vds={vdd},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time={time_str}]"
        e = f"[vds,Id(vgs={vdd}),vbs=0.00000,ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr=0,Time={time_str}]"

        # 對應 VBA 邏輯計算區段長度
        g1 = 14
        while g1 + 1 < len(df_idvg) and pd.notna(df_idvg.iloc[g1 + 1, 1]):
            g1 += 1
        num = g1 - 14

        g2 = 17
        while g2 + 1 < len(df_idvd) and pd.notna(df_idvd.iloc[g2 + 1, 1]):
            g2 += 1
        num2 = g2 - 17

        # G4 = Cells(1,1).End(xlDown).Row -> num3 = G4 - num
        g4 = 0
        while g4 + 1 < len(df_idvg) and pd.notna(df_idvg.iloc[g4 + 1, 0]):
            g4 += 1
        num3 = max(0, g4 - num)

        g5 = 0
        while g5 + 1 < len(df_idvd) and pd.notna(df_idvd.iloc[g5 + 1, 0]):
            g5 += 1
        num4 = max(0, g5 - num2)

        lines = [""] * (26 + 2 * num + num2)
        lines[0] = "//Copyright (C) 2010-2020 Primarius Technologies Co., Ltd"
        lines[1] = "// All rights reserved."
        lines[3] = "ObjInfo{}"
        lines[4] = "ProbeInfo{}"
        lines[5] = "Version{4.1}"
        lines[6] = "ModelType{DC}"
        lines[7] = b
        lines[8] = "Delimiter{,}"
        lines[9] = "WorkingMode{Forward}"
        lines[10] = a
        lines[11] = "Node{D|7|0.000000,G|8|0.000000,S|6|0.000000,B|5|0.000000}"
        lines[12] = "MeasurementSetup{PDA}"
        lines[13] = "Input{vgs,vbs,vds}"
        lines[14] = "Output{Id,Isub}"
        lines[15] = "{MOS_Id_Vgs_Vbs_DC}"
        lines[16] = c

        lines[18 + num] = "{MOS_Id_Vgs_Vbs_DC}"
        lines[19 + num] = d
        lines[23 + 2 * num] = "{MOS_Id_Vds_Vgs_DC}"
        lines[24 + 2 * num] = e

        for i in range(num + 1):
            target_idx = num3 + i
            if target_idx < len(df_idvg):
                v1 = self.format_val(df_idvg.iloc[target_idx, 0])
                v2 = self.format_val(df_idvg.iloc[target_idx, 1])
                v4 = self.format_val(df_idvg.iloc[target_idx, 3])
                lines[17 + i] = f"{v1},{v2}"
                lines[20 + num + i] = f"{v1},{v4}"

        for j in range(num2 + 1):
            target_idx = num4 + j
            if target_idx < len(df_idvd):
                v1 = self.format_val(df_idvd.iloc[target_idx, 0])
                v2 = self.format_val(df_idvd.iloc[target_idx, 1])
                lines[25 + 2 * num + j] = f"{v1},{v2}"

        return [line for line in lines if line is not None]

    def isub_igate_iv(self, df_fitting, df_ibig):
        w = self.format_val(df_fitting.iloc[8, 1])
        l_val = self.format_val(df_fitting.iloc[9, 1])
        t_val = self.format_val(df_fitting.iloc[7, 1])
        vdd = self.format_val(df_ibig.iloc[2, 1])
        mos_type = "NMOS" if float(df_fitting.iloc[2, 1]) > 0 else "PMOS"

        a = f"Instance{{W={w},L={l_val},T={t_val}}}"
        b = f"DataType{{{mos_type}}}"

        # G1 = Cells(33, 1).End(xlDown).Row -> vg_num = G1 - 34 (Python index 33 為 Excel 第 34 列)
        g1 = 33
        while g1 + 1 < len(df_ibig) and pd.notna(df_ibig.iloc[g1 + 1, 0]):
            g1 += 1
        vg_num = g1 - 33

        lines = [""] * (22 + vg_num * 2)
        lines[0] = "//Copyright (C) 2010-2020 Primarius Technologies Co., Ltd"
        lines[1] = "// All rights reserved."
        lines[3] = "ObjInfo{}"
        lines[4] = "ProbeInfo{}"
        lines[5] = "Version{4.1}"
        lines[6] = "ModelType{DC}"
        lines[7] = b
        lines[8] = "Delimiter{,}"
        lines[9] = "WorkingMode{Forward}"
        lines[10] = a
        lines[11] = "Node{D|7|0.000000,G|8|0.000000,S|6|0.000000,B|5|0.000000}"
        lines[12] = "MeasurementSetup{PDA}"
        lines[13] = "Input{vgs,vbs,vds}"
        lines[14] = "Output{Id,Isub}"
        lines[15] = "{MOS_Isub_Vgs_Vds_DC}"
        lines[16] = f"[vgs,Isub(vds={vdd}),vbs=0.00000,ref_vs=0.00000]"
        lines[18 + vg_num] = "{Igate_Vgs_Vds_DC}"
        lines[19 + vg_num] = f"[vgs,Ig(vds={vdd}),vbs=0.00000,ref_vs=0.00000]"

        for i in range(vg_num + 1):
            row_idx = 33 + i
            if row_idx < len(df_ibig):
                v1 = self.format_val(df_ibig.iloc[row_idx, 0])
                v5 = self.format_val(df_ibig.iloc[row_idx, 4])
                v3 = self.format_val(df_ibig.iloc[row_idx, 2])
                lines[17 + i] = f"{v1},{v5}"
                lines[20 + vg_num + i] = f"{v1},{v3}"

        return [line for line in lines if line is not None]

    def run_save_dat(self):
        """執行 SAVE IV DATA 按鈕邏輯"""
        target_dir = self.path_entry.get().strip()
        if not target_dir or not os.path.isdir(target_dir):
            messagebox.showerror("錯誤", "請先選擇有效的資料夾路徑！")
            return

        if not self.trans_results:
            messagebox.showwarning("提示", "目前沒有可儲存的轉檔資料，請先執行 'HCI IV DATA' 轉檔！")
            return

        self.log("=== 開始儲存 .dat 檔案 ===")
        results_folder = os.path.join(target_dir, "after_trans")
        os.makedirs(results_folder, exist_ok=True)

        saved_count = 0
        for sheet_name, lines in self.trans_results.items():
            file_path = os.path.join(results_folder, f"{sheet_name}.dat")
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    for line in lines:
                        f.write(str(line) + "\n")
                self.log(f"已儲存: {sheet_name}.dat")
                saved_count += 1
            except Exception as e:
                self.log(f"儲存 {sheet_name}.dat 失敗: {str(e)}")

        self.log(f"=== 儲存完成！共輸出 {saved_count} 個檔案至 after_trans 資料夾 ===")
        messagebox.showinfo("成功", f"成功儲存 {saved_count} 個 .dat 檔案到 after_trans 目錄！")

if __name__ == "__main__":
    root = tk.Tk()
    app = IVTransApp(root)
    root.mainloop()
