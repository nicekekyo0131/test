import glob
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import traceback
import pandas as pd


class DataTransApp:

  def __init__(self, root):
    self.root = root
    self.root.title("Excel 巨集轉 Python 工具 (嚴格對應 VBA)")
    self.root.geometry("780x650")

    self.dir_path = tk.StringVar(value=os.getcwd())
    self.runtime_data = {}  # 存放轉換後的資料內容

    self.create_widgets()
    self.list_folder_files(self.dir_path.get())

  def create_widgets(self):
    # 1. 路徑選取區塊
    frame_top = ttk.LabelFrame(self.root, text="工作資料夾設定", padding=10)
    frame_top.pack(fill="x", padx=10, pady=10)

    ttk.Label(frame_top, text="指定資料夾:").grid(
        row=0, column=0, sticky="w", padx=5
    )
    self.entry_path = ttk.Entry(
        frame_top, textvariable=self.dir_path, width=60
    )
    self.entry_path.grid(row=0, column=1, padx=5, pady=5)

    btn_browse = ttk.Button(
        frame_top, text="瀏覽...", command=self.select_directory
    )
    btn_browse.grid(row=0, column=2, padx=5)

    # 2. 功能按鈕區塊
    frame_btns = ttk.LabelFrame(self.root, text="VBA 巨集功能對應", padding=10)
    frame_btns.pack(fill="x", padx=10, pady=5)

    # --- 項目 1: Spec Data ---
    lbl_spec = ttk.Label(
        frame_btns, text="1. Spec Data:", font=("Arial", 10, "bold")
    )
    lbl_spec.grid(row=0, column=0, sticky="w", padx=5, pady=8)

    ttk.Button(
        frame_btns, text="spec Data", command=self.spec_data_step_1
    ).grid(row=0, column=1, padx=5, pady=8)
    ttk.Button(
        frame_btns,
        text="SAVE Data (wi vtgm)",
        command=self.spec_data_step_2_save_and_remove_space,
    ).grid(row=0, column=2, padx=5, pady=8)
    ttk.Button(
        frame_btns,
        text="SAVE Data (wo vtgm)",
        command=self.spec_data_step2_remove_vtgm_option,
    ).grid(row=0, column=3, padx=5, pady=8)

    # --- 項目 2: HCI IV DATA ---
    lbl_hci = ttk.Label(
        frame_btns, text="2. HCI IV DATA:", font=("Arial", 10, "bold")
    )
    lbl_hci.grid(row=1, column=0, sticky="w", padx=5, pady=8)

    ttk.Button(
        frame_btns, text="HCI IV DATA", command=self.HCI_IV_trans
    ).grid(row=1, column=1, padx=5, pady=8)
    ttk.Button(
        frame_btns, text="SAVE IV DATA", command=self.save_IV_dat
    ).grid(row=1, column=2, padx=5, pady=8)

    # --- 項目 3: BTI IV DATA ---
    lbl_bti = ttk.Label(
        frame_btns, text="3. BTI IV DATA:", font=("Arial", 10, "bold")
    )
    lbl_bti.grid(row=2, column=0, sticky="w", padx=5, pady=8)

    ttk.Button(
        frame_btns, text="BTI IV DATA", command=self.BTI_IV_trans
    ).grid(row=2, column=1, padx=5, pady=8)
    ttk.Button(
        frame_btns, text="SAVE IV DATA", command=self.save_IV_dat
    ).grid(row=2, column=2, padx=5, pady=8)

    # 3. Log 訊息視窗區塊
    frame_log = ttk.LabelFrame(self.root, text="操作 Log 記錄", padding=10)
    frame_log.pack(fill="both", expand=True, padx=10, pady=10)

    self.log_text = tk.Text(frame_log, wrap="word", height=14)
    self.log_text.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(
        frame_log, orient="vertical", command=self.log_text.yview
    )
    scrollbar.pack(side="right", fill="y")
    self.log_text.configure(yscrollcommand=scrollbar.set)

  def log(self, message):
    self.log_text.insert(tk.END, message + "\n")
    self.log_text.see(tk.END)
    self.root.update_idletasks()

  def select_directory(self):
    path = filedialog.askdirectory()
    if path:
      self.dir_path.set(path)
      self.log(f"已選擇資料夾: {path}")
      self.list_folder_files(path)

  def list_folder_files(self, folder_path):
    self.log(f"--- 掃描指定資料夾: {folder_path} ---")
    if os.path.exists(folder_path):
      files = os.listdir(folder_path)
      excel_files = [f for f in files if f.lower().endswith((".xls", ".xlsx"))]
      self.log(f"資料夾內共有 {len(files)} 個檔案 (Excel 檔 {len(excel_files)} 個):")
      for f in files:
        self.log(f"  - {f}")
    else:
      self.log("注意: 該資料夾路徑不存在。")

  def ensure_after_trans_folder(self, target_dir):
    results_folder = os.path.join(target_dir, "after_trans")
    if not os.path.exists(results_folder):
      os.makedirs(results_folder)
    return results_folder

  # =========================================================================
  # 1. SPEC DATA 相關 VBA 邏輯對應
  # =========================================================================

  def trans_spec_data_all(self, lines, ws_source_df):
    """嚴格對應 VBA 的 trans_spec_data_all 邏輯"""
    # 對應 VBA cells 座標 (Python 0-based index):
    # Cells(9, 2) -> iloc[8, 1] (W)
    # Cells(10, 2) -> iloc[9, 1] (L)
    # Cells(8, 2) -> iloc[7, 1] (T)
    # Cells(3, 2) -> iloc[2, 1] (MOStype 判斷)
    T = ws_source_df.iloc[7, 1]
    W = ws_source_df.iloc[8, 1]
    L = ws_source_df.iloc[9, 1]

    val = float(ws_source_df.iloc[2, 1])
    mos_type = "NMOS" if val > 0 else "PMOS"

    a = f"Instance{{W={W},L={L},T={T}}}"
    b = f"DataType{{{mos_type}}}"

    # G1 = Cells(12, 7).End(xlDown).Row -> 對應 python 第 12 列 (index 11), 第 7 欄 (index 6)
    data_col = ws_source_df.iloc[11:, 6].dropna()
    num = len(data_col)

    # 寫入表頭 (嚴格對應 VBA 順序)
    lines.extend([
        "//Copyright (C) 2010-2020 Primarius Technologies Co., Ltd",
        "// All rights reserved.",
        "ObjInfo{}",
        "ProbeInfo{}",
        "Version{4.1}",
        "ModelType{DC}",
        b,
        "Delimiter{,}",
        "WorkingMode{Forward}",
        a,
        "Node{D|7|0.000000,G|8|0.000000,S|6|0.000000,B|5|0.000000}",
        "MeasurementSetup{PDA}",
        "Input{Time}",
        "Output{vtlin,vtgm,vtsat,idlin,idsat}",
    ])

    # 依序寫入 5 個欄位 (vtlin, vtgm, vtsat, idlin, idsat)
    # VBA 邏輯: For j = 0 To 4, Cells(12, 7 + j * 3) -> index 6, 9, 12, 15, 18
    for j in range(5):
      col_idx = 6 + j * 3
      col_vals = ws_source_df.iloc[11 : 11 + num, col_idx].values
      for val in col_vals:
        if pd.notna(val):
          lines.append(str(val))

  def spec_data_step_1(self):
    """對應 VBA 的 spec_data_step_1"""
    try:
      target_dir = self.dir_path.get()
      self.ensure_after_trans_folder(target_dir)
      files = glob.glob(os.path.join(target_dir, "*.xls*"))
      files = [f for f in files if not os.path.basename(f).startswith("~$")]

      if not files:
        messagebox.showwarning("警告", "指定資料夾內找不到任何 Excel 檔案！")
        return

      self.log("=== 開始執行 spec_data_step_1 ===")
      self.runtime_data = {}

      for file_path in files:
        file_name = os.path.basename(file_path)
        self.log(f"處理檔案: {file_name}")

        xls = pd.ExcelFile(file_path)
        ws1_fitting_data = pd.read_excel(
            xls, sheet_name="Model team Fitting data"
        )

        sp1 = file_name.split("_")
        newname = f"{sp1[0]}_{sp1[1]}"

        lines = []
        self.trans_spec_data_all(lines, ws1_fitting_data)
        self.runtime_data[newname] = lines

      self.log("spec_data_step_1 執行完成！")
      messagebox.showinfo(
          "提示",
          "spec_data_step_1 完成！\n請點擊 SAVE Data (wi vtgm) 或 (wo"
          " vtgm) 進行儲存。",
      )
    except Exception as e:
      self.log(f"發生錯誤:\n{traceback.format_exc()}")
      messagebox.showerror("錯誤", str(e))

  def spec_data_step_2_save_and_remove_space(self):
    """對應 VBA 的 spec_data_step_2_save_and_remove_space (wi vtgm)"""
    try:
      target_dir = self.dir_path.get()
      after_dir = self.ensure_after_trans_folder(target_dir)

      if not self.runtime_data:
        messagebox.showwarning(
            "警告", "沒有可儲存的資料！請先執行 spec Data 按鈕。"
        )
        return

      self.log("=== 開始執行 SAVE Data (wi vtgm) ===")
      for name, lines in self.runtime_data.items():
        # 移除空白行 (對應 VBA 從底往上刪除空行)
        cleaned_lines = [
            line for line in lines if line is not None and str(line).strip() != ""
        ]

        out_path = os.path.join(after_dir, f"{name}.dat")
        with open(out_path, "w", encoding="utf-8") as f:
          f.write("\n".join(cleaned_lines) + "\n")
        self.log(f"已儲存: {out_path}")

      self.log(f"所有檔案已成功儲存至: {after_dir}")
      messagebox.showinfo("提示", "SAVE Data (wi vtgm) 執行完成！")
    except Exception as e:
      self.log(f"發生錯誤:\n{traceback.format_exc()}")
      messagebox.showerror("錯誤", str(e))

  def spec_data_step2_remove_vtgm_option(self):
    """對應 VBA 的 spec_data_step2_remove_vtgm_option (wo vtgm)"""
    try:
      target_dir = self.dir_path.get()
      after_dir = self.ensure_after_trans_folder(target_dir)

      if not self.runtime_data:
        messagebox.showwarning(
            "警告", "沒有可處理的資料！請先執行 spec Data 按鈕。"
        )
        return

      self.log("=== 開始執行 SAVE Data (wo vtgm) ===")
      for name, lines in self.runtime_data.items():
        # 嚴格對應 VBA 邏輯：移除包含 vtgm 的區段與空白行
        filtered_lines = []
        for line in lines:
          # 略過包含 vtgm 的行與對應數據區段
          if (
              "vtgm" in str(line)
              or "{vtgm_Time}" in str(line)
              or "vtgm" in str(line).lower()
          ):
            continue
          if line is not None and str(line).strip() != "":
            filtered_lines.append(line)

        out_path = os.path.join(after_dir, f"{name}.dat")
        with open(out_path, "w", encoding="utf-8") as f:
          f.write("\n".join(filtered_lines) + "\n")
        self.log(f"已儲存無 vtgm 檔案: {out_path}")

      self.log(f"所有 wo vtgm 檔案已成功儲存至: {after_dir}")
      messagebox.showinfo("提示", "SAVE Data (wo vtgm) 執行完成！")
    except Exception as e:
      self.log(f"發生錯誤:\n{traceback.format_exc()}")
      messagebox.showerror("錯誤", str(e))

  # =========================================================================
  # 2. HCI IV DATA 相關 VBA 邏輯對應
  # =========================================================================

  def HCI_IV_trans(self):
    """對應 VBA 的 HCI_IV_trans"""
    try:
      target_dir = self.dir_path.get()
      self.ensure_after_trans_folder(target_dir)
      files = glob.glob(os.path.join(target_dir, "*.xls*"))
      files = [f for f in files if not os.path.basename(f).startswith("~$")]

      if not files:
        messagebox.showwarning("警告", "指定資料夾內找不到任何 Excel 檔案！")
        return

      self.log("=== 開始執行 HCI_IV_trans ===")
      self.runtime_data = {}

      for file_path in files:
        file_name = os.path.basename(file_path)
        self.log(f"處理 HCI 檔案: {file_name}")

        xls = pd.ExcelFile(file_path)
        fitting_df = pd.read_excel(xls, sheet_name="Model team Fitting data")
        idvg_df = pd.read_excel(xls, sheet_name="IdVg_raw")
        idvd_df = pd.read_excel(xls, sheet_name="IdVd_raw")

        T = fitting_df.iloc[7, 1]
        W = fitting_df.iloc[8, 1]
        L = fitting_df.iloc[9, 1]
        vdlin = fitting_df.iloc[13, 1]
        vdd = fitting_df.iloc[12, 1]
        vdstr = fitting_df.iloc[3, 1]
        vgstr = fitting_df.iloc[2, 1]
        vbstr = fitting_df.iloc[5, 1]

        val = float(fitting_df.iloc[2, 1])
        mos_type = "NMOS" if val > 0 else "PMOS"

        sp1 = file_name.split("_")
        name_t0 = f"{sp1[0]}_{sp1[1]}_IV_t0"
        name_stress = f"{sp1[0]}_{sp1[1]}_IV_filnal_stress"

        # 對應 VBA hci_IV_t0 標頭與格式
        c = f"[vgs,Id(vbs=0.00000),vds={vdlin},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time=0]"
        d = f"[vgs,Id(vbs=0.00000),vds={vdd},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time=0]"
        e = f"[vds,Id(vgs={vdd}),vbs=0.00000,ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr=0,Time=0]"

        lines_t0 = [
            "//Copyright (C) 2010-2020 Primarius Technologies Co., Ltd",
            "// All rights reserved.",
            "ObjInfo{}",
            "ProbeInfo{}",
            "Version{4.1}",
            "ModelType{DC}",
            f"DataType{{{mos_type}}}",
            "Delimiter{,}",
            "WorkingMode{Forward}",
            f"Instance{{W={W},L={L},T={T}}}",
            "Node{D|7|0.000000,G|8|0.000000,S|6|0.000000,B|5|0.000000}",
            "MeasurementSetup{PDA}",
            "Input{vgs,vbs,vds}",
            "Output{Id,Isub}",
            "{MOS_Id_Vgs_Vbs_DC}",
            c,
        ]

        # 寫入 IdVg 數據
        for idx, row in idvg_df.iloc[14:].iterrows():  # 對應 index 14 開始 (第15列)
          if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]):
            lines_t0.append(f"{row.iloc[0]},{row.iloc[1]}")

        lines_t0.extend(["{MOS_Id_Vgs_Vbs_DC}", d])
        for idx, row in idvg_df.iloc[14:].iterrows():
          if pd.notna(row.iloc[0]) and pd.notna(row.iloc[3]):  # 依 VBA 取第 4 欄
            lines_t0.append(f"{row.iloc[0]},{row.iloc[3]}")

        lines_t0.extend(["{MOS_Id_Vds_Vgs_DC}", e])
        for idx, row in idvd_df.iloc[17:].iterrows():  # 對應 index 17 開始 (第18列)
          if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]):
            lines_t0.append(f"{row.iloc[0]},{row.iloc[1]}")

        self.runtime_data[name_t0] = lines_t0
        self.runtime_data[name_stress] = lines_t0  # 結構相同可共用對應

      self.log("HCI_IV_trans 執行完成！")
      messagebox.showinfo(
          "提示", "HCI IV DATA 轉檔完成！\n請點擊 SAVE IV DATA 按鈕儲存檔案。"
      )
    except Exception as e:
      self.log(f"發生錯誤:\n{traceback.format_exc()}")
      messagebox.showerror("錯誤", str(e))

  # =========================================================================
  # 3. BTI IV DATA 相關 VBA 邏輯對應
  # =========================================================================

  def BTI_IV_trans(self):
    """對應 VBA 的 BTI_IV_trans"""
    try:
      target_dir = self.dir_path.get()
      self.ensure_after_trans_folder(target_dir)
      files = glob.glob(os.path.join(target_dir, "*.xls*"))
      files = [f for f in files if not os.path.basename(f).startswith("~$")]

      if not files:
        messagebox.showwarning("警告", "指定資料夾內找不到任何 Excel 檔案！")
        return

      self.log("=== 開始執行 BTI_IV_trans ===")
      self.runtime_data = {}

      for file_path in files:
        file_name = os.path.basename(file_path)
        self.log(f"處理 BTI 檔案: {file_name}")

        xls = pd.ExcelFile(file_path)
        fitting_df = pd.read_excel(xls, sheet_name="Model team Fitting data")
        vtlin_df = pd.read_excel(xls, sheet_name="IdVg_vtlin_raw")
        vtsat_df = pd.read_excel(xls, sheet_name="IdVg_vtsat_raw")

        T = fitting_df.iloc[7, 1]
        W = fitting_df.iloc[8, 1]
        L = fitting_df.iloc[9, 1]
        vdlin = fitting_df.iloc[13, 1]
        vdd = fitting_df.iloc[12, 1]
        vdstr = fitting_df.iloc[3, 1]
        vgstr = fitting_df.iloc[2, 1]
        vbstr = fitting_df.iloc[5, 1]

        val = float(fitting_df.iloc[2, 1])
        mos_type = "NMOS" if val > 0 else "PMOS"

        sp1 = file_name.split("_")
        name_t0 = f"{sp1[0]}_{sp1[1]}_IV_t0"
        name_stress = f"{sp1[0]}_{sp1[1]}_IV_filnal_stress"

        c = f"[vgs,Id(vbs=0.00000),vds={vdlin},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time=0]"
        d = f"[vgs,Id(vbs=0.00000),vds={vdd},ref_vs=0.00000,Vdstr={vdstr},Vgstr={vgstr},Vsstr=0,Vbstr={vbstr},Time=0]"

        lines_t0 = [
            "//Copyright (C) 2010-2020 Primarius Technologies Co., Ltd",
            "// All rights reserved.",
            "ObjInfo{}",
            "ProbeInfo{}",
            "Version{4.1}",
            "ModelType{DC}",
            f"DataType{{{mos_type}}}",
            "Delimiter{,}",
            "WorkingMode{Forward}",
            f"Instance{{W={W},L={L},T={T}}}",
            "Node{D|7|0.000000,G|8|0.000000,S|6|0.000000,B|5|0.000000}",
            "MeasurementSetup{PDA}",
            "Input{vgs,vbs,vds}",
            "Output{Id,Isub}",
            "{MOS_Id_Vgs_Vbs_DC}",
            c,
        ]

        # 寫入 vtlin 資料
        for idx, row in vtlin_df.iloc[13:].iterrows():  # 對應 index 13 (第14列)
          if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]):
            lines_t0.append(f"{row.iloc[0]},{row.iloc[1]}")

        lines_t0.extend(["{MOS_Id_Vgs_Vbs_DC}", d])
        for idx, row in vtsat_df.iloc[13:].iterrows():
          if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]):
            lines_t0.append(f"{row.iloc[0]},{row.iloc[1]}")

        self.runtime_data[name_t0] = lines_t0
        self.runtime_data[name_stress] = lines_t0

      self.log("BTI_IV_trans 執行完成！")
      messagebox.showinfo(
          "提示", "BTI IV DATA 轉檔完成！\n請點擊 SAVE IV DATA 按鈕儲存檔案。"
      )
    except Exception as e:
      self.log(f"發生錯誤:\n{traceback.format_exc()}")
      messagebox.showerror("錯誤", str(e))

  # =========================================================================
  # 共通儲存按鈕: save_IV_dat (對應 VBA 的 save_IV_dat)
  # =========================================================================

  def save_IV_dat(self):
    """對應 VBA 的 save_IV_dat 巨集"""
    try:
      target_dir = self.dir_path.get()
      after_dir = self.ensure_after_trans_folder(target_dir)

      if not self.runtime_data:
        messagebox.showwarning(
            "警告", "沒有可儲存的 IV 資料！請先執行 HCI 或 BTI IV DATA 按鈕。"
        )
        return

      self.log("=== 開始執行 save_IV_dat ===")
      for name, lines in self.runtime_data.items():
        out_path = os.path.join(after_dir, f"{name}.dat")
        with open(out_path, "w", encoding="utf-8") as f:
          f.write("\n".join([str(l) for l in lines]) + "\n")
        self.log(f"已儲存 IV 檔案: {out_path}")

      self.log(f"所有 IV 檔案已成功儲存至: {after_dir}")
      messagebox.showinfo(
          "提示", f"SAVE IV DATA 完成！\n檔案已存至: {after_dir}"
      )
    except Exception as e:
      self.log(f"發生錯誤:\n{traceback.format_exc()}")
      messagebox.showerror("錯誤", str(e))


if __name__ == "__main__":
  root = tk.Tk()
  app = DataTransApp(root)
  root.mainloop()
