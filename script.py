import tkinter as tk
from tkinter import ttk, messagebox
import time
import xlwings as xw
import logging
import pyvisa
import os

# ====== Log 設定 ======
DEBUG = True
LOG_TO_FILE = True
VERIFY_WRITE = False  # True: 每次寫入後嘗試 *OPC? 或 SYST:ERR? 檢查
SCPI_USE_ALT = True  # True 時，若主指令失敗會嘗試通用寫法
ex_content = []
if LOG_TO_FILE:
    logging.basicConfig(
        filename="app.log",
        filemode="a",
        level=logging.DEBUG if DEBUG else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",   # 時間,等級,訊息

    )
else:
    logging.basicConfig(
        level=logging.DEBUG if DEBUG else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )


def open_excel_wb(filename="執行檔.xlsx",
                  visible=False,
                  add_book=False,
                  max_sheets=11,
                  show_error_dialog=True):
    """
    開啟指定 Excel 檔並回傳 (xw_app, workbook, sheets, err_text)
    不使用型別註解與任何新語法，最保守的相容版本。
    """
    try:
        import xlwings as xw
        from tkinter import messagebox
    except Exception as e:
        _log("error", "[Excel] 載入套件失敗：%s" % e)
        return None, None, [], "載入 xlwings 或 tkinter 失敗：%s" % e

    app = None
    try:
        app = xw.App(visible=visible, add_book=add_book)
        try:
            app.display_alerts = False
            app.screen_updating = False
        except Exception:
            pass

        wb = app.books.open(filename)
        total = len(wb.sheets)
        count = min(max_sheets, total)
        sheets = [wb.sheets[i] for i in range(count)]

        _log("info", "[Excel] 已開啟檔案：%s，工作表數: %d（取前 %d 張）"
             % (os.path.abspath(filename), total, len(sheets)))
        if sheets:
            _log("debug", "[Excel] 第一張表名：%s" % sheets[0].name)

        return app, wb, sheets, None

    except Exception as e:
        _log("error", "[Excel] 開啟或列舉工作表失敗：%s" % e)
        if show_error_dialog:
            try:
                from tkinter import messagebox
                messagebox.showerror("Excel 讀取錯誤",
                                     "%s\n\n檔案：%s" % (e, os.path.abspath(filename)))
            except Exception:
                pass
        try:
            if app is not None:
                app.quit()
        except Exception:
            pass
        return None, None, [], str(e)

def _log(level, msg):
    """統一印出到 console + logger"""
    ts = time.strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    getattr(logging, level)(msg)
#log

def _to_num(v):
    """可將 Excel 讀到的數值(含 3.0 字串/浮點)轉為 int/float；失敗則原值返回"""
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except Exception:
        return v


def generate_prog_data(sheet):
    """從指定 sheet 讀 B/C 欄，產生 PROG:DATA 指令串列"""
    global ex_content
    ex_content.clear()
    result = []       #回傳的 SCPI 指令字串
    prefix ="PROG:DATA "
    try:
        sheet_name = sheet.name
        _log("info", f"[Excel] 開始讀取工作表：{sheet_name}") #讀取工作表名稱

        # 以 B 欄最後使用列為準（避免 B1 是標題導致 .end('down') 撞牆）
        last_row = sheet.range("B" + str(sheet.cells.last_cell.row)).end("up").row
        _log("debug", f"[Excel] B 欄最後列 = {last_row}")

        if last_row < 2:
            _log("warning", "[Excel] 找不到資料（B2 以下無內容）")
            return result

        arr = sheet.range(f"B2:C{last_row}").value  # [[B2,C2], [B3,C3], ...]
        if not isinstance(arr, list):
            arr = [arr]

        _log("info", f"[Excel] 讀到筆數: {len(arr)}")
        print(ex_content)
        _log("info", f"[Excel] 讀到筆數: {ex_content}")

        for row in arr:   #b2下只有一欄
            if not row or len(row) < 2:
                continue
            b_raw, c_raw = row[0], row[1]
            if b_raw is None or c_raw is None:
                continue
            #將 excel 讀到的值盡可能轉成數字（3.0 -> 3；其餘維持原樣）
            b_val = _to_num(b_raw)
            c_val = _to_num(c_raw)
            ex_content.append(c_val)

            new_str = prefix + f"{b_val},LIST,{c_val},0,{c_val}"
            result.append(new_str)

        _log("info", f"[Excel] 產生指令數: {len(result)}")
        if result:
            _log("debug", f"[Excel] 首條指令: {result[0]}")
            _log("debug", f"[Excel] 末條指令: {result[-1]}")

    except Exception as e:
        _log("error", f"[Excel 讀取錯誤] {e}")
    #回傳完整的指令清單
    return result
def generate_prog_data2(sheet, row_count=None):
    # 從指定的Excel (xlwings 的 sheet 物件) 讀取 B~F 欄位的參數，
    # 由 B2 開始往下，共 row_count 列，並組成儀器可用的`PROG:DATA:LIST ...` SCPI 指令字串清單後回傳。
    #  row_count  : 從 B2 開始要讀取的列數。
     #result     : list[str]，每個元素是一條完整的 "PROG:DATA:LIST ..." 指令
    prefix = "PROG:DATA:LIST "
    result = []

    # 自動偵測列數（從 B2 向下到最後一筆連續資料）
    if row_count is None:
        last_row = sheet.cells(2, "B").end("down").row  # 回到最後一筆
        row_count = max(0, last_row - 1)  # 扣掉表頭 B1

    if row_count <= 0:
        return result

    # === 一次把要用到的區塊抓出來 ===
    # 從 B2 起，往下 row_count 列，橫向讀 5 欄（B~F）
    # 取得的 arr 會是 list of lists，例如:
    #   [[B2, C2, D2, E2, F2], [B3, C3, D3, E3, F3], ...]
    arr = sheet.range("B2").resize(row_count, 5).value

    for row in arr:
        # row = [B, C, D, E, F]
        if not row or any(v is None for v in row):
            continue

        vals = []
        for v in row:
            if isinstance(v, (int, float)) and float(v).is_integer():
                vals.append(str(int(v)))
            else:
                vals.append(str(v).strip())

        # 組字串
        new_str = (f"{prefix} {vals[0]},{vals[1]},AUTO,CC,2,"
                   f"{vals[2]},{vals[3]},{vals[3]},{vals[4]},-1,-1,-1,-1,-1,-1,1")
        result.append(new_str)

    return result

def try_init_rm():
    """依序嘗試 NI-VISA 與 pyvisa-py，成功即回傳 (rm, backend_tag, errs_text)"""
    # 先嘗試載入 pyvisa；若失敗，整個流程無法進行，直接回傳錯誤文字
    # try:
    #     import pyvisa
    #     _log("info", f"[VISA] pyvisa 版本: {getattr(pyvisa, '__version__', 'unknown')}")
    # except Exception as e:
    #     # rm 與 backend_tag 都無法提供，因此回傳 (None, None, 錯誤訊息)
    #     return None, None, f"未安裝 pyvisa：{e}"

    # 預設後端嘗試順序：
    #   1) "@ni"：NI-VISA 後端（需安裝 NI-VISA Runtime）
    #   2) "@py"：pyvisa-py 純 Python 後端（常見於跨平台；需相依 libusb/pyusb 等）
    #   3) None：使用 pyvisa 的預設後端解析（視環境可能仍落到 @ni 或 @py）
    backend_candidates = ["@ni", "@py", None]

    # 用來累積各後端失敗的錯誤訊息（最後統一組合成 errs_text）
    errs = []

    # 逐一嘗試每個後端
    for be in backend_candidates:
        try:
            # 建立 ResourceManager：
            #   - 若 be 有值（"@ni" 或 "@py"），則指名使用該後端
            #   - 若 be 為 None，則使用 pyvisa.ResourceManager() 預設行為
            rm = pyvisa.ResourceManager(be) if be else pyvisa.ResourceManager()

            # 以 list_resources() 做為「此後端可正常運作」的驗證
            # 若此步驟丟出例外，代表該後端不可用或環境未備妥
            res = rm.list_resources()

            # 將 None（預設）後端標記為 "default"，其餘沿用實際字串
            tag = be or "default"

            # 紀錄可用後端與當前可見的資源清單（方便診斷）
            _log("info", f"[VISA] 後端可用: {tag}，目前資源: {res}")

            # 走到這裡表示此後端可用，直接回傳成功三元組
            return rm, tag, None

        except Exception as e:
            # 此後端嘗試失敗：把錯誤訊息加入 errs，並以 warning 級別記錄
            msg = f"{be or 'default'}: {e}"
            _log("warning", f"[VISA] 後端失敗 => {msg}")
            errs.append(msg)

    # 三個後端都失敗時，將累積的錯誤訊息用 " / " 串接回傳
    return None, None, " / ".join(errs)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("LOAD測試")
        self.root.geometry("680x350")

        self.rm = None
        self.inst = None
        self.backend = None

        # Excel/工作簿保存為屬性，並在關閉視窗時釋放
        # try:
        #     self.xw_app = xw.App(visible=False, add_book=False)
        #     self.wb = self.xw_app.books.open("執行檔.xlsx")
        #     self.sheets = [self.wb.sheets[i] for i in range(min(11, len(self.wb.sheets)))]
        #     _log("info", f"[Excel] 已開啟檔案，工作表數: {len(self.sheets)}")
        #     if self.sheets:
        #         _log("debug", f"[Excel] 第一張表名：{self.sheets[0].name}")
        # except Exception as e:
        #     _log("error", f"[Excel] 開啟或列舉工作表失敗：{e}")
        #     messagebox.showerror("Excel 讀取錯誤", str(e))
            # 也允許 GUI 照常起來，但 set_load 會因 sheets 不存在而提醒
        # self.xw_app, self.wb, self.sheets, err =open_excel_wb(
        #     filename="執行檔.xlsx",
        #     visible=False,
        #     add_book=False,
        #     max_sheets=11,
        #     show_error_dialog=True  # 想用你原本的 messagebox UX 就留 True
        # )
        self.make_ui()
        self.init_backend_and_scan()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def send_parameter_by_index(self, sheet_idx_1based, count_idx_1based):
        """依指定工作表與 ex_content 的索引，送出 PROG:DATA:LIST 並儲存程式"""
        # 1) 基本檢查
        if self.inst is None:
            self._set_status("尚未連線儀器")
            return
        if not hasattr(self, "sheets") or not self.sheets:
            self._set_status("Excel 尚未載入或無工作表")
            return

       # idx = sheet_idx_1based - 1
       #  if sheet_idx_1based < 0 or sheet_idx_1based >= len(self.sheets):
       #      self._set_status(f"工作表索引超出範圍：{sheet_idx_1based}")
       #      return

        sheet = self.sheets[sheet_idx_1based-1]
        _log("info", f"[send_parameter] 目標工作表：{sheet.name}（index={sheet_idx_1based}）")

        # 2) 取得 row_count：優先用 ex_content[count_idx]，不足時自動偵測
        global ex_content
        if len(ex_content) < count_idx_1based:
            _log("info", "[send_parameter] ex_content 不足，先從該表產生一次 PROG:DATA 以取得資料")
            prog_data_list= generate_prog_data(sheet)  # 這步會填充 ex_content
            for line in prog_data_list:
                cmd = line.strip()
                if cmd:  # 跳過空行
                    self.inst.write(cmd)
                   # print(f"已送出: {cmd}")

        if len(ex_content) >= count_idx_1based:
            row_count = int(_to_num(ex_content[count_idx_1based - 1]))
            _log("info", f"[send_parameter] 來自 ex_content  = {ex_content}")
            _log("info", f"[send_parameter] 來自 ex_content 的 row_count = {row_count}")
        else:
            # 後備：用 B 欄最後一列推算
            last_row = sheet.range("B" + str(sheet.cells.last_cell.row)).end("up").row
            row_count = max(0, last_row - 1)  # 扣掉表頭
            _log("warning", f"[send_parameter] ex_content 仍不足，改用自動列數 row_count={row_count}")

        if row_count <= 0:
            self._set_status("row_count 為 0，沒有可送出的資料")
            return

        # 3) 產生 PROG:DATA:LIST 指令並送出
        set_sheet = self.sheets[sheet_idx_1based]
        prog_data_list = generate_prog_data2(set_sheet,row_count)
        total = len(prog_data_list)
        if total == 0:
            self._set_status("沒有可送出的 PROG:DATA:LIST 指令")
            return

        self._set_status(f"開始載入數據，共 {total} 條")
        ok_count = 0
        for i, line in enumerate(prog_data_list, 1):
            cmd = (line or "").strip()
            if not cmd:
                continue
            if self._visa_write(cmd):
                ok_count += 1
            if i % 20 == 0 or i == total:
                self._set_status(f"載入進度：{i}/{total}")

        # 4) 儲存程式
        self._visa_write("PROG:SAV")
        self._set_status(f"載入完成，成功 {ok_count}/{total} 條，已下達 PROG:SAV")
        _log("info", "[send_parameter] 程式已儲存 (PROG:SAV)")

    def unlock_excel_and_reload(self):
        """解除 Excel 檔案鎖（關閉 wb/xw_app）並重新開啟活頁簿，更新選單"""
        # 1) 關閉活頁簿
        try:
            if getattr(self, "wb", None):
                self.wb.close()
                self.wb = None
                _log("debug", "[Excel] 工作簿已關閉（解除鎖定第一步）")
        except Exception as e:
            _log("warning", f"[Excel] 關閉工作簿例外：{e}")

        # 2) 關閉整個 Excel 應用（確保檔案鎖釋放）
        try:
            if getattr(self, "xw_app", None):
                self.xw_app.quit()
                self.xw_app = None
                _log("debug", "[Excel] Excel 應用已結束（鎖應該釋放）")
        except Exception as e:
            _log("warning", f"[Excel] 關閉 Excel 應用例外：{e}")

    def send_parameter(self):
        for i in range(1, 11):
         self.inst.write(f"PROG:SEQ:CLE {i}")
        self.send_parameter_by_index(1, 1)  # 有設備
        time.sleep(0.2)
        self.send_parameter_by_index(2, 2)  # 有設備
        time.sleep(0.2)
        self.send_parameter_by_index(3, 3)  # 有設備
        time.sleep(0.2)
        self.send_parameter_by_index(4, 4)  # 有設備
        time.sleep(0.2)
        self.send_parameter_by_index(5, 5)  # 有設備
        time.sleep(0.2)
        self.send_parameter_by_index(6, 6)  # 有設備
        time.sleep(0.2)
        self.send_parameter_by_index(7, 7)  # 有設備
        time.sleep(0.2)
        self.send_parameter_by_index(8, 8)  # 有設備
        time.sleep(0.2)
        self.send_parameter_by_index(9, 9)  # 有設備
        time.sleep(0.2)
        self.send_parameter_by_index(10, 10)  # 有設備
        time.sleep(0.2)
        self.send_parameter_by_index(11, 11)  # 有設備
        time.sleep(0.2)
       # self.send_parameter_by_index(3, 3)  # 有設備
       #  time.sleep(0.2)
    def SET_Meth1(self):
         # resp = self.inst.query("PROG:DATA:LIST? 1,1")
         # print("程式1序列1參數:", resp.strip())
         # resp = self.inst.inst.query("PROG:DATA:LIST? 1,2")
         # print("程式1序列1參數:", resp.strip())
         # resp = self.inst.inst.query("PROG:NSEL?")
         # print("程式1序列1參數:", resp.strip())
         self.inst.write("PROG:NSEL 1")
         self.inst.write("PROG:RUN")
         self.inst.write("LOAD ON")
         self.inst.write("SYST:LOC")

    def SET_Meth2(self):
         # resp = self.inst.query("PROG:DATA:LIST? 2,1")
         # print("程式1序列1參數:", resp.strip())
         # resp = self.inst.inst.query("PROG:DATA:LIST? 2,2")
         # print("程式1序列1參數:", resp.strip())
         # resp = self.inst.inst.query("PROG:NSEL?")
         # print("程式1序列1參數:", resp.strip())
         self.inst.write("PROG:NSEL 2")
         self.inst.write("PROG:RUN")
         self.inst.write("LOAD ON")
         self.inst.write("SYST:LOC")
    def SET_Meth3(self):
         self.inst.write("PROG:NSEL 3")
         self.inst.write("PROG:RUN")
         self.inst.write("LOAD ON")
         self.inst.write("SYST:LOC")
    def SET_Meth4(self):
         self.inst.write("PROG:NSEL 4")
         self.inst.write("PROG:RUN")
         self.inst.write("LOAD ON")
         self.inst.write("SYST:LOC")

    def SET_Meth5(self):
        self.inst.write("PROG:NSEL 5")
        self.inst.write("PROG:RUN")
        self.inst.write("LOAD ON")
        self.inst.write("SYST:LOC")

    def SET_Meth6(self):
        self.inst.write("PROG:NSEL 6")
        self.inst.write("PROG:RUN")
        self.inst.write("LOAD ON")
        self.inst.write("SYST:LOC")

    def SET_Meth7(self):
        self.inst.write("PROG:NSEL 7")
        self.inst.write("PROG:RUN")
        self.inst.write("LOAD ON")
        self.inst.write("SYST:LOC")

    def SET_Meth8(self):
        self.inst.write("PROG:NSEL 8")
        self.inst.write("PROG:RUN")
        self.inst.write("LOAD ON")
        self.inst.write("SYST:LOC")

    def SET_Meth9(self):
        self.inst.write("PROG:NSEL 9")
        self.inst.write("PROG:RUN")
        self.inst.write("LOAD ON")
        self.inst.write("SYST:LOC")

    def SET_Meth10(self):
        self.inst.write("PROG:NSEL 10")
        self.inst.write("PROG:RUN")
        self.inst.write("LOAD ON")
        self.inst.write("SYST:LOC")

    def on_close(self):
        _log("info", "[App] 關閉程式，釋放資源...")
        # 釋放 Excel 資源
        try:
            if getattr(self, "wb", None):
                self.wb.close()
                _log("debug", "[Excel] 工作簿已關閉")
        except Exception as e:
            _log("warning", f"[Excel] 關閉工作簿例外：{e}")
        try:
            if getattr(self, "xw_app", None):
                self.xw_app.quit()
                _log("debug", "[Excel] Excel 應用已結束")
        except Exception as e:
            _log("warning", f"[Excel] 關閉 Excel 應用例外：{e}")
        self.root.destroy()

    def make_ui(self):
        frm = tk.Frame(self.root)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        # 資源下拉
        tk.Label(frm, text="VISA 資源：", font=("Times New Roman", 14)).grid(row=0, column=0, sticky="w")
        self.cmb = ttk.Combobox(frm, width=62, state="readonly")
        self.cmb.grid(row=0, column=1, columnspan=2, sticky="we", pady=2)

        # ── 這段放在 make_ui() 的同一個 frm 裡 ──
        # 重新掃描 / 連線：仍放在第 0 列右側
        tk.Button(frm, text="重新掃描",command=self.scan_resources) .grid(row=0, column=3, padx=2, sticky="e")
        tk.Button(frm, text="連線", command=self.connect_selected) .grid(row=0, column=4, padx=2, sticky="e")
        tk.Button(frm, text="解除工作表鎖定", font=("Times New Roman",10),command=self.unlock_excel_and_reload) .grid(row=10,column=0, sticky="e", pady=50)


        # 一個容器裝兩排步驟按鈕
        steps = tk.Frame(frm)
        steps.grid(row=3, column=0, columnspan=6, sticky="ew", pady=(8, 4))

        btn_kwargs = dict(font=("Times New Roman", 12))

        labels_and_cmds = [
            ("載入數據", self.send_parameter),
            ("第一步", self.SET_Meth1),
            ("第二步", self.SET_Meth2),
            ("第三步", self.SET_Meth3),
            ("第四步", self.SET_Meth4),
            ("第五步", self.SET_Meth5),
        ]
        labels_and_cmds2 = [
            ("第六步", self.SET_Meth6),
            ("第七步", self.SET_Meth7),
            ("第八步", self.SET_Meth8),
            ("第九步", self.SET_Meth9),
            ("第十步", self.SET_Meth10),
        ]

        # 第一排：row=0
        for col, (txt, cmd) in enumerate(labels_and_cmds):
            tk.Button(steps, text=txt, command=cmd, **btn_kwargs) \
                .grid(row=0, column=col, padx=2, pady=2, sticky="ew")

        # 第二排：row=1
        for col, (txt, cmd) in enumerate(labels_and_cmds2):
            tk.Button(steps, text=txt, command=cmd, **btn_kwargs) \
                .grid(row=1, column=col, padx=2, pady=2, sticky="ew")

        # 讓每一欄等寬、可隨窗寬伸縮（取兩排的最大欄數）
        max_cols = max(len(labels_and_cmds), len(labels_and_cmds2))
        for col in range(max_cols):
            steps.grid_columnconfigure(col, weight=1, uniform="btns")

        # 父容器對被跨到的欄也給權重，整排能一起拉伸（視你的 frm 欄數調整）
        for col in range(6):
            frm.grid_columnconfigure(col, weight=1)
        tk.Button(frm, text="關掉負載",
                  font=("Times New Roman", 12),
                  command=self.set_load_off).grid(row=11,column=0, sticky="w", pady=7)


        # 狀態
        self.status = tk.Label(frm, text="未連線", anchor="w")
        self.status.grid(row=1, column=0, columnspan=5, sticky="we", pady=4)



        frm.grid_columnconfigure(1, weight=1)
        frm.grid_columnconfigure(2, weight=1)

    def _set_status(self, msg):
        self.status.config(text=msg)
        _log("info", f"[Status] {msg}")

    def init_backend_and_scan(self):
        rm, be, err = try_init_rm()
        if rm is None:
            self._set_status(f"ResourceManager 初始化失敗")
            messagebox.showerror(
                "錯誤",
                "VISA 後端初始化失敗：\n"
                f"{err}\n\n請確認：\n1) 已安裝 NI-VISA Runtime 或 pyvisa-py + libusb\n"
                "2) USB 驅動/裝置出現在裝置管理員\n3) 同廠儀器驅動正確"
            )
            return
        self.rm, self.backend = rm, be
        self._set_status(f"VISA 後端：{self.backend}")
        self.scan_resources(auto_connect=True)

    def scan_resources(self, auto_connect=False):
        #掃描目前VISA後端可見的資源，並把結果放進下拉選單(self.cmb)。
        auto_connect = True
        try:
            # 向 ResourceManager 查詢所有可見資源（例如 USB0::...、ASRL3::...、TCPIP::...）
            res = self.rm.list_resources()
            _log("info", f"[VISA] 掃描到資源: {res}")
        except Exception as e:
            # 若掃描失敗（後端沒裝好、驅動問題、權限不足等），更新狀態並清空下拉選單
            self._set_status(f"列資源失敗：{e}")
            self.cmb["values"] = []
            return

        self.cmb["values"] = res
        if res:
            self.cmb.current(0)
            self._set_status(f"共找到 {len(res)} 個資源")
            if auto_connect:
                self.connect_selected()
        else:
            self._set_status("找不到任何 VISA 資源。請檢查驅動/後端/線材。")

    def unlock_excel_and_reload(self):
        """解除 Excel 檔案鎖（關閉 wb/xw_app）並重新開啟活頁簿，更新選單"""
        # 1) 關閉活頁簿
        try:
            if getattr(self, "wb", None):
                self.wb.close()
                self.wb = None
                _log("debug", "[Excel] 工作簿已關閉（解除鎖定第一步）")
        except Exception as e:
            _log("warning", f"[Excel] 關閉工作簿例外：{e}")

        # 2) 關閉整個 Excel 應用（確保檔案鎖釋放）
        try:
            if getattr(self, "xw_app", None):
                self.xw_app.quit()
                self.xw_app = None
                _log("debug", "[Excel] Excel 應用已結束（鎖應該釋放）")
        except Exception as e:
            _log("warning", f"[Excel] 關閉 Excel 應用例外：{e}")

    def connect_selected(self):
        ex_content.clear()
        if not self.cmb["values"]:
            self._set_status("沒有資源可連線")
            return
        addr = self.cmb.get()
        try:
            self.inst = self.rm.open_resource(addr)
            self.inst.timeout = 5000
            self.inst.write_termination = "\n"
            self.inst.read_termination = "\n"
            try:
                idn = self.inst.query("*IDN?")
            except Exception:
                idn = ""
            self._set_status(f"已連線：{addr}  {idn.strip() if idn else ''}")
            _log("info", f"[VISA] 連線成功：{addr}；IDN='{idn.strip() if idn else 'N/A'}'")

            self.xw_app, self.wb, self.sheets, err = open_excel_wb(
            filename="執行檔.xlsx",
            visible=False,
            add_book=False,
            max_sheets=11,
            show_error_dialog=True  # 想用你原本的 messagebox UX 就留 True
            )
        except Exception as e:
            self.inst = None
            self._set_status(f"連線失敗：{e}")
            _log("error", f"[VISA] 連線失敗：{addr}；錯誤：{e}")
            messagebox.showerror("連線失敗", str(e))

    def _visa_write(self, cmd):
        """包一層 write，含除錯與可選驗證；回傳 True/False"""
        try:
            self.inst.write(cmd)
            _log("debug", f"[WRITE] -> {cmd}")
            if VERIFY_WRITE:
                # 儀器若支援 *OPC? 則回 '1'；不支援則嘗試 SYST:ERR?
                try:
                    ok = self.inst.query("*OPC?").strip()
                    _log("debug", f"[WRITE-VERIFY] *OPC? = '{ok}'")
                except Exception:
                    try:
                        err = self.inst.query("SYST:ERR?").strip()
                        _log("debug", f"[WRITE-VERIFY] SYST:ERR? = '{err}'")
                    except Exception:
                        _log("debug", "[WRITE-VERIFY] 無法驗證（不支援 *OPC?/SYST:ERR?）")
            return True
        except Exception as e:
            _log("error", f"[WRITE-FAIL] {cmd} | 錯誤: {e}")
            return False

    def set_load_off(self):
        if self.inst is None:
            self._set_status("尚未連線儀器")
            return
        try:
            try:
                self._visa_write("LOAD OFF")
            except Exception:
                self._visa_write("INP OFF")
            self._set_status("已關閉負載")
            self._visa_write("SYST:LOC")
        except Exception as e:
            self._set_status(f"關閉失敗：{e}")
            _log("error", f"[WRITE-FAIL] 關閉負載錯誤：{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()



