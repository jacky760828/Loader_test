import tkinter as tk
import pyvisa
import time
from tkinter import ttk, messagebox
import sys, logging, time
import os
ex_content=[]
# ====== Log 設定（console + 檔案）======
LOG_TO_FILE = False
DEBUG_LEVEL = logging.DEBUG  # 可改 logging.INFO

logger = logging.getLogger("visa_app")#用 Logger代換
logger.setLevel(DEBUG_LEVEL)
_fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
# 設定輸出格式：時間戳 + 等級 + 訊息

_ch = logging.StreamHandler(sys.stdout)    #建立「主控台」輸出用的處理器
_ch.setLevel(DEBUG_LEVEL)
_ch.setFormatter(_fmt)     # 設定輸出格式（時間、等級、訊息）
if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
    logger.addHandler(_ch)

if LOG_TO_FILE:
    _fh = logging.FileHandler("app.log", encoding="utf-8")
    _fh.setLevel(DEBUG_LEVEL)
    _fh.setFormatter(_fmt)
    if not any(isinstance(h, logging.FileHandler) for h in logger.handlers):
        logger.addHandler(_fh)


def open_excel_wb(filename="case1.xlsx",
                  visible=False,
                  add_book=False,
                  max_sheets=11,#最多抓11張
                  show_error_dialog=True):
    """開啟 Excel 檔並回傳 (xw_app, workbook, sheets, err_text)"""
    try:
        import xlwings as xw
    except Exception as e:
        logger.error("[Excel] 載入 xlwings 失敗：%s", e)
        return None, None, [], "載入 xlwings 失敗：%s" % e

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

        logger.info("[Excel] 已開啟檔案：%s，工作表數: %d（取前 %d 張）",
                    os.path.abspath(filename), total, len(sheets))
        if sheets:
            logger.debug("[Excel] 第一張表名：%s", sheets[0].name)

        return app, wb, sheets, None

    except Exception as e:
        logger.error("[Excel] 開啟或列舉工作表失敗：%s", e)
        if show_error_dialog:
            try:
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


def _fmt_exc(e: Exception) -> str:
    return f"{type(e).__name__}: {e}"


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
        result = []  # 回傳的 SCPI 指令字串
        prefix = "PROG:DATA "
        try:
            sheet_name = sheet.name
            logger.info("[Excel] 開始讀取工作表：%s", sheet_name)

            # 以 B 欄最後使用列為準（避免 B1 是標題導致 .end('down') 撞牆）
            last_row = sheet.range("B" + str(sheet.cells.last_cell.row)).end("up").row
            logger.debug("[Excel] B 欄最後列 = %d", last_row)

            if last_row < 2:
                logger.info( "[Excel] 找不到資料（B2 以下無內容）")
                return result

            arr = sheet.range(f"B2:C{last_row}").value  # [[B2,C2], [B3,C3], ...]
            if not isinstance(arr, list):
                arr = [arr]

            logger.info("[Excel] 讀到筆數:%d",len(arr))
            print(ex_content)
            logger.info("[Excel] 長度:%d",len(ex_content))
            logger.debug("[Excel] ex_content 範例：前3=%s 後3=%s",
             ex_content[:3],
             ex_content[-3:] if len(ex_content) > 3 else ex_content)

            for row in arr:  # b2下只有一欄
                if not row or len(row) < 2:
                    continue
                b_raw, c_raw = row[0], row[1]
                if b_raw is None or c_raw is None:
                    continue
                # 將 excel 讀到的值盡可能轉成數字（3.0 -> 3；其餘維持原樣）
                b_val = _to_num(b_raw)
                c_val = _to_num(c_raw)
                ex_content.append(c_val)

                new_str = prefix + f"{b_val},LIST,{c_val},0,{c_val}"
                result.append(new_str)

            logger.info("[Excel]產生指令數:%d", len(result))
            if result:
                logger.debug("[Excel] 首條指令: %s", result[0])
                logger.debug("[Excel] 末條指令: %s", result[-1])

        except Exception as e:
            logger.info("error[Excel 讀取錯誤]")
        # 回傳完整的指令清單
        return result
def generate_prog_data2(sheet, row_count=None):
    # 從指定的Excel (xlwings 的 sheet 物件) 讀取 B~F 欄位的參數，
    # 由 B2 開始往下，共 row_count 列，並組成儀器可用的`PROG:DATA:LIST ...` SCPI 指令字串清單後回傳。
    #  row_count  : 從 B2 開始要讀取的列數。
    # result     : list[str]，每個元素是一條完整的 "PROG:DATA:LIST ..." 指令
    prefix = "PROG:DATA:LIST "
    result = []

    # --- LOG: 進函式 ---
    try:
        sheet_name = getattr(sheet, "name", None)
    except Exception:
        sheet_name = None
    logger.debug("[generate_prog_data2] start: sheet=%s, row_count(arg)=%s", sheet_name, row_count)

    # 自動偵測列數（從 B2 向下到最後一筆連續資料）
    if row_count is None:
        last_row = sheet.cells(2, "B").end("down").row  # 回到最後一筆
        row_count = max(0, last_row - 1)  # 扣掉表頭 B1
        logger.debug("[generate_prog_data2] auto-detect: last_row=%s -> row_count=%s", last_row, row_count)

    if row_count <= 0:
        logger.warning("[generate_prog_data2] row_count<=0，無資料可讀；直接回傳空清單")
        return result

    # === 一次把要用到的區塊抓出來 ===
    # 從 B2 起，往下 row_count 列，橫向讀 5 欄（B~F）
    # 取得的 arr 會是 list of lists，例如:
    #   [[B2, C2, D2, E2, F2], [B3, C3, D3, E3, F3], ...]
    arr = sheet.range("B2").resize(row_count, 5).value
    n = len(arr) if isinstance(arr, list) else 0
    logger.debug("[generate_prog_data2] read block: rows=%s", n)
    if n:
        logger.debug("[generate_prog_data2] sample head: %s", arr[:2])
        if n > 2:
            logger.debug("[generate_prog_data2] sample tail: %s", arr[-2:])

    for row in arr:
        # row = [B, C, D, E, F]
        if not row or any(v is None for v in row):
            logger.debug("[generate_prog_data2] skip row (empty or has None): %s", row)
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

    # logger.info("[generate_prog_data2] generated commands: %d", len(result))
    # if result:
        # logger.debug("[generate_prog_data2] first cmd: %s", result[0])
        # logger.debug("[generate_prog_data2] last  cmd: %s", result[-1])

    return result




class App:
    def __init__(self, root):
        self.root = root
        self.root.title("同步抽載")
        self.rm = None
        self.resources = []
        # 面板狀態（含各自的 Excel 選擇與物件）
        #共兩組
        self.panels = [   #addr該連線的資源位址字串
            {"inst": None, "addr": None,
             "excel_app": None, "excel_wb": None, "excel_sheets": [],
             "excel_file": tk.StringVar(self.root, value="執行檔1.xlsx")},
            {"inst": None, "addr": None,
             "excel_app": None, "excel_wb": None, "excel_sheets": [],
             "excel_file": tk.StringVar(self.root, value="執行檔1.xlsx")},
        ]
        logger.info("應用程式啟動")
        self.make_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ===== 基本 VISA =====
    def _ensure_rm(self):
        if self.rm is None:
            try:
                logger.debug("正在建立 ResourceManager …")
                t0 = time.perf_counter()
                self.rm = pyvisa.ResourceManager()  # 需要可改 '@ni' 或 '@py'
                dt = (time.perf_counter() - t0) * 1000
                logger.info("ResourceManager 建立成功（%.1f ms）：%s", dt, self.rm)
            except Exception as e:
                logger.exception("建立 ResourceManager 失敗")
                messagebox.showerror("錯誤", f"建立 ResourceManager 失敗：{e}")
                return False
        return True

    def load_and_send_i(self, i):
        self.load_excel_i(i)
        self.send_parameter_i(i)

    def send_parameter_by_index(self, panel_idx, sheet_idx_1based=1, count_idx_1based=1):
        """
        依指定工作表與 ex_content 的索引，送出 PROG:DATA:LIST 並儲存程式
        ※ 每次呼叫都會重建 ex_content（從該面板的第1張表讀 B/C）
        """
        logger.debug(
            "send_parameter_by_index(panel_idx=%s, sheet_idx_1based=%s, count_idx_1based=%s)",
            panel_idx, sheet_idx_1based, count_idx_1based
        )

        p = self.panels[panel_idx]
        inst = p.get("inst")
        logger.debug("[面板%s] inst=%r, type=%s, addr=%r",
                     panel_idx, inst, type(inst), p.get("addr"))
        try:
            rn = getattr(inst, "resource_name", None)
            logger.debug("[面板%s] VISA resource_name=%r", panel_idx, rn)
        except Exception as e:
            logger.debug("[面板%s] 讀取 resource_name 失敗：%s", panel_idx, e)

        if not inst:
            p["status"].config(text="尚未連線")
            logger.warning("[面板%s] 尚未連線，inst 為空", panel_idx)
            return

        sheets = p.get("excel_sheets", [])
        logger.debug("[面板%s] excel_sheets loaded=%s, count=%s, names=%s",
                     panel_idx, bool(sheets), (len(sheets) if sheets else 0),
                     ([s.name for s in sheets] if sheets else []))
        if not sheets:
            p["status"].config(text="尚未載入執行檔")
            logger.warning("[面板%s] 尚未載入執行檔", panel_idx)
            return

        # 目標表（1-based → 0-based），用該面板的 sheets
        idx = sheet_idx_1based - 1
        if idx < 0 or idx >= len(sheets):
            p["status"].config(text=f"工作表索引超出範圍：{sheet_idx_1based}")
            logger.error("[面板%s] 工作表索引超出範圍：%s（可用 1..%s）",
                         panel_idx, sheet_idx_1based, len(sheets))
            return
        sheet = sheets[idx]
        logger.info("[面板%s] 目標工作表：%s（index=%s）", panel_idx, sheet.name, sheet_idx_1based)

        # 每次都重建 ex_content：從面板的第1張表讀 B/C
        # global ex_content
        # logger.debug("[面板%s] ex_content 清空（原 len=%d）", panel_idx, len(ex_content))
        # ex_content.clear()
        if sheet_idx_1based==1:
           global ex_content
           logger.debug("[面板%s] ex_content 清空（原 len=%d）", panel_idx, len(ex_content))
           ex_content.clear()
           count_sheet = sheets[0]
           prog_data_list = generate_prog_data(count_sheet)
           for line in prog_data_list:
             cmd = line.strip()
             if cmd:  # 跳過空行
              inst.write(cmd)
            # print(f"已送出: {cmd}")
        #dt = (time.perf_counter() - t0) * 1000
        logger.info("[面板%s] ex_content 重建完成，len=%d；前2=%s；後2=%s",
                    panel_idx, len(ex_content),
                    ex_content[:2],
                    (ex_content[-2:] if len(ex_content) > 2 else ex_content))

        # 由 ex_content 取 row_count；不足就用目標表 B 欄自動偵測
        if len(ex_content) >= count_idx_1based:
            row_count = int(_to_num(ex_content[count_idx_1based - 1]))
            logger.debug("[面板%s] row_count 來自 ex_content[%d] -> %s",
                         panel_idx, count_idx_1based - 1, row_count)
        else:
            last_row = sheet.range("B" + str(sheet.cells.last_cell.row)).end("up").row
            row_count = max(0, last_row - 1)
            logger.warning("[面板%s] ex_content 筆數不足（len=%d < %d），改用自動偵測 row_count=%d",
                           panel_idx, len(ex_content), count_idx_1based, row_count)

        if row_count <= 0:
            p["status"].config(text="row_count 為 0，沒有可送出的資料")
            logger.error("[面板%s] row_count=0，無可送資料", panel_idx)
            return
        inst.write("SYST:LOC")

        set_sheet=sheets[sheet_idx_1based]
        # 產生 + 下發 PROG:DATA:LIST
        #t1 = time.perf_counter()
        prog_list= generate_prog_data2(set_sheet,row_count)
        #
        #
        # #gen_ms = (time.perf_counter() - t1) * 1000
        logger.info("[面板%s] 已產生 PROG:DATA:LIST：%d 條",
                     panel_idx, len(prog_list))
        if not prog_list:
             p["status"].config(text="沒有可送出的 PROG:DATA:LIST 指令")
             logger.warning("[面板%s] PROG:DATA:LIST 為空", panel_idx)
             return
        #
        # p["status"].config(text=f"開始載入數據，共 {len(prog_list)} 條")
        ok = 0
        total = len(prog_list)
        # t2 = time.perf_counter()
        for i, line in enumerate(prog_list, 1):
            cmd = (line or "").strip()
            if not cmd:
                logger.debug("[面板%s] 第 %d 條為空白，略過", panel_idx, i)
                continue
            # 只對前 3 / 每 50 / 後 3 條做詳細 log，避免過量
            if i <= 3 or i % 50 == 0 or i > total - 3:
                logger.debug("[面板%s] CMD %d/%d -> %s", panel_idx, i, total, cmd)

            if inst.write(cmd):
                #ok += 1
                time.sleep(0.1)

            if i % 20 == 0 or i == total:
                p["status"].config(text=f"載入進度：{i}/{total}")
        inst.write ("PROG:SAV")
        # p["status"].config(text=f"載入完成，成功 {ok}/{len(prog_list)} 條，已下達 PROG:SAV")
        # logger.info("[面板%s] PROG:SAV 已送出", panel_idx)
        # inst.write("SYST:LOC")

    def set_load_off_panel(self, panel_idx: int):
        """只關閉某一個面板（panel_idx）的負載"""
        p = self.panels[panel_idx]
        inst = p.get("inst")
        if not inst:
            # 沒連線就提示並返回
            p["status"].config(text="尚未連線儀器")
            return False

        try:
            try:
                inst.write("LOAD OFF")
            except Exception:
                # 有些儀器用 INP OFF
                inst.write("INP OFF")

            p["status"].config(text="已關閉負載")

            # 回到本機控制（若儀器支援）
            try:
                inst.write("SYST:LOC")
            except Exception:
                pass  # 不支援就略過

            return True
        except Exception as e:
            p["status"].config(text=f"關閉失敗：{e}")
            # 若你有 logger，可加：
            # logger.error("[面板%s] 關閉負載錯誤：%s", panel_idx, e)
            return False








    def send_parameter_i(self, panel_idx):
        """針對第 panel_idx 面板執行：清程式序列 + 載入/下發資料"""
        logger.debug("send_parameter_i(panel_idx=%s)", panel_idx)

        p = self.panels[panel_idx]
        inst = p.get("inst")

        # >>> 想看 inst 是什麼：把 repr / type / 地址等都印出來
        logger.debug("[面板%s] inst=%r, type=%s, addr=%r",
                     panel_idx, inst, type(inst), p.get("addr"))

        # 也順便把一些常見屬性抓出來（若存在）
        try:
            rn = getattr(inst, "resource_name", None)
            timeout = getattr(inst, "timeout", None)
            wt = getattr(inst, "write_termination", None)
            rt = getattr(inst, "read_termination", None)
            logger.debug("[面板%s] VISA info -> resource_name=%r, timeout=%r, write_term=%r, read_term=%r",
                         panel_idx, rn, timeout, wt, rt)
        except Exception as e:
            logger.debug("[面板%s] 讀取 inst 屬性失敗：%s", panel_idx, e)

        if not inst:
            p["status"].config(text="尚未連線")
            logger.warning("[面板%s] 尚未連線，inst 為空", panel_idx)
            return  # <- 確保 return 在 if 內部

        # （可選）確認有載入對應 Excel：需在別處把 sheets 放到 p["excel_sheets"]
        sheets = p.get("excel_sheets")
        logger.debug("[面板%s] excel_sheets loaded=%s, count=%s",
                     panel_idx, bool(sheets), (len(sheets) if sheets else 0))
        if not sheets:
            p["status"].config(text="尚未載入執行檔")
            logger.warning("[面板%s] 尚未載入執行檔", panel_idx)
            return  # <- 同樣把 return 放進 if

        # 下面這段若你原本就有可保留；示範清 1~10 序列 test
        for seq in range(1, 11):
            cmd = f"PROG:SEQ:CLE {seq}"
            try:
                inst.write(cmd)
                #logger.debug("[面板%s] WRITE -> %s  [OK]", panel_idx, cmd)
            except Exception as e:
                logger.error("[面板%s] WRITE 失敗 -> %s | %s", panel_idx, cmd, e)
                p["status"].config(text=f"清序列失敗：{e}")
                return

        logger.debug("清除完成")
        #
        # 接著送資料（依你的既有流程）
        self.send_parameter_by_index(panel_idx, sheet_idx_1based=1, count_idx_1based=1)
        time.sleep(0.2)
        self.send_parameter_by_index(panel_idx, sheet_idx_1based=2, count_idx_1based=2)
        time.sleep(0.2)
        self.send_parameter_by_index(panel_idx, sheet_idx_1based=3, count_idx_1based=3)
        time.sleep(0.2)
        self.send_parameter_by_index(panel_idx, sheet_idx_1based=4, count_idx_1based=4)
        time.sleep(0.2)
        self.send_parameter_by_index(panel_idx, sheet_idx_1based=5, count_idx_1based=5)
        time.sleep(0.2)
        self.send_parameter_by_index(panel_idx, sheet_idx_1based=6, count_idx_1based=6)
        time.sleep(0.2)
        self.send_parameter_by_index(panel_idx, sheet_idx_1based=7, count_idx_1based=7)
        time.sleep(0.2)
        self.send_parameter_by_index(panel_idx, sheet_idx_1based=8, count_idx_1based=8)
        time.sleep(0.2)
        self.send_parameter_by_index(panel_idx, sheet_idx_1based=9, count_idx_1based=9)
        time.sleep(0.2)
        self.send_parameter_by_index(panel_idx, sheet_idx_1based=10, count_idx_1based=10)
    def scan_resources_all(self):
        logger.info("開始掃描 VISA 資源")
        if not self._ensure_rm():
            return
        try:
            t0 = time.perf_counter()
            self.resources = list(self.rm.list_resources())
           # dt = (time.perf_counter() - t0) * 1000
            #logger.info("掃描完成（%.1f ms），找到資源：%s", dt, self.resources or "（無）")
        except Exception as e:
            self.resources = []
            logger.error("掃描失敗：%s", _fmt_exc(e))
            messagebox.showerror("錯誤", f"掃描失敗：{e}")

        for idx, p in enumerate(self.panels):
            p["cmb"]["values"] = self.resources
            if self.resources:
                p["cmb"].current(0)
            p["status"].config(text="尚未連線")
            p["idn_val"].set("")
            p["addr"] = None
            if p["inst"]:
                try:
                    p["inst"].close()
                    logger.debug("[面板%d] 關閉舊連線", idx)
                except Exception as e:
                    logger.warning("[面板%d] 關閉舊連線例外：%s", idx, _fmt_exc(e))
            p["inst"] = None

    def connect_selected_i(self, i):
        p = self.panels[i]
        logger.info("[面板%d] 嘗試連線所選資源", i)
        if not self._ensure_rm():
            return
        vals = p["cmb"]["values"]
        if not vals:
            p["status"].config(text="無資源")
            logger.warning("[面板%d] 無資源可選", i)
            return
        addr = p["cmb"].get()
        logger.debug("[面板%d] 使用者選擇的地址：%r", i, addr)
        if not addr:
            p["status"].config(text="未選擇資源")
            logger.warning("[面板%d] 尚未選擇資源", i)
            return
        try:
            if p["inst"]:
                try:
                    p["inst"].close()
                    logger.debug("[面板%d] 關閉既有連線", i)
                except Exception as e:
                    logger.warning("[面板%d] 關閉既有連線例外：%s", i, _fmt_exc(e))
                p["inst"] = None

            logger.debug("[面板%d] open_resource -> %s", i, addr)
            t0 = time.perf_counter()
            inst = self.rm.open_resource(addr)
            inst.timeout = 5000
            inst.read_termination = "\n"
            inst.write_termination = "\n"
            open_dt = (time.perf_counter() - t0) * 1000
            logger.info("[面板%d] 開啟成功（%.1f ms）", i, open_dt)

            p["inst"] = inst
            p["addr"] = addr

            try:
                t1 = time.perf_counter()
                idn = inst.query("*IDN?").strip()
                qdt = (time.perf_counter() - t1) * 1000
                logger.debug("[面板%d] QUERY *IDN? （%.1f ms）<<< %s", i, qdt, idn)
            except Exception as e:
                idn = "(讀取 *IDN? 失敗)"
                logger.error("[面板%d] *IDN? 失敗：%s", i, _fmt_exc(e))
            p["idn_val"].set(idn)
            p["status"].config(text=f"已連線：{addr}")
        except Exception as e:
            p["status"].config(text=f"連線失敗")
            logger.exception("[面板%d] 連線失敗：%s", i, _fmt_exc(e))
            messagebox.showerror("連線失敗", str(e))

    def read_idn_i(self, i):
        p = self.panels[i]
        logger.info("[面板%d] 讀取 *IDN?", i)
        if not p["inst"]:
            p["status"].config(text="尚未連線")
            logger.warning("[面板%d] 尚未連線，無法讀取 *IDN?", i)
            return
        try:
            t0 = time.perf_counter()
            idn = p["inst"].query("*IDN?").strip()
            dt = (time.perf_counter() - t0) * 1000
            p["idn_val"].set(idn)
            p["status"].config(text="*IDN? 讀取完成")
            logger.debug("[面板%d] QUERY *IDN? （%.1f ms）<<< %s", i, dt, idn)
        except Exception as e:
            p["status"].config(text="讀取失敗")
            logger.error("[面板%d] 讀取 *IDN? 失敗：%s", i, _fmt_exc(e))
            messagebox.showerror("讀取失敗", str(e))

    # ===== Excel：每面板用 Radiobutton 選檔 + 載入 =====
    def _run_prog(self, panel_idx: int, prog_n: int):
        """在指定面板 panel_idx 上：
           選擇第 prog_n 號程序 -> RUN -> LOAD ON -> (可選) 回到本機 SYST:LOC
        """
       # logger.debug("_run_prog(panel_idx=%s, prog_n=%s) 進入", panel_idx, prog_n)
        p = self.panels[panel_idx]
        inst = p.get("inst")
       # logger.debug("[面板%s] inst=%r, type=%s, addr=%r",
         #        panel_idx, inst, type(inst), p.get("addr"))
        if not inst:
            p["status"].config(text="尚未連線")
            logger.warning("[面板%s] 尚未連線，無法執行程序 %s", panel_idx, prog_n)
            return
    
        try:
            logger.debug("[面板%s] WRITE -> PROG:NSEL %s", panel_idx, prog_n)
            inst.write(f"PROG:NSEL {prog_n}")
            logger.debug("[面板%s] WRITE -> PROG:RUN", panel_idx)
            inst.write("PROG:RUN")
            logger.debug("[面板%s] WRITE -> LOAD ON", panel_idx)
            inst.write("LOAD ON")
            logger.debug("[面板%s] WRITE -> SYST:LOC", panel_idx)
            inst.write("SYST:LOC")
    
            p["status"].config(text=f"程序 {prog_n} 執行中（LOAD ON）")
        except Exception as e:
            p["status"].config(text=f"程序 {prog_n} 執行失敗：{e}")
            # 視需要可用 logger 記一筆或彈出視窗
            # logger.error(f"[面板{panel_idx}] 程序 {prog_n} 執行失敗: {e}")
            # messagebox.showerror("錯誤", str(e))

    def set_load_off_panel(self, panel_idx: int):
        """只關閉某一個面板（panel_idx）的負載"""
        p = self.panels[panel_idx]
        inst = p.get("inst")
        if not inst:
            # 沒連線就提示並返回
            p["status"].config(text="尚未連線儀器")
            return False

        try:
            try:
                inst.write("LOAD OFF")
            except Exception:
                # 有些儀器用 INP OFF
                inst.write("INP OFF")

            p["status"].config(text="已關閉負載")

            # 回到本機控制（若儀器支援）
            try:
                inst.write("SYST:LOC")
            except Exception:
                pass  # 不支援就略過

            return True
        except Exception as e:
            p["status"].config(text=f"關閉失敗：{e}")
            # 若你有 logger，可加：
            # logger.error("[面板%s] 關閉負載錯誤：%s", panel_idx, e)
            return False

    def set_load_off_both(self):
        """同時關閉 A/B 兩個面板的負載（依序下發）"""
        # 依序對 0、1 面板執行；若未連線會在子函式內處理
        for idx in range(min(2, len(self.panels))):
            self.set_load_off_panel(idx)
    def run_prog_both(self, prog_n: int):
        """讓 A、B 兩個面板同時執行第 prog_n 號程序（序列化下發）"""
        # 這裡直接重用你已經寫好的 _run_prog，
        # 會依序對 panel 0、panel 1 下指令（若未連線會自動略過並顯示狀態）
        for idx in range(min(2, len(self.panels))):
            try:
                self._run_prog(idx, prog_n)
            except Exception as e:
                # 保險：任何一台出錯都不會影響另一台
                try:
                    self.panels[idx]["status"].config(text=f"程序 {prog_n} 失敗：{e}")
                except Exception:
                    pass
    # 五個對應按鈕的包裝（按鈕綁定呼叫這些）
    def SET_Meth1(self, panel_idx): self._run_prog(panel_idx, 1)
    def SET_Meth2(self, panel_idx): self._run_prog(panel_idx, 2)
    def SET_Meth3(self, panel_idx): self._run_prog(panel_idx, 3)
    def SET_Meth4(self, panel_idx): self._run_prog(panel_idx, 4)
    def SET_Meth5(self, panel_idx): self._run_prog(panel_idx, 5)
    def SET_Meth6(self, panel_idx): self._run_prog(panel_idx, 6)
    def SET_Meth7(self, panel_idx): self._run_prog(panel_idx, 7)
    def SET_Meth8(self, panel_idx): self._run_prog(panel_idx, 8)
    def SET_Meth9(self, panel_idx): self._run_prog(panel_idx, 9)
    def SET_Meth10(self, panel_idx): self._run_prog(panel_idx, 10)


    # def load_excel_i(self, i, path=None, max_sheets=11):
    #     """
    #     載入第 i 個面板要用的 Excel，並把 sheets 存到 self.panels[i]["excel_sheets"]。
    #     i: 0=設備A, 1=設備B
    #     """
    #     p = self.panels[i]
    #
    #     # 選路徑：你也可以改成用檔案選擇器，或依 RadioButton 決定不同檔名
    #     if path is None:
    #         # 例：A用 執行檔A.xlsx、B用 執行檔B.xlsx（請依實際需求修改）
    #         path = "case1.xlsx" if i == 0 else "case2_m.xlsx"
    #
    #     # 關掉舊的 Excel 連結（如果曾經載入過）
    #     try:
    #         if p.get("excel_wb"):
    #             p["excel_wb"].close()
    #         if p.get("excel_app"):
    #             p["excel_app"].quit()
    #     except Exception:
    #         pass
    #
    #     # 打開新的
    #     try:
    #         import xlwings as xw
    #     except Exception as e:
    #         p["status"].config(text="未安裝 xlwings")
    #         from tkinter import messagebox
    #         messagebox.showerror("錯誤", f"未安裝 xlwings：{e}")
    #         return
    #
    #     try:
    #         app = xw.App(visible=False, add_book=False)
    #         wb = app.books.open(path)
    #         total = len(wb.sheets)
    #         cnt = min(max_sheets, total)
    #         sheets = [wb.sheets[j] for j in range(cnt)]
    #     except Exception as e:
    #         p["status"].config(text="載入 Excel 失敗")
    #         from tkinter import messagebox
    #         messagebox.showerror("Excel 讀取錯誤", f"{e}\n\n檔案：{path}")
    #         try:
    #             app.quit()
    #         except Exception:
    #             pass
    #         return
    #
    #     # ★★★ 關鍵：把工作表清單「塞進面板」 ★★★
    #     p["excel_app"] = app
    #     p["excel_wb"] = wb
    #     p["excel_sheets"] = sheets
    #
    #     # UI 狀態提示
    #     p["status"].config(text=f"已載入：{path}（{len(sheets)} 張表）")     #目前沒問題
    # ===== UI 佈局 =====
    def load_excel_i(self, i, path=None, max_sheets=11, base_dir=None):
        """
        載入第 i 個面板要用的 Excel，並把 sheets 存到 self.panels[i]["excel_sheets"]。
        i: 0=設備A, 1=設備B
        path: 若為 None，將使用該面板 RadioButton 選到的檔名（self.panels[i]["excel_file"].get()）
        base_dir: 可選；若給了且 path 不是絕對路徑，會用 base_dir 拼出完整路徑
        """
        p = self.panels[i]

        # 1) 來源路徑：優先用參數，其次用 RadioButton 選擇；最後備援給預設
        if path is None:
            sel = p.get("excel_file")
            if hasattr(sel, "get"):
                path = sel.get()
            else:
                path = sel or ("case1.xlsx" if i == 0 else "case2_m.xlsx")

        if base_dir and not os.path.isabs(path):
            path = os.path.join(base_dir, path)

        # 2) 關掉舊的 Excel 連結（如果曾經載入過）
        try:
            if p.get("excel_wb"):
                p["excel_wb"].close()
            if p.get("excel_app"):
                p["excel_app"].quit()
        except Exception:
            pass

        # 3) 開新檔
        try:
            import xlwings as xw
        except Exception as e:
            p["status"].config(text="未安裝 xlwings")
            messagebox.showerror("錯誤", f"未安裝 xlwings：{e}")
            return

        if not os.path.exists(path):
            p["status"].config(text=f"找不到檔案：{path}")
            messagebox.showerror("Excel 讀取錯誤", f"找不到檔案：\n{path}")
            return

        try:
            app = xw.App(visible=False, add_book=False)
            try:
                app.display_alerts = False
                app.screen_updating = False
            except Exception:
                pass

            wb = app.books.open(path)
            total = len(wb.sheets)
            cnt = min(max_sheets, total)
            sheets = [wb.sheets[j] for j in range(cnt)]
        except Exception as e:
            p["status"].config(text="載入 Excel 失敗")
            messagebox.showerror("Excel 讀取錯誤", f"{e}\n\n檔案：{path}")
            try:
                app.quit()
            except Exception:
                pass
            return

        # 4) 存回面板狀態
        p["excel_app"] = app
        p["excel_wb"] = wb
        p["excel_sheets"] = sheets

        # 5) UI 提示
        p["status"].config(text=f"已載入：{path}（{len(sheets)} 張表）")

    def _build_device_panel(self, parent, title, index):
        # 1) 有內邊距的面板
        lf = ttk.LabelFrame(parent, text=title, padding=(10, 8))
        lf.grid(row=0, column=index, sticky="nsew", padx=8, pady=8)

        # 2) 讓 1~5 欄可伸縮、等寬
        for c in range(1, 6):
            lf.grid_columnconfigure(c, weight=1, uniform="lfcols")

        # 資源選擇
        ttk.Label(lf, text="VISA 資源：").grid(row=0, column=0, sticky="w", padx=6, pady=2)
        cmb = ttk.Combobox(lf, state="readonly")  # width 可以拿掉、交給伸縮
        cmb.grid(row=0, column=1, columnspan=2, sticky="we", padx=4, pady=2)

        ttk.Button(lf, text="重新掃描", command=self.scan_resources_all) \
            .grid(row=0, column=3, padx=4, pady=2, sticky="e")
        ttk.Button(lf, text="連線", command=lambda i=index: self.connect_selected_i(i)) \
            .grid(row=0, column=4, padx=4, pady=2, sticky="e")

        # 狀態列
        ttk.Label(lf, text="狀態：").grid(row=1, column=0, sticky="w", padx=6, pady=2)
        status = ttk.Label(lf, text="尚未連線", anchor="w")
        status.grid(row=1, column=1, columnspan=4, sticky="we", padx=4, pady=2)

        # *IDN?
        ttk.Button(lf, text="讀取 *IDN?", command=lambda i=index: self.read_idn_i(i)) \
            .grid(row=2, column=0, padx=6, pady=2, sticky="w")
        idn_val = tk.StringVar(value="")
        # ← 修正：把 Entry 放到 row=2（和按鈕同一列），不跟狀態列重疊
        ttk.Entry(lf, textvariable=idn_val) \
            .grid(row=2, column=1, columnspan=4, sticky="we", padx=4, pady=2)

        # 執行檔 Radiobutton 列
        # ttk.Label(lf, text="執行檔：").grid(row=3, column=0, sticky="w", padx=6, pady=(6, 2))
        # rb_var = self.panels[index]["excel_file"]
        # files = ["case1.xlsx", "case2_m.xlsx", "case2_s.xlsx", "case3_m.xlsx", "case3_s.xlsx"]
        # if rb_var.get() not in files:
        #     rb_var.set(files[0])
        #
        # filebar = ttk.Frame(lf)
        # filebar.grid(row=3, column=1, columnspan=4, sticky="ew", padx=4, pady=(6, 2))
        # for c in range(len(files)):
        #     filebar.grid_columnconfigure(c, weight=1, uniform="files")
        # 用一個子框包起來（放在 row=3，吃掉整列）
        fileline = ttk.Frame(lf)
        fileline.grid(row=3, column=0, columnspan=5, sticky="w", padx=1, pady=(4, 2))

        # 子框內：第 0 欄放標籤、第 1 欄放 radiobutton 群組
        ttk.Label(fileline, text="執行檔：").grid(row=0, column=0, sticky="w", padx=(0, 4), pady=0)

        filebar = ttk.Frame(fileline)
        filebar.grid(row=0, column=1, sticky="w", padx=0, pady=0)

        # 讓 radiobutton 這一欄可拉伸（如果你想讓它吃寬）
        fileline.grid_columnconfigure(1, weight=1)
        rb_var = self.panels[index]["excel_file"]
        files = ["case1.xlsx", "case2_m.xlsx","case2_s.xlsx", "case3_m.xlsx", "case3_s.xlsx"]
        # radiobuttons 放在 filebar 裡
        for col, name in enumerate(files):
            ttk.Radiobutton(filebar, text=name, variable=rb_var, value=name) \
                .grid(row=0, column=col, sticky="w", padx=1, pady=0)

        # # 單列排法（若想兩列，見下方註解）
        # for col, name in enumerate(files):
        #     ttk.Radiobutton(filebar, text=name, variable=rb_var, value=name) \
        #         .grid(row=0, column=col, sticky="w", padx=2, pady=2)

        # 右側動作鈕
        ttk.Button(lf, text="寫入程序",
                   command=lambda i=index: self.load_and_send_i(i)) \
            .grid(row=3, column=5, padx=4, pady=(6, 2), sticky="e")

        # 程序控制按鈕列（五顆等寬）
        btn_row = 5
        btns = ttk.Frame(lf)
        btns.grid(row=btn_row, column=0, columnspan=6, sticky="ew", padx=4, pady=(8, 2))
        for c in range(5):
            btns.grid_columnconfigure(c, weight=1, uniform="progbtns")

        labels_cmds = [
            ("第一步", self.SET_Meth1),
            ("第二步", self.SET_Meth2),
            ("第三步", self.SET_Meth3),
            ("第四步", self.SET_Meth4),
            ("第五步", self.SET_Meth5),
        ]
        for col, (txt, fn) in enumerate(labels_cmds):
            ttk.Button(btns, text=txt, command=lambda i=index, f=fn: f(i)) \
                .grid(row=0, column=col, padx=4, pady=(6, 2), sticky="ew")

        # 保存元件
        self.panels[index]["cmb"] = cmb
        self.panels[index]["status"] = status
        self.panels[index]["idn_val"] = idn_val

    def _sync_prog(self):
       sync = ttk.LabelFrame(self.root, text="雙機同步")
       sync.pack(fill="x", padx=12, pady=(6, 12))

       ttk.Label(sync, text="同步程序控制：").grid(row=0, column=0, sticky="w", padx=2, pady=(6, 2))

       for col, n in enumerate([1, 2, 3, 4, 5,6,7,8,9,10], start=1):
         ttk.Button(sync, text=f"同步第{n}步", command=lambda n=n: self.run_prog_both(n)) \
             .grid(row=0, column=col, padx=2, pady=(6, 2), sticky="ew")

       ttk.Button(sync, text="同步關閉負載", command=self.set_load_off_both) \
           .grid(row=1, column=0, columnspan=6, padx=2, pady=(6, 2), sticky="ew")



    def make_ui(self):
        rootfrm = ttk.Frame(self.root)
        rootfrm.pack(fill="both", expand=True, padx=12, pady=12)
        rootfrm.columnconfigure(0, weight=1)
        rootfrm.columnconfigure(1, weight=1)

        self._build_device_panel(rootfrm, "設備 A", 0)
        self._build_device_panel(rootfrm, "設備 B", 1)
        self._sync_prog()
        self.scan_resources_all()

    def _on_close(self):
        for idx, p in enumerate(self.panels):
            if p["inst"]:
                try:
                    p["inst"].close()
                    logger.debug("[面板%d] 關閉 VISA 連線", idx)
                except Exception as e:
                    logger.warning("[面板%d] 關閉 VISA 連線例外：%s", idx, _fmt_exc(e))
            try:
                if p["excel_wb"]:
                    p["excel_wb"].close()
                if p["excel_app"]:
                    p["excel_app"].quit()
                logger.debug("[面板%d] 關閉 Excel", idx)
            except Exception as e:
                logger.warning("[面板%d] 關閉 Excel 例外：%s", idx, _fmt_exc(e))
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
