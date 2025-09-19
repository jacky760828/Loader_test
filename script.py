import pyvisa
import tkinter
import data_read
import time
import xlwings as xw

# 初始化 Resource Manager


def send_parameter_by_index( sheet_idx_1based,count_idx_1based,inst=None):
    sheet = sheets[sheet_idx_1based ]
    #sheet = sheets[10]
    row_count = data_read.ex_content[count_idx_1based - 1]
    #row_count = data_read.ex_content[9]
    prog_data_list = data_read.generate_prog_data2(sheet, row_count)

    for line in prog_data_list:
      print(line)
      cmd = (line or "").strip()
      cmd = line.strip()
      if cmd:  # 跳過空行
       if inst is not None:
          inst.write(cmd)
       print(f"已送出: {cmd}")

    inst.write("PROG:SAV")
    print("程式已儲存")

def send_parameter():
  send_parameter_by_index(1, 1, inst)   # 有設備
  time.sleep(0.2)
  send_parameter_by_index(2, 2,inst)
  time.sleep(0.2)
  send_parameter_by_index(3, 3,inst)
  time.sleep(0.2)
  send_parameter_by_index(4, 4,inst)
  time.sleep(0.2)
  send_parameter_by_index(5, 5,inst)
  time.sleep(0.2)
  send_parameter_by_index(6, 6,inst)
  time.sleep(0.2)
  send_parameter_by_index(7, 7,inst)
  time.sleep(0.2)
  send_parameter_by_index(8, 8,inst)
  time.sleep(0.2)
  send_parameter_by_index(9, 9,inst)
  time.sleep(0.2)
  send_parameter_by_index(10, 10,inst)
def SET_Meth1():
        resp = inst.query("PROG:DATA:LIST? 1,1")
        print("程式1序列1參數:",resp.strip())
        resp = inst.query("PROG:DATA:LIST? 1,2")
        print("程式1序列1參數:",resp.strip())
        resp = inst.query("PROG:NSEL?")
        print("程式1序列1參數:", resp.strip())
        inst.write("PROG:NSEL 1")
        inst.write("PROG:RUN")
        inst.write("LOAD ON")
        inst.write("SYST:LOC")


def SET_Meth2():
    resp = inst.query("PROG:DATA:LIST? 2,1")
    print("程式1序列1參數:", resp.strip())
    resp = inst.query("PROG:DATA:LIST? 2,2")
    print("程式1序列1參數:", resp.strip())

    resp = inst.query("PROG:NSEL?")
    print("程式1序列1參數:", resp.strip())
    inst.write("PROG:NSEL 2")
    inst.write("PROG:RUN")
    inst.write("LOAD ON")
    # print(f"已設定 {current} A 拉載並啟動")
    inst.write("SYST:LOC")


def SET_Meth3():
    resp = inst.query("PROG:DATA:LIST? 3,1")
    print("程式1序列1參數:", resp.strip())
    resp = inst.query("PROG:DATA:LIST? 3,2")
    print("程式1序列1參數:", resp.strip())
    resp = inst.query("PROG:DATA:LIST? 3,3")
    print("程式1序列1參數:", resp.strip())

    resp = inst.query("PROG:NSEL?")
    print("程式1序列1參數:", resp.strip())
    inst.write("PROG:NSEL 3")
    inst.write("PROG:RUN")
    inst.write("LOAD ON")
    # print(f"已設定 {current} A 拉載並啟動")
    inst.write("SYST:LOC")


def SET_Meth4():
    resp = inst.query("PROG:DATA:LIST? 4,1")
    print("程式1序列1參數:", resp.strip())
    resp = inst.query("PROG:DATA:LIST? 4,2")
    print("程式1序列1參數:", resp.strip())
    resp = inst.query("PROG:DATA:LIST? 4,3")
    print("程式1序列1參數:", resp.strip())

    resp = inst.query("PROG:NSEL?")
    print("程式1序列1參數:", resp.strip())
    inst.write("PROG:NSEL 4")
    inst.write("PROG:RUN")
    inst.write("LOAD ON")
    # print(f"已設定 {current} A 拉載並啟動")
    inst.write("SYST:LOC")

def SET_Meth5():
    resp = inst.query("PROG:DATA:LIST? 5,1")
    print("程式1序列1參數:", resp.strip())
    resp = inst.query("PROG:DATA:LIST? 5,2")
    print("程式1序列1參數:", resp.strip())
    resp = inst.query("PROG:DATA:LIST? 5,3")
    print("程式1序列1參數:", resp.strip())

    resp = inst.query("PROG:NSEL?")
    print("程式1序列1參數:", resp.strip())
    inst.write("PROG:NSEL 5")
    inst.write("PROG:RUN")
    inst.write("LOAD ON")
    # print(f"已設定 {current} A 拉載並啟動")
    inst.write("SYST:LOC")


def SET_Meth6():
    resp = inst.query("PROG:DATA:LIST? 6,1")
    print("程式1序列1參數:", resp.strip())
    resp = inst.query("PROG:DATA:LIST? 6,2")
    print("程式1序列1參數:", resp.strip())
    resp = inst.query("PROG:DATA:LIST? 6,3")
    print("程式1序列1參數:", resp.strip())

    resp = inst.query("PROG:NSEL?")
    print("程式1序列1參數:", resp.strip())
    inst.write("PROG:NSEL 6")
    inst.write("PROG:RUN")
    inst.write("LOAD ON")
    # print(f"已設定 {current} A 拉載並啟動")
    inst.write("SYST:LOC")


def SET_Meth7():
    resp = inst.query("PROG:DATA:LIST? 7,1")
    print("程式1序列1參數:", resp.strip())
    resp = inst.query("PROG:DATA:LIST? 7,2")
    print("程式1序列1參數:", resp.strip())
    resp = inst.query("PROG:DATA:LIST? 7,3")
    print("程式1序列1參數:", resp.strip())

    resp = inst.query("PROG:NSEL?")
    print("程式1序列1參數:", resp.strip())
    inst.write("PROG:NSEL 7")
    inst.write("PROG:RUN")
    inst.write("LOAD ON")
    # print(f"已設定 {current} A 拉載並啟動")
    inst.write("SYST:LOC")

def SET_Meth8():
    resp = inst.query("PROG:DATA:LIST? 1,1")
    print("程式1序列1參數:", resp.strip())

    resp = inst.query("PROG:NSEL?")
    print("程式1序列1參數:", resp.strip())
    inst.write("PROG:NSEL 8")
    inst.write("PROG:RUN")
    inst.write("LOAD ON")
    # print(f"已設定 {current} A 拉載並啟動")
    inst.write("SYST:LOC")


def SET_Meth9():
    resp = inst.query("PROG:DATA:LIST? 1,1")
    print("程式1序列1參數:", resp.strip())

    resp = inst.query("PROG:NSEL?")
    print("程式1序列1參數:", resp.strip())
    inst.write("PROG:NSEL 9")
    inst.write("PROG:RUN")
    inst.write("LOAD ON")
    # print(f"已設定 {current} A 拉載並啟動")
    inst.write("SYST:LOC")

def SET_Meth10():
    resp = inst.query("PROG:DATA:LIST? 1,1")
    print("程式1序列1參數:", resp.strip())

    resp = inst.query("PROG:NSEL?")
    print("程式1序列1參數:", resp.strip())
    inst.write("PROG:NSEL 10")
    inst.write("PROG:RUN")
    inst.write("LOAD ON")
    # print(f"已設定 {current} A 拉載並啟動")
    inst.write("SYST:LOC")
def set_load_OFF():
    inst.write("LOAD 0")
    inst.write("SYST:LOC")

global inst
try:
 rm = pyvisa.ResourceManager()
 resources = rm.list_resources()
 print("偵測到的 VISA 裝置：", resources)
 usb_str = resources[0]
 print(f"使用裝置: {usb_str}")
 inst = rm.open_resource(usb_str)
except Exception as e:
 print("讀取裝置失敗，錯誤訊息：", e)
 resources = []  # 確保後續程式能用
# 顯示可用裝置

root = tkinter.Tk()
root.title("LOAD測試")
root.geometry("800x400")
app = xw.App(visible=False, add_book=False)
wb = app.books.open("執行檔.xlsx")
sheets = [wb.sheets[i] for i in range(min(11, len(wb.sheets)))]
# 開啟第一個偵測到的裝置


for i in range(1, 11):
 if inst is not None:
    inst.write(f"PROG:SEQ:CLE {i}")
    print(f"[送出]PROG:SEQ:CLE {i}")
 else:
    print(f"[模擬]PROG:SEQ:CLE {i}")





prog_data_list = data_read.generate_prog_data(sheets[0])
for line in prog_data_list:
    cmd =line.strip()
    if cmd:  # 跳過空行
     inst.write(cmd)
     #print(f"已送出: {cmd}")
    #print(data_read.ex_content)
# 按鈕事件函式
#wb = app.books.open("執行檔.xlsx")

# UI 控件
#entry = tkinter.Entry(width=20, font=("Times New Roman", 18))
#entry.place(x=10, y=10)
button = tkinter.Button(root, text="載入數據", font=("Times New Roman", 12), command=send_parameter)
button.place(x=10, y=10)

button = tkinter.Button(root, text="設定第一步", font=("Times New Roman", 12), command=SET_Meth1)
button.place(x=10, y=60)
button = tkinter.Button(root, text="設定第二步", font=("Times New Roman", 12), command=SET_Meth2)
button.place(x=110, y=60)
button = tkinter.Button(root, text="設定第三步", font=("Times New Roman", 12), command=SET_Meth3)
button.place(x=210, y=60)
button = tkinter.Button(root, text="設定第四步", font=("Times New Roman", 12), command=SET_Meth4)
button.place(x=310, y=60)
button = tkinter.Button(root, text="設定第五步", font=("Times New Roman", 12), command=SET_Meth5)
button.place(x=410, y=60)
button = tkinter.Button(root, text="設定第六步", font=("Times New Roman", 12), command=SET_Meth6)
button.place(x=510, y=60)
button = tkinter.Button(root, text="設定第七步", font=("Times New Roman", 12), command=SET_Meth7)
button.place(x=610, y=60)
button = tkinter.Button(root, text="設定第八步", font=("Times New Roman", 12), command=SET_Meth8)
button.place(x=10, y=100)
button = tkinter.Button(root, text="設定第九步", font=("Times New Roman", 12), command=SET_Meth9)
button.place(x=110, y=100)
button = tkinter.Button(root, text="設定第10步", font=("Times New Roman", 12), command=SET_Meth10)
button.place(x=210, y=100)
button = tkinter.Button(root, text="關掉負載", font=("Times New Roman", 18), command=set_load_OFF)
button.place(x=10, y=180)

root.mainloop()
# 查詢目前作用通道