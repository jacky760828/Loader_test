import pyvisa
import tkinter

# 初始化 Resource Manager
root = tkinter.Tk()
root.title("LOAD測試")
root.geometry("800x400")

rm = pyvisa.ResourceManager()
resources = rm.list_resources()

# 顯示可用裝置
print("偵測到的 VISA 裝置：", resources)

# 開啟第一個偵測到的裝置
usb_str = resources[0]
print(f"使用裝置: {usb_str}")
inst = rm.open_resource(usb_str)


# 按鈕事件函式
def set_load():
    try:
        # 從輸入框抓取文字並轉成數字
        value = entry.get().strip()
        if not value:
            print("請輸入電流值！")
            return

        current = float(value)  # 轉成浮點數 (允許小數點)

        # 設定拉載
        inst.write(f"CURRent:STATic:L1 {current}")
        inst.write("MODE CCH")
        inst.write("LOAD ON")
        print(f"已設定 {current} A 拉載並啟動")
    except ValueError:
        print("輸入錯誤！請輸入數字")
    except Exception as e:
        print(f"錯誤：{e}")

def set_load_OFF():
    inst.write("LOAD 0")


# UI 控件
entry = tkinter.Entry(width=20, font=("Times New Roman", 18))
entry.place(x=10, y=10)

button = tkinter.Button(root, text="設定電流並啟動拉載", font=("Times New Roman", 18), command=set_load)
button.place(x=10, y=60)
button = tkinter.Button(root, text="關掉負載", font=("Times New Roman", 18), command=set_load_OFF)
button.place(x=10, y=180)

root.mainloop()
# 查詢目前作用通道
