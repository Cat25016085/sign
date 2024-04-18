import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 設置 Google 試算表 API 憑證
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(credentials)

# 打開 Google 試算表
spreadsheet_id = '1NHgXUoz1Dg1ov8iI81h0ifc3_2uRobxhiWp2Czl6Zw8'
spreadsheet = client.open_by_key(spreadsheet_id)
sheet = spreadsheet.sheet1

# 創建主視窗
root = tk.Tk()
root.title("簽到系統")

# 輸入資訊並簽到函數
def sign_in():
    class_name = class_var.get()
    seat_number = seat_var.get()
    name = name_entry.get()
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 檢查是否存在相同的資料
    existing_data = sheet.get_all_records()
    for row in existing_data:
        if row["班級"] == class_name and row["座號"] == seat_number:
            messagebox.showinfo("資料核對", f"已存在相同的資料：\n班級：{class_name}\n座號：{seat_number}\n姓名：{row['姓名']}\n簽到時間：{row['簽到時間']}")
            return

    # 將資料寫入試算表
    data = [class_name, seat_number, name, current_time]
    sheet.append_row(data)
    messagebox.showinfo("成功簽到", "簽到成功")

# UI設計
class_label = tk.Label(root, text="班級：")
class_label.place(relx=0.5, rely=0.4, anchor=tk.CENTER)
class_var = tk.StringVar()
class_dropdown = ttk.Combobox(root, textvariable=class_var, values=[f"{i}" for i in range(301, 321)])
class_dropdown.place(relx=0.6, rely=0.4, anchor=tk.CENTER)

seat_label = tk.Label(root, text="座號：")
seat_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
seat_var = tk.StringVar()
seat_dropdown = ttk.Combobox(root, textvariable=seat_var, values=[f"{i}" for i in range(1, 51)])
seat_dropdown.place(relx=0.6, rely=0.5, anchor=tk.CENTER)

name_label = tk.Label(root, text="姓名：")
name_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)
name_entry = tk.Entry(root)
name_entry.place(relx=0.6, rely=0.6, anchor=tk.CENTER)

sign_in_button = tk.Button(root, text="簽到", command=sign_in)
sign_in_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

# 設置視窗大小和位置
window_width = 400
window_height = 300
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_coordinate = (screen_width / 2) - (window_width / 2)
y_coordinate = (screen_height / 2) - (window_height / 2)
root.geometry(f'{window_width}x{window_height}+{int(x_coordinate)}+{int(y_coordinate)}')

# 啟動主視窗
root.mainloop()
