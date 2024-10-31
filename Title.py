import tkinter as tk
from tkinter import messagebox

def generate_report_id():
    # 获取输入框内容
    report_name = report_name_entry.get()
    reporter = reporter_entry.get()
    date = date_entry.get()
    time = time_entry.get()
    location = location_entry.get()
    report_number = report_number_entry.get()
    
    # 拼接字符串
    report_id = f"{report_name}_{reporter}_{date}_{time}_{location}_{report_number}"
    
    # 显示结果
    messagebox.showinfo("报告ID", report_id)
    
    # 将结果复制到剪贴板
    root.clipboard_clear()
    root.clipboard_append(report_id)

# 创建主窗口
root = tk.Tk()
root.title("报告信息生成器")

# 创建输入框和标签
tk.Label(root, text="报告名称：").grid(row=0, column=0)
report_name_entry = tk.Entry(root)
report_name_entry.grid(row=0, column=1)

tk.Label(root, text="报告人：").grid(row=1, column=0)
reporter_entry = tk.Entry(root)
reporter_entry.grid(row=1, column=1)

tk.Label(root, text="日期号：").grid(row=2, column=0)
date_entry = tk.Entry(root)
date_entry.grid(row=2, column=1)

tk.Label(root, text="具体时间：").grid(row=3, column=0)
time_entry = tk.Entry(root)
time_entry.grid(row=3, column=1)

tk.Label(root, text="地点：").grid(row=4, column=0)
location_entry = tk.Entry(root)
location_entry.grid(row=4, column=1)

tk.Label(root, text="报告期号：").grid(row=5, column=0)
report_number_entry = tk.Entry(root)
report_number_entry.grid(row=5, column=1)

# 创建按钮
generate_button = tk.Button(root, text="生成报告ID", command=generate_report_id)
generate_button.grid(row=6, column=0, columnspan=2)

# 运行主循环
root.mainloop()