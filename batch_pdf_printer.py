# -*- coding: utf-8 -*-
"""
file: batch_pdf_printer.py
create: 2023-12-28 17:39
author: s0ing
summary: 
"""
import os
import tempfile

import requests
import win32print
import win32api
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def print_pdf(printer, file_path):
    """
    打印单个PDF文件
    """
    try:
        win32api.ShellExecute(0, "print", file_path, f'/d:"{printer}"', ".", 0)
    except Exception as exp:
        messagebox.showinfo("出错了~", f"打印文件 {file_path} 时出错: {str(exp)} \ntips: 请确保打印机已连接并正常工作/或者打开文件是否正常显示。")


def select_folder():
    """
    选择包含文件的文件夹并更新文件列表
    """
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_path.set(folder_selected)
        update_file_list()


def start_printing():
    """
    批量打印选定文件夹中的PDF文件
    """
    folder = folder_path.get()
    printer_name = win32print.GetDefaultPrinter()  # 获取默认打印机
    files_to_print = [file for file in os.listdir(folder) if file.lower().endswith('.pdf')]

    if not files_to_print:
        messagebox.showinfo("信息", "文件夹中没有PDF文件")
        return

    threading.Thread(target=print_files, args=(printer_name, folder, files_to_print), daemon=True).start()


def print_files(printer_name, folder, files):
    """
    执行打印任务并在完成后显示通知
    """
    for file in files:
        file_path = os.path.join(folder, file)
        print_pdf(printer_name, file_path)
    root.after(0, notify_printing_done)


def notify_printing_done():
    """
    在所有文件打印完成后显示通知
    """
    messagebox.showinfo("成功", "任务已发送，请稍后查看打印机状态。")


def update_file_list():
    """
    更新文件列表显示文件夹中的PDF文件
    """
    folder = folder_path.get()
    file_list.delete(0, tk.END)
    count = 0
    for file in os.listdir(folder):
        if file.lower().endswith('.pdf'):
            file_list.insert(tk.END, file)
            count += 1
    file_count_label.config(text=f"文件总数: {count}", foreground="green")


def download_icon():
    """
    从给定的URL下载图标并返回图标的文件路径。
    """
    try:
        response = requests.get('https://gitee.com/s0ing/imgs/raw/master/ati3o-7gcvk-001.ico', timeout=2)
        if response.status_code == 200:
            # 创建一个临时文件
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.ico')
            temp_file.write(response.content)
            temp_file.close()
            return temp_file.name
    except:
        pass

    return None  # 如果下载失败，则返回None


fonts = "楷体"
LARGE_FONT = (fonts, 12)
NORMAL_FONT = (fonts, 10)
BUTTON_FONT = (fonts, 10, "bold")
BG_COLOR = "#f0f0f0"
BUTTON_COLOR = "#00bfff"
LABEL_COLOR = "green"

# 创建主窗口
root = tk.Tk()
root.title("Batch PDF Printer")
root.configure(bg=BG_COLOR)

# 下载并设置图标
icon_path = download_icon()
if icon_path:
    root.iconbitmap(default=icon_path)

# 创建并配置样式
style = ttk.Style()
style.configure("BW.TLabel", foreground="black", background="white")
style.configure("Large.TLabelframe.Label", font=LARGE_FONT)  # 设置 LabelFrame 标题的字体
style.configure("TButton", font=BUTTON_FONT, background=BUTTON_COLOR)
style.configure("TLabel", font=NORMAL_FONT, background=BG_COLOR)

# 窗口初始化时置顶
root.attributes('-topmost', True)
root.after(500, lambda: root.attributes('-topmost', False))

# 获取屏幕尺寸并计算窗口尺寸
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = screen_width // 3
window_height = screen_height // 3

# 设置窗口初始尺寸和位置
root.geometry(f'{window_width}x{window_height}+{screen_width // 2 - window_width // 2}+{screen_height // 2 - window_height // 2}')

# 文件夹路径
folder_path = tk.StringVar()

# 文件列表框架
file_list_frame = ttk.LabelFrame(root, text="选择的文件夹中的PDF文件", style="Large.TLabelframe")
file_list_frame.pack(padx=10, pady=10, fill='both', expand=True)

# 文件列表
file_list = tk.Listbox(file_list_frame)
file_list.pack(side=tk.LEFT, fill='both', expand=True)
scrollbar = ttk.Scrollbar(file_list_frame, orient='vertical', command=file_list.yview)
scrollbar.pack(side=tk.LEFT, fill='y')
file_list.config(yscrollcommand=scrollbar.set)

# 控制按钮和文件总数标签的框架
control_frame = ttk.Frame(root)
control_frame.pack(padx=10, pady=10, fill='x', expand=True)

# 使用Grid布局管理器
control_frame.columnconfigure(0, weight=1)
control_frame.columnconfigure(1, weight=1)
control_frame.columnconfigure(2, weight=1)

select_folder_button = ttk.Button(control_frame, text="选择文件夹", command=select_folder)
select_folder_button.grid(row=0, column=0, sticky='ew')

file_count_label = ttk.Label(control_frame, text="文件总数: 0")
file_count_label.grid(row=0, column=1)

print_button = ttk.Button(control_frame, text="开始打印", command=start_printing)
print_button.grid(row=0, column=2, sticky='ew')

# 运行主循环
root.mainloop()
