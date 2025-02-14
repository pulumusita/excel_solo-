
import tkinter as tk
from tkinter import filedialog, messagebox
from otherSoloExcel import *
import os

# 用户点击执行按钮后执行的操作
def submit_action():

    # 获取用户输入，output_path为为汇总文件保存的路径
    global first_search_text, folder_path, sheet_name, output_path
    first_search_text = entry_text.get()
    folder_path = entry_folder.get()
    #sheet_name = entry_text.get() # 用户输入需要处理的sheet_name
    
    # 使用用户输入的文字和文件夹路径
    print("搜索文字:", first_search_text)
    print("文件夹路径:", folder_path)

    # 清空输入框
    entry_text.delete(0, tk.END)
    entry_folder.delete(0, tk.END)


    # 调用check_summary_file_exists函数检查汇总文件是否存在，如果存在则提示用户删除后重新运行程序
    output_path = check_summary_file_exists()
    if output_path is None:
        return
    

    # 汇总.xls文件，则调用create_or_open_excel函数创建汇总文件到上面检查过的文件夹
    create_or_open_excel(output_path)

    # 执行主处理函数，对需要处理的文件夹内的xls和xlsx文件进行处理
    traverse_folder_and_test(folder_path, first_search_text,output_path)
    
    # 执行求和汇总函数
    sum_xlsx_columns(output_path)
    
    # 清空所有输入框
    entry_text.config(state='normal')# 恢复输入框原始状态

    init_file_choose()# 初始化文件选择框
    

# 初始化文件选择框
def init_file_choose():
    global entry_text, entry_folder
    entry_folder.config(state='normal')  # 设置文本框为可编辑状态
    entry_folder.delete(0, tk.END)
    entry_folder.config(width=25)  # 恢复原有宽度
    entry_folder.config(state='readonly')  # 设置文本框为只读



# 选择文件夹方法
def choose_folder():
    global folder_path
    # 使用filedialog模块的askdirectory函数打开文件夹选择对话框，并将选择的文件夹路径赋值给folder_path
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_folder.config(state='normal')  # 设置文本框为可编辑状态
        entry_folder.delete(0, tk.END)
        entry_folder.insert(0, folder_path)
        entry_folder.config(state='readonly')  # 设置文本框为只读
    path_length = len(folder_path)
    # 计算文件夹路径的长度
    entry_folder.config(width=path_length)  # 设置文本框的宽度为文件夹路径的长度

# 用户选择文件夹并检查汇总文件是否存在，如果不存在则说明该路径可用并返回路径
def check_summary_file_exists():
    # 弹出文件夹选择对话框
    folder_path = filedialog.askdirectory()
    # 检查用户是否选择了文件夹
    if not folder_path:
        print("没有选择文件夹")
        messagebox.showerror("错误", "没有选择文件夹，请选择一个文件夹用于保存汇总文件")
        return

    # 检查文件夹中是否存在汇总.xlsx文件
    summary_file_path = os.path.join(folder_path, '汇总.xlsx')
    summary_xls_file_path = os.path.join(folder_path, '汇总.xls')
    if os.path.exists(summary_file_path) or os.path.exists(summary_xls_file_path):
        messagebox.showerror("错误", "汇总文件已存在，请删除后重新运行程序")
        return

     # 确保folder_path以/结尾
    if not folder_path.endswith(os.path.sep):
        folder_path += os.path.sep
    # 如果汇总文件不存在，则返回文件夹路径
    print(folder_path)
    return folder_path



# 创建主窗口
root = tk.Tk()
# 将窗口居中
root.eval('tk::PlaceWindow %s center' % root.winfo_toplevel())
root.title("数据处理程序第一代")
root.geometry("500x500")  # 设置窗口大小为400x200像素

prompt_text = "警告：在使用本程序时请不要打开需要处理的表格文件，否则会导致程序错误"
tk.Label(root, text=prompt_text, fg='red', wraplength=300).pack(side=tk.TOP, fill=tk.X,pady=10)


# 创建标签和输入框用于搜索文本
tk.Label(root, text="输入要搜索的字段，例如：期货持仓汇总:").pack()
entry_text = tk.Entry(root,width=25)
entry_text.pack(pady=10)

# # 这里如果需要选择sheet_name可以将这个页面加进去（创建标签和输入框用于用户输入sheet_name）我这里默认第一个sheet
# tk.Label(root, text="输入要处理的sheet_name,例如客户交易结算日报:").pack()
# entry_sheet_name = tk.Entry(root)
# entry_sheet_name.pack(pady=10)

# 创建标签和按钮用于选择要处理的文件夹路径
tk.Label(root, text="请选择要处理的xls所在的文件夹:").pack()
entry_folder = tk.Entry(root,width=25)
entry_folder.pack(pady=10)
entry_folder.config(state='readonly')# 将entry_folder设置为只读状态，用户不能直接在输入框中修改内容
choose_folder_btn = tk.Button(root, text="选择文件夹", command=choose_folder)
choose_folder_btn.pack()

# 创建按钮，点击后获取输入并关闭窗口
submit_btn = tk.Button(root, text="处理数据并保存文件", command=submit_action)
submit_btn.pack(side=tk.BOTTOM,pady=10)


# 运行主循环
root.mainloop()





