import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, Toplevel, Radiobutton, StringVar, CENTER
import pandas as pd
import matplotlib.pyplot as plt
import mplcursors  # 确保已安装 mplcursors 库

def center_window(window):
    """将窗口居中显示"""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")

def select_file():
    file_path = filedialog.askopenfilename(
        title="选择Excel文件",
        filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))
    )
    return file_path

def input_columns(num_columns):
    root.withdraw()
    column_names = []
    
    column_input_window = Toplevel(root)
    column_input_window.title("输入列名")
    
    entries = []
    for i in range(num_columns):
        tk.Label(column_input_window, text=f"列名 {i+1}:").grid(row=i, column=0, padx=10, pady=5)
        entry = tk.Entry(column_input_window)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entries.append(entry)
    
    def save_columns():
        for entry in entries:
            column_name = entry.get().strip()
            if column_name:
                column_names.append(column_name)
        column_input_window.destroy()
    
    tk.Button(column_input_window, text="确定", command=save_columns).grid(row=num_columns, column=0, columnspan=2, pady=10)
    center_window(column_input_window)
    column_input_window.wait_window()
    
    return column_names if len(column_names) == num_columns else None

def input_max_rows():
    root.withdraw()
    max_rows = simpledialog.askinteger("最大行数", "请输入需要读取的最大行数（输入0表示无限制）:", parent=root, minvalue=0)
    return max_rows

def input_x_axis_column(column_names):
    """让用户选择一个列名作为 X 轴"""
    x_axis_column = StringVar()
    x_axis_window = Toplevel(root)
    x_axis_window.title("选择 X 轴列名")
    
    tk.Label(x_axis_window, text="请选择一个列名作为 X 轴:").pack(pady=10)
    for column in column_names:
        tk.Radiobutton(x_axis_window, text=column, variable=x_axis_column, value=column).pack(anchor="w")
    
    def confirm_x_axis():
        if x_axis_column.get():
            x_axis_window.destroy()
        else:
            messagebox.showwarning("警告", "请选择一个列名作为 X 轴！")
    
    tk.Button(x_axis_window, text="确定", command=confirm_x_axis).pack(pady=10)
    center_window(x_axis_window)
    x_axis_window.wait_window()
    
    return x_axis_column.get()

def plot_line_chart(file_path, column_names, max_rows):
    try:
        if max_rows > 0:
            df = pd.read_excel(file_path, usecols=column_names, nrows=max_rows)
        else:
            df = pd.read_excel(file_path, usecols=column_names)
        
        if df.empty:
            messagebox.showwarning("警告", "读取的数据为空，请检查列名是否正确。")
            return
        
        # 选择 X 轴列名
        x_axis_column = input_x_axis_column(column_names)
        if not x_axis_column:
            return
        
        # 检查 X 轴列是否有重复数据
        if df[x_axis_column].duplicated().any():
            result = messagebox.askyesno("重复数据", f"X 轴列 '{x_axis_column}' 存在重复数据，是否需要聚合数据？")
            if result:
                # 对重复的 X 轴数据取均值
                df = df.groupby(x_axis_column).mean().reset_index()
            else:
                return

        x = df[x_axis_column]  # 使用聚合后的 X 轴数据
        plt.figure(figsize=(20, 8))
        for column in column_names:
            if column != x_axis_column:  # 排除 X 轴列
                plt.plot(x, df[column], label=column, marker='o')

        plt.xlabel(x_axis_column)
        plt.ylabel('Values')
        plt.title('Goal Chart')
        plt.legend()
        plt.grid(True)

        # 设置 X 轴标签旋转角度
        plt.xticks(rotation=45)

        # 添加鼠标悬停显示坐标功能
        cursor = mplcursors.cursor(hover=True)
        cursor.connect("add", lambda sel: sel.annotation.set_text(
            f"{x_axis_column}: {sel.target[0]}\nValue: {sel.target[1]:.2f}"
        ))

        plt.show()
        
    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {e}")

def plot_bar_chart(file_path, column_names, max_rows):
    try:
        if max_rows > 0:
            df = pd.read_excel(file_path, usecols=column_names, nrows=max_rows)
        else:
            df = pd.read_excel(file_path, usecols=column_names)
        
        if df.empty:
            messagebox.showwarning("警告", "读取的数据为空，请检查列名是否正确。")
            return
        
        plt.figure(figsize=(10, 6))
        sums = df.sum()
        plt.bar(column_names, sums, color='skyblue')
        plt.xlabel('Columns')
        plt.ylabel('Total Values')
        plt.title('Columns Total Values Bar Chart')
        plt.show()
        
    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {e}")

def plot_pie_chart(file_path, column_names, max_rows):
    try:
        if max_rows > 0:
            df = pd.read_excel(file_path, usecols=column_names, nrows=max_rows)
        else:
            df = pd.read_excel(file_path, usecols=column_names)
        
        if df.empty:
            messagebox.showwarning("警告", "读取的数据为空，请检查列名是否正确。")
            return
        
        sums = df.sum()
        plt.figure(figsize=(8, 8))
        plt.pie(sums, labels=sums.index, autopct='%1.1f%%', startangle=140)
        plt.title('Columns Total Values Pie Chart')
        plt.show()
        
    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {e}")

def main_ui():
    file_path = select_file()
    if not file_path:
        return
    num_columns = simpledialog.askinteger("列数", "请输入需要读取的列名数量:", parent=root, minvalue=1)
    if num_columns is None:
        return
    column_names = input_columns(num_columns)
    if column_names is None:
        return
    max_rows = input_max_rows()
    if max_rows is None:
        return
    
    chart_type_window = Toplevel(root)
    chart_type_window.title("选择图表类型")
    
    chart_type = StringVar(value="line")
    tk.Radiobutton(chart_type_window, text="折线图", variable=chart_type, value="line").pack(pady=5)
    tk.Radiobutton(chart_type_window, text="条形图", variable=chart_type, value="bar").pack(pady=5)
    tk.Radiobutton(chart_type_window, text="饼状图", variable=chart_type, value="pie").pack(pady=5)
    
    def confirm_chart_type():
        chart_type_window.destroy()
    
    tk.Button(chart_type_window, text="确定", command=confirm_chart_type).pack(pady=10)
    center_window(chart_type_window)
    chart_type_window.wait_window()
    
    if chart_type.get() == "line":
        plot_line_chart(file_path, column_names, max_rows)
    elif chart_type.get() == "bar":
        plot_bar_chart(file_path, column_names, max_rows)
    else:
        plot_pie_chart(file_path, column_names, max_rows)

root = tk.Tk()
root.title("图表工具")

main_ui()
root.destroy()