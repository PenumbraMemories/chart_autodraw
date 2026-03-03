import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
import mplcursors  # 确保已安装 mplcursors 库
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

# 解决matplotlib显示中文问题
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

# 设置图表字体大小和样式
plt.rcParams['font.size'] = 14
plt.rcParams['axes.titlesize'] = 20
plt.rcParams['axes.labelsize'] = 16
plt.rcParams['xtick.labelsize'] = 14
plt.rcParams['ytick.labelsize'] = 14
plt.rcParams['legend.fontsize'] = 14
plt.rcParams['figure.titlesize'] = 22

# 设置应用程序样式
APP_COLOR = {
    "bg": "#f5f7fa",
    "fg": "#2c3e50",
    "button_bg": "#3498db",
    "button_fg": "white",
    "button_hover": "#2980b9",
    "select_bg": "#e3f2fd",
    "title_bg": "#2c3e50",
    "title_fg": "white",
    "frame_bg": "#ffffff",
    "entry_bg": "#ffffff",
    "entry_fg": "#2c3e50",
    "label_fg": "#34495e",
    "radio_bg": "#f5f7fa",
    "radio_select": "#3498db"
}

# 设置字体
APP_FONT = {
    "title": ("Microsoft YaHei", 22, "bold"),
    "normal": ("Microsoft YaHei", 14),
    "button": ("Microsoft YaHei", 14, "bold"),
    "label": ("Microsoft YaHei", 14),
    "tree": ("Microsoft YaHei", 12),
    "combo": ("Microsoft YaHei", 12),
    "listbox": ("Microsoft YaHei", 12)
}

# 全局变量
file_path = None
df = None
selected_columns = []
max_rows = 0

class ChartApp:
    def __init__(self, root):
        self.root = root
        self.root.title("数据可视化工具")
        self.root.state("zoom")  # Windows系统下最大化窗口
        self.root.configure(bg=APP_COLOR["bg"])

        # 创建主框架
        self.main_frame = tk.Frame(root, bg=APP_COLOR["bg"])
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建可拉伸分割面板
        self.paned_window = tk.PanedWindow(self.main_frame, orient=tk.HORIZONTAL, bg=APP_COLOR["bg"],
                                         sashwidth=8, sashrelief=tk.RAISED, showhandle=True,
                                         handlesize=8, opaqueresize=False)
        self.paned_window.pack(fill=tk.BOTH, expand=True)

        # 创建左侧控制面板
        self.control_panel = tk.LabelFrame(self.paned_window, text="控制面板", 
                                          font=APP_FONT["title"], 
                                          bg=APP_COLOR["frame_bg"], fg=APP_COLOR["title_fg"])
        # 设置控制面板初始宽度
        self.control_panel.config(width=350)

        # 创建右侧图表显示区域
        self.chart_panel = tk.LabelFrame(self.paned_window, text="图表显示", 
                                        font=APP_FONT["title"], 
                                        bg=APP_COLOR["frame_bg"], fg=APP_COLOR["title_fg"])
        # 将两个面板添加到分割窗口
        self.paned_window.add(self.control_panel, minsize=250)
        self.paned_window.add(self.chart_panel, minsize=400)
        
        # 绑定面板大小改变事件，使用延迟刷新避免卡顿
        self.paned_window.bind("<ButtonRelease-1>", lambda e: self.schedule_preview_update())
        self.paned_window.bind("<B1-Motion>", lambda e: self.schedule_preview_update())

        # 初始化控制面板
        self.init_control_panel()

        # 初始化图表区域
        self.init_chart_panel()

        # 状态栏
        self.status_bar = tk.Label(root, text="就绪", bd=1, relief=tk.SUNKEN, anchor=tk.W,
                                 bg=APP_COLOR["title_bg"], fg=APP_COLOR["title_fg"])
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def init_control_panel(self):
        # 文件选择区域
        file_frame = tk.Frame(self.control_panel, bg=APP_COLOR["frame_bg"])
        file_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(file_frame, text="数据文件:", font=APP_FONT["label"], 
                bg=APP_COLOR["frame_bg"], fg=APP_COLOR["label_fg"]).pack(anchor=tk.W)

        self.file_path_var = tk.StringVar()
        self.file_entry = tk.Entry(file_frame, textvariable=self.file_path_var, 
                                  bg=APP_COLOR["entry_bg"], fg=APP_COLOR["entry_fg"])
        self.file_entry.pack(fill=tk.X, pady=5)

        tk.Button(file_frame, text="浏览...", command=self.select_file,
                 bg=APP_COLOR["button_bg"], fg=APP_COLOR["button_fg"],
                 font=APP_FONT["button"]).pack(fill=tk.X, pady=5)

        # 数据预览区域
        preview_frame = tk.Frame(self.control_panel, bg=APP_COLOR["frame_bg"])
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tk.Label(preview_frame, text="数据预览:", font=APP_FONT["label"],
                bg=APP_COLOR["frame_bg"], fg=APP_COLOR["label_fg"]).pack(anchor=tk.W)

        # 创建预览表格
        self.preview_frame = tk.Frame(preview_frame, bg=APP_COLOR["frame_bg"])
        self.preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # 使用Treeview显示数据，增大字体
        style = ttk.Style()
        style.configure("Treeview", font=APP_FONT["tree"])
        style.configure("Treeview.Heading", font=(APP_FONT["tree"][0], APP_FONT["tree"][1], "bold"))

        self.tree = ttk.Treeview(self.preview_frame)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 添加滚动条，初始时隐藏
        self.scrollbar_y = ttk.Scrollbar(self.preview_frame, orient="vertical", command=self.tree.yview)
        self.scrollbar_x = ttk.Scrollbar(self.preview_frame, orient="horizontal", command=self.tree.xview)
        
        # 初始时隐藏滚动条
        self.tree.configure(yscrollcommand=self.on_y_scroll, xscrollcommand=self.on_x_scroll)

        # 图表设置区域
        settings_frame = tk.Frame(self.control_panel, bg=APP_COLOR["frame_bg"])
        settings_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(settings_frame, text="图表设置:", font=APP_FONT["label"],
                bg=APP_COLOR["frame_bg"], fg=APP_COLOR["label_fg"]).pack(anchor=tk.W)

        # 图表类型选择
        chart_type_frame = tk.Frame(settings_frame, bg=APP_COLOR["frame_bg"])
        chart_type_frame.pack(fill=tk.X, pady=5)

        tk.Label(chart_type_frame, text="图表类型:", font=APP_FONT["label"],
                bg=APP_COLOR["frame_bg"], fg=APP_COLOR["label_fg"]).pack(anchor=tk.W)

        self.chart_type_var = tk.StringVar(value="line")
        tk.Radiobutton(chart_type_frame, text="折线图", variable=self.chart_type_var, value="line",
                      bg=APP_COLOR["radio_bg"], selectcolor=APP_COLOR["radio_select"]).pack(anchor=tk.W)
        tk.Radiobutton(chart_type_frame, text="条形图", variable=self.chart_type_var, value="bar",
                      bg=APP_COLOR["radio_bg"], selectcolor=APP_COLOR["radio_select"]).pack(anchor=tk.W)
        tk.Radiobutton(chart_type_frame, text="饼图", variable=self.chart_type_var, value="pie",
                      bg=APP_COLOR["radio_bg"], selectcolor=APP_COLOR["radio_select"]).pack(anchor=tk.W)

        # X轴选择
        x_axis_frame = tk.Frame(settings_frame, bg=APP_COLOR["frame_bg"])
        x_axis_frame.pack(fill=tk.X, pady=5)

        tk.Label(x_axis_frame, text="X轴:", font=APP_FONT["label"],
                bg=APP_COLOR["frame_bg"], fg=APP_COLOR["label_fg"]).pack(anchor=tk.W)

        self.x_axis_var = tk.StringVar()

        # 设置Combobox样式
        combo_style = ttk.Style()
        combo_style.configure("TCombobox", fieldbackground=APP_COLOR["entry_bg"], 
                             font=APP_FONT["combo"])

        self.x_axis_combo = ttk.Combobox(x_axis_frame, textvariable=self.x_axis_var, state="readonly")
        self.x_axis_combo.pack(fill=tk.X, pady=2)

        # Y轴选择
        y_axis_frame = tk.Frame(settings_frame, bg=APP_COLOR["frame_bg"])
        y_axis_frame.pack(fill=tk.X, pady=5)

        tk.Label(y_axis_frame, text="Y轴(可多选):", font=APP_FONT["label"],
                bg=APP_COLOR["frame_bg"], fg=APP_COLOR["label_fg"]).pack(anchor=tk.W)

        # 创建一个框架来容纳列表框和滚动条
        listbox_frame = tk.Frame(y_axis_frame, bg=APP_COLOR["frame_bg"])
        listbox_frame.pack(fill=tk.X, pady=2)

        # 创建列表框用于多选Y轴
        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.y_axis_listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, 
                                         yscrollcommand=scrollbar.set,
                                         bg=APP_COLOR["entry_bg"], fg=APP_COLOR["entry_fg"])
        self.y_axis_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.y_axis_listbox.yview)

        # 最大行数设置
        max_rows_frame = tk.Frame(settings_frame, bg=APP_COLOR["frame_bg"])
        max_rows_frame.pack(fill=tk.X, pady=5)

        tk.Label(max_rows_frame, text="最大行数(0=无限制):", font=APP_FONT["label"],
                bg=APP_COLOR["frame_bg"], fg=APP_COLOR["label_fg"]).pack(anchor=tk.W)

        self.max_rows_var = tk.IntVar(value=0)
        self.max_rows_spinbox = tk.Spinbox(max_rows_frame, from_=0, to=10000, 
                                          textvariable=self.max_rows_var,
                                          bg=APP_COLOR["entry_bg"], fg=APP_COLOR["entry_fg"])
        self.max_rows_spinbox.pack(fill=tk.X, pady=2)

        # 创建图表按钮
        tk.Button(settings_frame, text="创建图表", command=self.create_chart,
                 bg=APP_COLOR["button_bg"], fg=APP_COLOR["button_fg"],
                 font=APP_FONT["button"]).pack(fill=tk.X, pady=10)

    def init_chart_panel(self):
        # 创建matplotlib图形，进一步增大默认尺寸
        self.fig, self.ax = plt.subplots(figsize=(12, 8))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_panel)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 添加工具栏
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.chart_panel)
        self.toolbar.update()
        
        # 绑定图表面板大小改变事件，使用延迟刷新避免卡顿
        self.chart_panel.bind("<Configure>", lambda e: self.schedule_chart_update())

        # 初始时显示一个空白图表
        self.ax.text(0.5, 0.5, "请选择数据文件并设置参数", 
                    horizontalalignment='center', verticalalignment='center',
                    transform=self.ax.transAxes, fontsize=18)
        self.canvas.draw()

    def select_file(self):
        global file_path, df, selected_columns

        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))
        )

        if file_path:
            self.file_path_var.set(file_path)
            self.status_bar.config(text=f"已选择文件: {file_path}")

            try:
                # 读取Excel文件
                df = pd.read_excel(file_path)

                # 更新数据预览
                self.update_preview()

                # 更新轴选择选项
                self.update_axis_options()

                self.status_bar.config(text=f"成功加载数据，共 {len(df)} 行，{len(df.columns)} 列")

            except Exception as e:
                messagebox.showerror("错误", f"读取文件时出错: {str(e)}")
                self.status_bar.config(text="读取文件失败")

    def on_y_scroll(self, *args):
        """处理垂直滚动事件，根据需要显示或隐藏滚动条"""
        # 获取当前显示区域高度和内容总高度
        bbox = self.tree.bbox("all")
        if bbox:
            content_height = bbox[3]  # 内容总高度
            view_height = self.tree.winfo_height()  # 显示区域高度
            
            # 如果内容高度大于显示区域高度，显示垂直滚动条
            if content_height > view_height:
                self.scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
                self.scrollbar_y.set(*args)
            else:
                self.scrollbar_y.pack_forget()
        else:
            self.scrollbar_y.pack_forget()
    
    def on_x_scroll(self, *args):
        """处理水平滚动事件，根据需要显示或隐藏滚动条"""
        # 获取当前显示区域宽度和内容总宽度
        bbox = self.tree.bbox("all")
        if bbox:
            content_width = bbox[2]  # 内容总宽度
            view_width = self.tree.winfo_width()  # 显示区域宽度
            
            # 如果内容宽度大于显示区域宽度，显示水平滚动条
            if content_width > view_width:
                self.scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
                self.scrollbar_x.set(*args)
            else:
                self.scrollbar_x.pack_forget()
        else:
            self.scrollbar_x.pack_forget()
    
    def update_preview(self):
        global df

        # 清除现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)

        if df is not None:
            # 设置列
            self.tree["columns"] = list(df.columns)

            # 隐藏默认的第一列
            self.tree["show"] = "headings"
            
            # 格式化列
            total_width = self.preview_frame.winfo_width() if self.preview_frame.winfo_width() > 1 else 300
            col_width = max(80, total_width // len(df.columns))
            
            for col in df.columns:
                self.tree.column(col, anchor=tk.CENTER, width=col_width, minwidth=60)
                self.tree.heading(col, text=col)

            # 添加数据（限制显示前20行）
            for i, row in df.head(20).iterrows():
                self.tree.insert("", tk.END, values=list(row))

            # 如果数据超过20行，添加提示
            if len(df) > 20:
                self.tree.insert("", tk.END, values=["..."] * len(df.columns))
                self.tree.insert("", tk.END, values=[f"共 {len(df)} 行"] + [""] * (len(df.columns)-1))
            
            # 延迟检查滚动条，确保数据已渲染
            self.root.after(100, self.check_scrollbars)
    
    def schedule_preview_update(self):
        """延迟更新预览，避免频繁刷新导致卡顿"""
        # 取消之前的定时任务
        if hasattr(self, "_preview_update_job"):
            self.root.after_cancel(self._preview_update_job)
        
        # 设置新的定时任务
        self._preview_update_job = self.root.after(200, self.update_preview)
    
    def schedule_chart_update(self):
        """延迟更新图表，避免频繁刷新导致卡顿"""
        # 取消之前的定时任务
        if hasattr(self, "_chart_update_job"):
            self.root.after_cancel(self._chart_update_job)
        
        # 设置新的定时任务
        self._chart_update_job = self.root.after(200, self.refresh_chart)
    
    def check_scrollbars(self):
        """检查是否需要显示滚动条"""
        # 检查垂直滚动条
        bbox = self.tree.bbox("all")
        if bbox:
            content_height = bbox[3]  # 内容总高度
            view_height = self.tree.winfo_height()  # 显示区域高度
            
            # 如果内容高度大于显示区域高度，显示垂直滚动条
            if content_height > view_height:
                self.scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
            else:
                self.scrollbar_y.pack_forget()
            
            # 检查水平滚动条
            content_width = bbox[2]  # 内容总宽度
            view_width = self.tree.winfo_width()  # 显示区域宽度
            
            # 如果内容宽度大于显示区域宽度，显示水平滚动条
            if content_width > view_width:
                self.scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
            else:
                self.scrollbar_x.pack_forget()

    def update_axis_options(self):
        global df

        if df is not None:
            columns = list(df.columns)
            self.x_axis_combo['values'] = columns

            # 清除列表框中的旧内容
            self.y_axis_listbox.delete(0, tk.END)

            # 添加所有列到列表框
            for col in columns:
                self.y_axis_listbox.insert(tk.END, col)

            # 默认选择第一列作为X轴，第二列作为Y轴（如果有多列）
            if len(columns) > 0:
                self.x_axis_var.set(columns[0])
            if len(columns) > 1:
                self.y_axis_listbox.selection_set(1)  # 选择第二项作为默认Y轴

    def create_chart(self):
        global file_path, df, selected_columns, max_rows

        if df is None:
            messagebox.showwarning("警告", "请先选择数据文件")
            return

        # 获取设置
        chart_type = self.chart_type_var.get()
        x_col = self.x_axis_var.get()
        # 获取选中的Y轴列
        selected_indices = self.y_axis_listbox.curselection()
        y_cols = [self.y_axis_listbox.get(i) for i in selected_indices]
        max_rows = self.max_rows_var.get()

        if not x_col:
            messagebox.showwarning("警告", "请选择X轴数据")
            return

        # 根据最大行数设置读取数据
        if max_rows > 0:
            plot_df = df.head(max_rows)
        else:
            plot_df = df.copy()

        # 清除当前图形
        self.ax.clear()

        try:
            if chart_type == "line":
                # 检查 X 轴列是否有重复数据
                if plot_df[x_col].duplicated().any():
                    result = messagebox.askyesno("重复数据", f"X 轴列 '{x_col}' 存在重复数据，是否需要聚合数据？")
                    if result:
                        # 对重复的 X 轴数据取均值
                        plot_df = plot_df.groupby(x_col).mean().reset_index()
                    # 无论选择是或否，都继续生成图表

                x = plot_df[x_col]  # 使用聚合后的 X 轴数据

                # 获取Y轴列
                if not y_cols:
                    # 如果没有选择Y轴，则使用除X轴外的所有列
                    y_columns = [col for col in plot_df.columns if col != x_col]
                else:
                    y_columns = y_cols

                for column in y_columns:
                    self.ax.plot(x, plot_df[column], label=column, marker='o')

                self.ax.set_xlabel(x_col)
                self.ax.set_ylabel('值')
                self.ax.set_title('折线图')
                self.ax.legend()
                self.ax.grid(True)

                # 设置 X 轴标签旋转角度
                self.fig.autofmt_xdate()

                # 添加鼠标悬停显示坐标功能
                cursor = mplcursors.cursor(self.ax, hover=True)
                cursor.connect("add", lambda sel: sel.annotation.set_text(
                    f"{x_col}: {sel.target[0]:.2f}\nValue: {sel.target[1]:.2f}"
                ))

            elif chart_type == "bar":
                if not y_cols:
                    messagebox.showwarning("警告", "条形图需要选择Y轴数据")
                    return

                # 计算每列的总和
                if x_col in y_cols and len(y_cols) == 1:
                    # 如果X轴和Y轴相同，则显示该列的值
                    self.ax.bar(plot_df[x_col], plot_df[y_cols[0]])
                    self.ax.set_xlabel(x_col)
                    self.ax.set_ylabel(y_cols[0])
                else:
                    # 否则，显示每列的总和
                    sums = plot_df[y_cols].sum()
                    self.ax.bar(sums.index, sums, color='skyblue')
                    self.ax.set_xlabel('列')
                    self.ax.set_ylabel('总值')

                self.ax.set_title('条形图')

                # 如果数据点太多，旋转X轴标签
                if len(plot_df) > 10:
                    self.fig.autofmt_xdate()

            else:  # pie chart
                if not y_cols:
                    messagebox.showwarning("警告", "饼图需要选择Y轴数据")
                    return

                if x_col in y_cols and len(y_cols) == 1:
                    # 如果X轴和Y轴相同，则显示该列的值
                    self.ax.pie(plot_df[y_cols[0]], labels=plot_df[x_col], autopct='%1.1f%%', startangle=140)
                else:
                    # 否则，显示每列的总和
                    sums = plot_df[y_cols].sum()
                    self.ax.pie(sums, labels=sums.index, autopct='%1.1f%%', startangle=140)

                self.ax.set_title('饼图')

            # 更新画布
            self.canvas.draw()

            self.status_bar.config(text=f"已创建{chart_type}图表")

        except Exception as e:
            messagebox.showerror("错误", f"创建图表时出错: {str(e)}")
            self.status_bar.config(text="创建图表失败")
    
    def refresh_chart(self):
        """刷新图表，用于延迟更新机制"""
        if hasattr(self, "canvas"):
            self.canvas.draw()

if __name__ == "__main__":
    root = tk.Tk()
    app = ChartApp(root)
    root.mainloop()
