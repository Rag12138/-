import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
import os


class WaveformAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("波形分析工具 - 列数据平均值计算")
        self.root.geometry("1000x700")

        # 初始化变量
        self.df = None  # 存储表格数据
        self.col_data = None  # 存储选中列的数据
        self.fig, self.ax = plt.subplots(figsize=(8, 4), dpi=100)  # 绘图画布
        self.canvas = None  # tkinter中的matplotlib画布
        self.selection_rect = None  # 选取区域的蒙版
        self.start_idx = 0  # 选取起始索引
        self.end_idx = 20000  # 扩大初始蒙版范围（避免重叠）
        self.is_dragging = False  # 是否正在拖动蒙版
        self.drag_type = None  # 拖动类型：start/end/move

        # 创建GUI布局
        self._create_widgets()

    def _create_widgets(self):
        # 顶部操作区
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)

        # 打开文件按钮
        ttk.Button(top_frame, text="打开表格文件", command=self.load_file).grid(row=0, column=0, padx=5)

        # 列选择下拉框
        self.col_var = tk.StringVar()
        self.col_combobox = ttk.Combobox(top_frame, textvariable=self.col_var, state="disabled")
        self.col_combobox.grid(row=0, column=1, padx=5)
        self.col_combobox.bind("<<ComboboxSelected>>", self.load_column_data)

        # 平均值显示区
        self.mean_var = tk.StringVar(value="选中区域平均值：--")
        ttk.Label(top_frame, textvariable=self.mean_var).grid(row=0, column=2, padx=20)

        # 复制按钮
        ttk.Button(
            top_frame,
            text="将结果复制到剪贴板",
            command=self.copy_result_to_clipboard
        ).grid(row=0, column=3, padx=10)

        # 新增：复制提示文本标签（初始为空）
        self.copy_hint_var = tk.StringVar(value="")
        self.copy_hint_label = ttk.Label(
            top_frame,
            textvariable=self.copy_hint_var,
            foreground="green"  # 提示文本设为绿色
        )
        self.copy_hint_label.grid(row=0, column=4, padx=5)

        # 绘图区域
        plot_frame = ttk.Frame(self.root, padding="10")
        plot_frame.pack(fill=tk.BOTH, expand=True)

        # 初始化matplotlib画布
        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 绑定鼠标事件
        self.canvas.mpl_connect("button_press_event", self.on_mouse_press)
        self.canvas.mpl_connect("motion_notify_event", self.on_mouse_move)
        self.canvas.mpl_connect("button_release_event", self.on_mouse_release)

    def load_file(self):
        """打开文件对话框，读取Excel/CSV表格"""
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("表格文件", "*.xlsx *.xls *.csv"),
                ("Excel文件", "*.xlsx *.xls"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        if not file_path:
            return

        # 读取表格
        try:
            if file_path.endswith((".xlsx", ".xls")):
                self.df = pd.read_excel(file_path)
            elif file_path.endswith(".csv"):
                self.df = pd.read_csv(file_path)
            else:
                messagebox.showerror("错误", "不支持的文件格式！")
                return

            # 更新列选择下拉框
            self.col_combobox["state"] = "readonly"
            self.col_combobox["values"] = list(self.df.columns)
            messagebox.showinfo("成功", f"已加载文件，共{len(self.df)}行数据，{len(self.df.columns)}列")

        except Exception as e:
            messagebox.showerror("读取失败", f"文件读取出错：{str(e)}")

    def load_column_data(self, event=None):
        """加载选中列的数据并绘制波形图"""
        if self.df is None:
            return

        col_name = self.col_var.get()
        if not col_name:
            return

        # 获取列数据并清理无效值
        self.col_data = self.df[col_name].dropna().values
        if len(self.col_data) == 0:
            messagebox.warning("警告", "该列无有效数据！")
            return

        # 重置选取范围（根据数据长度设置合理初始范围）
        self.start_idx = 0
        self.end_idx = min(20000, len(self.col_data) - 1)  # 初始取前2000个点（避免太短）

        # 绘制波形图
        self.ax.clear()
        self.ax.plot(self.col_data, color="blue", linewidth=1)
        self.ax.set_title(f"波形图 - 列：{col_name}", fontsize=12)
        self.ax.set_xlabel("数据索引")
        self.ax.set_ylabel("数值")
        self.ax.grid(True, alpha=0.3)

        # 绘制初始选取蒙版
        self._draw_selection_rect()

        # 计算初始平均值
        self._calculate_mean()

        # 更新画布
        self.canvas.draw()

    def _draw_selection_rect(self):
        """绘制选取区域的蒙版（修复边界判断）"""
        # 移除旧的蒙版
        if self.selection_rect is not None:
            try:
                self.selection_rect.remove()
            except:
                pass  # 防止重复移除报错
            self.selection_rect = None

        # 确保起始索引小于结束索引
        if self.start_idx >= self.end_idx:
            self.end_idx = self.start_idx + 1  # 强制拉开距离

        # 获取Y轴范围，用于绘制蒙版
        y_min, y_max = self.ax.get_ylim()
        if y_min == y_max:  # 避免数据全部相同导致的问题
            y_min -= 1
            y_max += 1

        # 绘制半透明蒙版（只保留一段选取）
        self.selection_rect = self.ax.axvspan(
            self.start_idx, self.end_idx,
            facecolor="orange", alpha=0.3, edgecolor="red", linewidth=1
        )

    def _calculate_mean(self):
        """计算选中区域的平均值并更新显示"""
        if self.col_data is None:
            return

        # 确保索引有效
        start = max(0, self.start_idx)
        end = min(len(self.col_data) - 1, self.end_idx)
        if start >= end:
            self.mean_var.set("选中区域平均值：--")
            return

        # 计算平均值
        selected_data = self.col_data[start:end + 1]
        mean_val = np.mean(selected_data)
        self.mean_var.set(f"选中区域平均值：{mean_val:.4f}")

    def on_mouse_press(self, event):
        """鼠标按下事件：判断是否点击蒙版边缘（优化容错范围）"""
        if self.col_data is None or event.inaxes != self.ax:
            return

        # 检测是否点击起始边缘（±10个像素的容错，适配大数据量）
        if abs(event.xdata - self.start_idx) < 1000:
            self.is_dragging = True
            self.drag_type = "start"
        # 检测是否点击结束边缘
        elif abs(event.xdata - self.end_idx) < 1000:
            self.is_dragging = True
            self.drag_type = "end"
        # 点击蒙版内部：拖动整个蒙版
        elif self.start_idx <= event.xdata <= self.end_idx:
            self.is_dragging = True
            self.drag_type = "move"
            self.drag_offset = event.xdata - self.start_idx  # 记录偏移量

    def on_mouse_move(self, event):
        """鼠标移动事件：调整蒙版位置/大小（优化边界限制）"""
        if not self.is_dragging or self.col_data is None or event.inaxes != self.ax:
            return

        # 限制X轴范围在0到数据长度之间
        new_x = max(0, min(len(self.col_data) - 1, event.xdata))

        # 根据拖动类型更新索引
        if self.drag_type == "start":
            # 起始索引不能超过结束索引
            self.start_idx = min(int(new_x), self.end_idx - 1)
        elif self.drag_type == "end":
            # 结束索引不能小于起始索引
            self.end_idx = max(int(new_x), self.start_idx + 1)
        elif self.drag_type == "move":
            # 拖动整个蒙版
            new_start = max(0, min(len(self.col_data) - (self.end_idx - self.start_idx), new_x - self.drag_offset))
            new_end = new_start + (self.end_idx - self.start_idx)
            self.start_idx = int(new_start)
            self.end_idx = int(new_end)

        # 重新绘制蒙版并计算平均值
        self._draw_selection_rect()
        self._calculate_mean()
        self.canvas.draw()

    def on_mouse_release(self, event):
        """鼠标释放事件：结束拖动"""
        self.is_dragging = False
        self.drag_type = None

    def copy_result_to_clipboard(self):
        """将当前计算的平均值复制到系统剪贴板（无弹窗，文本提示）"""
        # 清空之前的提示
        self.copy_hint_var.set("")

        # 获取显示的平均值文本
        mean_text = self.mean_var.get()
        # 提取纯数字部分（去掉前缀）
        if "：--" in mean_text:
            self.copy_hint_var.set("暂无有效数据")
            self.copy_hint_label.config(foreground="red")  # 错误提示设为红色
            # 3秒后清空提示
            self.root.after(3000, lambda: self.copy_hint_var.set(""))
            return

        # 截取数值部分
        mean_value = mean_text.split("：")[1].strip()
        try:
            # 将数值复制到剪贴板
            self.root.clipboard_clear()  # 清空剪贴板
            self.root.clipboard_append(mean_value)  # 添加数值
            self.root.update()  # 确保剪贴板更新

            # 显示成功提示
            self.copy_hint_var.set("复制成功")
            self.copy_hint_label.config(foreground="green")  # 成功提示设为绿色
            # 3秒后自动清空提示
            self.root.after(3000, lambda: self.copy_hint_var.set(""))

        except Exception as e:
            # 显示失败提示
            self.copy_hint_var.set("复制失败")
            self.copy_hint_label.config(foreground="red")
            # 3秒后清空提示
            self.root.after(3000, lambda: self.copy_hint_var.set(""))

if __name__ == "__main__":
    # 解决matplotlib中文显示问题
    plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
    plt.rcParams["axes.unicode_minus"] = False  # 解决负号显示问题

    root = tk.Tk()
    app = WaveformAnalyzer(root)
    root.mainloop()