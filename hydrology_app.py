import sys
import os
import numpy as np
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.widgets import Button, TextBox
from geomdl import fitting
import pandas as pd
from tkinter import filedialog

import pyperclip
import scipy as sp
from matplotlib.ticker import MultipleLocator, AutoMinorLocator
import mpl_toolkits
import openpyxl
import xlsxwriter
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False


def load_hydrological_data():
    """按水位排序加载数据（返回年份和测次）"""
    try:
        file_path = filedialog.askopenfilename(
            title="选择水文数据文件",
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        if not file_path:
            return None, None, None, None, None

        # 读取数据
        df = pd.read_excel(file_path, header=0)

        # 数据清洗（保留年份和测次）
        df_clean = df.dropna(subset=['水位', '流量'])
        df_clean['流量'] = pd.to_numeric(df_clean['流量'], errors='coerce')
        df_clean['水位'] = pd.to_numeric(df_clean['水位'], errors='coerce')
        df_clean = df_clean.dropna(subset=['流量', '水位'])

        if df_clean.empty:
            raise ValueError("清洗后无有效数据")

        # 按水位排序
        df_sorted = df_clean.sort_values(by='水位')

        # 提取各列数据并确保类型正确
        years = df_sorted['年份'].astype(int).values
        tests = df_sorted['测次'].astype(int).values
        discharges = df_sorted['流量'].astype(float).values
        water_levels = df_sorted['水位'].astype(float).values

        # 检查数据长度一致性
        if len(years) != len(tests) or len(years) != len(discharges) or len(years) != len(water_levels):
            raise ValueError("数据列长度不一致")

        return years, tests, discharges, water_levels, file_path

    except Exception as e:
        print(f"数据加载失败: {str(e)}")
        return None, None, None, None, None


class CurveEditor:
    def __init__(self, years, tests, discharges, water_levels, file_path, degree=3, control_points_size=15):
        # ==== 数据校验 ====
        if len(years) == 0 or len(tests) == 0 or len(discharges) == 0 or len(water_levels) == 0:
            raise ValueError("输入数据不能为空")
        if len(years) != len(tests) or len(years) != len(discharges) or len(years) != len(water_levels):
            raise ValueError("输入数据长度不一致")
        # ==== 参数类型强制转换 ====
        self.degree = int(degree)  # 确保为整数
        self.control_points_size = int(control_points_size)  # 确保为整数
        # ==== 历史记录属性初始化 ====
        self.history = []  # 存储历史控制点
        self.history_index = -1  # 当前历史索引
        # ==== 数据存储 ====
        self.years = years
        self.tests = tests
        self.discharges = np.array(discharges, dtype=float)
        self.water_levels = np.array(water_levels, dtype=float)
        self.file_path = file_path
        # ==== 新增关键属性初始化 ====
        self.degree = degree  # B样条曲线阶数
        self.control_points_size = control_points_size  # 控制点数量
        # ==== 其他初始化代码 ====
        self.discharges = np.array(discharges).astype(float)
        self.water_levels = np.array(water_levels).astype(float)
        # ==== 曲线初始化（必须放在绘图前）====
        # 曲线初始化（使用已定义的属性）
        data_points = [[q, wl] for q, wl in zip(self.discharges, self.water_levels)]  # 明确生成二维列表
        # ==== 曲线初始化 ====
        try:
            self.curve = fitting.approximate_curve(
                data_points,
                degree=self.degree,
                ctrlpts_size=self.control_points_size
            )
            self.curve.sample_size = 2000
        except Exception as e:
            raise RuntimeError(f"曲线初始化失败: {str(e)}")
        self.curve.sample_size = 2000  # 确保足够多的评估点

        # ==== 新增标题属性初始化 ====
        self.file_path = file_path
        self.base_title = self._generate_title()  # 初始化时生成标题
        # ==== 图形初始化 ====
        self.fig, self.ax = plt.subplots()
        self.ax.set_title(self._generate_title(), fontsize=22, fontweight='bold', color='blue', pad=20)

        # ==== 绘图元素初始化 ====
        self.data_plot, = self.ax.plot(self.discharges, self.water_levels, 'ro', label='实测数据')
        self.ctrl_plot, = self.ax.plot([], [], 'g--', marker='s', markersize=8, linewidth=1, label='控制多边形')
        self.curve_plot, = self.ax.plot([], [], 'b-', linewidth=2, label='拟合曲线')

        # ==== 必须在曲线初始化后调用 ====
        self.update_control_points_plot()
        self.update_curve_plot()
        # 设置主次刻度
        self.ax.xaxis.set_major_locator(MultipleLocator(400))  # x轴主刻度50
        self.ax.xaxis.set_minor_locator(AutoMinorLocator(5))  # 每个主刻度分5个次刻度（10间隔）
        self.ax.yaxis.set_major_locator(MultipleLocator(1))  # y轴主刻度1
        self.ax.yaxis.set_minor_locator(AutoMinorLocator(5))  # 每个主刻度分5个次刻度（0.2间隔）

        # 启用次刻度
        self.ax.minorticks_on()

        # 设置网格样式
        self.ax.grid(which='major', linestyle='-', linewidth=1.5, color='#888888')
        self.ax.grid(which='minor', linestyle=':', linewidth=2.0, color='#DDDDDD')

        # 事件绑定
        self.selected_point = None
        self.operation_mode = 'view'
        self.fig.canvas.mpl_connect('button_press_event', self.on_click)
        self.fig.canvas.mpl_connect('motion_notify_event', self.on_drag)
        self.fig.canvas.mpl_connect('button_release_event', self.on_release)

        # 添加控制按钮
        self.add_buttons()

    def _generate_title(self):
        """生成基础标题（新增方法）"""
        if self.file_path:
            filename = os.path.splitext(os.path.basename(self.file_path))[0]
            return f"{filename} 水位流量关系曲线"
        return "水位流量关系曲线"
    def update_control_points_plot(self):
        """更新控制点显示"""
        ctrlpts = np.array(self.curve.ctrlpts)
        if ctrlpts.ndim == 2 and ctrlpts.shape[1] == 2:
            self.ctrl_plot.set_data(ctrlpts[:, 0], ctrlpts[:, 1])
        else:
            print("警告: 控制点格式异常")

    def update_curve_plot(self):
        """更新曲线显示"""
        try:
            evalpts = np.array(self.curve.evalpts)
            if evalpts.ndim == 2 and evalpts.shape[1] == 2:
                self.curve_plot.set_data(evalpts[:, 0], evalpts[:, 1])
                self.fig.canvas.draw_idle()
            else:
                print("警告: 曲线评估点格式异常")
        except AttributeError:
            print("错误: 曲线未正确初始化")

    def on_click(self, event):
        """点击事件处理"""
        if self.operation_mode != 'edit' or event.inaxes != self.ax or event.button != 1:
            return

        try:
            ctrlpts = np.array(self.curve.ctrlpts)
            if ctrlpts.ndim != 2 or ctrlpts.shape[1] != 2:
                print("控制点结构异常")
                return

            # 计算最近控制点
            distances = np.hypot(ctrlpts[:, 0] - event.xdata, ctrlpts[:, 1] - event.ydata)
            self.selected_point = np.argmin(distances)

            if distances[self.selected_point] < 0.05:  # 有效选择阈值
                self.ctrl_plot.set_markerfacecolor('yellow')
                self.fig.canvas.draw_idle()
        except Exception as e:
            print(f"点击事件处理失败: {str(e)}")

    def on_drag(self, event):
        """拖动事件处理"""
        if self.operation_mode == 'edit' and self.selected_point is not None:
            if event.xdata is None or event.ydata is None:
                return

            try:
                new_ctrlpts = [list(pt) for pt in self.curve.ctrlpts]
                new_ctrlpts[self.selected_point] = [event.xdata, event.ydata]
                self.curve.ctrlpts = new_ctrlpts
                self.update_control_points_plot()
                self.update_curve_plot()
            except IndexError:
                print("错误: 控制点索引越界")
            except Exception as e:
                print(f"拖动事件处理失败: {str(e)}")


    def toggle_mode(self, event):
        """修复后的模式切换方法"""
        self.operation_mode = 'edit' if self.operation_mode == 'view' else 'view'
        new_color = 'red' if self.operation_mode == 'edit' else 'blue'
        new_label = '编辑模式' if self.operation_mode == 'edit' else '视图模式'
        # 使用通过方法生成的标题
        self.ax.set_title(self.base_title,
                          fontsize=22,
                          fontweight='bold',
                          color=new_color,
                          pad=20)

        # 更新按钮标签
        btn = self.buttons['mode_toggle']
        btn.label.set_text(new_label)
        self.fig.canvas.draw_idle()

    def process_query(self, event):
        try:
            input_str = self.txt_query.text
            values = [float(v.strip()) for v in input_str.split(',') if v.strip()]
            results = []
            for v in values:
                if '流量' in self.ax.get_xlabel():
                    eval_res = self.curve.evaluate_single(v)
                    results.append([v, eval_res[1]])
                else:
                    q_values = np.linspace(min(self.discharges), max(self.discharges), 1000)
                    closest_q = min(q_values, key=lambda q: abs(self.curve.evaluate_single(q)[1] - v))
                    results.append([v, closest_q])
            self.query_table.set_cellText([[str(x) for x in row] for row in results])
            self.fig.canvas.draw_idle()
        except Exception as e:
            print(f"查询失败: {str(e)}")

    # (标签, x位置, y位置, 宽度, 高度, 回调函数)
    def add_buttons(self):
        button_config = [
            # 格式：(键名, x位置, y位置, 宽度, 高度, 回调函数, 显示文本)
            ('test_result', 0.00, 0.90, 0.08, 0.05, self.show_test_results, '查看检验'),
            ('mode_toggle', 0.00, 0.02, 0.08, 0.05, self.toggle_mode, '视图模式'),
            ('three_test', 0.00, 0.84, 0.08, 0.05, self.export_three_tests, '检验导出'),
            ('data_points', 0.00, 0.78, 0.08, 0.05, self.toggle_data_points, '隐藏测点'),
            ('ctrl_points', 0.00, 0.72, 0.08, 0.05, self.toggle_control_points, '隐藏控点'),
            ('redo_btn', 0.00, 0.66, 0.08, 0.05, self.redo, '重做'),
            ('undo_btn', 0.00, 0.60, 0.08, 0.05, self.undo, '撤销'),
            ('adjust_ctrl', 0.00, 0.54, 0.08, 0.05, self.adjust_control_points, '调整点数'),
            ('export_wl', 0.00, 0.48, 0.08, 0.05, self.export_relationship, 'H~Q导出'),
            ('export_data', 0.00, 0.42, 0.08, 0.05, self.export_data, 'Q~H导出'),
            ('reset_btn', 0.00, 0.36, 0.08, 0.05, self.reset_curve, 'Reset'),
            # ('generate_formula', 0.00, 0.30, 0.08, 0.05, self.generate_formula, '生成公式')
        ]

        self.buttons = {}
        for (key, x, y, w, h, func, label) in button_config:
            ax = self.fig.add_axes([x, y, w, h])
            btn = Button(ax, label,
                         color='lightblue' if key == 'mode_toggle' else 'lightgoldenrodyellow',
                         hovercolor='deepskyblue' if key == 'mode_toggle' else 'lightyellow')
            btn.label.set_fontsize(14)
            btn.label.set_fontweight('bold')
            btn.on_clicked(func)
            self.buttons[key] = btn

    # 以下是所有方法实现（保持正确缩进）
    def toggle_data_points(self, event):
        try:
            visible = self.data_plot.get_visible()
            self.data_plot.set_visible(not visible)
            # 使用正确的键名 'data_points'
            btn = self.buttons['data_points']
            btn.label.set_text('显示测点' if visible else '隐藏测点')
            self.fig.canvas.draw_idle()
        except Exception as e:
            print(f"切换实测点可见性失败: {str(e)}")
    def update_curve_plot(self):
        evalpts = np.array(self.curve.evalpts)
        self.curve_plot.set_data(evalpts[:, 0], evalpts[:, 1])
        self.fig.canvas.draw_idle()

    def on_release(self, event):
        """释放事件处理"""
        if self.selected_point is not None:
            self.push_history()
            self.selected_point = None
            self.ctrl_plot.set_markerfacecolor('green')
            self.update_curve_plot()

    def reset_curve(self, event):
        """重置曲线（修复后）"""
        data_points = np.column_stack([self.discharges, self.water_levels]).tolist()
        self.curve = fitting.approximate_curve(
            data_points,
            degree=self.degree,  # 使用已初始化的属性
            ctrlpts_size=self.control_points_size  # 使用已初始化的属性
        )
        self.update_control_points_plot()
        self.update_curve_plot()

    def toggle_control_points(self, event):
        try:
            visible = self.ctrl_plot.get_visible()
            self.ctrl_plot.set_visible(not visible)
            # 使用正确的键名 'ctrl_points'
            btn = self.buttons['ctrl_points']
            btn.label.set_text('显示控点' if visible else '隐藏控点')
            self.fig.canvas.draw_idle()
        except Exception as e:
            print(f"切换控制点可见性失败: {str(e)}")

    def push_history(self):
        """保存当前状态到历史记录"""
        if self.history_index < len(self.history) - 1:
            self.history = self.history[:self.history_index + 1]
        current_ctrlpts = np.array(self.curve.ctrlpts).copy()
        self.history.append(current_ctrlpts)
        self.history_index += 1

    def undo(self, event):
        """撤销操作（修复后）"""
        if self.history_index > 0:
            self.history_index -= 1
            self.curve.ctrlpts = self.history[self.history_index].tolist()
            self.update_control_points_plot()
            self.update_curve_plot()

    def redo(self, event):
        """重做操作（修复后）"""
        if self.history_index < len(self.history) - 1:
            self.history_index += 1
            self.curve.ctrlpts = self.history[self.history_index].tolist()
            self.update_control_points_plot()
            self.update_curve_plot()

    def adjust_control_points(self, event):
        """调整控制点数量（修复后）"""
        self.ctrl_fig = plt.figure("调整控制点数量", figsize=(4, 2))
        ax = self.ctrl_fig.add_subplot(111)
        # 使用已初始化的属性
        self.txt_ctrl_num = TextBox(ax, '新点数:', initial=str(self.control_points_size))
        self.txt_ctrl_num.on_submit(self.update_ctrlpts_num)
        plt.show()

    def update_ctrlpts_num(self, text):
        try:
            new_num = int(text)
            if new_num < 3:
                print("控制点数量不能少于3")
                return
            self.control_points_size = new_num
            self.reset_curve(None)  # 触发曲线重拟合
            plt.close("调整控制点数量")  # 关闭调整窗口
            self.fig.canvas.draw_idle()  # 刷新主界面
        except ValueError:
            print("请输入有效整数")

    def export_relationship(self, event):
        """修复后的水位流量关系导出"""
        try:
            # 确保曲线已更新
            self.curve.evaluate()
            evalpts = np.array(self.curve.evalpts)

            # 获取原始数据的水位范围
            min_wl = np.min(self.water_levels)
            max_wl = np.max(self.water_levels)

            # 按水位排序评估点并去重
            sorted_indices = np.argsort(evalpts[:, 1])
            sorted_evalpts = evalpts[sorted_indices]
            unique_wl, unique_indices = np.unique(sorted_evalpts[:, 1], return_index=True)
            unique_q = sorted_evalpts[unique_indices, 0]

            # 生成0.01米步长的水位序列
            new_wl = np.arange(min_wl, max_wl + 0.01, 0.01)

            # 创建插值函数（水位→流量）
            from scipy.interpolate import interp1d
            interp_func = interp1d(
                unique_wl,
                unique_q,
                kind='linear',
                bounds_error=False,
                fill_value='extrapolate'  # 外推以包含边界
            )

            # 计算对应流量
            new_q = interp_func(new_wl)

            # 构建DataFrame
            df = pd.DataFrame({
                '水位（米）': new_wl,
                '流量（m³/s）': new_q
            })

            # 弹出保存对话框
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")],
                title="保存水位流量关系数据"
            )

            if file_path:
                df.to_excel(file_path, index=False)
                print(f"成功导出{len(df)}条数据至：{file_path}")

        except Exception as e:
            print(f"导出失败: {str(e)}")

    def export_three_tests(self, event):
        """导出三项检验数据（含统计结果）"""
        try:
            # 获取检验数据
            test_data = self.calculate_three_tests()
            if test_data is None:
                raise ValueError("无有效检验数据")

            # 计算统计指标
            n = len(test_data['sorted_q'])
            K = np.nansum(test_data['k_values'])
            sum_deviation = np.nansum(test_data['pi_deviation'])
            sum_pi_squared = np.nansum(test_data['pi_squared'])

            # 计算各项指标
            try:
                u = (abs(K - 0.5 * n) - 0.5) / (0.5 * np.sqrt(n)) if n > 0 else np.nan
                S = np.sqrt(sum_deviation / (n - 1)) if n > 1 else np.nan
                SP = S / np.sqrt(n) if S is not np.nan and n > 0 else np.nan
                pi_mean = abs(np.nanmean(test_data['pi_values']))
                t = pi_mean / SP if SP not in (0, np.nan) else np.nan
                se = np.sqrt(sum_pi_squared / (n - 2)) if n > 2 else np.nan
                # 计算适线检验 U
                U = (0.5 * (n - 1) - K - 0.5) / (0.5 * np.sqrt(n - 1))
            except Exception as e:
                print(f"统计指标计算失败: {str(e)}")
                u = S = SP = t = se = U = np.nan

            # 创建主数据表格
            main_df = pd.DataFrame({
                '年份': self.years,  # 新增A列
                '测次': self.tests,  # 新增B列
                '水位（米）': test_data['sorted_wl'],
                '流量Qi（m³/s）': test_data['sorted_q'],
                '查线流量Qci（m³/s）': test_data['qci_values'],
                'Pi (%)': test_data['pi_values'],
                'Pi离均差': test_data['pi_deviation'],
                'Pi²': test_data['pi_squared'],
                'K': test_data['k_values']
            })

            # 创建统计结果表格
            stats_df = pd.DataFrame({
                '统计指标': [
                    '数据组数 n',
                    '符号统计 K',
                    '符号检验值 u',
                    '标准差 S (%)',
                    '标准差 SP (%)',
                    '标准差 SE (%)',
                    't 检验值',
                    '适线检验 U'
                ],
                '数值': [
                    n,
                    int(K),
                    f"{u:.2f}" if not np.isnan(u) else "N/A",
                    f"{S:.2f}" if not np.isnan(S) else "N/A",
                    f"{SP:.2f}" if not np.isnan(SP) else "N/A",
                    f"{se:.2f}" if not np.isnan(se) else "N/A",
                    f"{t:.2f}" if not np.isnan(t) else "N/A",
                    f"{U:.2f}" if not np.isnan(U) else "N/A"  # 新增适线检验 U
                ]
            })

            # 合并数据
            combined_df = pd.concat([main_df, stats_df], ignore_index=True)

            # 弹出保存对话框
            file_path = filedialog.asksaveasfilename(
                initialdir=os.path.expanduser("~/Desktop"),
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx")],
                title="保存三项检验数据"
            )

            if file_path:
                # 使用ExcelWriter设置格式
                with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                    combined_df.to_excel(writer, index=False, sheet_name='三项检验')

                    # 获取工作表对象
                    workbook = writer.book
                    worksheet = writer.sheets['三项检验']

                    # 设置统计结果行的格式
                    stats_format = workbook.add_format({
                        'bold': True,
                        'bg_color': '#FFFF00',
                        'border': 1
                    })

                    # 应用格式（最后8行）
                    for row in range(len(main_df) + 1, len(combined_df) + 1):
                        worksheet.set_row(row, None, stats_format)

                print(f"三项检验数据（含统计结果）已导出至：{file_path}")

        except Exception as e:
            self.show_error_message("导出错误", str(e))

    def calculate_three_tests(self):
        """统一计算三项检验所需数据"""
        try:
            sorted_idx = np.argsort(self.water_levels)
            sorted_wl = self.water_levels[sorted_idx]
            sorted_q = self.discharges[sorted_idx]

            # 获取当前曲线评估点
            self.curve.evaluate()
            evalpts = np.array(self.curve.evalpts)
            curve_x = evalpts[:, 0]
            curve_y = evalpts[:, 1]

            # 计算查线流量（使用线性插值）
            from scipy.interpolate import interp1d
            f_interp = interp1d(curve_y, curve_x, kind='linear', bounds_error=False, fill_value='extrapolate')
            qci_values = f_interp(sorted_wl)

            # 计算各项指标
            with np.errstate(divide='ignore', invalid='ignore'):
                pi_values = (sorted_q - qci_values) / qci_values * 100
                pi_values = np.nan_to_num(pi_values, nan=0.0)

            pi_mean = np.mean(pi_values)
            pi_deviation = (pi_values - pi_mean) ** 2
            pi_squared = pi_values ** 2

            # 计算K值
            k_values = np.zeros(len(pi_values))
            for i in range(1, len(pi_values)):
                if pi_values[i] * pi_values[i - 1] < 0:
                    k_values[i] = 1

            return {
                'sorted_wl': sorted_wl,
                'sorted_q': sorted_q,
                'qci_values': qci_values,
                'pi_values': pi_values,
                'pi_deviation': pi_deviation,
                'pi_squared': pi_squared,
                'k_values': k_values
            }
        except Exception as e:
            print(f"计算失败: {str(e)}")
            return None

    def show_test_results(self, event):
        try:
            test_data = self.calculate_three_tests()
            if test_data is None:
                raise ValueError("无法计算检验数据")

            # 后续显示逻辑与原show_test_results相同
            n = len(test_data['sorted_q'])
            K = np.nansum(test_data['k_values'])

            # ==== 符号检验值计算 ====
            try:
                # 修正后的符号检验公式
                numerator =( abs(K - 0.5 * n) - 0.5)
                denominator = 0.5 * np.sqrt(n)
                u = numerator / denominator if denominator != 0 else np.nan
            except:
                u = np.nan

            # ==== 标准差计算 ====
            sum_deviation = np.nansum(test_data['pi_deviation'])
            try:
                S = np.sqrt(sum_deviation / (n - 1)) if K > 1 else np.nan
            except:
                S = np.nan

            try:
                SP = S / np.sqrt(n) if K > 0 else np.nan
            except:
                SP = np.nan

            # ==== t值计算 ====
            pi_mean = abs(np.nanmean(test_data['pi_values']))
            try:
                t = pi_mean / SP if not np.isnan(SP) and SP != 0 else np.nan
            except:
                t = np.nan

            sum_pi_squared = np.nansum(test_data['pi_squared'])
            try:
                se = np.sqrt(sum_pi_squared / (n - 2))
            except ZeroDivisionError:
                se = np.nan

            try:
                # sum_pi_squared = np.nansum(test_data['pi_squared'])
                U = (0.5 * (n - 1) - K - 0.5) / (0.5 * np.sqrt(n - 1))
            except:
                U = np.nan
            # ==== 构建显示文本 ====
            result_text = [
                f"数据组数 n = {n}",
                f"符号统计 K = {int(K)}",
                f"符号检验值 u = {u:.2f}" if not np.isnan(u) else "符号检验值 u = 无法计算",
                "偏离数值检验",
                f"P的标准差 S = {S:.2f}%",
                f"P的标准差 SP = {SP:.2f}%",
                f"标准差 SE = {se:.2f}%",
                f"（学生氏）t = {t:.2f}",
                "适线检验",  # 第9行
                f"适线检验 U = {U:.2f}"  # 第10行
            ]

            # ==== 创建结果显示窗口 ====
            result_fig = plt.figure("海南水文提醒：导出三项检验数据请重新校核结果", figsize=(10, 5))
            ax = result_fig.add_subplot(111)
            ax.axis('off')

            # 设置专业排版样式
            text_params = {
                'fontsize': 14,
                'fontfamily': 'Microsoft YaHei',
                'verticalalignment': 'top',
                'horizontalalignment': 'left'
            }

            # 添加带项目符号的文本
            y_pos = 0.95
            for line in result_text:
                # 添加项目符号
                ax.text(0.05, y_pos, "•", **text_params, color='darkblue')
                ax.text(0.08, y_pos - 0.01, line, **text_params)
                y_pos -= 0.15  # 调整行间距

            # 添加分隔线
            ax.axhline(y=0.65, xmin=0.05, xmax=0.95, color='gray', linewidth=0.8)

            plt.tight_layout()
            plt.show()

        except Exception as e:
            self.show_error_message("计算错误", f"{str(e)}\n请确保已完成曲线拟合")

    # 辅助方法：显示错误对话框
    def show_error_message(self, title, message):
        plt.figure(title)
        plt.text(0.5, 0.7, message,
                 ha='center', va='center',
                 fontsize=12, color='red')
        plt.axis('off')
        plt.show()

    # 新增信息提示方法
    def show_info_message(self, title, message):
        plt.figure(title, figsize=(6, 2))
        plt.text(0.5, 0.7, message,
                 ha='center', va='center',
                 fontsize=12, color='green')
        plt.axis('off')
        plt.show()

    def export_data(self, event):
        try:
            eval_data = self.calculate_evaluation()
            # 过滤无效数据（NaN）
            valid_data = [row for row in eval_data if not np.isnan(row[2])]

            df = pd.DataFrame(valid_data,
                              columns=['流量(m³/s)', '实测水位', '拟合水位', '绝对误差'])

            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")]
            )

            if file_path:
                if file_path.endswith('.csv'):
                    df.to_csv(file_path, index=False, encoding='utf_8_sig')
                else:
                    df.to_excel(file_path, index=False)
        except Exception as e:
            print(f"导出失败: {str(e)}")

    def calculate_metrics(self):
        """计算三项检验指标"""
        from scipy import stats

        eval_data = self.calculate_evaluation()
        valid_data = [row for row in eval_data if not np.isnan(row[2])]

        if len(valid_data) < 2:
            return [['错误', '至少需要2个有效数据点']]

        try:
            measured = [row[1] for row in valid_data]
            fitted = [row[2] for row in valid_data]
            errors = [row[3] for row in valid_data]

            # 计算相关系数
            corr_coef, _ = stats.pearsonr(measured, fitted)
            # 计算纳什效率系数
            nse = 1 - (np.var(errors) / np.var(measured))
            return [
                ['平均绝对误差', f"{np.mean(errors):.2f} m"],
                ['相关系数', f"{corr_coef:.2f}"],
                ['纳什效率系数', f"{nse:.2f}"]
            ]
        except Exception as e:
            return [['计算错误', str(e)]]

    def create_table(self, ax, data, headers, colWidths=None):
        """创建表格（完全兼容空数据）"""
        ax.clear()

        if not data:
            ax.text(0.5, 0.5, '无数据', ha='center', va='center', fontsize=12)
            ax.axis('off')
            return

        # 验证数据格式
        for row in data:
            if len(row) != len(headers):
                raise ValueError(f"数据项 {row} 与表头 {headers} 长度不匹配")

        # 创建表格
        table = ax.table(
            cellText=data,
            colLabels=headers,
            loc='center',
            colWidths=colWidths if colWidths else [0.25] * len(headers)
        )
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        ax.axis('off')



    def calculate_evaluation(self):
        """计算拟合误差（流量→水位）Q-H导出"""
        results = []
        try:
            # 确保曲线已正确生成评估点
            if not hasattr(self.curve, 'evalpts'):
                self.curve.evaluate()  # 手动触发评估点计算

            evalpts = np.array(self.curve.evalpts)
            if evalpts.size == 0:
                print("警告: 曲线评估点为空")
                return []

            curve_q = evalpts[:, 0]
            curve_wl = evalpts[:, 1]

            for q, wl in zip(self.discharges, self.water_levels):
                try:
                    idx = np.argmin(np.abs(curve_q - q))
                    fit_wl = curve_wl[idx]
                    error = abs(wl - fit_wl)
                    results.append([q, wl, fit_wl, error])
                except Exception as e:
                    print(f"处理数据点 ({q}, {wl}) 失败: {str(e)}")
                    results.append([q, wl, np.nan, np.nan])
        except Exception as e:
            print(f"评估曲线时发生严重错误: {str(e)}")
            results = []  # 确保返回空列表而非None
        return sorted(results, key=lambda x: x[0]) if results else []



def start_interface():
    """创建开始界面"""
    fig = plt.figure("水位流量关系曲线拟合小程序", figsize=(10, 6))
    ax = fig.add_subplot(111)
    ax.axis('off')  # 隐藏坐标轴

    # 添加标题（正中靠上）
    fig.text(
        0.5, 0.85, "水位流量综合曲线拟合小程序",
        ha='center', va='center',
        fontsize=24, fontweight='bold', color='navy'
    )

    # 添加单位信息（左下）
    fig.text(
        0.05, 0.05, "编制单位：海南省水文水资源勘测局东部大队",
        ha='left', va='bottom',
        fontsize=12, color='gray'
    )

    # 添加导入按钮（居中）
    ax_button = fig.add_axes([0.4, 0.4, 0.2, 0.15])  # [left, bottom, width, height]
    btn_import = Button(
        ax_button, '导入综合水位流量数据',
        color='lightgreen', hovercolor='limegreen'
    )
    btn_import.label.set_fontsize(14)
    btn_import.label.set_fontweight('bold')

    # 设置按钮点击回调
    def on_import_clicked(event):
        plt.close(fig)  # 关闭开始界面
        load_and_show_main_app()

    btn_import.on_clicked(on_import_clicked)

    plt.show()


def load_and_show_main_app():
    """加载数据并显示主程序"""
    years, tests, discharges, water_levels, file_path = load_hydrological_data()
    if discharges is None:
        print("程序终止：未选择有效数据文件")
        return

    try:
        editor = CurveEditor(years, tests, discharges, water_levels, file_path)
        plt.show()
    except Exception as e:
        print(f"初始化失败: {str(e)}")

# ... 主程序 ...

if __name__ == '__main__':
    start_interface()  # 显示开始界面
    # 加载数据
    years, tests, discharges, water_levels, file_path = load_hydrological_data()
    if discharges is None:
        print("程序终止：未选择有效数据文件")
        exit()

    # 初始化编辑器
    try:
        editor = CurveEditor(years, tests, discharges, water_levels, file_path)
        plt.show()
    except Exception as e:
        print(f"初始化失败: {str(e)}")
