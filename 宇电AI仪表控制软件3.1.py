import tkinter as tk
from tkinter import ttk, messagebox, filedialog, colorchooser
import serial
import serial.tools.list_ports
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import threading
import time
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
import json
import os
import pystray
from PIL import Image, ImageDraw
from pystray import MenuItem as item
import sys
import traceback
import logging
import gc

# --- 1. 日志配置 ---
logging.basicConfig(
    filename='run_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    err_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    logging.critical("致命错误捕获:\n" + err_msg)
sys.excepthook = handle_exception

# --- 2. 全局配置 ---
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial']
matplotlib.rcParams['axes.unicode_minus'] = False

DB_FILE = "multi_channel_history.db"
CONFIG_FILE = "instruments_config_v5.json"
SCHEDULE_FILE = "schedule_tasks.json" 
DATA_RETENTION_DAYS = 7     
MAX_DISPLAY_POINTS = 800  

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("AI-708 实验室温控系统 (v5.2 终极完整版)")
        self.root.state('zoomed') 
        
        # 拦截关闭事件：点击X进入后台
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

        # 样式配置
        self.root.option_add('*TCombobox*Listbox.font', ("Arial", 14))
        self.root.option_add('*Menu.font', ("微软雅黑", 12)) 
        style = ttk.Style()
        style.theme_use('clam') 
        style.configure("Treeview.Heading", font=("微软雅黑", 12, "bold"), rowheight=30)
        style.configure("Treeview", font=("Arial", 11), rowheight=25)
        style.configure("TNotebook.Tab", font=("微软雅黑", 14, "bold"), padding=[10, 5])
        
        # --- 变量初始化 ---
        self.is_running = True
        self.is_visible = True
        self.serial_conn = None 
        self.selected_port = tk.StringVar()
        self.protocol_type = tk.StringVar(value="AIBUS")
        self.heartbeat_var = tk.StringVar(value="●")
        
        # 加载配置
        self.instruments = self.load_config() 
        self.schedule_tasks = self.load_schedule() 
        self.lines = {} 
        self.sch_canvases = {} 

        # 导出相关变量
        self.export_mode = tk.StringVar(value="recent") 
        self.recent_hours = tk.StringVar(value="5")
        now_dt = datetime.now()
        start_dt = now_dt - timedelta(hours=5)
        self.start_time_str = tk.StringVar(value=start_dt.strftime("%Y-%m-%d %H:%M"))
        self.end_time_str = tk.StringVar(value=now_dt.strftime("%Y-%m-%d %H:%M"))

        self.plot_duration_val = tk.StringVar(value="60") 
        self.plot_duration_unit = tk.StringVar(value="分钟") 

        # 数据库初始化
        self.init_db()
        self.cleanup_old_data()

        # --- 界面构建 ---
        self.create_menu()
        self.setup_main_layout() 
        
        self.refresh_ports() # 初始刷新
        
        # 启动托盘图标 (独立线程)
        threading.Thread(target=self.init_tray_icon, daemon=True).start()
        
        # 数据采集线程
        self.thread = threading.Thread(target=self.data_loop, daemon=True)
        self.thread.start()
        
        self.root.after(1000, self.update_ui_loop)

    # --- 配置文件加载 (核心修复) ---
    def load_config(self):
        default_inst = [{"name": "1号仪表", "addr": 1, "color": "#ff0000", "decimal": 0.1, "k": 1.0, "l": 0.0}]
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for inst in data:
                        # 仅当字段缺失时补充默认值，绝不覆盖用户设置
                        if "decimal" not in inst: inst["decimal"] = 0.1 
                        if "k" not in inst: inst["k"] = 1.0
                        if "l" not in inst: inst["l"] = 0.0
                    return data
            except: pass
        return default_inst
    
    def save_config(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.instruments, f, indent=2, ensure_ascii=False)
        self.refresh_schedule_tabs()

    def load_schedule(self):
        if os.path.exists(SCHEDULE_FILE):
            try:
                with open(SCHEDULE_FILE, 'r', encoding='utf-8') as f: return json.load(f)
            except: pass
        return []

    def save_schedule(self):
        with open(SCHEDULE_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.schedule_tasks, f, indent=2, ensure_ascii=False)

    # ================= 界面布局 =================
    def setup_main_layout(self):
        top_frame = tk.Frame(self.root, bg="#f0f0f0", height=50)
        top_frame.pack(fill="x", side="top")
        
        # 端口选择
        tk.Label(top_frame, text="USB接口选择:", bg="#f0f0f0", font=("微软雅黑", 12)).pack(side="left", padx=5)
        self.cb_ports = ttk.Combobox(top_frame, textvariable=self.selected_port, width=15)
        self.cb_ports.pack(side="left", padx=5)
        self.cb_ports.bind('<Button-1>', lambda e: self.refresh_ports())
        
        # 模式选择
        tk.Label(top_frame, text="模式选择:", bg="#f0f0f0", font=("微软雅黑", 12)).pack(side="left", padx=15)
        self.cb_proto = ttk.Combobox(top_frame, textvariable=self.protocol_type, values=["AIBUS", "MODBUS"], width=10, state="readonly")
        self.cb_proto.pack(side="left", padx=5)

        self.lbl_status = tk.Label(top_frame, text="就绪", fg="gray", bg="#f0f0f0", font=("微软雅黑", 12))
        self.lbl_status.pack(side="left", padx=20)
        tk.Label(top_frame, textvariable=self.heartbeat_var, fg="red", bg="#f0f0f0", font=("Arial", 16)).pack(side="right", padx=10)

        # 选项卡
        self.main_notebook = ttk.Notebook(self.root)
        self.main_notebook.pack(fill="both", expand=True, padx=5, pady=5)

        self.tab_monitor = tk.Frame(self.main_notebook)
        self.main_notebook.add(self.tab_monitor, text="  📊 数据监测  ")
        self.setup_monitor_tab(self.tab_monitor)

        self.tab_schedule = tk.Frame(self.main_notebook)
        self.main_notebook.add(self.tab_schedule, text="  ⏱️ 数据设定 (定时)  ")
        self.setup_schedule_main_tab(self.tab_schedule)

    def refresh_ports(self):
        current_port = self.selected_port.get()
        ports = sorted(list(set([p.device for p in serial.tools.list_ports.comports()] + ["COM1","COM2","COM3","COM4"])))
        self.cb_ports['values'] = ports
        if current_port not in ports and ports:
            self.cb_ports.current(0)
        elif not current_port and ports:
            self.cb_ports.current(0)

    def setup_monitor_tab(self, parent):
        paned = tk.PanedWindow(parent, orient="horizontal", sashwidth=5, bg="#ddd")
        paned.pack(fill="both", expand=True)

        # 左侧：绘图
        left_frame = tk.Frame(paned)
        paned.add(left_frame, width=900, stretch="always")
        
        graph_frame = tk.Frame(left_frame)
        graph_frame.pack(side="top", fill="both", expand=True)
        self.fig, self.ax = plt.subplots(figsize=(5, 4), dpi=100)
        self.fig.subplots_adjust(bottom=0.1, left=0.08, right=0.95, top=0.92)
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        self.init_plot_lines()

        # 底部控制区 (包含导出)
        ctrl_frame = tk.LabelFrame(left_frame, text="显示与导出", font=("微软雅黑", 10, "bold"), bg="white", padx=5, pady=5)
        ctrl_frame.pack(side="bottom", fill="x", padx=5, pady=5)
        
        row1 = tk.Frame(ctrl_frame, bg="white")
        row1.pack(fill="x", pady=2)
        tk.Label(row1, text="【绘图】显示过去:", bg="white").pack(side="left")
        v_cmd = (self.root.register(self.validate_number), '%P')
        tk.Entry(row1, textvariable=self.plot_duration_val, width=5, validate="key", validatecommand=v_cmd).pack(side="left")
        ttk.Combobox(row1, textvariable=self.plot_duration_unit, values=["分钟", "小时"], width=5, state="readonly").pack(side="left")
        tk.Button(row1, text="更新绘图", command=self.force_redraw_plot, bg="#ddd").pack(side="left", padx=10)

        row2 = tk.Frame(ctrl_frame, bg="white")
        row2.pack(fill="x", pady=5)
        tk.Label(row2, text="【导出】", bg="white", font=("bold")).pack(side="left")
        
        f_recent = tk.Frame(row2, bg="white")
        f_recent.pack(side="left", padx=10)
        ttk.Radiobutton(f_recent, text="最近", variable=self.export_mode, value="recent").pack(side="left")
        tk.Entry(f_recent, textvariable=self.recent_hours, width=4).pack(side="left")
        tk.Label(f_recent, text="小时", bg="white").pack(side="left")

        f_range = tk.Frame(row2, bg="white")
        f_range.pack(side="left", padx=10)
        ttk.Radiobutton(f_range, text="时间段", variable=self.export_mode, value="range").pack(side="left")
        tk.Entry(f_range, textvariable=self.start_time_str, width=16).pack(side="left")
        tk.Label(f_range, text="至", bg="white").pack(side="left")
        tk.Entry(f_range, textvariable=self.end_time_str, width=16).pack(side="left")

        tk.Button(row2, text="导出Excel/CSV", bg="#4CAF50", fg="white", font=("微软雅黑", 10, "bold"), command=self.export_data).pack(side="right", padx=10)

        # 右侧：表格
        right_frame = tk.Frame(paned)
        paned.add(right_frame, stretch="always")
        tree_scroll = tk.Frame(right_frame)
        tree_scroll.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_scroll, show="headings")
        vsb = ttk.Scrollbar(tree_scroll, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.setup_tree_columns()

    def setup_schedule_main_tab(self, parent):
        tk.Label(parent, text="校准逻辑说明: 实际值(Y) = (原始值 × 读数倍率) × k + l", font=("微软雅黑", 10), fg="#555", pady=5).pack(side="top")
        self.sch_notebook = ttk.Notebook(parent)
        self.sch_notebook.pack(fill="both", expand=True, padx=10, pady=5)
        self.refresh_schedule_tabs()

    def refresh_schedule_tabs(self):
        for tab in self.sch_notebook.tabs():
            self.sch_notebook.forget(tab)
        
        self.sch_canvases = {} 
        
        if not self.instruments:
            lbl = tk.Label(self.sch_notebook, text="请先在【配置】菜单中添加仪表", font=("微软雅黑", 14))
            self.sch_notebook.add(lbl, text="无仪表")
            return

        for inst in self.instruments:
            addr = inst['addr']
            name = inst['name']
            tab_frame = tk.Frame(self.sch_notebook)
            self.sch_notebook.add(tab_frame, text=f" [{addr}] {name} ")
            self.setup_single_instrument_schedule_ui(tab_frame, addr)

    def setup_single_instrument_schedule_ui(self, parent, target_addr):
        paned = tk.PanedWindow(parent, orient="horizontal", sashwidth=5)
        paned.pack(fill="both", expand=True, padx=5, pady=5)
        
        left_frame = tk.Frame(paned)
        paned.add(left_frame, width=500)
        
        ctrl_box = ttk.LabelFrame(left_frame, text="添加任务", padding=5)
        ctrl_box.pack(fill="x", pady=5)
        
        tk.Label(ctrl_box, text="时间:").pack(side="left")
        # 默认当前时间
        time_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d %H:%M"))
        tk.Entry(ctrl_box, textvariable=time_var, width=16).pack(side="left", padx=2)
        
        tk.Label(ctrl_box, text="实际目标值:").pack(side="left", padx=5)
        # 默认值10
        val_var = tk.StringVar(value="10")
        tk.Entry(ctrl_box, textvariable=val_var, width=6).pack(side="left", padx=2)
        
        cols = ("time", "real_val", "inst_val", "status")
        tree = ttk.Treeview(left_frame, columns=cols, show="headings")
        tree.heading("time", text="触发时间")
        tree.heading("real_val", text="实际目标(Y)")
        tree.heading("inst_val", text="仪表写入(X)")
        tree.heading("status", text="状态")
        tree.column("time", width=140, anchor="center")
        tree.column("real_val", width=100, anchor="center")
        tree.column("inst_val", width=100, anchor="center")
        tree.column("status", width=80, anchor="center")
        
        vsb = ttk.Scrollbar(left_frame, orient="vertical", command=tree.yview)
        tree.configure(yscroll=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        
        def refresh_view():
            tree.delete(*tree.get_children())
            my_tasks = [t for t in self.schedule_tasks if t['addr'] == target_addr]
            my_tasks.sort(key=lambda x: x['time'])
            
            inst_cfg = next((i for i in self.instruments if i['addr'] == target_addr), None)
            k = inst_cfg.get('k', 1.0)
            l = inst_cfg.get('l', 0.0)
            
            times = []
            values = []
            
            for task in my_tasks:
                status = "✅已执行" if task.get('done') else "⏳等待"
                real_val = float(task['val'])
                # 反向计算展示
                try:
                    inst_val = (real_val - l) / k
                except: inst_val = 0
                
                tree.insert("", "end", values=(task['time'], real_val, f"{inst_val:.2f}", status))
                try:
                    dt = datetime.strptime(task['time'], "%Y-%m-%d %H:%M")
                    times.append(dt)
                    values.append(real_val)
                except: pass
            
            # 台阶图
            ax_step.clear()
            ax_step.set_title(f"地址{target_addr} 实际值设定走势", fontsize=10)
            ax_step.grid(True, linestyle='--', alpha=0.5)
            
            if times:
                if len(times) > 0:
                    last_time = max(times[-1], datetime.now())
                    future_time = last_time + timedelta(hours=2)
                    times_plot = times + [future_time]
                    values_plot = values + [values[-1]] 
                    ax_step.step(times_plot, values_plot, where='post', color='blue', linewidth=2)
                    ax_step.plot(times, values, 'o', color='red', markersize=4)
                fig_step.autofmt_xdate()
            canvas_step.draw()

        def add_task_action():
            try:
                t_str = time_var.get().strip()
                v = float(val_var.get())
                datetime.strptime(t_str, "%Y-%m-%d %H:%M") 
                self.schedule_tasks.append({
                    "time": t_str, "addr": target_addr, "val": v, "done": False
                })
                self.save_schedule()
                refresh_view()
            except ValueError:
                messagebox.showerror("错误", "时间格式(YYYY-MM-DD HH:MM)或数值错误")

        def del_task_action():
            sel = tree.selection()
            if not sel: return
            if messagebox.askyesno("确认", "删除选中任务？"):
                for item_id in reversed(sel):
                    val = tree.item(item_id, "values")
                    t_str, t_val = val[0], float(val[1])
                    for i, task in enumerate(self.schedule_tasks):
                        if task['addr'] == target_addr and task['time'] == t_str and float(task['val']) == t_val:
                            del self.schedule_tasks[i]
                            break
                self.save_schedule()
                refresh_view()

        tk.Button(ctrl_box, text="➕添加", command=add_task_action, bg="#eef").pack(side="left", padx=10)
        tk.Button(ctrl_box, text="🗑️删除", command=del_task_action, bg="#fee").pack(side="left")

        right_frame = tk.Frame(paned)
        paned.add(right_frame)
        
        fig_step, ax_step = plt.subplots(figsize=(4, 3), dpi=100)
        canvas_step = FigureCanvasTkAgg(fig_step, master=right_frame)
        canvas_step.get_tk_widget().pack(fill="both", expand=True)
        
        self.sch_canvases[target_addr] = (fig_step, ax_step, canvas_step)
        refresh_view()

    # --- 菜单栏 ---
    def create_menu(self):
        menubar = tk.Menu(self.root)
        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(label="仪表参数与校准", command=self.open_settings_window)
        menubar.add_cascade(label="⚙️ 配置", menu=config_menu)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于", command=lambda: messagebox.showinfo("关于", "宇电温控系统 v5.2"))
        menubar.add_cascade(label="ℹ️ 帮助", menu=help_menu)
        self.root.config(menu=menubar)

    def force_redraw_plot(self):
        self.init_plot_lines()
        self.update_plot_fast()

    # --- 仪表配置窗口 (带预览) ---
    def open_settings_window(self):
        win = tk.Toplevel(self.root)
        win.title("仪表参数与流量校准")
        win.geometry("900x600") 
        
        FONT_UI = ("微软雅黑", 12)
        list_frame = tk.Frame(win, padx=10, pady=10)
        list_frame.pack(side="left", fill="y")
        lb = tk.Listbox(list_frame, font=("Arial", 12), width=25, height=20)
        lb.pack(fill="y", expand=True, pady=5)
        
        edit_frame = tk.Frame(win, padx=20, pady=20)
        edit_frame.pack(side="left", fill="both", expand=True)
        
        # 变量
        vars_map = {
            "name": tk.StringVar(),
            "addr": tk.StringVar(),
            "color": tk.StringVar(value="#ff0000"),
            "dec": tk.StringVar(value="0.1"),
            "k": tk.StringVar(value="1.0"),
            "l": tk.StringVar(value="0.0")
        }
        
        # 实时预览标签
        preview_lbl = tk.Label(edit_frame, text="计算预览: 请先选择仪表", fg="blue", font=("Arial", 10))

        def calc_preview(*args):
            try:
                # 模拟一个原始整数值 40
                raw_sample = 40
                dec = float(vars_map["dec"].get())
                k = float(vars_map["k"].get())
                l = float(vars_map["l"].get())
                disp = raw_sample * dec
                real = disp * k + l
                preview_lbl.config(text=f"测试: 假设仪表传回整数40 -> 乘以倍率{dec} = {disp} -> 实际值 = {real}")
            except: 
                preview_lbl.config(text="计算预览: 输入格式错误")

        # 绑定变量变化
        for key in ["dec", "k", "l"]:
            vars_map[key].trace("w", calc_preview)

        # 布局
        r = 0
        layout = [("名称:", "name"), ("地址(Addr):", "addr"), ("颜色:", "color"), 
                  ("读数倍率:", "dec"), ("斜率 k:", "k"), ("截距 l:", "l")]
        
        for label_text, key in layout:
            tk.Label(edit_frame, text=label_text, font=FONT_UI).grid(row=r, column=0, pady=5, sticky="e")
            if key == "color":
                tk.Button(edit_frame, text="■", command=lambda: vars_map["color"].set(colorchooser.askcolor()[1] or vars_map["color"].get())).grid(row=r, column=1, sticky="w")
            else:
                tk.Entry(edit_frame, textvariable=vars_map[key], width=15, font=("Arial", 12)).grid(row=r, column=1, sticky="w")
            r += 1
        
        preview_lbl.grid(row=r, column=0, columnspan=2, pady=10)

        def refresh_list():
            lb.delete(0, tk.END)
            for inst in self.instruments:
                lb.insert(tk.END, f"[{inst['addr']}] {inst['name']}")
            if self.instruments:
                lb.selection_set(0)
                on_select(None)
        
        def on_select(evt):
            if not lb.curselection(): return
            idx = lb.curselection()[0]
            d = self.instruments[idx]
            vars_map["name"].set(d['name'])
            vars_map["addr"].set(str(d['addr']))
            vars_map["color"].set(d['color'])
            # 兼容旧键名
            vars_map["dec"].set(str(d.get('dec', d.get('decimal', 0.1)))) 
            vars_map["k"].set(str(d.get('k', 1.0)))
            vars_map["l"].set(str(d.get('l', 0.0)))
            calc_preview()

        lb.bind('<<ListboxSelect>>', on_select)
        
        def save_change(action):
            try:
                addr = int(vars_map["addr"].get().strip())
                new_data = {
                    "name": vars_map["name"].get(),
                    "addr": addr,
                    "color": vars_map["color"].get(),
                    "decimal": float(vars_map["dec"].get()),
                    "k": float(vars_map["k"].get()),
                    "l": float(vars_map["l"].get())
                }
                
                if action == "add":
                    self.instruments.append(new_data)
                elif action == "update":
                    if not lb.curselection(): return
                    idx = lb.curselection()[0]
                    self.instruments[idx] = new_data
                elif action == "delete":
                    if not lb.curselection(): return
                    if not messagebox.askyesno("确认", "删除此仪表？"): return
                    del self.instruments[lb.curselection()[0]]
                
                self.save_config() 
                refresh_list()
                self.setup_tree_columns()
                self.force_redraw_plot()
                messagebox.showinfo("成功", "设置已更新")
            except Exception as e:
                messagebox.showerror("错误", str(e))

        refresh_list()
        btn_frame = tk.Frame(edit_frame, pady=20)
        btn_frame.grid(row=r+1, column=0, columnspan=2)
        tk.Button(btn_frame, text="新增", bg="#81C784", width=8, command=lambda: save_change("add")).pack(side="left", padx=5)
        tk.Button(btn_frame, text="保存修改", bg="#64B5F6", width=8, command=lambda: save_change("update")).pack(side="left", padx=5)
        tk.Button(btn_frame, text="删除", bg="#E57373", width=8, command=lambda: save_change("delete")).pack(side="left", padx=5)

    def pick_color(self, var, btn):
        color = colorchooser.askcolor(title="选择颜色")[1]
        if color: var.set(color); btn.config(bg=color)

    # --- 通讯逻辑 (含校准) ---
    def write_sv(self, addr, real_target_value):
        inst = next((i for i in self.instruments if i['addr'] == addr), None)
        if not inst: return False
        
        k = inst.get('k', 1.0)
        l = inst.get('l', 0.0)
        # 兼容两种键名
        dec = inst.get('dec', inst.get('decimal', 0.1))
        
        # 反向计算: Raw = ((Real - l) / k) / dec
        try:
            raw_float = ((real_target_value - l) / k) / dec
            raw_int = int(round(raw_float))
        except: return False

        proto = self.protocol_type.get()
        if proto == "AIBUS": return self.write_aibus_sv_raw(addr, raw_int)
        else: return self.write_modbus_sv_raw(addr, raw_int)

    def write_aibus_sv_raw(self, addr, val_int):
        if self.serial_conn is None: return False
        try:
            param_id = 0x00 
            if val_int < 0: val_int += 65536 
            data_low = val_int & 0xFF
            data_high = (val_int >> 8) & 0xFF
            chk_val = 0x43 + param_id + data_low + (data_high * 256) + addr
            chk_low = chk_val & 0xFF
            chk_high = (chk_val >> 8) & 0xFF
            cmd = bytes([0x80 + addr, 0x80 + addr, 0x43, param_id, data_low, data_high, chk_low, chk_high])
            self.serial_conn.flushInput()
            self.serial_conn.write(cmd)
            time.sleep(0.1) 
            return True
        except Exception as e:
            logging.error(f"AIBUS Write Error: {e}")
            return False

    def write_modbus_sv_raw(self, addr, val_int):
        if self.serial_conn is None: return False
        try:
            reg_addr = 0x0000 
            if val_int < 0: val_int += 65536
            base_cmd = bytes([addr, 0x06, (reg_addr >> 8) & 0xFF, reg_addr & 0xFF, (val_int >> 8) & 0xFF, val_int & 0xFF])
            crc = 0xFFFF
            for byte in base_cmd:
                crc ^= byte
                for _ in range(8):
                    if crc & 1: crc = (crc >> 1) ^ 0xA001
                    else: crc >>= 1
            cmd = base_cmd + bytes([crc & 0xFF, (crc >> 8) & 0xFF])
            self.serial_conn.flushInput()
            self.serial_conn.write(cmd)
            time.sleep(0.1) 
            return True
        except Exception as e:
            logging.error(f"Modbus Write Error: {e}")
            return False

    # --- 后台数据线程 ---
    def data_loop(self):
        logging.info("数据线程启动")
        while self.is_running:
            start_ts = time.time()
            try:
                if self.is_visible:
                    try: self.heartbeat_var.set("●" if self.heartbeat_var.get() == "○" else "○")
                    except: pass
                
                port = self.selected_port.get()
                if port and self.open_serial(port):
                    now = datetime.now()
                    now_str = now.strftime("%Y-%m-%d %H:%M")
                    
                    # 1. 任务
                    for task in self.schedule_tasks:
                        if not task.get('done', False) and task['time'] == now_str:
                            logging.info(f"执行定时任务: 地址{task['addr']} -> 目标{task['val']}")
                            success = self.write_sv(task['addr'], float(task['val']))
                            if success:
                                task['done'] = True
                                self.save_schedule()
                            else:
                                logging.error(f"任务失败: {task}")
                    
                    # 2. 读取 (应用校准)
                    has_data = False
                    for inst in self.instruments:
                        addr = inst['addr']
                        try:
                            # 获取原始整数
                            raw_val = self.read_temp_raw(addr)
                            if raw_val is not None:
                                dec = inst.get('dec', inst.get('decimal', 0.1))
                                k = inst.get('k', 1.0)
                                l = inst.get('l', 0.0)
                                
                                # 核心计算公式
                                display_val = raw_val * dec
                                real_val = display_val * k + l
                                
                                self.cursor.execute("INSERT INTO records VALUES (?, ?, ?, ?, ?)", 
                                    (now.timestamp(), now.strftime('%Y-%m-%d'), now.strftime('%H:%M:%S'), addr, real_val))
                                has_data = True
                        except: pass
                    if has_data:
                        try: self.conn.commit()
                        except: pass
                else:
                    time.sleep(2)
            except Exception as e:
                logging.error(f"Loop Error: {e}")
                time.sleep(1)
            elapsed = time.time() - start_ts
            if elapsed < 1.0:
                time.sleep(1.0 - elapsed)

    def update_ui_loop(self):
        if not self.is_running: return
        if self.is_visible:
            self.update_table()
            self.update_plot_fast()
        if int(time.time()) % 60 == 0: gc.collect()
        self.root.after(1000, self.update_ui_loop)

    def update_plot_fast(self):
        try:
            data_map, unit, limit_val = self.get_plot_data_optimized()
            if not data_map: return
            has_data = False
            for addr, d in data_map.items():
                if addr in self.lines:
                    self.lines[addr].set_data(d['x'], d['y'])
                    if len(d['x']) > 0: has_data = True
            if has_data:
                self.ax.relim(); self.ax.autoscale_view(); self.ax.set_xlim(-limit_val, 0)
            self.canvas.draw()
        except Exception as e: logging.error(f"Plot Error: {e}")

    def init_plot_lines(self):
        self.ax.clear()
        self.ax.set_title("实时趋势 (实际校准值)", fontsize=12)
        self.ax.set_xlabel("时间", fontsize=10)
        self.ax.set_ylabel("值", fontsize=10)
        self.ax.grid(True, linestyle='--', alpha=0.5)
        self.lines = {}
        for inst in self.instruments:
            line, = self.ax.plot([], [], color=inst['color'], label=inst['name'], linewidth=1.5)
            self.lines[inst['addr']] = line
        self.ax.legend(loc='upper left', fontsize=9)

    def get_plot_data_optimized(self):
        try:
            val = int(self.plot_duration_val.get()); unit = self.plot_duration_unit.get()
            if val <= 0: val = 60
            delta = timedelta(minutes=val) if unit == "分钟" else timedelta(hours=val)
            start_ts = (datetime.now() - delta).timestamp()
            self.cursor.execute("SELECT timestamp, address, temperature FROM records WHERE timestamp > ? ORDER BY timestamp ASC", (start_ts,))
            rows = self.cursor.fetchall()
            if not rows: return {}, unit, val
            data_map = {i['addr']: {'x': [], 'y': []} for i in self.instruments}
            now_ts = time.time()
            total_points = len(rows)
            step = 1
            if total_points > MAX_DISPLAY_POINTS * len(self.instruments):
                step = total_points // (MAX_DISPLAY_POINTS * len(self.instruments))
            for i in range(0, total_points, step):
                r = rows[i]; addr = r[1]
                if addr in data_map:
                    diff = r[0] - now_ts
                    x_val = diff / 60.0 if unit == "分钟" else diff / 3600.0
                    data_map[addr]['x'].append(x_val); data_map[addr]['y'].append(r[2])
            return data_map, unit, val
        except: return {}, "分钟", 60

    def update_table(self):
        try:
            self.cursor.execute("SELECT time_str, address, temperature FROM records ORDER BY timestamp DESC LIMIT 50")
            rows = self.cursor.fetchall()
            display_data = {}; ordered_times = []
            for r in rows:
                t_str, addr, temp = r[0], r[1], r[2]
                if t_str not in display_data: display_data[t_str] = {}; ordered_times.append(t_str)
                display_data[t_str][addr] = temp
            self.tree.delete(*self.tree.get_children())
            for t in ordered_times[:15]: 
                row_vals = [t]
                for inst in self.instruments:
                    val = display_data[t].get(inst['addr'], "--")
                    row_vals.append(f"{val:.1f}" if isinstance(val, float) else val)
                self.tree.insert("", "end", values=row_vals)
        except: pass

    # --- 基础函数 ---
    def open_serial(self, port):
        try:
            if self.serial_conn is not None:
                if self.serial_conn.port == port and self.serial_conn.is_open: return True 
                else: self.serial_conn.close()
            self.serial_conn = serial.Serial(port=port, baudrate=9600, timeout=0.2)
            self.lbl_status.config(text=f"运行中: {port}", fg="green")
            return True
        except: 
            self.serial_conn = None; self.lbl_status.config(text="连接断开", fg="red"); return False
    
    # 纯净读取函数：只返回原始整数，不带任何缩放
    def read_temp_raw(self, addr):
        return self.read_aibus_raw(addr) if self.protocol_type.get() == "AIBUS" else self.read_modbus_raw(addr)
    
    def read_aibus_raw(self, addr):
        if self.serial_conn is None: return None
        try:
            chk = 0x52 + addr; cmd = bytes([0x80 + addr, 0x80 + addr, 0x52, 0x00, 0x00, 0x00, chk & 0xFF, (chk >> 8) & 0xFF])
            self.serial_conn.flushInput(); self.serial_conn.write(cmd); resp = self.serial_conn.read(10)
            if len(resp) < 10: return None
            pv = resp[0] + (resp[1] << 8)
            if pv > 32767: pv -= 65536
            return pv
        except: return None
    
    def read_modbus_raw(self, addr):
        if self.serial_conn is None: return None
        try:
            def calc_crc(d):
                c = 0xFFFF
                for b in d:
                    c ^= b
                    for _ in range(8): c = (c >> 1) ^ 0xA001 if c & 1 else c >> 1
                return c
            base = bytes([addr, 0x03, 0x00, 0x00, 0x00, 0x04]); crc = calc_crc(base)
            self.serial_conn.flushInput(); self.serial_conn.write(base + bytes([crc & 0xFF, crc >> 8])); resp = self.serial_conn.read(13)
            if len(resp) < 13: return None
            pv = (resp[3] << 8) + resp[4]
            if pv > 32767: pv -= 65536
            return pv
        except: return None

    def init_db(self):
        self.conn = sqlite3.connect(DB_FILE, check_same_thread=False)
        self.cursor = self.conn.cursor()
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS records (timestamp REAL, date_str TEXT, time_str TEXT, address INTEGER, temperature REAL)''')
        self.cursor.execute('CREATE INDEX IF NOT EXISTS idx_ts ON records(timestamp)')
        self.conn.commit()
    
    def validate_number(self, val): return val.isdigit() or val == ""
    
    def cleanup_old_data(self):
        try: t = (datetime.now()-timedelta(days=DATA_RETENTION_DAYS)).timestamp(); self.cursor.execute("DELETE FROM records WHERE timestamp < ?", (t,)); self.conn.commit()
        except: pass

    # --- 完整的导出功能 ---
    def export_data(self):
        mode = self.export_mode.get()
        start_dt, end_dt = None, None
        try:
            if mode == "recent":
                h = float(self.recent_hours.get())
                end_dt = datetime.now()
                start_dt = end_dt - timedelta(hours=h)
            else:
                fmt = "%Y-%m-%d %H:%M"
                start_dt = datetime.strptime(self.start_time_str.get(), fmt)
                end_dt = datetime.strptime(self.end_time_str.get(), fmt)
            
            query = "SELECT date_str, time_str, address, temperature FROM records WHERE timestamp BETWEEN ? AND ? ORDER BY timestamp ASC"
            df = pd.read_sql_query(query, self.conn, params=(start_dt.timestamp(), end_dt.timestamp()))
            
            if df.empty: 
                messagebox.showwarning("提示", "该时间段内无数据")
                return
            
            pivot_df = df.pivot_table(index=['date_str', 'time_str'], columns='address', values='temperature', aggfunc='first')
            new_cols = []
            name_map = {i['addr']: i['name'] for i in self.instruments}
            for addr in pivot_df.columns:
                new_cols.append(name_map.get(addr, f"地址_{addr}"))
            pivot_df.columns = new_cols
            pivot_df.reset_index(inplace=True)
            
            default_name = f"监测数据_{start_dt.strftime('%Y%m%d_%H%M')}.csv"
            fname = filedialog.asksaveasfilename(initialfile=default_name, filetypes=[("CSV文件", "*.csv")])
            
            if fname:
                pivot_df.to_csv(fname, index=False, encoding="utf-8-sig")
                messagebox.showinfo("成功", f"成功导出 {len(pivot_df)} 条记录")
                
        except Exception as e:
            messagebox.showerror("导出错误", f"详细信息: {str(e)}")

    def create_image(self):
        image = Image.new('RGB', (64, 64), "blue")
        ImageDraw.Draw(image).rectangle((16, 16, 48, 48), fill="white")
        return image

    def init_tray_icon(self):
        try:
            menu = (item('显示监控窗口', self.show_window), item('彻底退出系统', self.quit_app))
            self.icon = pystray.Icon("name", self.create_image(), "温控系统(运行中)", menu)
            self.icon.run()
        except Exception as e:
            logging.error(f"Tray Icon Error: {e}")

    def hide_window(self):
        self.root.withdraw()
        self.is_visible = False 
        try: self.icon.notify("软件仍在后台运行记录数据", "已最小化到托盘")
        except: pass
        gc.collect() 

    def show_window(self):
        self.root.deiconify()
        self.root.state('zoomed') 
        self.is_visible = True 
        self.force_redraw_plot() 

    def quit_app(self):
        if messagebox.askyesno("确认退出", "退出后将停止记录数据，确定吗？"):
            self.is_running = False
            try: self.icon.stop()
            except: pass
            try: self.conn.close()
            except: pass
            try: self.root.destroy()
            except: pass
            os._exit(0)

    def setup_tree_columns(self):
        cols = ["time"] + [f"addr_{i['addr']}" for i in self.instruments]
        self.tree["columns"] = cols
        self.tree.heading("time", text="时间"); self.tree.column("time", width=180, anchor="center")
        for inst in self.instruments:
            self.tree.heading(f"addr_{inst['addr']}", text=inst['name']); self.tree.column(f"addr_{inst['addr']}", width=100, anchor="center")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = App(root)
        root.mainloop()
    except Exception as e:
        import traceback
        error_msg = traceback.format_exc()
        with open("crash_error.txt", "w", encoding="utf-8") as f:
            f.write(error_msg)
        print("!"*30)
        print("致命错误：")
        print(error_msg)
        print("!"*30)
        input("按回车退出...")