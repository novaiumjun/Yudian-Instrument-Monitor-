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

# --- 1. 日志与异常捕获配置 ---
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
CONFIG_FILE = "instruments_config.json"
DATA_RETENTION_DAYS = 7     
MAX_DISPLAY_POINTS = 800  # 绘图最大点数（超过抽稀）

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("AI-708 实验室温控系统 (最终完整版)")
        self.root.state('zoomed') 
        
        # 【交互逻辑】点击窗口X -> 隐藏（后台运行）
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

        # 样式配置
        self.root.option_add('*TCombobox*Listbox.font', ("Arial", 14))
        self.root.option_add('*Menu.font', ("微软雅黑", 12)) 
        style = ttk.Style()
        style.theme_use('clam') 
        style.configure("Treeview.Heading", font=("微软雅黑", 12, "bold"), rowheight=30)
        style.configure("Treeview", font=("Arial", 11), rowheight=25)
        
        # --- 变量初始化 ---
        self.is_running = True
        self.is_visible = True  # 控制绘图开关
        self.serial_conn = None 
        self.selected_port = tk.StringVar()
        self.protocol_type = tk.StringVar(value="AIBUS")
        self.heartbeat_var = tk.StringVar(value="●")
        
        # 加载仪表配置
        self.instruments = self.load_config() 
        self.lines = {} # 存储绘图线条对象

        # 导出相关变量
        self.export_mode = tk.StringVar(value="recent") 
        self.recent_hours = tk.StringVar(value="5")
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        past_str = (datetime.now() - timedelta(hours=5)).strftime("%Y-%m-%d %H:%M")
        self.start_time_str = tk.StringVar(value=past_str)
        self.end_time_str = tk.StringVar(value=now_str)

        # 绘图设置
        self.plot_duration_val = tk.StringVar(value="60") 
        self.plot_duration_unit = tk.StringVar(value="分钟") 

        # 数据库初始化
        self.init_db()
        self.cleanup_old_data()

        # --- 界面构建 ---
        self.create_menu()
        self.setup_ui()
        
        # --- 启动逻辑 ---
        self.refresh_ports()
        
        # 启动托盘 (Daemon线程)
        threading.Thread(target=self.init_tray_icon, daemon=True).start()
        
        # 数据采集线程 (始终运行)
        self.thread = threading.Thread(target=self.data_loop, daemon=True)
        self.thread.start()
        
        # UI刷新循环
        self.root.after(1000, self.update_ui_loop)

    def setup_ui(self):
        # 1. 顶部状态栏
        top_frame = tk.Frame(self.root, bg="#f0f0f0", height=50)
        top_frame.pack(fill="x", side="top")
        
        tk.Label(top_frame, text="端口:", bg="#f0f0f0", font=("微软雅黑", 12)).pack(side="left", padx=5)
        self.cb_ports = ttk.Combobox(top_frame, textvariable=self.selected_port, width=10)
        self.cb_ports.pack(side="left", padx=5)
        
        tk.Label(top_frame, text="协议:", bg="#f0f0f0", font=("微软雅黑", 12)).pack(side="left", padx=5)
        self.cb_proto = ttk.Combobox(top_frame, textvariable=self.protocol_type, values=["AIBUS", "MODBUS"], width=8, state="readonly")
        self.cb_proto.pack(side="left", padx=5)

        tk.Button(top_frame, text="刷新", command=self.refresh_ports).pack(side="left", padx=5)
        self.lbl_status = tk.Label(top_frame, text="就绪", fg="gray", bg="#f0f0f0", font=("微软雅黑", 12))
        self.lbl_status.pack(side="left", padx=10)
        tk.Label(top_frame, textvariable=self.heartbeat_var, fg="red", bg="#f0f0f0", font=("Arial", 16)).pack(side="right", padx=10)

        # 2. 主分割区
        paned = tk.PanedWindow(self.root, orient="horizontal", sashwidth=5, bg="#ddd")
        paned.pack(fill="both", expand=True)

        # === 左侧: 绘图 + 导出 ===
        left_frame = tk.Frame(paned)
        paned.add(left_frame, width=900, stretch="always")
        
        # 绘图区域
        graph_frame = tk.Frame(left_frame)
        graph_frame.pack(side="top", fill="both", expand=True)
        self.fig, self.ax = plt.subplots(figsize=(5, 4), dpi=100)
        self.fig.subplots_adjust(bottom=0.1, left=0.08, right=0.95, top=0.92)
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        
        # 初始化绘图线条
        self.init_plot_lines()

        # 底部控制区 (集成导出功能)
        ctrl_frame = tk.LabelFrame(left_frame, text="设置与导出", font=("微软雅黑", 10, "bold"), bg="white", pady=5)
        ctrl_frame.pack(side="bottom", fill="x", padx=5, pady=5)
        
        # 第一排：绘图范围
        row1 = tk.Frame(ctrl_frame, bg="white")
        row1.pack(fill="x", pady=2)
        tk.Label(row1, text="显示范围:", bg="white").pack(side="left")
        v_cmd = (self.root.register(self.validate_number), '%P')
        tk.Entry(row1, textvariable=self.plot_duration_val, width=5, validate="key", validatecommand=v_cmd).pack(side="left")
        ttk.Combobox(row1, textvariable=self.plot_duration_unit, values=["分钟", "小时"], width=5, state="readonly").pack(side="left")
        tk.Button(row1, text="立即应用", command=self.force_redraw_plot, bg="#ddd").pack(side="left", padx=10)

        # 第二排：导出设置
        row2 = tk.Frame(ctrl_frame, bg="white")
        row2.pack(fill="x", pady=5)
        
        # 模式1
        f_recent = tk.Frame(row2, bg="white")
        f_recent.pack(side="left", padx=5)
        ttk.Radiobutton(f_recent, text="最近", variable=self.export_mode, value="recent").pack(side="left")
        tk.Entry(f_recent, textvariable=self.recent_hours, width=4).pack(side="left")
        tk.Label(f_recent, text="小时", bg="white").pack(side="left")

        # 模式2
        f_range = tk.Frame(row2, bg="white")
        f_range.pack(side="left", padx=15)
        ttk.Radiobutton(f_range, text="范围", variable=self.export_mode, value="range").pack(side="left")
        tk.Entry(f_range, textvariable=self.start_time_str, width=14).pack(side="left")
        tk.Label(f_range, text="至", bg="white").pack(side="left")
        tk.Entry(f_range, textvariable=self.end_time_str, width=14).pack(side="left")

        tk.Button(row2, text="导出数据", bg="#4CAF50", fg="white", font=("微软雅黑", 10, "bold"), command=self.export_data).pack(side="right", padx=10)

        # === 右侧: 表格 ===
        right_frame = tk.Frame(paned)
        paned.add(right_frame, stretch="always")
        
        tree_scroll = tk.Frame(right_frame)
        tree_scroll.pack(fill="both", expand=True)
        
        self.tree = ttk.Treeview(tree_scroll, show="headings")
        vsb = ttk.Scrollbar(tree_scroll, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_scroll, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)
        
        self.setup_tree_columns()

    # ================= 3. 核心功能区 =================

    # --- 仪表配置窗口 (从第一版移回) ---
    def open_settings_window(self):
        win = tk.Toplevel(self.root)
        win.title("仪表参数配置")
        win.geometry("700x500") 
        
        FONT_UI = ("微软雅黑", 12)
        
        # 左侧列表
        list_frame = tk.Frame(win, padx=10, pady=10)
        list_frame.pack(side="left", fill="y")
        tk.Label(list_frame, text="仪表列表", font=("微软雅黑", 12, "bold")).pack()
        lb = tk.Listbox(list_frame, font=("Arial", 12), width=25, height=20)
        lb.pack(fill="y", expand=True, pady=5)
        
        # 右侧编辑
        edit_frame = tk.Frame(win, padx=20, pady=20)
        edit_frame.pack(side="left", fill="both", expand=True)
        
        name_var = tk.StringVar()
        addr_var = tk.StringVar()
        color_var = tk.StringVar(value="#ff0000")
        
        tk.Label(edit_frame, text="名称:", font=FONT_UI).grid(row=0, column=0, pady=10, sticky="e")
        tk.Entry(edit_frame, textvariable=name_var, font=FONT_UI).grid(row=0, column=1, sticky="w")
        
        tk.Label(edit_frame, text="地址(Addr):", font=FONT_UI).grid(row=1, column=0, pady=10, sticky="e")
        tk.Entry(edit_frame, textvariable=addr_var, font=FONT_UI).grid(row=1, column=1, sticky="w")
        
        tk.Label(edit_frame, text="曲线颜色:", font=FONT_UI).grid(row=2, column=0, pady=10, sticky="e")
        color_btn = tk.Button(edit_frame, text="■ 选择颜色", bg=color_var.get(), command=lambda: self.pick_color(color_var, color_btn))
        color_btn.grid(row=2, column=1, sticky="w")

        def refresh_list():
            lb.delete(0, tk.END)
            for inst in self.instruments:
                lb.insert(tk.END, f"[{inst['addr']}] {inst['name']}")
        
        def on_select(evt):
            if not lb.curselection(): return
            idx = lb.curselection()[0]
            data = self.instruments[idx]
            name_var.set(data['name'])
            addr_var.set(str(data['addr']))
            color_var.set(data['color'])
            color_btn.config(bg=data['color'])
        
        lb.bind('<<ListboxSelect>>', on_select)
        
        def save_change(action):
            try:
                addr = int(addr_var.get().strip())
                new_data = {"name": name_var.get(), "addr": addr, "color": color_var.get()}
                
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
                
                # 保存并刷新所有界面
                self.save_config()
                refresh_list()
                self.setup_tree_columns() # 刷新表格列
                self.force_redraw_plot()  # 刷新绘图线条
                messagebox.showinfo("成功", "设置已更新")
            except Exception as e:
                messagebox.showerror("错误", str(e))

        refresh_list()
        
        btn_frame = tk.Frame(edit_frame, pady=20)
        btn_frame.grid(row=3, column=0, columnspan=2)
        tk.Button(btn_frame, text="新增", bg="#81C784", width=8, command=lambda: save_change("add")).pack(side="left", padx=5)
        tk.Button(btn_frame, text="保存修改", bg="#64B5F6", width=8, command=lambda: save_change("update")).pack(side="left", padx=5)
        tk.Button(btn_frame, text="删除", bg="#E57373", width=8, command=lambda: save_change("delete")).pack(side="left", padx=5)

    def pick_color(self, var, btn):
        color = colorchooser.askcolor(title="选择颜色")[1]
        if color: 
            var.set(color)
            btn.config(bg=color)

    # --- 完整数据导出功能 (从第一版移回) ---
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
            
            # 数据透视：行是时间，列是各仪表
            pivot_df = df.pivot_table(index=['date_str', 'time_str'], columns='address', values='temperature', aggfunc='first')
            
            # 重命名列头
            new_cols = []
            name_map = {i['addr']: i['name'] for i in self.instruments}
            for addr in pivot_df.columns:
                new_cols.append(name_map.get(addr, f"地址_{addr}"))
            pivot_df.columns = new_cols
            pivot_df.reset_index(inplace=True)
            
            default_name = f"温度数据_{start_dt.strftime('%Y%m%d_%H%M')}.csv"
            fname = filedialog.asksaveasfilename(initialfile=default_name, filetypes=[("CSV文件", "*.csv")])
            
            if fname:
                pivot_df.to_csv(fname, index=False, encoding="utf-8-sig")
                messagebox.showinfo("成功", f"成功导出 {len(pivot_df)} 条记录")
                
        except Exception as e:
            messagebox.showerror("导出错误", f"详细信息: {str(e)}")

    # --- 数据采集循环 (后台稳定版) ---
    def data_loop(self):
        logging.info("数据线程启动")
        while self.is_running:
            start_ts = time.time()
            try:
                # 仅当前台显示时更新界面心跳变量
                if self.is_visible:
                    self.heartbeat_var.set("●" if self.heartbeat_var.get() == "○" else "○")
                
                port = self.selected_port.get()
                if port and self.open_serial(port):
                    now = datetime.now()
                    has_data = False
                    for inst in self.instruments:
                        addr = inst['addr']
                        try:
                            temp = self.read_temp(addr)
                            if temp > -99.0:
                                self.cursor.execute("INSERT INTO records VALUES (?, ?, ?, ?, ?)", 
                                    (now.timestamp(), now.strftime('%Y-%m-%d'), now.strftime('%H:%M:%S'), addr, temp))
                                has_data = True
                        except: pass
                    
                    if has_data:
                        try: self.conn.commit()
                        except: pass
                else:
                    time.sleep(2) # 串口未连接时，休眠久一点
            except Exception as e:
                logging.error(f"Loop Error: {e}")
                time.sleep(1)

            elapsed = time.time() - start_ts
            if elapsed < 1.0:
                time.sleep(1.0 - elapsed)

    # --- UI 刷新循环 (智能省电版) ---
    def update_ui_loop(self):
        if not self.is_running: return
        
        # 只有在窗口显示时，才更新绘图和表格，节省资源
        if self.is_visible:
            self.update_table()
            self.update_plot_fast()
        
        # 必须定期GC，防止长期运行内存缓慢泄露
        if int(time.time()) % 60 == 0:
            gc.collect()

        self.root.after(1000, self.update_ui_loop)

    # --- 高性能绘图 (set_data 更新法) ---
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
                self.ax.relim()
                self.ax.autoscale_view()
                self.ax.set_xlim(-limit_val, 0)
            
            self.canvas.draw()
        except Exception as e:
            logging.error(f"Plot Error: {e}")

    # --- 辅助功能 ---
    def init_plot_lines(self):
        """初始化绘图对象，在启动和修改配置时调用"""
        self.ax.clear()
        self.ax.set_title("实时温度趋势", fontsize=12)
        self.ax.set_xlabel("时间", fontsize=10)
        self.ax.set_ylabel("温度 (°C)", fontsize=10)
        self.ax.grid(True, linestyle='--', alpha=0.5)
        self.lines = {}
        for inst in self.instruments:
            line, = self.ax.plot([], [], color=inst['color'], label=inst['name'], linewidth=1.5)
            self.lines[inst['addr']] = line
        self.ax.legend(loc='upper left', fontsize=9)

    def get_plot_data_optimized(self):
        """获取数据并进行智能降采样(LOD)，防止点数过多卡顿"""
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
            
            # 智能抽稀算法
            total_points = len(rows)
            step = 1
            if total_points > MAX_DISPLAY_POINTS * len(self.instruments):
                step = total_points // (MAX_DISPLAY_POINTS * len(self.instruments))
            
            for i in range(0, total_points, step):
                r = rows[i]
                addr = r[1]
                if addr in data_map:
                    diff = r[0] - now_ts
                    # X轴转换为相对时间(负数)
                    x_val = diff / 60.0 if unit == "分钟" else diff / 3600.0
                    data_map[addr]['x'].append(x_val)
                    data_map[addr]['y'].append(r[2])
            return data_map, unit, val
        except: return {}, "分钟", 60

    def force_redraw_plot(self):
        self.init_plot_lines()
        self.update_plot_fast()

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

    # --- 托盘与退出逻辑 ---
    def create_image(self):
        # 绘制一个简单的图标
        image = Image.new('RGB', (64, 64), "blue")
        ImageDraw.Draw(image).rectangle((16, 16, 48, 48), fill="white")
        return image

    def init_tray_icon(self):
        menu = (item('显示监控窗口', self.show_window), item('彻底退出系统', self.quit_app))
        self.icon = pystray.Icon("name", self.create_image(), "温控系统(运行中)", menu)
        self.icon.run()

    def hide_window(self):
        """进入后台模式"""
        self.root.withdraw()
        self.is_visible = False # 停止绘图
        gc.collect() # 清理内存

    def show_window(self):
        """恢复前台模式"""
        self.root.deiconify()
        self.root.state('zoomed') 
        self.is_visible = True # 恢复绘图
        self.force_redraw_plot() # 立即刷新

    def quit_app(self):
        """彻底退出"""
        if messagebox.askyesno("确认退出", "退出后将停止记录数据，确定吗？"):
            self.is_running = False
            try: self.icon.stop()
            except: pass
            try: self.conn.close()
            except: pass
            try: self.root.destroy()
            except: pass
            logging.info("程序正常退出")
            os._exit(0)

    # --- 基础架构函数 ---
    def create_menu(self):
        menubar = tk.Menu(self.root)
        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(label="仪表参数配置", command=self.open_settings_window)
        menubar.add_cascade(label="配置", menu=config_menu)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于", command=lambda: messagebox.showinfo("关于", "宇电温控系统\n版本: Final\n制作: Dalian Institute of Chemical Physics"))
        menubar.add_cascade(label="帮助", menu=help_menu)
        self.root.config(menu=menubar)

    def setup_tree_columns(self):
        cols = ["time"] + [f"addr_{i['addr']}" for i in self.instruments]
        self.tree["columns"] = cols
        self.tree.heading("time", text="时间"); self.tree.column("time", width=180, anchor="center")
        for inst in self.instruments:
            self.tree.heading(f"addr_{inst['addr']}", text=inst['name']); self.tree.column(f"addr_{inst['addr']}", width=100, anchor="center")

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

    def read_temp(self, addr):
        return self.read_aibus_temp(addr) if self.protocol_type.get() == "AIBUS" else self.read_modbus_temp(addr)

    def read_aibus_temp(self, addr):
        if self.serial_conn is None: return -100.0
        try:
            chk = 0x52 + addr; cmd = bytes([0x80 + addr, 0x80 + addr, 0x52, 0x00, 0x00, 0x00, chk & 0xFF, (chk >> 8) & 0xFF])
            self.serial_conn.flushInput(); self.serial_conn.write(cmd); resp = self.serial_conn.read(10)
            if len(resp) < 10: return -100.0 
            pv = resp[0] + (resp[1] << 8); return (pv - 65536 if pv > 32767 else pv) / 10.0
        except: return -100.0

    def read_modbus_temp(self, addr):
        if self.serial_conn is None: return -100.0
        try:
            def calc_crc(d):
                c = 0xFFFF
                for b in d:
                    c ^= b
                    for _ in range(8): c = (c >> 1) ^ 0xA001 if c & 1 else c >> 1
                return c
            base = bytes([addr, 0x03, 0x00, 0x00, 0x00, 0x04]); crc = calc_crc(base)
            self.serial_conn.flushInput(); self.serial_conn.write(base + bytes([crc & 0xFF, crc >> 8])); resp = self.serial_conn.read(13)
            if len(resp) < 13: return -100.0
            pv = (resp[3] << 8) + resp[4]; return (pv - 65536 if pv > 32767 else pv) / 10.0
        except: return -100.0

    def init_db(self):
        self.conn = sqlite3.connect(DB_FILE, check_same_thread=False)
        self.cursor = self.conn.cursor()
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS records (timestamp REAL, date_str TEXT, time_str TEXT, address INTEGER, temperature REAL)''')
        self.cursor.execute('CREATE INDEX IF NOT EXISTS idx_ts ON records(timestamp)')
        self.conn.commit()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f: return json.load(f)
            except: pass
        return [{"name": "1号仪表", "addr": 1, "color": "#ff0000"}]
    
    def save_config(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(self.instruments, f, indent=2)

    def validate_number(self, val): return val.isdigit() or val == ""
    
    def refresh_ports(self):
        ports = sorted(list(set([p.device for p in serial.tools.list_ports.comports()] + ["COM1","COM2","COM3","COM4"])))
        self.cb_ports['values'] = ports; 
        if ports: self.cb_ports.current(0)
    
    def cleanup_old_data(self):
        try: t = (datetime.now()-timedelta(days=DATA_RETENTION_DAYS)).timestamp(); self.cursor.execute("DELETE FROM records WHERE timestamp < ?", (t,)); self.conn.commit()
        except: pass

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = App(root)
        root.mainloop()
    except Exception as e:
        logging.critical(f"Main Crash: {e}")