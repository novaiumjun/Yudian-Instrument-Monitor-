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

# --- 字体与显示配置 ---
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial']
matplotlib.rcParams['axes.unicode_minus'] = False

# 绘图区字体大小
PLOT_TITLE_SIZE = 32
PLOT_LABEL_SIZE = 24
PLOT_TICK_SIZE = 20

# 全局常量
DB_FILE = "multi_channel_history.db"
CONFIG_FILE = "instruments_config.json"
DATA_RETENTION_DAYS = 7     # 【修改】保留7天数据
MAX_PLOT_POINTS = 1000      

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("AI-708 实验室温控系统 (后台运行版)")
        self.root.state('zoomed') 

        # 拦截关闭事件（点击X时触发）
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

        # 字体与样式设置
        self.root.option_add('*TCombobox*Listbox.font', ("Arial", 24))
        self.root.option_add('*Menu.font', ("微软雅黑", 20)) 
        self.root.option_add('*Dialog.msg.font', ("微软雅黑", 14))

        style = ttk.Style()
        style.theme_use('clam') 
        style.configure('TCombobox', arrowsize=40)
        style.configure("Treeview.Heading", font=("微软雅黑", 20, "bold"), rowheight=70)
        style.configure("Treeview", font=("Arial", 18), rowheight=50)
        style.configure("Big.TRadiobutton", font=("微软雅黑", 24)) 

        # --- 变量初始化 ---
        self.is_running = True
        self.serial_conn = None 
        self.selected_port = tk.StringVar()
        self.protocol_type = tk.StringVar(value="AIBUS") # 默认AIBUS
        
        # 加载仪表配置
        self.instruments = self.load_config() 

        # 导出相关变量
        self.export_mode = tk.StringVar(value="recent") 
        self.recent_hours = tk.StringVar(value="5") # 【修改】默认5小时
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        past_str = (datetime.now() - timedelta(hours=5)).strftime("%Y-%m-%d %H:%M")
        self.start_time_str = tk.StringVar(value=past_str)
        self.end_time_str = tk.StringVar(value=now_str)

        self.plot_duration_val = tk.StringVar(value="60") 
        self.plot_duration_unit = tk.StringVar(value="分钟") 

        # 数据库
        self.init_db()
        self.cleanup_old_data()

        # ================= 菜单栏 =================
        self.create_menu()

        # ================= 界面布局 =================
        # 1. 顶部栏
        top_frame = tk.Frame(root, bg="#e0e0e0", height=100)
        top_frame.pack(fill="x", side="top", pady=5)
        
        tk.Label(top_frame, text="端口:", font=("微软雅黑", 24), bg="#e0e0e0").pack(side="left", padx=10)
        self.cb_ports = ttk.Combobox(top_frame, textvariable=self.selected_port, font=("Arial", 24), width=10)
        self.cb_ports.pack(side="left", padx=5)
        
        # 协议选择
        tk.Label(top_frame, text="协议:", font=("微软雅黑", 24), bg="#e0e0e0").pack(side="left", padx=10)
        self.cb_proto = ttk.Combobox(top_frame, textvariable=self.protocol_type, values=["AIBUS", "MODBUS"], font=("Arial", 24), width=8, state="readonly")
        self.cb_proto.pack(side="left", padx=5)

        tk.Button(top_frame, text="刷新", font=("微软雅黑", 20), command=self.refresh_ports).pack(side="left", padx=10)
        self.lbl_status = tk.Label(top_frame, text="初始化...", font=("微软雅黑", 24, "bold"), fg="red", bg="#e0e0e0")
        self.lbl_status.pack(side="left", padx=20)

        # 2. 主体区域
        self.paned = tk.PanedWindow(root, orient="horizontal", sashrelief="raised", sashwidth=15, bg="#ccc")
        self.paned.pack(fill="both", expand=True)

        # === 左侧: 绘图 + 导出 ===
        left_frame = tk.Frame(self.paned)
        self.paned.add(left_frame, width=1300, stretch="always") 

        # 2.1 绘图区
        graph_frame = tk.Frame(left_frame)
        graph_frame.pack(side="top", fill="both", expand=True)
        self.fig, self.ax = plt.subplots()
        self.fig.subplots_adjust(bottom=0.15, left=0.08, right=0.95, top=0.90) 
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # 2.2 控制与导出区
        ctrl_frame = tk.LabelFrame(left_frame, text="数据导出与设置", font=("微软雅黑", 30, "bold"), bg="white")
        ctrl_frame.pack(side="bottom", fill="x", padx=10, pady=10)
        
        font_label = ("微软雅黑", 24); font_entry = ("Arial", 24)
        
        # 显示设置
        tk.Label(ctrl_frame, text="绘图范围:", font=font_label, bg="white").grid(row=0, column=0, sticky="e", padx=10, pady=15)
        v_cmd = (root.register(self.validate_number), '%P')
        tk.Entry(ctrl_frame, textvariable=self.plot_duration_val, font=font_entry, width=4, validate="key", validatecommand=v_cmd).grid(row=0, column=1, sticky="w")
        ttk.Combobox(ctrl_frame, textvariable=self.plot_duration_unit, values=["分钟", "小时"], font=("微软雅黑", 22), width=4, state="readonly").grid(row=0, column=2, sticky="w", padx=5)
        
        # 导出模式
        ttk.Radiobutton(ctrl_frame, text="模式1: 最近", variable=self.export_mode, value="recent", style="Big.TRadiobutton").grid(row=1, column=0, sticky="w", padx=10)
        tk.Entry(ctrl_frame, textvariable=self.recent_hours, font=font_entry, width=4).grid(row=1, column=1, sticky="w")
        tk.Label(ctrl_frame, text="小时", font=font_label, bg="white").grid(row=1, column=2, sticky="w", padx=5)

        ttk.Radiobutton(ctrl_frame, text="模式2: 范围", variable=self.export_mode, value="range", style="Big.TRadiobutton").grid(row=2, column=0, sticky="w", padx=10)
        range_frame = tk.Frame(ctrl_frame, bg="white")
        range_frame.grid(row=2, column=1, columnspan=3, sticky="w")
        tk.Entry(range_frame, textvariable=self.start_time_str, font=("Arial", 20), width=16).pack(side="left")
        tk.Label(range_frame, text=" 至 ", font=font_label, bg="white").pack(side="left")
        tk.Entry(range_frame, textvariable=self.end_time_str, font=("Arial", 20), width=16).pack(side="left")

        btn_export = tk.Button(ctrl_frame, text="导出数据", font=("微软雅黑", 28, "bold"), bg="#4CAF50", fg="white", command=self.export_data)
        btn_export.grid(row=0, column=5, rowspan=3, padx=60, sticky="nsew", pady=20)

        # === 右侧: 动态表格 ===
        right_frame = tk.Frame(self.paned)
        self.paned.add(right_frame, stretch="always")

        tree_label = tk.Label(right_frame, text="实时多路数据流", font=("微软雅黑", 26, "bold"))
        tree_label.pack(side="top", pady=10)

        tree_container = tk.Frame(right_frame)
        tree_container.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_container, show="headings")
        
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        self.setup_tree_columns()

        # ================= 启动 =================
        self.refresh_ports()
        
        # 启动托盘图标线程
        threading.Thread(target=self.init_tray_icon, daemon=True).start()

        # 启动数据线程
        self.thread = threading.Thread(target=self.data_loop, daemon=True)
        self.thread.start()
        
        self.root.after(1000, self.update_ui)

    # ================= 托盘与后台运行逻辑 =================
    def create_image(self):
        """生成一个简单的图标 (避免依赖外部ico文件)"""
        width = 64
        height = 64
        color1 = "blue"
        color2 = "white"
        image = Image.new('RGB', (width, height), color1)
        dc = ImageDraw.Draw(image)
        dc.rectangle((width // 2, 0, width, height // 2), fill=color2)
        dc.rectangle((0, height // 2, width // 2, height), fill=color2)
        return image

    def init_tray_icon(self):
        """初始化托盘图标"""
        image = self.create_image()
        menu = (item('显示窗口', self.show_window), item('退出系统', self.quit_app))
        self.icon = pystray.Icon("name", image, "实验室温控系统", menu)
        self.icon.run()

    def hide_window(self):
        """点击关闭按钮时，隐藏窗口而不是退出"""
        self.root.withdraw()
        # 气泡提示 (Win7可能不支持，但不影响功能)
        try:
            self.icon.notify("软件仍在后台运行记录数据", "已最小化到托盘")
        except: pass

    def show_window(self):
        """从托盘恢复窗口"""
        self.root.deiconify()
        self.root.state('zoomed') 

    def quit_app(self):
        """真正的退出"""
        self.is_running = False
        if self.icon:
            self.icon.stop()
        self.conn.close()
        self.root.quit()
        import sys
        sys.exit(0)

    # ================= 核心通讯函数 =================
    def open_serial(self, port):
        """根据协议打开串口"""
        try:
            if self.serial_conn is not None:
                if self.serial_conn.port == port and self.serial_conn.is_open:
                    return True 
                else:
                    self.serial_conn.close()
            
            baud = 9600 
            
            self.serial_conn = serial.Serial(
                port=port,
                baudrate=baud, 
                bytesize=8,
                parity=serial.PARITY_NONE,
                stopbits=1,
                timeout=0.2 
            )
            self.lbl_status.config(text=f"串口已开: {port} ({self.protocol_type.get()})", fg="green")
            return True
        except Exception as e:
            self.lbl_status.config(text=f"串口错误: {e}", fg="red")
            self.serial_conn = None
            return False

    def read_temp(self, addr):
        proto = self.protocol_type.get()
        if proto == "AIBUS":
            return self.read_aibus_temp(addr)
        else:
            return self.read_modbus_temp(addr)

    def read_aibus_temp(self, addr):
        if self.serial_conn is None: return -100.0
        try:
            chk = 0x52 + addr 
            chk_low = chk & 0xFF
            chk_high = (chk >> 8) & 0xFF
            header_byte = 0x80 + addr
            cmd = bytes([header_byte, header_byte, 0x52, 0x00, 0x00, 0x00, chk_low, chk_high])
            self.serial_conn.flushInput()
            self.serial_conn.write(cmd)
            resp = self.serial_conn.read(10)
            if len(resp) < 10: return -100.0 
            pv_raw = resp[0] + (resp[1] << 8)
            if pv_raw > 32767: pv_raw -= 65536
            return pv_raw / 10.0
        except: return -100.0

    def read_modbus_temp(self, addr):
        if self.serial_conn is None: return -100.0
        try:
            def calc_crc(data):
                crc = 0xFFFF
                for pos in data:
                    crc ^= pos
                    for i in range(8):
                        if (crc & 1) != 0: crc >>= 1; crc ^= 0xA001
                        else: crc >>= 1
                return crc
            base_cmd = bytes([addr, 0x03, 0x00, 0x00, 0x00, 0x04])
            crc = calc_crc(base_cmd)
            cmd = base_cmd + bytes([crc & 0xFF, (crc >> 8) & 0xFF])
            self.serial_conn.flushInput()
            self.serial_conn.write(cmd)
            resp = self.serial_conn.read(13) 
            if len(resp) < 13: return -100.0
            pv_raw = (resp[3] << 8) + resp[4]
            if pv_raw > 32767: pv_raw -= 65536
            return pv_raw / 10.0
        except: return -100.0

    # ================= 配置相关功能 =================
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except: pass
        return [{"name": "1号仪表", "addr": 1, "color": "#ff0000"}]

    def save_config(self):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.instruments, f, ensure_ascii=False, indent=2)

    def create_menu(self):
        menubar = tk.Menu(self.root)
        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(label="仪表参数设置", command=self.open_settings_window)
        menubar.add_cascade(label="配置", menu=config_menu)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="导出数据", command=lambda: (self.export_mode.set("recent"), self.export_data()))
        menubar.add_cascade(label="文件", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="帮我", command=self.show_help)
        help_menu.add_command(label="关于", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=help_menu)

        self.root.config(menu=menubar)

    def open_settings_window(self):
        win = tk.Toplevel(self.root); win.title("仪表参数配置"); win.geometry("1000x700") 
        FONT_UI = ("微软雅黑", 18); FONT_INPUT = ("Arial", 18)
        list_frame = tk.Frame(win, padx=20, pady=20); list_frame.pack(side="left", fill="y")
        tk.Label(list_frame, text="仪表列表", font=("微软雅黑", 18, "bold")).pack()
        lb = tk.Listbox(list_frame, font=FONT_INPUT, width=25, height=20, selectmode=tk.SINGLE, exportselection=False)
        lb.pack(fill="y", expand=True, pady=10)
        edit_frame = tk.Frame(win, padx=40, pady=40); edit_frame.pack(side="left", fill="both", expand=True)
        name_var = tk.StringVar(); addr_var = tk.StringVar(); color_var = tk.StringVar(value="#ff0000")
        
        tk.Label(edit_frame, text="仪表名称:", font=FONT_UI).grid(row=0, column=0, pady=15, sticky="e")
        tk.Entry(edit_frame, textvariable=name_var, font=FONT_INPUT, width=20).grid(row=0, column=1, sticky="w", padx=10)
        tk.Label(edit_frame, text="通讯地址 (Addr):", font=FONT_UI).grid(row=1, column=0, pady=15, sticky="e")
        tk.Entry(edit_frame, textvariable=addr_var, font=FONT_INPUT, width=10).grid(row=1, column=1, sticky="w", padx=10)
        tk.Label(edit_frame, text="绘图颜色:", font=FONT_UI).grid(row=2, column=0, pady=15, sticky="e")
        color_btn = tk.Button(edit_frame, text="■ 点击选择颜色", font=FONT_UI, bg=color_var.get(), width=15, command=lambda: self.pick_color(color_var, color_btn))
        color_btn.grid(row=2, column=1, sticky="w", padx=10)

        def refresh_list(select_idx=None):
            lb.delete(0, tk.END)
            for inst in self.instruments: lb.insert(tk.END, f"[{inst['addr']}] {inst['name']}")
            if select_idx is not None and select_idx < lb.size(): lb.selection_set(select_idx); lb.activate(select_idx)
        def on_select(evt):
            if not lb.curselection(): return
            idx = lb.curselection()[0]; data = self.instruments[idx]
            name_var.set(data['name']); addr_var.set(str(data['addr'])); color_var.set(data['color']); color_btn.config(bg=data['color'])
        lb.bind('<<ListboxSelect>>', on_select)
        
        def add_inst():
            try:
                addr_val = int(addr_var.get().strip())
                self.instruments.append({"name": name_var.get(), "addr": addr_val, "color": color_var.get()})
                self.save_config(); refresh_list(len(self.instruments)-1); self.setup_tree_columns(); messagebox.showinfo("成功", "已添加")
            except ValueError: messagebox.showerror("错误", "地址错误")
        def update_inst():
            if not lb.curselection(): return
            idx = lb.curselection()[0]
            try:
                self.instruments[idx] = {"name": name_var.get(), "addr": int(addr_var.get().strip()), "color": color_var.get()}
                self.save_config(); refresh_list(idx); self.setup_tree_columns(); messagebox.showinfo("成功", "已保存")
            except ValueError: messagebox.showerror("错误", "地址错误")
        def del_inst():
            if not lb.curselection(): return
            if messagebox.askyesno("确认", "删除?"): del self.instruments[lb.curselection()[0]]; self.save_config(); refresh_list(); self.setup_tree_columns()

        refresh_list()
        btn_frame = tk.Frame(edit_frame, pady=50); btn_frame.grid(row=3, column=0, columnspan=2)
        tk.Button(btn_frame, text="新增", command=add_inst, font=("微软雅黑", 16), bg="#aaf", width=8).pack(side="left", padx=15)
        tk.Button(btn_frame, text="修改保存", command=update_inst, font=("微软雅黑", 16), bg="#afa", width=10).pack(side="left", padx=15)
        tk.Button(btn_frame, text="删除", command=del_inst, font=("微软雅黑", 16), bg="#faa", width=8).pack(side="left", padx=15)

    def pick_color(self, var, btn):
        color = colorchooser.askcolor(title="选择线条颜色")[1]
        if color: var.set(color); btn.config(bg=color)

    def setup_tree_columns(self):
        cols = ["time"] + [f"addr_{i['addr']}" for i in self.instruments]
        self.tree["columns"] = cols
        self.tree.heading("time", text="时间"); self.tree.column("time", width=220, anchor="center")
        for inst in self.instruments:
            col_id = f"addr_{inst['addr']}"
            self.tree.heading(col_id, text=inst['name']); self.tree.column(col_id, width=150, anchor="center")

    # ================= 数据逻辑 =================
    def init_db(self):
        self.conn = sqlite3.connect(DB_FILE, check_same_thread=False)
        self.cursor = self.conn.cursor()
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS records (timestamp REAL, date_str TEXT, time_str TEXT, address INTEGER, temperature REAL)''')
        self.cursor.execute('CREATE INDEX IF NOT EXISTS idx_ts ON records(timestamp)')
        self.conn.commit()

    def data_loop(self):
        while self.is_running:
            start_ts = time.time(); now = datetime.now()
            port = self.selected_port.get()
            
            if port and self.open_serial(port):
                for inst in self.instruments:
                    addr = inst['addr']
                    temp = self.read_temp(addr) 
                    try:
                        self.cursor.execute("INSERT INTO records VALUES (?, ?, ?, ?, ?)", 
                            (now.timestamp(), now.strftime('%Y-%m-%d'), now.strftime('%H:%M:%S'), addr, temp))
                    except: pass
                try: self.conn.commit()
                except: pass
            else:
                time.sleep(1)

            elapsed = time.time() - start_ts
            if elapsed < 1.0: time.sleep(1.0 - elapsed)

    def get_plot_data(self):
        try:
            val = int(self.plot_duration_val.get()); unit = self.plot_duration_unit.get()
            if val <= 0: val = 60
            delta = timedelta(minutes=val) if unit == "分钟" else timedelta(hours=val)
            start_ts = (datetime.now() - delta).timestamp()
            self.cursor.execute("SELECT timestamp, address, temperature FROM records WHERE timestamp > ? ORDER BY timestamp ASC", (start_ts,))
            rows = self.cursor.fetchall()
            if not rows: return {}, unit, val

            data_map = {i['addr']: {'x': [], 'y': [], 'color': i['color'], 'name': i['name']} for i in self.instruments}
            now_ts = time.time(); step = 1
            if len(rows) > MAX_PLOT_POINTS * len(self.instruments): step = len(rows) // (MAX_PLOT_POINTS * len(self.instruments))
            for i in range(0, len(rows), step):
                r = rows[i]; addr = r[1]
                if addr in data_map:
                    diff = r[0] - now_ts
                    x_val = diff / 60.0 if unit == "分钟" else diff / 3600.0
                    data_map[addr]['x'].append(x_val); data_map[addr]['y'].append(r[2])
            return data_map, unit, val
        except: return {}, "分钟", 60

    def update_ui(self):
        try:
            self.cursor.execute("SELECT time_str, address, temperature FROM records ORDER BY timestamp DESC LIMIT 100")
            rows = self.cursor.fetchall(); display_data = {}; ordered_times = []
            for r in rows:
                t_str, addr, temp = r[0], r[1], r[2]
                if t_str not in display_data: display_data[t_str] = {}; ordered_times.append(t_str)
                display_data[t_str][addr] = temp
            self.tree.delete(*self.tree.get_children())
            for t in ordered_times[:20]:
                row_vals = [t]
                for inst in self.instruments:
                    val = display_data[t].get(inst['addr'], "--")
                    if isinstance(val, float): row_vals.append(f"{val:.1f}")
                    else: row_vals.append(val)
                self.tree.insert("", "end", values=row_vals)
        except: pass

        data_map, unit, limit_val = self.get_plot_data()
        self.ax.clear()
        self.ax.set_title(f"多路温度趋势 (最近{limit_val}{unit})", fontsize=PLOT_TITLE_SIZE, pad=15)
        self.ax.set_xlabel(f"时间 ({unit}前)", fontsize=PLOT_LABEL_SIZE)
        self.ax.set_ylabel("温度 (°C)", fontsize=PLOT_LABEL_SIZE)
        self.ax.tick_params(labelsize=PLOT_TICK_SIZE); self.ax.grid(True, linestyle='--', alpha=0.5)
        has_data = False
        for addr, d in data_map.items():
            if d['x']: has_data = True; self.ax.plot(d['x'], d['y'], color=d['color'], label=d['name'], linewidth=1.5)
        if has_data: self.ax.legend(loc='upper left', fontsize=12, ncol=3); self.ax.set_xlim(-limit_val, 0)
        self.canvas.draw()
        if self.is_running: self.root.after(1000, self.update_ui)

    def export_data(self):
        mode = self.export_mode.get()
        start_dt, end_dt = None, None
        try:
            if mode == "recent":
                h = float(self.recent_hours.get()); end_dt = datetime.now(); start_dt = end_dt - timedelta(hours=h)
            else:
                fmt = "%Y-%m-%d %H:%M"; start_dt = datetime.strptime(self.start_time_str.get(), fmt); end_dt = datetime.strptime(self.end_time_str.get(), fmt)
            query = "SELECT date_str, time_str, address, temperature FROM records WHERE timestamp BETWEEN ? AND ? ORDER BY timestamp ASC"
            df = pd.read_sql_query(query, self.conn, params=(start_dt.timestamp(), end_dt.timestamp()))
            if df.empty: messagebox.showwarning("空", "无数据"); return
            df['DateTime'] = df['date_str'] + " " + df['time_str']
            pivot_df = df.pivot_table(index=['date_str', 'time_str'], columns='address', values='temperature', aggfunc='first')
            new_cols = []; name_map = {i['addr']: i['name'] for i in self.instruments}
            for addr in pivot_df.columns: new_cols.append(name_map.get(addr, f"Addr_{addr}"))
            pivot_df.columns = new_cols; pivot_df.reset_index(inplace=True) 
            fname = filedialog.asksaveasfilename(initialfile=f"{start_dt.strftime('%Y%m%d %H%M')}-{end_dt.strftime('%Y%m%d %H%M')} 多路温度.csv", filetypes=[("CSV", "*.csv")])
            if fname: pivot_df.to_csv(fname, index=False, encoding="utf-8-sig"); messagebox.showinfo("成功", "导出成功")
        except Exception as e: messagebox.showerror("错误", str(e))

    def validate_number(self, val): return val.isdigit() or val == ""
    def show_help(self): messagebox.showinfo("说明", "后台运行版\n点击右上角关闭按钮会最小化到托盘\n右键托盘图标可退出系统")
    def show_about(self):
        top = tk.Toplevel(self.root); top.geometry("600x350")
        tk.Label(top, text="宇电多路记录软件", font=("微软雅黑", 32, "bold")).pack(pady=40)
        tk.Label(top, text="中国科学院大连化学物理研究所，马军制作", font=("微软雅黑", 18)).pack()
        tk.Label(top, text="novaium@qq.com", font=("Arial", 14), fg="blue").pack(pady=10)
    def refresh_ports(self):
        ports = sorted(list(set([p.device for p in serial.tools.list_ports.comports()] + ["COM1","COM2","COM3","COM4"])))
        self.cb_ports['values'] = ports; 
        if ports: self.cb_ports.current(0)
    def cleanup_old_data(self):
        try: t = (datetime.now()-timedelta(days=DATA_RETENTION_DAYS)).timestamp(); self.cursor.execute("DELETE FROM records WHERE timestamp < ?", (t,)); self.conn.commit()
        except: pass

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()