#Updated with themocouple selection and live plotting
#Developed and Created by:Richard Manimtim |RE|Eastwood City PH
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, Menu
from ttkthemes import ThemedStyle
import asyncio
import queue
import threading
import pyvisa
import time
import logging
import math
import re
import csv
from datetime import datetime
import os
import sys
import traceback
import cv2
from PIL import Image, ImageTk
from enum import Enum
from collections import deque
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.dates as mdates
from matplotlib.animation import FuncAnimation
import configparser

config = configparser.ConfigParser()
config.read('config.ini')

ver = config.get('DEFAULT', 'version', fallback="v3.12")

# Configure logging
log_directory = config.get('paths', 'log_directory', fallback='logs')
os.makedirs(log_directory, exist_ok=True)
log_file = os.path.join(log_directory, f'temperature_monitor_{datetime.now().strftime("%Y%m%d")}.log')
logging.basicConfig(filename=log_file, level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

class ConnectionState(Enum):
    DISCONNECTED = 0
    CONNECTING = 1
    CONNECTED = 2
    RECONNECTING = 3

class VisaCommunication:
    def __init__(self, resource_name):
        self.resource_name = resource_name
        self.inst = None
        self.state = ConnectionState.DISCONNECTED
        self.lock = asyncio.Lock()
        self.last_heartbeat = 0
        self.heartbeat_interval = config.getint('connection', 'heartbeat_interval', fallback=5)

    async def connect(self):
        async with self.lock:
            if self.state != ConnectionState.DISCONNECTED:
                return
            
            self.state = ConnectionState.CONNECTING
            try:
                rm = pyvisa.ResourceManager()
                self.inst = await asyncio.to_thread(rm.open_resource, self.resource_name)
                await asyncio.to_thread(self.inst.write, "*CLS")
                self.state = ConnectionState.CONNECTED
                self.last_heartbeat = asyncio.get_event_loop().time()
            except Exception as e:
                self.state = ConnectionState.DISCONNECTED
                raise ConnectionError(f"Failed to connect: {str(e)}")

    async def disconnect(self):
        async with self.lock:
            if self.state == ConnectionState.DISCONNECTED:
                return
            
            try:
                if self.inst:
                    await asyncio.to_thread(self.inst.close)
            finally:
                self.inst = None
                self.state = ConnectionState.DISCONNECTED

    async def _perform_operation(self, operation, command, max_retries=3):
        async with self.lock:
            if self.state != ConnectionState.CONNECTED:
                raise ConnectionError("Not connected to the instrument")

            current_time = asyncio.get_event_loop().time()
            if current_time - self.last_heartbeat >= self.heartbeat_interval:
                try:
                    await asyncio.to_thread(self.inst.query, "*OPC?")
                    self.last_heartbeat = current_time
                except Exception as e:
                    self.state = ConnectionState.DISCONNECTED
                    raise ConnectionError(f"Heartbeat failed: {str(e)}")

            for attempt in range(max_retries):
                try:
                    result = await asyncio.to_thread(operation, command)
                    self.last_heartbeat = asyncio.get_event_loop().time()
                    return result
                except Exception as e:
                    if attempt == max_retries - 1:
                        self.state = ConnectionState.DISCONNECTED
                        raise ConnectionError(f"Operation failed after {max_retries} attempts: {str(e)}")
                    await asyncio.sleep(1)

    async def query(self, command, max_retries=3):
        return await self._perform_operation(self.inst.query, command, max_retries)

    async def write(self, command, max_retries=3):
        return await self._perform_operation(self.inst.write, command, max_retries)

class TemperatureMonitorApp(tk.Tk):
    def __init__(self, loop):
        super().__init__()
        self.title(f"Base Plate Temperature Monitoring System ({ver}) Copyright (c) 2024, Reliability Engineering")
        self.geometry("1200x950")
        self.configure(bg='#f0f0f0')
        self.protocol("WM_DELETE_WINDOW", self.on_exit)
        
        self.loop = loop
        self.running = True
        
        self.visa_comm = None
        self.monitoring_task = None
        self.stop_video_event = threading.Event()
        self.frame_queue = queue.Queue(maxsize=1)
        self.monitoring_flag = asyncio.Event()
        self.stop_instrument_event = asyncio.Event()
        self.data_queue = asyncio.Queue()
        
        self.csv_file = None
        self.csv_writer = None
        self.last_save_time = time.time()
        self.save_interval = config.getint('monitoring', 'save_interval', fallback=30)
        self.gui_update_interval = config.getfloat('monitoring', 'gui_update_interval', fallback=0.5)
        self.logging_task = None
        
        default_channels_str = config.get('channels', 'default_temp_channels', fallback='101, 102, 103')
        self.channels = [int(c.strip()) for c in default_channels_str.split(',')]
        self.channel_vars = []
        self.thermocouple_vars = {}
        self.fan_channel_var = tk.IntVar(value=config.getint('channels', 'default_fan_channel', fallback=203))
        
        self.status_var = tk.StringVar(value="Ready")
        
        self.rotating_video = config.get('paths', 'rotating_video', fallback='videos/rotating_fan.mp4')
        self.stopped_video = config.get('paths', 'stopped_video', fallback='videos/stopped_fan.mp4')
        
        if not os.path.exists(self.rotating_video) or not os.path.exists(self.stopped_video):
            messagebox.showwarning("Video Files Missing", "One or both video files are missing. The fan animation may not work correctly.")
        
        self.fan_video_frame = None
        self.video_label = None
        self.video_thread = None
        self.current_video = None
        
        self.is_monitoring = False
        self.set_temperature = None
        self.sleep_interval = None
        
        self.connection_status_var = tk.StringVar(value="Disconnected")
        self.reconnection_attempts = 0
        self.max_reconnection_attempts = config.getint('connection', 'max_reconnection_attempts', fallback=5)
        self.reconnection_timeout = config.getint('connection', 'reconnection_timeout', fallback=30)
        self.communication_error_threshold = config.getint('connection', 'communication_error_threshold', fallback=3)
        self.consecutive_errors = 0
        self.is_reconnecting = False
        
        self.max_errors = 5
        self.error_count = 0
        self.reconnection_task = None
        self.heartbeat_task = None

        self.plot_data = {'time': deque(maxlen=config.getint('monitoring', 'max_plot_points', fallback=100))}
        
        self.create_menu()
        self.create_widgets()
        self.start_asyncio_tasks()
        
        sys.excepthook = self.handle_exception

        self.ani = FuncAnimation(self.fig, self.animate_plot, interval=1000, blit=False, cache_frame_data=False)

    def start_asyncio_tasks(self):
        self.loop.create_task(self.update_video_frame_async())
        self.loop.create_task(self.process_queue_async())

    async def update_video_frame_async(self):
        while self.running:
            self.update_video_frame()
            await asyncio.sleep(0.03)

    async def process_queue_async(self):
        while self.running:
            try:
                message = await self.data_queue.get()
                self.update_gui(message)
            except asyncio.CancelledError:
                break
            except Exception as e:
                logging.error(f"Error in process_queue_async: {e}")

    def create_menu(self):
        menubar = Menu(self)
        self.config(menu=menubar)

        theme_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Theme", menu=theme_menu)
        theme_menu.add_command(label="Light", command=lambda: self.set_theme("radiance"))
        theme_menu.add_command(label="Dark", command=lambda: self.set_theme("clam"))

        about_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="About", menu=about_menu)
        about_menu.add_command(label="About This Program", command=self.show_about)

    def show_about(self):
        about_text = f"""
        Base Plate Temperature Monitoring System ({ver})

        The Base Plate Temperature Monitoring System is a simple python gui thermal management solution engineered for industries that demand accuracy, reliability, and automation. Designed to monitor and regulate base plate temperatures in real time, this system ensures optimal operating conditions and safeguards critical components from thermal stress.

        Key Benefits:
        
        🔍 Real-Time Monitoring
        Instantly track temperature fluctuations with high-precision sensors and a responsive data logger.

        🌡️ Smart Thermal Control
        Automatically activates cooling fans based on user-defined thresholds to maintain ideal temperature levels.

        📊 Comprehensive Data Logging
        Capture and store temperature data for performance analysis, compliance, and traceability.

        🖥️ User-Friendly Interface
        Intuitive display and controls make system operation simple and efficient.

        Copyright (c) 2024, Reliability Engineering
        Developer: Richard Manimtim, RE, Eastwood City, Philippines
        """
        messagebox.showinfo(f"About Base Plate Temperature Monitoring System", about_text)

    def create_widgets(self):
        self.style = ThemedStyle(self)
        
        self.style.configure('TLabel', font=('Helvetica', 10))
        self.style.configure('TButton', font=('Helvetica', 10))
        self.style.configure('TLabelframe.Label', font=('Helvetica', 11, 'bold'))

        # Create a Notebook (tabbed interface)
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create frames for each tab
        monitor_tab = ttk.Frame(notebook)
        instructions_tab = ttk.Frame(notebook)

        notebook.add(monitor_tab, text="Monitoring")
        notebook.add(instructions_tab, text="Instructions")

        # --- Populate Monitoring Tab ---
        paned_window = ttk.PanedWindow(monitor_tab, orient=tk.VERTICAL)
        paned_window.pack(fill=tk.BOTH, expand=True, padx=0, pady=5)

        top_frame = ttk.Frame(paned_window, padding="10 10 10 10")
        paned_window.add(top_frame, weight=0)

        bottom_frame = ttk.Frame(paned_window)
        paned_window.add(bottom_frame, weight=1)

        status_frame = ttk.Frame(top_frame)
        status_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(status_frame, text="Connection Status:").pack(side=tk.LEFT)
        ttk.Label(status_frame, textvariable=self.connection_status_var).pack(side=tk.LEFT, padx=(5, 0))

        input_frame = ttk.LabelFrame(top_frame, text="Input Parameters", padding="10 5 10 5")
        input_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(input_frame, text="Set temperature (0-200°C):").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.entry_set_temp = ttk.Entry(input_frame, width=20)
        self.entry_set_temp.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(input_frame, text="Sleep interval (e.g., 500ms, 10s, 2m):").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.entry_sleep_interval = ttk.Entry(input_frame, width=20)
        self.entry_sleep_interval.grid(row=1, column=1, padx=5, pady=5)

        button_frame = ttk.Frame(top_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)

        self.btn_connect = ttk.Button(button_frame, text="Connect to Data Logger", command=self.connect_to_data_logger, width=25)
        self.btn_connect.pack(side=tk.LEFT, padx=5)

        self.btn_start_monitoring = ttk.Button(button_frame, text="Start Monitoring", command=self.start_monitoring, state=tk.DISABLED, width=20)
        self.btn_start_monitoring.pack(side=tk.LEFT, padx=5)

        self.btn_stop_monitoring = ttk.Button(button_frame, text="Stop Monitoring", command=self.stop_monitoring, state=tk.DISABLED, width=20)
        self.btn_stop_monitoring.pack(side=tk.LEFT, padx=5)

        self.connection_address_label = ttk.Label(top_frame, text="Connection Address: Not Connected")
        self.connection_address_label.pack(pady=5)

        channel_frame = ttk.LabelFrame(top_frame, text="Channel Selection", padding="10 5 10 5")
        channel_frame.pack(fill=tk.X, padx=5, pady=5)

        horizontal_frame = ttk.Frame(channel_frame)
        horizontal_frame.pack(fill=tk.X, expand=True)

        fan_control_frame = ttk.LabelFrame(horizontal_frame, text="Fan Control", padding="5 5 5 5")
        fan_control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10), anchor='n')

        ttk.Label(fan_control_frame, text="Fan Channel:").pack(anchor=tk.W, pady=(0,2))
        
        fan_channels = list(range(config.getint('channels', 'fan_channels_start', fallback=201),
                                  config.getint('channels', 'fan_channels_end', fallback=215) + 1))
        self.fan_channel_combo = ttk.Combobox(
            fan_control_frame,
            textvariable=self.fan_channel_var,
            values=fan_channels,
            state="readonly",
            width=8
        )
        self.fan_channel_combo.pack(anchor=tk.W)
        self.fan_channel_combo.bind("<<ComboboxSelected>>", self.on_fan_channel_selected)

        self.temp_channels_frame = ttk.Frame(horizontal_frame)
        self.temp_channels_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        temp_channels_start = config.getint('channels', 'temp_channels_start', fallback=101)
        temp_channels_end = config.getint('channels', 'temp_channels_end', fallback=120)
        for i, channel in enumerate(range(temp_channels_start, temp_channels_end + 1)):
            var = tk.BooleanVar(value=channel in self.channels)
            cb = ttk.Checkbutton(self.temp_channels_frame, text=f"Ch {channel}", variable=var,
                                 command=lambda ch=channel, v=var: self.update_channels(ch, v.get()))
            cb.grid(row=i // 5, column=i % 5 * 3, padx=2, pady=2, sticky=tk.W)
            
            tc_var = tk.StringVar(value="T")
            tc_combo = ttk.Combobox(self.temp_channels_frame, textvariable=tc_var, values=["T", "K"], width=3)
            tc_combo.grid(row=i // 5, column=i % 5 * 3 + 1, padx=2, pady=2)
            
            self.channel_vars.append((channel, var, cb))
            self.thermocouple_vars[channel] = tc_var

        self.create_video_frame(horizontal_frame)

        ttk.Button(channel_frame, text="Reset to Default", 
                   command=self.reset_to_default_channels).pack(side=tk.BOTTOM, padx=5, pady=5)

        temp_frame = ttk.LabelFrame(top_frame, text="Temperature Readings", padding="10 5 10 5")
        temp_frame.pack(fill=tk.X, padx=5, pady=5)

        self.temp_labels_frame = ttk.Frame(temp_frame)
        self.temp_labels_frame.pack(fill=tk.X, expand=True)

        self.temperature_labels = {}
        self.update_temperature_labels()

        avg_fan_frame = ttk.Frame(top_frame)
        avg_fan_frame.pack(fill=tk.X, padx=5, pady=5)

        self.average_temp_label = ttk.Label(avg_fan_frame, text="Average Temperature: N/A", font=("Helvetica", 18, "bold"))
        self.average_temp_label.pack(side=tk.LEFT, padx=10)

        self.fan_status_label = ttk.Label(avg_fan_frame, text="Fan Status: N/A", font=("Helvetica", 16))
        self.fan_status_label.pack(side=tk.RIGHT, padx=10)
        
        self.fan_indicator = tk.Label(avg_fan_frame, width=2, height=1, bg="gray")
        self.fan_indicator.pack(side=tk.RIGHT, padx=5)

        self.create_plot(bottom_frame)

        # --- Populate Instructions Tab ---
        self.create_instructions_tab(instructions_tab)

        self.status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.current_theme = config.get('display', 'theme', fallback='radiance')
        self.set_theme(self.current_theme, initial_load=True)

    def create_plot(self, parent_frame):
        plot_frame = ttk.LabelFrame(parent_frame, text="Live Temperature Trend", padding="10 5 10 5")
        plot_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.fig = Figure(figsize=(5, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def animate_plot(self, i):
        is_dark = self.current_theme == 'clam'
        bg_color = '#333333' if is_dark else '#f0f0f0'
        text_color = 'white' if is_dark else 'black'
        grid_color = 'gray' if is_dark else 'lightgray'

        self.fig.patch.set_facecolor(bg_color)
        self.ax.set_facecolor(bg_color)

        self.ax.clear()
        time_data = list(self.plot_data['time'])
        
        if 'average' in self.plot_data:
            avg_data = list(self.plot_data['average'])
            if len(time_data) == len(avg_data) and len(time_data) > 0:
                self.ax.plot(time_data, avg_data, label='Average Temp', color='cyan' if is_dark else 'black', linewidth=2, marker='o', markersize=3)

        for channel in self.channels:
            if channel in self.plot_data:
                ch_data = list(self.plot_data[channel])
                if len(time_data) == len(ch_data) and len(time_data) > 0:
                    self.ax.plot(time_data, ch_data, label=f'Ch {channel}', marker='o', markersize=3)

        handles, labels = self.ax.get_legend_handles_labels()
        if handles:
            legend = self.ax.legend(loc='upper left', fontsize='small', facecolor=bg_color)
            for text in legend.get_texts():
                text.set_color(text_color)
        
        self.ax.set_xlabel("Time", color=text_color)
        self.ax.set_ylabel("Temperature (°C)", color=text_color)
        self.ax.grid(True, color=grid_color, linestyle='--', linewidth=0.5)
        
        self.ax.tick_params(axis='x', colors=text_color)
        self.ax.tick_params(axis='y', colors=text_color)
        self.ax.spines['bottom'].set_color(text_color)
        self.ax.spines['top'].set_color(text_color)
        self.ax.spines['left'].set_color(text_color)
        self.ax.spines['right'].set_color(text_color)

        if len(time_data) > 0:
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
            self.fig.autofmt_xdate()

        try:
            self.fig.subplots_adjust(bottom=0.25, top=0.9, left=0.1, right=0.95)
        except Exception:
            pass

    def create_instructions_tab(self, parent_frame):
        text_area = scrolledtext.ScrolledText(parent_frame, wrap=tk.WORD, relief=tk.FLAT, bg='#FFFFFF')
        text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Define styles for a cleaner, more modern look
        base_font = "Segoe UI"
        text_area.tag_configure("h1", font=(base_font, 18, "bold"), foreground="#2c3e50", spacing3=15)
        text_area.tag_configure("h2", font=(base_font, 14, "bold"), foreground="#34495e", spacing1=10, spacing3=5)
        text_area.tag_configure("bold", font=(base_font, 10, "bold"), foreground="#34495e")
        text_area.tag_configure("body", font=(base_font, 10), foreground="#34495e", lmargin1=20, lmargin2=20, spacing3=5)
        text_area.tag_configure("code", font=("Consolas", 9), foreground="#e74c3c", background="#ecf0f1", lmargin1=20, lmargin2=20)
        text_area.tag_configure("list", font=(base_font, 10), lmargin1=20, lmargin2=40, spacing3=5)

        # --- Add content ---
        text_area.insert(tk.END, "User Guide\n", "h1")

        text_area.insert(tk.END, "1. First-Time Setup\n", "h2")
        text_area.insert(tk.END, "To connect the data logger to your PC, you will need a GPIB-USB adapter. Adapters from National Instruments or Keysight are recommended.\n\n", "body")
        text_area.insert(tk.END, "Before running the application for the first time, ensure all software dependencies are installed. Open a terminal or command prompt in the application's directory and run:\n", "body")
        text_area.insert(tk.END, "pip install -r requirements.txt\n\n", "code")

        text_area.insert(tk.END, "2. Configuration\n", "h2")
        text_area.insert(tk.END, "The application uses a ", "body")
        text_area.insert(tk.END, "config.ini", "bold")
        text_area.insert(tk.END, " file to manage settings. You can edit this file to change default behaviors without modifying the source code. Key settings include:\n", "body")
        text_area.insert(tk.END, "  • Default temperature and fan channels\n", "list")
        text_area.insert(tk.END, "  • File paths for logs and videos\n", "list")
        text_area.insert(tk.END, "  • Connection and monitoring parameters\n\n", "list")
        text_area.insert(tk.END, "If ", "body")
        text_area.insert(tk.END, "config.ini", "bold")
        text_area.insert(tk.END, " is deleted, the application will run with default, built-in settings.\n\n", "body")

        text_area.insert(tk.END, "3. Connecting to the Instrument\n", "h2")
        text_area.insert(tk.END, "The application automatically detects and connects to a compatible data logger (Keysight DAQ970A or HP 34970A). The connection status is displayed at the top left. If it fails, check the physical connection and ensure the instrument is powered on.\n\n", "body")

        text_area.insert(tk.END, "4. Running a Test\n", "h2")
        text_area.insert(tk.END, "  • ", "list")
        text_area.insert(tk.END, "Set Temperature:", "bold")
        text_area.insert(tk.END, " The threshold (0-200°C) for activating the fan.\n", "list")
        text_area.insert(tk.END, "  • ", "list")
        text_area.insert(tk.END, "Sleep Interval:", "bold")
        text_area.insert(tk.END, " The frequency of measurements (e.g., 500ms, 10s, 2m). 5 minutes is the max settings.\n", "list")
        text_area.insert(tk.END, "  • ", "list")
        text_area.insert(tk.END, "Channel Selection:", "bold")
        text_area.insert(tk.END, " Select temperature channels (101-120) and the fan control channel (201-215), channel 203 is the default fan control channel. Use the dropdown to select channels from the dropdown.\n\n", "list")
        text_area.insert(tk.END, "Click ", "body")
        text_area.insert(tk.END, "Start Monitoring", "bold")
        text_area.insert(tk.END, " to begin. Live data will be displayed and logged. Click ", "body")
        text_area.insert(tk.END, "Stop Monitoring", "bold")
        text_area.insert(tk.END, " to end the test.\n\n", "body")

        text_area.insert(tk.END, "5. Data and Logs\n", "h2")
        text_area.insert(tk.END, "  • ", "list")
        text_area.insert(tk.END, "CSV Logs:", "bold")
        text_area.insert(tk.END, " All test data is saved to a timestamped CSV file in the main directory.\n", "list")
        text_area.insert(tk.END, "  • ", "list")
        text_area.insert(tk.END, "Application Logs:", "bold")
        text_area.insert(tk.END, " Detailed logs for troubleshooting are saved in the 'logs' folder.\n", "list")

        text_area.config(state=tk.DISABLED)

    def create_video_frame(self, parent_frame):
        self.fan_video_frame = ttk.Frame(parent_frame, width=240, height=260)
        self.fan_video_frame.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)
        self.video_label = ttk.Label(self.fan_video_frame)
        self.video_label.pack(fill=tk.BOTH, expand=True)
        self.play_video(self.stopped_video)

    def play_video(self, video_path):
        if self.video_thread and self.video_thread.is_alive():
            if video_path == self.current_video:
                return
            self.stop_video_event.set()
            self.video_thread.join()
        
        self.stop_video_event.clear()
        self.current_video = video_path
        self.video_thread = threading.Thread(target=self._video_thread, args=(video_path,))
        self.video_thread.start()

    def stop_video_playback(self):
        self.stop_video_event.set()
        if self.video_thread:
            self.video_thread.join()
        self.video_thread = None

    def _video_thread(self, video_path):
        cap = cv2.VideoCapture(video_path)
        while not self.stop_video_event.is_set():
            ret, frame = cap.read()
            if not ret:
                cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
                continue
            frame = cv2.resize(frame, (440, 300))
            cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGBA)
            img = Image.fromarray(cv2image)
            self.frame_queue.put(img)
            time.sleep(0.03)
        cap.release()

    def update_video_frame(self):
        try:
            img = self.frame_queue.get_nowait()
            imgtk = ImageTk.PhotoImage(image=img)
            self.video_label.imgtk = imgtk
            self.video_label.configure(image=imgtk)
        except queue.Empty:
            pass

    def update_status(self, status):
        status_message = f"Program Status - {status}"
        self.status_var.set(status_message)
        self.status_bar.update_idletasks()

    def update_gui(self, message):
        if isinstance(message, dict):
            if 'status' in message:
                self.update_status(message['status'])
            if 'connection_status' in message:
                self.update_connection_status(message['connection_status'])
            if 'connection_address' in message:
                self.connection_address_label.config(text=message['connection_address'])
            if 'temperatures' in message:
                for channel, temp in message['temperatures'].items():
                    if channel in self.temperature_labels:
                        if temp is None:
                            self.temperature_labels[channel].config(text="N/A")
                        else:
                            self.temperature_labels[channel].config(text=f"{temp:.1f}°C")
            if 'average' in message:
                if message['average'] is None:
                    self.average_temp_label.config(text="Average Temperature: N/A")
                else:
                    self.average_temp_label.config(text=f"Average Temperature: {message['average']:.1f}°C")
            if 'fan_status' in message:
                fan_status = message['fan_status']
                self.fan_status_label.config(text=f"Fan Status: {fan_status}")
                if fan_status == "Fan Rotating":
                    self.fan_status_label.config(fg="green")
                    self.fan_indicator.config(bg="green")
                    if self.current_video != self.rotating_video:
                        self.play_video(self.rotating_video)
                else:
                    self.fan_status_label.config(fg="red")
                    self.fan_indicator.config(bg="red")
                    if self.current_video != self.stopped_video:
                        self.play_video(self.stopped_video)
            if 'error' in message:
                messagebox.showerror("Error", message['error'])
            if 'enable_connect_button' in message:
                self.btn_connect.config(state=tk.NORMAL if message['enable_connect_button'] else tk.DISABLED)
            if 'enable_start_button' in message:
                self.btn_start_monitoring.config(state=tk.NORMAL if message['enable_start_button'] else tk.DISABLED)
            if 'enable_stop_button' in message:
                self.btn_stop_monitoring.config(state=tk.NORMAL if message['enable_stop_button'] else tk.DISABLED)

        self.update_idletasks()

    def update_connection_status(self, status):
        self.connection_status_var.set(status)
        self.update_idletasks()

    def connect_to_data_logger(self):
        self.btn_connect.config(state=tk.DISABLED)
        self.update_status("Connecting to Data Logger...")
        self.update_connection_status("Connecting...")
        self.loop.create_task(self._connect_thread())

    async def _connect_thread(self):
        try:
            resource_name = await self.auto_negotiate_instrument()
            if not resource_name:
                raise ConnectionError("No compatible instrument found")
            
            self.visa_comm = VisaCommunication(resource_name)
            await self.visa_comm.connect()
            
            await self.data_queue.put({
                'status': "Connected",
                'connection_status': "Connected",
                'connection_address': f"Connected to: {resource_name}",
                'enable_start_button': True
            })
            logging.info(f"Connected to instrument at {resource_name}")
            
        except Exception as e:
            error_message = f"Connection failed: {str(e)}"
            await self.data_queue.put({
                'status': "Connection Failed",
                'connection_status': "Disconnected",
                'connection_address': "Data Logger: Not Connected",
                'error': error_message,
                'enable_start_button': False
            })
            logging.error(error_message)
        
        await self.data_queue.put({'enable_connect_button': True})

    async def auto_negotiate_instrument(self):
        try:
            gpib_file = config.get('paths', 'gpib_address_file', fallback='gpib_address.txt')
            with open(gpib_file, 'r') as f:
                selected_resource = f.read().strip()
        except FileNotFoundError:
            selected_resource = None

        rm = pyvisa.ResourceManager()
        resources = rm.list_resources()

        if selected_resource and selected_resource in resources:
            if await self.try_connect(rm, selected_resource):
                return selected_resource

        for resource in resources:
            if await self.try_connect(rm, resource):
                gpib_file = config.get('paths', 'gpib_address_file', fallback='gpib_address.txt')
                with open(gpib_file, 'w') as f:
                    f.write(resource)
                return resource

        return None

    async def try_connect(self, resource_manager, resource):
        try:
            inst = await self.loop.run_in_executor(None, resource_manager.open_resource, resource)
            identification = await self.loop.run_in_executor(None, inst.query, "*IDN?")
            if "Keysight Technologies,DAQ970A" in identification or "HEWLETT-PACKARD,34970A" in identification:
                logging.info(f"Connected to instrument: {resource}")
                await self.loop.run_in_executor(None, inst.close)
                return True
            await self.loop.run_in_executor(None, inst.close)
        except pyvisa.Error as e:
            logging.error(f"Error connecting to {resource}: {e}")
        return False

    async def handle_disconnection(self):
        if self.visa_comm.state == ConnectionState.RECONNECTING:
            return

        self.visa_comm.state = ConnectionState.RECONNECTING
        await self.data_queue.put({
            'status': "Connection lost. Attempting to reconnect...",
            'connection_status': "Reconnecting"
        })
        
        reconnected = await self._reconnect_coroutine()

        if reconnected:
            await self.data_queue.put({
                'status': "Reconnected successfully",
                'connection_status': "Connected"
            })
            self.error_count = 0
            if self.is_monitoring:
                self.monitoring_task = self.loop.create_task(self.monitor_temperature())
        else:
            await self.data_queue.put({
                'status': "Failed to reconnect. Please check the connection and restart the application.",
                'connection_status': "Disconnected"
            })
            self.stop_monitoring()

        self.visa_comm.state = ConnectionState.CONNECTED if reconnected else ConnectionState.DISCONNECTED


    async def _reconnect_coroutine(self):
        for attempt in range(self.max_reconnection_attempts):
            try:
                await self.visa_comm.disconnect()
                await self.visa_comm.connect()
                logging.info("Successfully reconnected to the instrument")
                await self.data_queue.put({
                    'status': "Reconnected",
                    'status_bar': "Reconnected to the instrument",
                    'gpib_address': f"Connected to: {self.visa_comm.resource_name}"
                })
                return True
            except Exception as e:
                logging.error(f"Reconnection attempt {attempt + 1} failed: {str(e)}")
                await self.data_queue.put({
                    'status': f"Reconnection failed. Retrying... ({attempt + 1}/{self.max_reconnection_attempts})",
                    'status_bar': f"Reconnection failed. Retrying... ({attempt + 1}/{self.max_reconnection_attempts})"
                })
            
            await asyncio.sleep(5)

        logging.critical("Failed to reconnect after multiple attempts")
        await self.data_queue.put({
            'status': "Reconnection failed. Please check the instrument and restart the application.",
            'status_bar': "Reconnection failed. Please restart the application.",
            'enable_start_button': False,
            'enable_stop_button': False
        })
        return False

    def start_monitoring(self):
        if not self.visa_comm or self.visa_comm.state != ConnectionState.CONNECTED:
            messagebox.showerror("Connection Error", "Please connect to the Data Logger first.")
            return

        try:
            set_temp_input = self.entry_set_temp.get().strip()
            if not set_temp_input:
                raise ValueError("Set temperature cannot be empty.")
            
            self.set_temperature = float(set_temp_input)
            if not 0 <= self.set_temperature <= 200:
                raise ValueError("Temperature must be between 0-200°C")
            
            self.sleep_interval = self.get_sleep_interval_in_seconds(self.entry_sleep_interval.get())
        except ValueError as e:
            messagebox.showerror("Invalid Input", str(e))
            logging.error(f"Invalid input: {e}")
            return

        self.btn_start_monitoring.config(state=tk.DISABLED)
        self.btn_stop_monitoring.config(state=tk.NORMAL)
        self.disable_channel_selection()
        self.update_status("Reading Measurements...")

        max_points = config.getint('monitoring', 'max_plot_points', fallback=100)
        self.plot_data = {'time': deque(maxlen=max_points)}
        for ch in self.channels:
            self.plot_data[ch] = deque(maxlen=max_points)
        self.plot_data['average'] = deque(maxlen=max_points)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_filename = f'temperature_log_{timestamp}.csv'
        self.csv_file = open(csv_filename, 'w', newline='')
        self.csv_writer = csv.writer(self.csv_file)
        header = ['Timestamp', 'Average Temperature'] + [f'Temp (Ch {ch} {self.thermocouple_vars[ch].get()})' for ch in self.channels] + ['Fan Status']
        self.csv_writer.writerow(header)

        self.is_monitoring = True
        self.monitoring_task = self.loop.create_task(self.monitor_temperature())
        self.logging_task = self.loop.create_task(self.log_data())
        logging.info("Monitoring started")
        print("Monitoring started")

    def stop_monitoring(self):
        self.is_monitoring = False
        if self.monitoring_task and not self.monitoring_task.done():
            self.monitoring_task.cancel()
        if self.logging_task and not self.logging_task.done():
            self.logging_task.cancel()
        
        self.btn_connect.config(state=tk.NORMAL)
        self.btn_start_monitoring.config(state=tk.NORMAL)
        self.btn_stop_monitoring.config(state=tk.DISABLED)
        self.enable_channel_selection()
        self.update_status("Monitoring stopped")

        if self.csv_file:
            self.csv_file.close()
            self.csv_file = None
            self.csv_writer = None

        logging.info("Monitoring stopped")
        print("Monitoring stopped")
        messagebox.showinfo("Operation Stopped", "Monitoring has been stopped")

    async def monitor_temperature(self):
        while self.is_monitoring:
            start_time = time.monotonic()
            try:
                temperature_values = {}
                for channel in self.channels:
                    temp = await self.read_temperature(channel)
                    temperature_values[channel] = temp

                average_temperature = self.get_average_temperature(list(temperature_values.values()))

                timestamp = datetime.now()
                
                self.plot_data['time'].append(timestamp)
                if average_temperature is not None:
                    self.plot_data['average'].append(average_temperature)
                else:
                    self.plot_data['average'].append(float('nan'))

                for ch in self.channels:
                    temp = temperature_values.get(ch)
                    if temp is not None:
                        self.plot_data[ch].append(temp)
                    else:
                        self.plot_data[ch].append(float('nan'))

                fan_status = "Fan Rotating" if average_temperature is not None and average_temperature > self.set_temperature else "Fan Stopped"
                
                await self.data_queue.put({
                    'temperatures': temperature_values,
                    'average': average_temperature,
                    'fan_status': fan_status,
                    'status_bar': f"Monitoring: Avg Temp {average_temperature:.1f}°C - {fan_status}" if average_temperature is not None else f"Monitoring: Avg Temp N/A - {fan_status}"
                })
                
                video_path = self.rotating_video if fan_status == "Fan Rotating" else self.stopped_video
                self.play_video(video_path)

                if self.csv_writer:
                    row = [timestamp.strftime("%Y-%m-%d %H:%M:%S"), f"{average_temperature:.1f}" if average_temperature is not None else "N/A"]
                    row.extend([f"{temperature_values.get(ch, 'N/A'):.1f}" if temperature_values.get(ch) is not None else 'N/A' for ch in self.channels])
                    row.append(fan_status)
                    self.csv_writer.writerow(row)
                    
                    current_time = time.time()
                    if current_time - self.last_save_time >= self.save_interval:
                        self.csv_file.flush()
                        os.fsync(self.csv_file.fileno())
                        self.last_save_time = current_time

                if average_temperature is not None:
                    fan_command = "CLOSE" if average_temperature > self.set_temperature else "OPEN"
                    await self.visa_comm.write(f"ROUTE:{fan_command} (@{self.fan_channel_var.get()})")

            except Exception as e:
                logging.error(f"Error in monitoring loop: {e}")
                await self.data_queue.put({
                    'status': f"Error: {e}",
                    'status_bar': f"Error occurred. Attempting to recover..."
                })
                await self.handle_disconnection()

            work_duration = time.monotonic() - start_time
            sleep_duration = self.sleep_interval - work_duration
            if sleep_duration > 0:
                await asyncio.sleep(sleep_duration)

    async def read_temperature(self, channel):
        try:
            tc_type = self.thermocouple_vars[channel].get()
            command = f"MEAS:TEMP? TC,{tc_type},(@{channel})"
            measurement = await self.visa_comm.query(command)
            temperature = float(measurement)
            if not (-200 <= temperature <= 1000):
                raise ValueError(f"Temperature out of range: {temperature}")
            return temperature
        except Exception as e:
            logging.error(f"Error reading temperature from channel {channel}: {e}")
            return None

    async def log_data(self):
        last_log_time = time.time()
        while self.is_monitoring:
            current_time = time.time()
            if current_time - last_log_time >= self.sleep_interval:
                try:
                    temperature_values = await asyncio.gather(*[self.read_temperature(channel) for channel in self.channels])
                    temperatures = dict(zip(self.channels, temperature_values))
                    average_temperature = self.get_average_temperature(list(temperatures.values()))
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    if self.csv_writer:
                        fan_status = "Fan Rotating" if average_temperature is not None and average_temperature > self.set_temperature else "Fan Stopped"
                        row = [timestamp, f"{average_temperature:.1f}" if average_temperature is not None else "N/A"]
                        row.extend([f"{temperatures.get(ch, 'N/A'):.1f}" if temperatures.get(ch) is not None else 'N/A' for ch in self.channels])
                        row.append(fan_status)
                        self.csv_writer.writerow(row)
                        
                        if current_time - self.last_save_time >= self.save_interval:
                            self.csv_file.flush()
                            os.fsync(self.csv_file.fileno())
                            self.last_save_time = current_time

                    last_log_time = current_time
                except Exception as e:
                    logging.error(f"Error in logging loop: {e}")

            await asyncio.sleep(0.1)

    def get_average_temperature(self, temperatures):
        valid_temperatures = [temp for temp in temperatures if temp is not None and isinstance(temp, (int, float)) and not math.isnan(temp)]
        return sum(valid_temperatures) / len(valid_temperatures) if valid_temperatures else None

    def get_sleep_interval_in_seconds(self, sleep_interval_str):
        sleep_interval_str = sleep_interval_str.lower().strip()
        if not sleep_interval_str:
            raise ValueError("Sleep interval cannot be empty.")

        match = re.match(r'^\s*(\d+)\s*(ms|s|m)?\s*$', sleep_interval_str)
        if not match:
            raise ValueError("Invalid format. Use a number followed by 'ms', 's', or 'm' (e.g., '500ms', '10s', '1m').")

        value = int(match.group(1))
        unit = match.group(2)

        if unit == 'ms':
            if not 200 <= value <= 999:
                raise ValueError("Millisecond value must be between 200 and 999.")
            return value / 1000.0
        elif unit == 's':
            if not 1 <= value <= 59:
                raise ValueError("Seconds value must be between 1 and 59.")
            return float(value)
        elif unit == 'm':
            if not 1 <= value <= 5:
                raise ValueError("Minute value must be between 1 and 5.")
            return float(value * 60)
        elif unit is None: # Default to seconds if no unit is provided
            if not 1 <= value <= 59:
                 raise ValueError("Default unit is seconds. Value must be between 1 and 59.")
            return float(value)
        
        raise ValueError("Invalid time unit. Use 'ms', 's', or 'm'.")

    def update_channels(self, channel, state):
        if state and channel not in self.channels:
            self.channels.append(channel)
            max_points = config.getint('monitoring', 'max_plot_points', fallback=100)
            self.plot_data[channel] = deque(maxlen=max_points)
        elif not state and channel in self.channels:
            self.channels.remove(channel)
            if channel in self.plot_data:
                del self.plot_data[channel]
        self.channels.sort()
        self.update_temperature_labels()
        tc_types = {ch: self.thermocouple_vars[ch].get() for ch in self.channels}
        self.update_status(f"Temperature channels updated: {', '.join([f'{ch}({tc_types[ch]})' for ch in self.channels])}")

    def update_fan_channel(self, channel):
        self.fan_channel_var.set(channel)
        self.update_status(f"Fan control channel updated to: {channel}")

    def on_fan_channel_selected(self, event):
        selected_channel = self.fan_channel_var.get()
        self.update_status(f"Fan control channel updated to: {selected_channel}")

    def reset_to_default_channels(self):
        default_channels_str = config.get('channels', 'default_temp_channels', fallback='101, 102, 103')
        self.channels = [int(c.strip()) for c in default_channels_str.split(',')]
        for channel, var, _ in self.channel_vars:
            var.set(channel in self.channels)
            self.thermocouple_vars[channel].set("T")
        
        default_fan_channel = config.getint('channels', 'default_fan_channel', fallback=203)
        self.update_fan_channel(default_fan_channel)
        self.fan_channel_var.set(default_fan_channel)

        self.update_temperature_labels()
        self.update_status(f"Channels reset to default: {', '.join(map(str, self.channels))}(T); Fan control: {default_fan_channel}")
        
        max_points = config.getint('monitoring', 'max_plot_points', fallback=100)
        self.plot_data = {'time': deque(maxlen=max_points)}
        for ch in self.channels:
            self.plot_data[ch] = deque(maxlen=max_points)
        self.plot_data['average'] = deque(maxlen=max_points)

    def update_temperature_labels(self):
        for widget in self.temp_labels_frame.winfo_children():
            widget.destroy()
        self.temperature_labels.clear()

        for i, channel in enumerate(self.channels):
            label_frame = ttk.Frame(self.temp_labels_frame)
            label_frame.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, pady=5)

            tc_type = self.thermocouple_vars[channel].get()
            label_text = f"Temp{i+1} Ch ({channel:03}) {tc_type}"
            ttk.Label(label_frame, text=label_text).pack()
            value_label = ttk.Label(label_frame, text="N/A", font=("Helvetica", 14, "bold"))
            value_label.pack()
            self.temperature_labels[channel] = value_label

    def disable_channel_selection(self):
        for _, _, cb in self.channel_vars:
            cb.config(state=tk.DISABLED)
        for widget in self.temp_channels_frame.winfo_children():
            if isinstance(widget, ttk.Combobox):
                widget.config(state='disabled')
        self.fan_channel_combo.config(state=tk.DISABLED)

    def enable_channel_selection(self):
        for _, _, cb in self.channel_vars:
            cb.config(state=tk.NORMAL)
        for widget in self.temp_channels_frame.winfo_children():
            if isinstance(widget, ttk.Combobox):
                widget.config(state='readonly')
        self.fan_channel_combo.config(state='readonly')

    def set_theme(self, theme_name, initial_load=False):
        self.style.set_theme(theme_name)
        self.current_theme = theme_name
        
        is_dark = theme_name == 'clam'
        bg_color = '#333333' if is_dark else '#f0f0f0'
        fg_color = 'white' if is_dark else 'black'
        
        self.configure(bg=bg_color)
        self.status_bar.configure(background=bg_color, foreground=fg_color)
        self.average_temp_label.configure(background=bg_color, foreground=fg_color)
        self.connection_address_label.configure(background=bg_color, foreground=fg_color)

        if not initial_load:
            # Update config file
            if not config.has_section('display'):
                config.add_section('display')
            config.set('display', 'theme', theme_name)
            with open('config.ini', 'w') as configfile:
                config.write(configfile)

    def handle_exception(self, exc_type, exc_value, exc_traceback):
        error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        logging.error("Uncaught exception: %s", error_msg)
        messagebox.showerror("Error", f"An unexpected error occurred:\n{str(exc_value)}\n\nPlease check the log file for details.")
        self.stop_monitoring()

    def on_exit(self):
        if messagebox.askyesno("Confirm Exit", "Are you sure you want to exit the application?"):
            self.running = False
            self.loop.create_task(self.shutdown())

    async def shutdown(self):
        logging.info("Application closing...")
        
        self.is_monitoring = False
        if self.monitoring_task and not self.monitoring_task.done():
            self.monitoring_task.cancel()
            try:
                await self.monitoring_task
            except asyncio.CancelledError:
                pass


        self.stop_video_playback()
        
        if self.visa_comm:
            try:
                await self.visa_comm.disconnect()
            except Exception as e:
                logging.error(f"Error closing instrument connection: {e}")

        tasks = [t for t in asyncio.all_tasks(self.loop) if t is not asyncio.current_task()]
        for task in tasks:
            task.cancel()
        await asyncio.gather(*tasks, return_exceptions=True)

        self.destroy()
        self.loop.stop()

def run_app():
    loop = asyncio.get_event_loop_policy().get_event_loop()
    app = TemperatureMonitorApp(loop)
    
    async def run_tk(app, interval=1/120):
        while app.running:
            app.update()
            await asyncio.sleep(interval)

    loop.create_task(run_tk(app))
    
    try:
        loop.run_forever()
    except Exception as e:
        logging.critical(f"Critical error: {e}")
        messagebox.showerror("Critical Error", f"A critical error occurred: {e}\nThe application will now close.")
    finally:
        if loop.is_running():
            loop.run_until_complete(loop.shutdown_asyncgens())
            loop.close()

if __name__ == "__main__":
    run_app()
