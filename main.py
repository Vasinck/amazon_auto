import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import threading
import os
import sys
import time

# Import the original scraping functionality
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import pandas as pd
import re

# Import default configuration
import config

class AmazonScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("亚马逊数据爬取工具")
        self.root.geometry("900x700")
        
        # Create variables for all configuration parameters
        self.setup_variables()
        
        # Create notebook for tabbed interface
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_browser_tab()
        self.create_amazon_tab()
        self.create_data_tab()
        self.create_execution_tab()
        
        # Create the log area
        self.create_log_area()
        
        # Create control buttons
        self.create_control_buttons()
        
        # Initialize the driver variable
        self.driver = None
        self.running = False
        
        # Redirect stdout to the log text area
        self.stdout_original = sys.stdout
        sys.stdout = self
        
    def setup_variables(self):
        """Setup all tkinter variables for configuration"""
        # Browser configuration
        self.chrome_driver_path = tk.StringVar(value=config.CHROME_DRIVER_PATH)
        self.extension_path = tk.StringVar(value=config.EXTENSION_PATH)
        
        # Browser options
        self.start_maximized = tk.BooleanVar(value=config.BROWSER_OPTIONS.get("start_maximized", True))
        self.no_sandbox = tk.BooleanVar(value=config.BROWSER_OPTIONS.get("no_sandbox", True))
        self.disable_dev_shm_usage = tk.BooleanVar(value=config.BROWSER_OPTIONS.get("disable_dev_shm_usage", True))
        self.disable_extensions_file_access_check = tk.BooleanVar(value=config.BROWSER_OPTIONS.get("disable_extensions_file_access_check", True))
        
        # Timeout settings
        self.page_load_timeout = tk.IntVar(value=config.PAGE_LOAD_TIMEOUT)
        self.element_wait_timeout = tk.IntVar(value=config.ELEMENT_WAIT_TIMEOUT)
        self.quick_wait_timeout = tk.IntVar(value=config.QUICK_WAIT_TIMEOUT)
        
        # Amazon site configuration
        self.amazon_site = tk.StringVar(value=config.AMAZON_SITE)
        self.delivery_zipcode = tk.StringVar(value=config.DELIVERY_ZIPCODE)
        
        # Blocked resources
        self.blocked_resources = tk.StringVar(value=", ".join(config.BLOCKED_RESOURCES))
        
        # File paths
        self.excel_path = tk.StringVar(value=config.EXCEL_PATH)
        self.screenshots_dir = tk.StringVar(value=config.SCREENSHOTS_DIR)
        
        # Data scraping configuration
        self.max_products = tk.IntVar(value=config.MAX_PRODUCTS)
        self.plugin_data_wait_time = tk.IntVar(value=config.PLUGIN_DATA_WAIT_TIME)
        self.plugin_initial_wait = tk.IntVar(value=config.PLUGIN_INITIAL_WAIT)
        self.search_result_initial_wait = tk.IntVar(value=config.SEARCH_RESULT_INITIAL_WAIT)
        self.plugin_data_processing_wait = tk.IntVar(value=config.PLUGIN_DATA_PROCESSING_WAIT)
        self.product_search_interval = tk.IntVar(value=config.PRODUCT_SEARCH_INTERVAL)
        self.min_product_search_interval = tk.IntVar(value=config.MIN_PRODUCT_SEARCH_INTERVAL)
        self.default_browser_close_wait = tk.IntVar(value=config.DEFAULT_BROWSER_CLOSE_WAIT)
        
        # 新增加的等待时间变量
        self.amazon_homepage_wait = tk.IntVar(value=getattr(config, "AMAZON_HOMEPAGE_WAIT", 10))  # 默认10秒
        self.delivery_location_wait = tk.IntVar(value=getattr(config, "DELIVERY_LOCATION_WAIT", 10))  # 默认10秒
    
    def create_browser_tab(self):
        """Create the browser configuration tab"""
        browser_frame = ttk.Frame(self.notebook)
        self.notebook.add(browser_frame, text="浏览器配置")
        
        # Add fields with labels and browsing buttons
        row = 0
        
        # Chrome driver path
        ttk.Label(browser_frame, text="Chrome驱动路径:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(browser_frame, textvariable=self.chrome_driver_path, width=50).grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(browser_frame, text="浏览...", command=lambda: self.browse_file(self.chrome_driver_path)).grid(row=row, column=2, padx=5, pady=5)
        row += 1
        
        # Extension path
        ttk.Label(browser_frame, text="插件路径:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(browser_frame, textvariable=self.extension_path, width=50).grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(browser_frame, text="浏览...", command=lambda: self.browse_directory(self.extension_path)).grid(row=row, column=2, padx=5, pady=5)
        row += 1
        
        # Browser options frame
        options_frame = ttk.LabelFrame(browser_frame, text="浏览器选项")
        options_frame.grid(row=row, column=0, columnspan=3, sticky='we', padx=5, pady=5)
        
        ttk.Checkbutton(options_frame, text="最大化窗口", variable=self.start_maximized).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        ttk.Checkbutton(options_frame, text="禁用沙箱", variable=self.no_sandbox).grid(row=0, column=1, sticky='w', padx=5, pady=5)
        ttk.Checkbutton(options_frame, text="禁用共享内存", variable=self.disable_dev_shm_usage).grid(row=1, column=0, sticky='w', padx=5, pady=5)
        ttk.Checkbutton(options_frame, text="禁用扩展文件访问检查", variable=self.disable_extensions_file_access_check).grid(row=1, column=1, sticky='w', padx=5, pady=5)
        row += 1
        
        # Timeout settings frame
        timeout_frame = ttk.LabelFrame(browser_frame, text="超时设置（秒）")
        timeout_frame.grid(row=row, column=0, columnspan=3, sticky='we', padx=5, pady=5)
        
        ttk.Label(timeout_frame, text="页面加载超时:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        ttk.Spinbox(timeout_frame, from_=1, to=60, textvariable=self.page_load_timeout, width=5).grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(timeout_frame, text="元素等待超时:").grid(row=0, column=2, sticky='w', padx=5, pady=5)
        ttk.Spinbox(timeout_frame, from_=1, to=60, textvariable=self.element_wait_timeout, width=5).grid(row=0, column=3, sticky='w', padx=5, pady=5)
        
        ttk.Label(timeout_frame, text="快速等待超时:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        ttk.Spinbox(timeout_frame, from_=1, to=10, textvariable=self.quick_wait_timeout, width=5).grid(row=1, column=1, sticky='w', padx=5, pady=5)
        row += 1
        
        # Blocked resources
        ttk.Label(browser_frame, text="屏蔽的资源（逗号分隔）:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(browser_frame, textvariable=self.blocked_resources, width=50).grid(row=row, column=1, columnspan=2, sticky='we', padx=5, pady=5)
    
    def create_amazon_tab(self):
        """Create the Amazon configuration tab"""
        amazon_frame = ttk.Frame(self.notebook)
        self.notebook.add(amazon_frame, text="亚马逊设置")
        
        row = 0
        
        # Amazon site
        ttk.Label(amazon_frame, text="亚马逊站点:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        ttk.Combobox(amazon_frame, textvariable=self.amazon_site, values=["amazon.com", "amazon.co.uk", "amazon.de", "amazon.fr", "amazon.it", "amazon.es", "amazon.co.jp", "amazon.ca"]).grid(row=row, column=1, sticky='w', padx=5, pady=5)
        row += 1
        
        # Delivery zipcode
        ttk.Label(amazon_frame, text="配送地址邮编:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(amazon_frame, textvariable=self.delivery_zipcode, width=15).grid(row=row, column=1, sticky='w', padx=5, pady=5)
    
    def create_data_tab(self):
        """Create the data configuration tab"""
        data_frame = ttk.Frame(self.notebook)
        self.notebook.add(data_frame, text="数据设置")
        
        row = 0
        
        # Excel path
        ttk.Label(data_frame, text="Excel文件路径:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(data_frame, textvariable=self.excel_path, width=50).grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(data_frame, text="浏览...", command=lambda: self.browse_file(self.excel_path, [("Excel文件", "*.xlsx *.xls")])).grid(row=row, column=2, padx=5, pady=5)
        row += 1
        
        # Screenshots directory
        ttk.Label(data_frame, text="截图保存目录:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        ttk.Entry(data_frame, textvariable=self.screenshots_dir, width=50).grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(data_frame, text="浏览...", command=lambda: self.browse_directory(self.screenshots_dir)).grid(row=row, column=2, padx=5, pady=5)
        row += 1
        
        # Max products
        ttk.Label(data_frame, text="最大爬取产品数:").grid(row=row, column=0, sticky='w', padx=5, pady=5)
        ttk.Spinbox(data_frame, from_=1, to=1000, textvariable=self.max_products, width=10).grid(row=row, column=1, sticky='w', padx=5, pady=5)
        row += 1
    
    def create_execution_tab(self):
        """Create the execution configuration tab"""
        execution_frame = ttk.Frame(self.notebook)
        self.notebook.add(execution_frame, text="执行设置")
        
        row = 0
        
        # Wait time settings frame
        wait_frame = ttk.LabelFrame(execution_frame, text="等待时间设置（秒）")
        wait_frame.grid(row=row, column=0, columnspan=3, sticky='we', padx=5, pady=5)
        
        # 新增加的等待时间设置
        ttk.Label(wait_frame, text="亚马逊首页加载后等待时间:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        ttk.Spinbox(wait_frame, from_=1, to=60, textvariable=self.amazon_homepage_wait, width=5).grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(wait_frame, text="设置邮编后等待时间:").grid(row=0, column=2, sticky='w', padx=5, pady=5)
        ttk.Spinbox(wait_frame, from_=1, to=60, textvariable=self.delivery_location_wait, width=5).grid(row=0, column=3, sticky='w', padx=5, pady=5)
        
        # 原有的等待时间设置，行号调整为从1开始
        ttk.Label(wait_frame, text="插件数据等待时间:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        ttk.Spinbox(wait_frame, from_=1, to=60, textvariable=self.plugin_data_wait_time, width=5).grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(wait_frame, text="插件初始化等待时间:").grid(row=1, column=2, sticky='w', padx=5, pady=5)
        ttk.Spinbox(wait_frame, from_=1, to=30, textvariable=self.plugin_initial_wait, width=5).grid(row=1, column=3, sticky='w', padx=5, pady=5)
        
        ttk.Label(wait_frame, text="搜索结果初始等待时间:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        ttk.Spinbox(wait_frame, from_=1, to=30, textvariable=self.search_result_initial_wait, width=5).grid(row=2, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(wait_frame, text="插件数据处理等待时间:").grid(row=2, column=2, sticky='w', padx=5, pady=5)
        ttk.Spinbox(wait_frame, from_=1, to=30, textvariable=self.plugin_data_processing_wait, width=5).grid(row=2, column=3, sticky='w', padx=5, pady=5)
        
        ttk.Label(wait_frame, text="产品搜索间隔时间:").grid(row=3, column=0, sticky='w', padx=5, pady=5)
        ttk.Spinbox(wait_frame, from_=1, to=30, textvariable=self.product_search_interval, width=5).grid(row=3, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(wait_frame, text="最小搜索间隔时间:").grid(row=3, column=2, sticky='w', padx=5, pady=5)
        ttk.Spinbox(wait_frame, from_=1, to=30, textvariable=self.min_product_search_interval, width=5).grid(row=3, column=3, sticky='w', padx=5, pady=5)
        
        ttk.Label(wait_frame, text="默认浏览器关闭等待时间:").grid(row=4, column=0, sticky='w', padx=5, pady=5)
        ttk.Spinbox(wait_frame, from_=5, to=300, textvariable=self.default_browser_close_wait, width=5).grid(row=4, column=1, sticky='w', padx=5, pady=5)
    
    def create_log_area(self):
        """Create the log output area"""
        log_frame = ttk.LabelFrame(self.root, text="日志输出")
        log_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create a scrolled text widget for log output
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, width=80, height=20)
        self.log_text.pack(fill='both', expand=True, padx=5, pady=5)
        self.log_text.config(state='disabled')
        
        # Add a progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(log_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill='x', padx=5, pady=5)
        
        # Progress label
        self.progress_label = ttk.Label(log_frame, text="准备就绪")
        self.progress_label.pack(anchor='w', padx=5)
    
    def create_control_buttons(self):
        """Create control buttons"""
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill='x', padx=10, pady=10)
        
        # Start button
        self.start_button = ttk.Button(button_frame, text="开始爬取", command=self.start_scraping)
        self.start_button.pack(side='left', padx=5)
        
        # Stop button
        self.stop_button = ttk.Button(button_frame, text="停止", command=self.stop_scraping, state='disabled')
        self.stop_button.pack(side='left', padx=5)
        
        # Clear log button
        self.clear_button = ttk.Button(button_frame, text="清除日志", command=self.clear_log)
        self.clear_button.pack(side='left', padx=5)
        
        # Save config button
        self.save_config_button = ttk.Button(button_frame, text="保存配置", command=self.save_config)
        self.save_config_button.pack(side='right', padx=5)
        
        # Load config button
        self.load_config_button = ttk.Button(button_frame, text="加载配置", command=self.load_config)
        self.load_config_button.pack(side='right', padx=5)
    
    def browse_file(self, string_var, file_types=None):
        """Open file browser and update the string variable"""
        if file_types is None:
            file_types = [("所有文件", "*.*")]
        
        filename = filedialog.askopenfilename(filetypes=file_types)
        if filename:
            string_var.set(filename)
    
    def browse_directory(self, string_var):
        """Open directory browser and update the string variable"""
        directory = filedialog.askdirectory()
        if directory:
            string_var.set(directory)
    
    def write(self, text):
        """Write text to the log area (used for redirecting stdout)"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        
        # Force update the UI
        self.root.update_idletasks()
    
    def flush(self):
        """Required for stdout redirection"""
        pass
    
    def clear_log(self):
        """Clear the log area"""
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
    
    def update_progress(self, current, total, status=None):
        """Update the progress bar and label"""
        progress = (current / total) * 100
        self.progress_var.set(progress)
        
        if status:
            self.progress_label.config(text=status)
        else:
            self.progress_label.config(text=f"进度: {current}/{total} ({progress:.1f}%)")
        
        # Force update the UI
        self.root.update_idletasks()
    
    def save_config(self):
        """Save current configuration to a file"""
        filename = filedialog.asksaveasfilename(defaultextension=".py", filetypes=[("Python文件", "*.py")])
        if not filename:
            return
        
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write("# 亚马逊数据爬取脚本的配置文件\n\n")
                
                # 浏览器配置
                f.write("# 浏览器配置\n")
                f.write(f"CHROME_DRIVER_PATH = r\"{self.chrome_driver_path.get()}\"\n")
                f.write(f"EXTENSION_PATH = r\"{self.extension_path.get()}\"\n\n")
                
                # 浏览器选项
                f.write("# 浏览器选项\n")
                f.write("BROWSER_OPTIONS = {\n")
                f.write(f"    \"start_maximized\": {self.start_maximized.get()},\n")
                f.write(f"    \"no_sandbox\": {self.no_sandbox.get()},\n")
                f.write(f"    \"disable_dev_shm_usage\": {self.disable_dev_shm_usage.get()},\n")
                f.write(f"    \"disable_extensions_file_access_check\": {self.disable_extensions_file_access_check.get()}\n")
                f.write("}\n\n")
                
                # 页面加载超时设置
                f.write("# 页面加载超时设置（秒）\n")
                f.write(f"PAGE_LOAD_TIMEOUT = {self.page_load_timeout.get()}\n")
                f.write(f"ELEMENT_WAIT_TIMEOUT = {self.element_wait_timeout.get()}\n")
                f.write(f"QUICK_WAIT_TIMEOUT = {self.quick_wait_timeout.get()}\n\n")
                
                # 亚马逊站点配置
                f.write("# 亚马逊站点配置\n")
                f.write(f"AMAZON_SITE = \"{self.amazon_site.get()}\"\n")
                f.write(f"DELIVERY_ZIPCODE = \"{self.delivery_zipcode.get()}\"\n\n")
                
                # 新增：亚马逊页面等待时间
                f.write("# 亚马逊页面等待时间（秒）\n")
                f.write(f"AMAZON_HOMEPAGE_WAIT = {self.amazon_homepage_wait.get()}\n")
                f.write(f"DELIVERY_LOCATION_WAIT = {self.delivery_location_wait.get()}\n\n")
                
                # 网络资源屏蔽列表
                f.write("# 网络资源屏蔽列表\n")
                resources = [r.strip() for r in self.blocked_resources.get().split(",")]
                f.write("BLOCKED_RESOURCES = [\n")
                for i, resource in enumerate(resources):
                    if resource:
                        if i < len(resources) - 1:
                            f.write(f"    \"{resource}\", \n")
                        else:
                            f.write(f"    \"{resource}\"\n")
                f.write("]\n\n")
                
                # 文件路径配置
                f.write("# 文件路径配置\n")
                f.write(f"EXCEL_PATH = \"{self.excel_path.get()}\"\n")
                f.write(f"SCREENSHOTS_DIR = \"{self.screenshots_dir.get()}\"\n\n")
                
                # 数据爬取配置
                f.write("# 数据爬取配置\n")
                f.write(f"MAX_PRODUCTS = {self.max_products.get()}\n")
                f.write(f"PLUGIN_DATA_WAIT_TIME = {self.plugin_data_wait_time.get()}\n")
                f.write(f"PLUGIN_INITIAL_WAIT = {self.plugin_initial_wait.get()}\n")
                f.write(f"SEARCH_RESULT_INITIAL_WAIT = {self.search_result_initial_wait.get()}\n")
                f.write(f"PLUGIN_DATA_PROCESSING_WAIT = {self.plugin_data_processing_wait.get()}\n")
                f.write(f"PRODUCT_SEARCH_INTERVAL = {self.product_search_interval.get()}\n")
                f.write(f"MIN_PRODUCT_SEARCH_INTERVAL = {self.min_product_search_interval.get()}\n\n")
                
                # 默认浏览器关闭等待时间
                f.write("# 默认浏览器关闭等待时间（秒）\n")
                f.write(f"DEFAULT_BROWSER_CLOSE_WAIT = {self.default_browser_close_wait.get()}\n")
            
            print(f"配置已保存到: {filename}")
        except Exception as e:
            print(f"保存配置时出错: {e}")
    
    def load_config(self):
        """Load configuration from a file"""
        filename = filedialog.askopenfilename(filetypes=[("Python文件", "*.py")])
        if not filename:
            return
        
        try:
            # Create a temporary module name
            import importlib.util
            spec = importlib.util.spec_from_file_location("temp_config", filename)
            temp_config = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(temp_config)
            
            # Update all variables from the loaded config
            self.chrome_driver_path.set(temp_config.CHROME_DRIVER_PATH)
            self.extension_path.set(temp_config.EXTENSION_PATH)
            
            self.start_maximized.set(temp_config.BROWSER_OPTIONS.get("start_maximized", True))
            self.no_sandbox.set(temp_config.BROWSER_OPTIONS.get("no_sandbox", True))
            self.disable_dev_shm_usage.set(temp_config.BROWSER_OPTIONS.get("disable_dev_shm_usage", True))
            self.disable_extensions_file_access_check.set(temp_config.BROWSER_OPTIONS.get("disable_extensions_file_access_check", True))
            
            self.page_load_timeout.set(temp_config.PAGE_LOAD_TIMEOUT)
            self.element_wait_timeout.set(temp_config.ELEMENT_WAIT_TIMEOUT)
            self.quick_wait_timeout.set(temp_config.QUICK_WAIT_TIMEOUT)
            
            self.amazon_site.set(temp_config.AMAZON_SITE)
            self.delivery_zipcode.set(temp_config.DELIVERY_ZIPCODE)
            
            # 新增：加载亚马逊页面等待时间（如果存在）
            if hasattr(temp_config, 'AMAZON_HOMEPAGE_WAIT'):
                self.amazon_homepage_wait.set(temp_config.AMAZON_HOMEPAGE_WAIT)
            
            if hasattr(temp_config, 'DELIVERY_LOCATION_WAIT'):
                self.delivery_location_wait.set(temp_config.DELIVERY_LOCATION_WAIT)
            
            self.blocked_resources.set(", ".join(temp_config.BLOCKED_RESOURCES))
            
            self.excel_path.set(temp_config.EXCEL_PATH)
            self.screenshots_dir.set(temp_config.SCREENSHOTS_DIR)
            
            self.max_products.set(temp_config.MAX_PRODUCTS)
            self.plugin_data_wait_time.set(temp_config.PLUGIN_DATA_WAIT_TIME)
            self.plugin_initial_wait.set(temp_config.PLUGIN_INITIAL_WAIT)
            self.search_result_initial_wait.set(temp_config.SEARCH_RESULT_INITIAL_WAIT)
            self.plugin_data_processing_wait.set(temp_config.PLUGIN_DATA_PROCESSING_WAIT)
            self.product_search_interval.set(temp_config.PRODUCT_SEARCH_INTERVAL)
            self.min_product_search_interval.set(temp_config.MIN_PRODUCT_SEARCH_INTERVAL)
            self.default_browser_close_wait.set(temp_config.DEFAULT_BROWSER_CLOSE_WAIT)
            
            print(f"配置已从 {filename} 加载")
        except Exception as e:
            print(f"加载配置时出错: {e}")
    
    def get_current_config(self):
        """Get current configuration as a dictionary"""
        blocked_resources = [r.strip() for r in self.blocked_resources.get().split(",") if r.strip()]
        
        return {
            "CHROME_DRIVER_PATH": self.chrome_driver_path.get(),
            "EXTENSION_PATH": self.extension_path.get(),
            "BROWSER_OPTIONS": {
                "start_maximized": self.start_maximized.get(),
                "no_sandbox": self.no_sandbox.get(),
                "disable_dev_shm_usage": self.disable_dev_shm_usage.get(),
                "disable_extensions_file_access_check": self.disable_extensions_file_access_check.get()
            },
            "PAGE_LOAD_TIMEOUT": self.page_load_timeout.get(),
            "ELEMENT_WAIT_TIMEOUT": self.element_wait_timeout.get(),
            "QUICK_WAIT_TIMEOUT": self.quick_wait_timeout.get(),
            "AMAZON_SITE": self.amazon_site.get(),
            "DELIVERY_ZIPCODE": self.delivery_zipcode.get(),
            "BLOCKED_RESOURCES": blocked_resources,
            "EXCEL_PATH": self.excel_path.get(),
            "SCREENSHOTS_DIR": self.screenshots_dir.get(),
            "MAX_PRODUCTS": self.max_products.get(),
            "PLUGIN_DATA_WAIT_TIME": self.plugin_data_wait_time.get(),
            "PLUGIN_INITIAL_WAIT": self.plugin_initial_wait.get(),
            "SEARCH_RESULT_INITIAL_WAIT": self.search_result_initial_wait.get(),
            "PLUGIN_DATA_PROCESSING_WAIT": self.plugin_data_processing_wait.get(),
            "PRODUCT_SEARCH_INTERVAL": self.product_search_interval.get(),
            "MIN_PRODUCT_SEARCH_INTERVAL": self.min_product_search_interval.get(),
            "DEFAULT_BROWSER_CLOSE_WAIT": self.default_browser_close_wait.get(),
            # 新增加的等待时间配置
            "AMAZON_HOMEPAGE_WAIT": self.amazon_homepage_wait.get(),
            "DELIVERY_LOCATION_WAIT": self.delivery_location_wait.get()
        }

    def setup_browser_with_specific_extension(self, config_dict):
        """设置并启动Chrome浏览器，加载特定的电商数据分析插件"""
        # 配置Chrome选项
        chrome_options = Options()
        
        # 从配置文件加载浏览器选项
        if config_dict["BROWSER_OPTIONS"].get("start_maximized"):
            chrome_options.add_argument("--start-maximized")
        if config_dict["BROWSER_OPTIONS"].get("no_sandbox"):
            chrome_options.add_argument("--no-sandbox")
        if config_dict["BROWSER_OPTIONS"].get("disable_dev_shm_usage"):
            chrome_options.add_argument("--disable-dev-shm-usage")
        if config_dict["BROWSER_OPTIONS"].get("disable_extensions_file_access_check"):
            chrome_options.add_argument("--disable-extensions-file-access-check")
        
        # 加载插件
        chrome_options.add_argument(f"--load-extension={config_dict['EXTENSION_PATH']}")
        
        # 防止检测自动化
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option("useAutomationExtension", False)
        
        # 创建浏览器实例
        service = Service(executable_path=config_dict["CHROME_DRIVER_PATH"])
        print("正在启动Chrome浏览器并加载电商数据分析插件...")
        driver = webdriver.Chrome(options=chrome_options, service=service)
        
        # 修改navigator.webdriver标志
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        # 设置页面加载超时时间
        driver.set_page_load_timeout(config_dict["PAGE_LOAD_TIMEOUT"])
        
        return driver
    
    def visit_amazon_homepage(self, driver, config_dict):
        """访问亚马逊主页"""
        url = f"https://www.{config_dict['AMAZON_SITE']}/"
        print(f"正在访问亚马逊主页: {url}")
        
        try:
            driver.get(url)
            print("亚马逊主页已加载")
        except TimeoutException:
            print(f"亚马逊主页加载超时({config_dict['PAGE_LOAD_TIMEOUT']}秒)，已中断加载")
            # 尝试停止页面加载
            driver.execute_script("window.stop();")
        
        try:
            # 只等待搜索框这个关键元素加载
            WebDriverWait(driver, config_dict["ELEMENT_WAIT_TIMEOUT"]).until(
                EC.presence_of_element_located((By.ID, "twotabsearchtextbox"))
            )
            print("亚马逊主页搜索框已加载，可以继续操作")
            return True
        except Exception as e:
            print(f"等待搜索框超时: {e}")
            return True  # 即使没有找到搜索框，也继续下一步
    
    def search_product(self, driver, keyword, config_dict):
        """在亚马逊搜索特定产品"""
        print(f"正在搜索产品: {keyword}")
        
        try:
            # 减少等待时间，快速定位搜索框
            try:
                search_box = WebDriverWait(driver, config_dict["ELEMENT_WAIT_TIMEOUT"]).until(
                    EC.element_to_be_clickable((By.ID, "twotabsearchtextbox"))
                )
                
                # 尝试使用JavaScript直接清空和设置值，这通常比send_keys快
                driver.execute_script("arguments[0].value = '';", search_box)
                search_box.send_keys(keyword)
                print(f"已输入搜索关键词: {keyword}")
                
                # 优先使用回车键提交搜索，这通常比点击按钮更快
                search_box.send_keys(Keys.RETURN)
                print("使用回车键提交了搜索")
            except TimeoutException:
                print("搜索框加载超时，尝试直接操作")
                try:
                    search_box = driver.find_element(By.ID, "twotabsearchtextbox")
                    driver.execute_script("arguments[0].value = '';", search_box)
                    search_box.send_keys(keyword)
                    search_box.send_keys(Keys.RETURN)
                except:
                    print("无法找到搜索框，继续下一步")
                    return True
            
            # 使用try-except处理页面加载超时
            try:
                # 快速等待搜索结果页面的某些标志性元素
                WebDriverWait(driver, config_dict["ELEMENT_WAIT_TIMEOUT"]).until(
                    EC.any_of(
                        EC.url_contains("s?k="),
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".s-result-list"))
                    )
                )
                print("搜索结果页面开始加载")
            except TimeoutException:
                print("搜索结果页面加载超时，继续下一步")
                # 尝试停止页面加载
                driver.execute_script("window.stop();")
            
            return True
        except Exception as e:
            print(f"搜索产品时出现错误: {e}")
            # 即使出错也尝试继续流程
            return True
    
    def set_delivery_location(self, driver, config_dict):
        """设置亚马逊配送地址"""
        zipcode = config_dict["DELIVERY_ZIPCODE"]
        print(f"正在尝试设置配送地址为: {zipcode}")
        
        try:
            # 等待并点击配送地址按钮
            try:
                WebDriverWait(driver, config_dict["ELEMENT_WAIT_TIMEOUT"]).until(
                    EC.element_to_be_clickable((By.ID, "glow-ingress-block"))
                ).click()
                print("点击了配送地址按钮")
            except TimeoutException:
                print("配送地址按钮加载超时，尝试直接点击")
                try:
                    driver.find_element(By.ID, "glow-ingress-block").click()
                except:
                    print("无法找到配送地址按钮，继续下一步")
                    return True
            
            # 只等待必要的元素 - 邮政编码输入框
            try:
                zipcode_input = WebDriverWait(driver, config_dict["ELEMENT_WAIT_TIMEOUT"]).until(
                    EC.presence_of_element_located((By.ID, "GLUXZipUpdateInput"))
                )
                zipcode_input.clear()
                zipcode_input.send_keys(zipcode)
                print(f"输入了邮政编码: {zipcode}")
                
                # 等待并点击设置按钮 (GLUXZipUpdate)
                try:
                    # 使用更长的超时时间确保按钮可点击
                    update_button = WebDriverWait(driver, config_dict["ELEMENT_WAIT_TIMEOUT"]).until(
                        EC.element_to_be_clickable((By.ID, "GLUXZipUpdate"))
                    )
                    update_button.click()
                    print("点击了设置按钮 (GLUXZipUpdate)")
                except Exception as e:
                    print(f"通过ID查找设置按钮失败: {e}")
                    # 备用方案：尝试通过XPath查找设置按钮
                    try:
                        xpath_selector = "//span[@id='GLUXZipUpdate-announce'][contains(text(), '设置')]/ancestor::span[@id='GLUXZipUpdate']"
                        update_button = WebDriverWait(driver, config_dict["ELEMENT_WAIT_TIMEOUT"]).until(
                            EC.element_to_be_clickable((By.XPATH, xpath_selector))
                        )
                        update_button.click()
                        print("通过XPath点击了设置按钮")
                    except Exception as e2:
                        print(f"通过XPath查找设置按钮失败: {e2}")
                        # 最后尝试使用JavaScript点击
                        try:
                            driver.execute_script("document.getElementById('GLUXZipUpdate').getElementsByTagName('input')[0].click();")
                            print("通过JavaScript点击了设置按钮")
                        except Exception as e3:
                            print(f"所有尝试点击设置按钮都失败: {e3}")
                            print("将继续尝试其他操作")
                
                # 减少等待对话框的时间
                time.sleep(config_dict["QUICK_WAIT_TIMEOUT"])  # 给一点时间让对话框显示
                
                # 尝试更快地找到并点击完成按钮
                try:
                    # 尝试通过多种选择器快速找到完成按钮
                    selectors = [
                        (By.NAME, "glowDoneButton"),
                        (By.ID, "GLUXConfirmClose"),
                        (By.XPATH, "//div[contains(@class, 'a-popover-footer')]//input"),
                        (By.XPATH, "//div[contains(@class, 'a-popover-footer')]//button"),
                        (By.XPATH, "//span[contains(text(), 'Done')]//parent::*")
                    ]
                    
                    for selector in selectors:
                        try:
                            elements = WebDriverWait(driver, config_dict["QUICK_WAIT_TIMEOUT"]).until(
                                EC.element_to_be_clickable(selector)
                            )
                            elements.click()
                            print(f"点击了完成按钮，使用选择器: {selector}")
                            break
                        except:
                            continue
                except:
                    print("无法找到完成按钮，尝试继续操作")
                
                # 减少等待页面刷新的时间
                time.sleep(config_dict["QUICK_WAIT_TIMEOUT"])
                
                return True
            except TimeoutException:
                print("邮政编码输入框加载超时，继续下一步")
                return True
                
        except Exception as e:
            print(f"设置配送地址时出现错误: {e}")
            return True  # 更改为True以不中断主流程
    
    def extract_keyword_data(self, driver, config_dict):
        """从插件数据面板中提取关键词数据
        返回一个字典，包含搜索转化率和点击转化率"""
        print("等待插件数据加载...")
        
        start_time = time.time()
        search_conversion_rate = "无"
        click_conversion_rate = "无"
        
        # 等待插件加载数据
        while time.time() - start_time < config_dict["PLUGIN_DATA_WAIT_TIME"]:
            try:
                # 首先检查是否有表格存在
                tables = driver.find_elements(By.XPATH, "//div[contains(@class, 'ant-table-content')]")
                if not tables:
                    print("未找到数据表格，继续等待...")
                    time.sleep(1)
                    continue
                    
                # 检查是否有"暂无数据"提示
                if len(driver.find_elements(By.XPATH, "//div[contains(@class, 'ant-empty-description') and contains(text(), '暂无数据')]")) > 0:
                    print("插件提示暂无数据")
                    return {"搜索转化率": "无", "点击转化率": "无"}
                
                # 多种方式尝试定位搜索转化率和点击转化率数据
                # 方法1: 通过表格标题行
                try:
                    # 查找表格标题行中"搜索转化率"和"点击转化率"的位置
                    headers = driver.find_elements(By.XPATH, "//th[contains(@class, 'ant-table-cell')]//div[contains(@class, 'ant-flex')]//div[not(contains(@class, 'sc-feUYzb'))]")
                    
                    search_conv_index = -1
                    click_conv_index = -1
                    
                    for i, header in enumerate(headers):
                        header_text = header.text.strip()
                        if "搜索转化率" in header_text:
                            search_conv_index = i
                        elif "点击转化率" in header_text:
                            click_conv_index = i
                    
                    # 如果找到了列索引，则尝试获取数据
                    if search_conv_index >= 0 and click_conv_index >= 0:
                        # 获取所有数据行
                        rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'ant-table-row')]")
                        
                        if rows:
                            # 获取第一行的单元格
                            cells = rows[0].find_elements(By.XPATH, ".//td")
                            
                            if len(cells) > search_conv_index:
                                search_conversion_rate = cells[search_conv_index].text.strip()
                                if search_conversion_rate:
                                    print(f"获取到搜索转化率: {search_conversion_rate}")
                                else:
                                    search_conversion_rate = "无"
                            
                            if len(cells) > click_conv_index:
                                click_conversion_rate = cells[click_conv_index].text.strip()
                                if click_conversion_rate:
                                    print(f"获取到点击转化率: {click_conversion_rate}")
                                else:
                                    click_conversion_rate = "无"
                except Exception as e:
                    print(f"通过表格标题尝试获取数据失败: {e}")
                
                # 方法2: 直接通过固定位置尝试
                if search_conversion_rate == "无" or click_conversion_rate == "无":
                    try:
                        # 假设搜索转化率在第5列，点击转化率在第6列（根据示例判断）
                        rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'ant-table-row')]")
                        if rows:
                            cells = rows[0].find_elements(By.XPATH, ".//td")
                            if len(cells) > 4:  # 第5列（索引为4）
                                search_conversion_rate = cells[4].text.strip()
                                if search_conversion_rate:
                                    print(f"通过固定位置获取到搜索转化率: {search_conversion_rate}")
                            
                            if len(cells) > 5:  # 第6列（索引为5）
                                click_conversion_rate = cells[5].text.strip()
                                if click_conversion_rate:
                                    print(f"通过固定位置获取到点击转化率: {click_conversion_rate}")
                    except Exception as e:
                        print(f"通过固定位置尝试获取数据失败: {e}")
                
                # 方法3: 尝试通过页面源码查找
                if search_conversion_rate == "无" or click_conversion_rate == "无":
                    try:
                        # 获取页面源码
                        page_source = driver.page_source
                        
                        # 使用正则表达式查找搜索转化率和点击转化率
                        search_pattern = r'搜索转化率.*?>([\d.]+%)<'
                        click_pattern = r'点击转化率.*?>([\d.]+%)<'
                        
                        search_match = re.search(search_pattern, page_source)
                        click_match = re.search(click_pattern, page_source)
                        
                        if search_match:
                            search_conversion_rate = search_match.group(1)
                            print(f"通过源码获取到搜索转化率: {search_conversion_rate}")
                        
                        if click_match:
                            click_conversion_rate = click_match.group(1)
                            print(f"通过源码获取到点击转化率: {click_conversion_rate}")
                    except Exception as e:
                        print(f"通过页面源码尝试获取数据失败: {e}")
                
                # 如果两个数据都获取到了，就可以返回结果了
                if search_conversion_rate != "无" and click_conversion_rate != "无":
                    return {
                        "搜索转化率": search_conversion_rate,
                        "点击转化率": click_conversion_rate
                    }
                
                # 如果已经有部分数据，就直接返回
                if search_conversion_rate != "无" or click_conversion_rate != "无":
                    print("只获取到部分数据，将继续使用")
                    return {
                        "搜索转化率": search_conversion_rate,
                        "点击转化率": click_conversion_rate
                    }
                    
            except Exception as e:
                print(f"提取数据时出现错误: {e}")
            
            # 短暂等待后再次尝试
            time.sleep(1)
            print(f"已等待 {int(time.time() - start_time)} 秒，继续尝试获取数据...")
        
        print(f"等待超时({config_dict['PLUGIN_DATA_WAIT_TIME']}秒)，无法获取完整数据")
        return {
            "搜索转化率": search_conversion_rate,
            "点击转化率": click_conversion_rate
        }
    
    def read_product_names_from_excel(self, excel_path, max_products):
        """从Excel文件中读取产品名"""
        try:
            # 检查文件是否存在
            if not os.path.exists(excel_path):
                print(f"错误: 找不到Excel文件: {excel_path}")
                return [], None
            
            # 读取Excel文件
            print(f"正在读取Excel文件: {excel_path}")
            df = pd.read_excel(excel_path)
            
            # 检查是否有"流量词"列
            if "流量词" not in df.columns:
                print(f"错误: Excel文件中没有'流量词'列")
                # 打印所有列名以便调试
                print(f"可用列: {df.columns.tolist()}")
                return [], None
            
            # 获取流量词列的前max_products个值
            product_names = df["流量词"].head(max_products).tolist()
            
            # 过滤掉空值和NaN值
            product_names = [name for name in product_names if pd.notna(name) and name.strip()]
            
            print(f"成功读取{len(product_names)}个产品名")
            return product_names, df
        
        except Exception as e:
            print(f"读取Excel文件时出现错误: {e}")
            return [], None
    
    def start_scraping(self):
        """开始数据爬取"""
        if self.running:
            print("爬取任务已在运行中...")
            return
        
        # 获取当前配置
        config_dict = self.get_current_config()
        
        # 检查Excel文件是否存在
        if not os.path.exists(config_dict["EXCEL_PATH"]):
            print(f"错误: 找不到Excel文件: {config_dict['EXCEL_PATH']}")
            return
        
        # 检查Chrome驱动是否存在
        if not os.path.exists(config_dict["CHROME_DRIVER_PATH"]):
            print(f"错误: 找不到Chrome驱动: {config_dict['CHROME_DRIVER_PATH']}")
            return
        
        # 禁用开始按钮，启用停止按钮
        self.start_button.config(state='disabled')
        self.stop_button.config(state='normal')
        
        # 重置进度条
        self.progress_var.set(0)
        self.progress_label.config(text="准备中...")
        
        # 创建并启动爬取线程
        self.running = True
        self.scrape_thread = threading.Thread(target=self.scrape_data, args=(config_dict,))
        self.scrape_thread.daemon = True
        self.scrape_thread.start()
    
    def stop_scraping(self):
        """停止数据爬取"""
        if not self.running:
            print("没有正在运行的爬取任务")
            return
        
        print("正在停止爬取任务...")
        self.running = False
        
        # 等待线程结束
        if hasattr(self, 'scrape_thread') and self.scrape_thread.is_alive():
            self.scrape_thread.join(2)  # 最多等待2秒
        
        # 关闭浏览器
        if self.driver:
            try:
                self.driver.quit()
                self.driver = None
                print("已关闭浏览器")
            except:
                print("关闭浏览器时出错")
        
        # 启用开始按钮，禁用停止按钮
        self.start_button.config(state='normal')
        self.stop_button.config(state='disabled')
        self.progress_label.config(text="已停止")
    
    def scrape_data(self, config_dict):
        """在单独的线程中运行爬取任务"""
        try:
            # 读取产品名
            product_names, df = self.read_product_names_from_excel(
                config_dict["EXCEL_PATH"], 
                config_dict["MAX_PRODUCTS"]
            )
            
            if not product_names:
                print("没有找到任何产品名，任务将退出")
                self.stop_scraping()
                return
            
            print(f"将搜索以下{len(product_names)}个产品")
            
            # 启动浏览器
            self.driver = self.setup_browser_with_specific_extension(config_dict)
            
            # 减少插件加载等待时间
            print("等待插件加载...")
            time.sleep(config_dict["PLUGIN_INITIAL_WAIT"])
            
            # 设置页面加载策略，但允许加载图片元素
            self.driver.execute_script("window.onbeforeunload = function() { return null; };")
            self.driver.execute_cdp_cmd('Page.setLifecycleEventsEnabled', {'enabled': True})
            # 只阻止非必要的第三方资源，保留图片加载
            self.driver.execute_cdp_cmd('Network.setBlockedURLs', {"urls": config_dict["BLOCKED_RESOURCES"]})
            self.driver.execute_cdp_cmd('Network.enable', {})
            
            # 直接访问亚马逊主页
            visit_success = self.visit_amazon_homepage(self.driver, config_dict)
            
            # 新增：访问亚马逊主页后等待指定时间
            homepage_wait = config_dict["AMAZON_HOMEPAGE_WAIT"]
            print(f"亚马逊主页加载后等待 {homepage_wait} 秒...")
            for i in range(homepage_wait, 0, -1):
                if not self.running:
                    break
                self.progress_label.config(text=f"亚马逊主页加载后等待 {i} 秒...")
                time.sleep(1)
            
            # 设置配送地址
            self.set_delivery_location(self.driver, config_dict)
            
            # 新增：设置配送地址后等待指定时间
            location_wait = config_dict["DELIVERY_LOCATION_WAIT"]
            print(f"设置配送地址后等待 {location_wait} 秒...")
            for i in range(location_wait, 0, -1):
                if not self.running:
                    break
                self.progress_label.config(text=f"设置配送地址后等待 {i} 秒...")
                time.sleep(1)
            
            # 收集每个关键词的数据
            results = {}
            
            print(f"开始依次搜索产品，每次搜索后将等待插件数据加载并提取数据")
            for i, product_name in enumerate(product_names, 1):
                # 检查是否停止运行
                if not self.running:
                    print("爬取任务被中断")
                    break
                
                # 更新进度条
                self.update_progress(i-1, len(product_names), f"正在搜索: {product_name} ({i}/{len(product_names)})")
                
                print(f"\n--- 第{i}/{len(product_names)}个搜索: {product_name} ---")
                
                # 设置每个操作的超时时间
                start_time = time.time()
                
                # 尝试搜索产品
                search_success = self.search_product(self.driver, product_name, config_dict)
                
                # 等待页面加载一些基本内容，然后尝试提取数据
                time.sleep(config_dict["SEARCH_RESULT_INITIAL_WAIT"])
                
                # 提取数据前等待插件完全加载
                print("等待插件加载和处理数据...")
                time.sleep(config_dict["PLUGIN_DATA_PROCESSING_WAIT"])
                
                # 尝试提取数据
                keyword_data = self.extract_keyword_data(self.driver, config_dict)
                results[product_name] = keyword_data
                
                # 截图保存当前页面状态（用于调试）
                try:
                    screenshot_dir = config_dict["SCREENSHOTS_DIR"]
                    if not os.path.exists(screenshot_dir):
                        os.makedirs(screenshot_dir)
                    screenshot_path = f"{screenshot_dir}/{product_name}_{time.strftime('%Y%m%d_%H%M%S')}.png"
                    self.driver.save_screenshot(screenshot_path)
                    print(f"页面截图已保存到: {screenshot_path}")
                except Exception as e:
                    print(f"保存截图时出现错误: {e}")
                
                # 计算实际搜索用时
                elapsed_time = time.time() - start_time
                print(f"本次搜索和数据收集用时: {elapsed_time:.2f}秒")
                
                if i < len(product_names) and self.running:
                    # 实际等待时间为设定间隔减去已花费的时间，但最少等待3秒
                    remaining_wait = max(config_dict["MIN_PRODUCT_SEARCH_INTERVAL"], 
                                        config_dict["PRODUCT_SEARCH_INTERVAL"] - int(elapsed_time))
                    print(f"等待{remaining_wait}秒后搜索下一个产品...")
                    
                    # 创建计数器进行倒计时
                    for j in range(remaining_wait, 0, -1):
                        if not self.running:
                            break
                        self.progress_label.config(text=f"等待 {j} 秒后继续...")
                        time.sleep(1)
            
            # 更新进度条到完成
            self.update_progress(len(product_names), len(product_names), "数据收集完成")
            
            # 将收集到的数据更新到Excel文件
            if df is not None and results and self.running:
                print("\n所有产品搜索和数据收集完成，正在更新Excel文件...")
                update_success = self.update_excel_with_data(df, results, config_dict["EXCEL_PATH"])
                if update_success:
                    self.progress_label.config(text="全部完成")
                else:
                    self.progress_label.config(text="数据收集完成，但Excel更新失败")
            
            # 完成所有搜索后，等待设定的时间再关闭浏览器
            close_wait = config_dict["DEFAULT_BROWSER_CLOSE_WAIT"]
            print(f"\n爬取任务结束。浏览器将在{close_wait}秒后自动关闭...")
            
            # 创建计数器进行倒计时
            for i in range(close_wait, 0, -1):
                if not self.running:
                    break
                self.progress_label.config(text=f"浏览器将在 {i} 秒后关闭...")
                time.sleep(1)
                
        except Exception as e:
            print(f"执行过程中遇到错误: {e}")
            import traceback
            traceback.print_exc()
            self.progress_label.config(text="发生错误")
        
        finally:
            # 程序结束前关闭浏览器
            if self.driver and self.running:
                print("即将关闭浏览器...")
                try:
                    self.driver.quit()
                    self.driver = None
                except:
                    print("关闭浏览器时出错")
            
            # 恢复按钮状态
            self.start_button.config(state='normal')
            self.stop_button.config(state='disabled')
            self.running = False
    
    def update_excel_with_data(self, df, results, excel_path):
        """将收集到的数据更新到Excel文件中"""
        try:
            print("正在更新Excel文件...")
            
            # 更新数据框
            for keyword, data in results.items():
                # 找到关键词所在的行
                mask = df["流量词"] == keyword
                if mask.any():
                    if "搜索转化率" in df.columns:
                        df.loc[mask, "搜索转化率"] = data["搜索转化率"]
                    else:
                        print("警告: Excel文件中没有'搜索转化率'列，无法更新")
                    
                    if "类目转化率" in df.columns:  # 点击转化率对应的列名是"类目转化率"
                        df.loc[mask, "类目转化率"] = data["点击转化率"]
                    else:
                        print("警告: Excel文件中没有'类目转化率'列，无法更新")
                else:
                    print(f"警告: 在Excel文件中找不到关键词: {keyword}")
            
            # 生成带时间戳的新文件名
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            output_path = excel_path.replace(".xlsx", f"_更新_{timestamp}.xlsx")
            
            # 保存到新文件
            df.to_excel(output_path, index=False)
            print(f"已将更新后的数据保存到: {output_path}")
            return True
        
        except Exception as e:
            print(f"更新Excel文件时出现错误: {e}")
            return False

def main():
    """主函数，启动GUI应用"""
    root = tk.Tk()
    app = AmazonScraperGUI(root)
    root.protocol("WM_DELETE_WINDOW", lambda: (app.stop_scraping(), root.destroy()))
    root.mainloop()

if __name__ == "__main__":
    main()