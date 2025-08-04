import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import psutil
import time
from datetime import datetime, timedelta
import traceback
import os
import csv
import logging
import sys
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Configure logging
logging.basicConfig(
    filename='microbot_manager.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def log_unhandled_exception(exc_type, exc_value, exc_traceback):
    """Log unhandled exceptions"""
    logging.error(
        "Unhandled exception",
        exc_info=(exc_type, exc_value, exc_traceback)
    )

sys.excepthook = log_unhandled_exception

# Import components
from components.batch_executor import BatchExecutor
from components.bot_manager import BotManager
from components.proxy_tester import ProxyTester
from components.finance_tracker import FinanceTracker
from components.cloud_integration import CloudIntegration
from components.bot_list_generator import BotListGenerator
from components.config_manager import ConfigManager
from components.runelite_profile_manager import RuneliteProfileManager

class MicrobotManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Microbot Manager Pro")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        try:
            self.config = ConfigManager()
            self.app_config = self.config.load_config("app_config.json")
            self.proxy_config = self.config.load_config("proxy_config.json")
            
            if "bot_jar_name" not in self.app_config:
                self.app_config["bot_jar_name"] = "microbot.jar"
                self.config.save_config("app_config.json", self.app_config)
            
            data_dir = os.path.join("data", "bot_processes")
            os.makedirs(data_dir, exist_ok=True)
            self.bot_manager = BotManager(data_dir=data_dir)
            
            self.batch_executor = BatchExecutor()
            self.proxy_tester = ProxyTester()
            self.finance_tracker = FinanceTracker()
            self.cloud_integration = CloudIntegration()
            self.bot_list_generator = BotListGenerator()
            self.runelite_profile_manager = RuneliteProfileManager()
            
            self.entries = []
            self.active_processes = {}
            self.create_widgets()
            self.setup_style()
            
            self.root.after(1000, self.monitor_processes)
            self.scan_running_bots()
            
            logging.info("Application initialized successfully")
            
        except Exception as e:
            logging.error(f"Failed to initialize application: {str(e)}")
            raise

    def setup_style(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabel', background='#f0f0f0', font=('Segoe UI', 9))
        style.configure('TButton', font=('Segoe UI', 9))
        style.configure('Treeview', font=('Segoe UI', 9), rowheight=25)
        style.configure('Treeview.Heading', font=('Segoe UI', 9, 'bold'))
        style.map('Treeview', background=[('selected', '#0078d7')])

    def create_widgets(self):
        # Notebook for multiple tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True)
        
        # Create tabs
        self.create_bat_generator_tab()
        self.create_process_manager_tab()
        self.create_proxy_tester_tab()
        self.create_finance_tracker_tab()
        self.create_cloud_tab()
        self.create_bot_list_tab()
        self.create_runelite_tab()
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief='sunken')
        self.status_bar.pack(fill='x')

    def create_bat_generator_tab(self):
        # BAT Generator Tab
        bat_tab = ttk.Frame(self.notebook)
        self.notebook.add(bat_tab, text="BAT Generator")
        
        # Frame: Jar Path & Folder Selection
        top_frame = ttk.Frame(bat_tab, padding="10")
        top_frame.pack(fill='x')
        
        ttk.Label(top_frame, text="Jar Location:").grid(row=0, column=0, sticky="w")
        self.jar_path_var = tk.StringVar()
        jar_entry = ttk.Entry(top_frame, textvariable=self.jar_path_var, width=50)
        jar_entry.grid(row=0, column=1, sticky="ew", padx=(5, 5))
        ttk.Button(top_frame, text="Browse", command=self.browse_jar).grid(row=0, column=2)
        
        ttk.Label(top_frame, text="Save Folder:").grid(row=1, column=0, sticky="w")
        self.save_folder_var = tk.StringVar()
        folder_entry = ttk.Entry(top_frame, textvariable=self.save_folder_var, width=50)
        folder_entry.grid(row=1, column=1, sticky="ew", padx=(5, 5))
        ttk.Button(top_frame, text="Choose", command=self.browse_folder).grid(row=1, column=2)
        
        # Frame: Input Area
        input_frame = ttk.LabelFrame(bat_tab, text="Add Entry Manually", padding="10")
        input_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(input_frame, text="Proxy (IP:Port:User:Pass):").grid(row=0, column=0, sticky="w")
        self.proxy_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.proxy_var).grid(row=0, column=1, sticky="ew")
        
        ttk.Label(input_frame, text="Profile Name:").grid(row=1, column=0, sticky="w")
        self.profile_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.profile_var).grid(row=1, column=1, sticky="ew")
        
        ttk.Button(input_frame, text="Add Entry", command=self.add_manual_entry).grid(row=2, column=0, columnspan=2, pady=5)
        
        # Treeview for entries
        tree_frame = ttk.Frame(bat_tab)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.tree = ttk.Treeview(tree_frame, columns=("Proxy", "Profile"), show="headings")
        self.tree.heading("Proxy", text="Proxy")
        self.tree.heading("Profile", text="Profile Name")
        self.tree.pack(side='left', fill='both', expand=True)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Buttons frame
        btn_frame = ttk.Frame(bat_tab)
        btn_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(btn_frame, text="Edit Selected", command=self.edit_selected).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Delete Selected", command=self.delete_selected).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Load from CSV", command=self.load_csv).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Generate .BAT Files", command=self.generate_bats).pack(side='right', padx=5)
        
        # Batch execution controls
        exec_frame = ttk.LabelFrame(bat_tab, text="Batch Execution", padding="10")
        exec_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(exec_frame, text="Target Folder:").grid(row=0, column=0, sticky="w")
        self.target_folder_var = tk.StringVar()
        ttk.Entry(exec_frame, textvariable=self.target_folder_var, width=50).grid(row=0, column=1, sticky="ew", padx=(5, 5))
        ttk.Button(exec_frame, text="Browse", command=self.browse_target_folder).grid(row=0, column=2)
        
        ttk.Label(exec_frame, text="Delay (seconds):").grid(row=1, column=0, sticky="w")
        self.delay_var = tk.StringVar(value="2")
        ttk.Entry(exec_frame, textvariable=self.delay_var, width=10).grid(row=1, column=1, sticky="w", padx=(5, 5))
        
        ttk.Button(exec_frame, text="Execute Batch Files", command=self.execute_batch_files).grid(row=1, column=2, sticky="e")

    def create_process_manager_tab(self):
        # Process Manager Tab
        proc_tab = ttk.Frame(self.notebook)
        self.notebook.add(proc_tab, text="Bot Manager")
        
        # Top control frame
        top_frame = ttk.Frame(proc_tab)
        top_frame.pack(fill='x', padx=10, pady=5)
        
        # Scan for running bots button
        ttk.Button(top_frame, text="Scan for Running Bots", 
                  command=self.scan_running_bots).pack(side='left', padx=5)
        
        # Cleanup old processes button
        ttk.Button(top_frame, text="Cleanup Stopped", 
                  command=self.cleanup_stopped_processes).pack(side='left', padx=5)
        
        # Refresh button
        ttk.Button(top_frame, text="Refresh", 
                  command=self.update_process_table).pack(side='right', padx=5)
        
        # Live processes table
        live_frame = ttk.LabelFrame(proc_tab, text="Active Bots", padding="10")
        live_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        columns = ("ID", "Profile", "Status", "Uptime", "CPU", "Memory", "Restarts", "Anti-Crash", "Type")
        self.process_tree = ttk.Treeview(live_frame, columns=columns, show="headings")
        
        for col in columns:
            self.process_tree.heading(col, text=col)
            self.process_tree.column(col, width=80, anchor='center')
        
        # Adjust column widths
        self.process_tree.column("ID", width=50)
        self.process_tree.column("Profile", width=120)
        self.process_tree.column("Status", width=100)
        self.process_tree.column("Uptime", width=100)
        self.process_tree.column("Type", width=80)
        
        self.process_tree.pack(side='left', fill='both', expand=True)
        
        scrollbar = ttk.Scrollbar(live_frame, orient="vertical", command=self.process_tree.yview)
        scrollbar.pack(side='right', fill='y')
        self.process_tree.configure(yscrollcommand=scrollbar.set)
        
        # Control buttons
        control_frame = ttk.Frame(proc_tab)
        control_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(control_frame, text="Start Selected", command=self.start_selected_bots).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Stop Selected", command=self.stop_selected_bots).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Restart Selected", command=self.restart_selected_bots).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Toggle Anti-Crash", command=self.toggle_anti_crash).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Remove Selected", command=self.remove_selected_bots).pack(side='left', padx=5)
        
        # Status colors
        self.process_tree.tag_configure('running', background='#e6f7e6')
        self.process_tree.tag_configure('crashed', background='#ffebeb')
        self.process_tree.tag_configure('stopped', background='#f0f0f0')
        self.process_tree.tag_configure('manual', background='#fff2cc')

    def create_proxy_tester_tab(self):
        # Proxy Tester Tab
        proxy_tab = ttk.Frame(self.notebook)
        self.notebook.add(proxy_tab, text="Proxy Tester")
        
        # Proxy input area
        input_frame = ttk.Frame(proxy_tab, padding="10")
        input_frame.pack(fill='x')
        
        ttk.Label(input_frame, text="Proxy (IP:Port:User:Pass):").grid(row=0, column=0, sticky="w")
        self.proxy_test_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.proxy_test_var, width=50).grid(row=0, column=1, sticky="ew", padx=(5, 5))
        ttk.Button(input_frame, text="Add Proxy", command=self.add_test_proxy).grid(row=0, column=2)
        
        ttk.Label(input_frame, text="Test URL:").grid(row=1, column=0, sticky="w")
        self.test_url_var = tk.StringVar(value="https://www.google.com")
        ttk.Entry(input_frame, textvariable=self.test_url_var).grid(row=1, column=1, sticky="ew", padx=(5, 5))
        ttk.Button(input_frame, text="Test Proxies", command=self.test_proxies).grid(row=1, column=2)
        
        # Results table
        results_frame = ttk.LabelFrame(proxy_tab, text="Proxy Test Results", padding="10")
        results_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        columns = ("Proxy", "Status", "Response Time", "Success Rate", "Last Test")
        self.proxy_tree = ttk.Treeview(results_frame, columns=columns, show="headings")
        
        for col in columns:
            self.proxy_tree.heading(col, text=col)
            self.proxy_tree.column(col, width=120, anchor='center')
        
        self.proxy_tree.pack(side='left', fill='both', expand=True)
        
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.proxy_tree.yview)
        scrollbar.pack(side='right', fill='y')
        self.proxy_tree.configure(yscrollcommand=scrollbar.set)

    def create_finance_tracker_tab(self):
        # Finance Tracker Tab - Enhanced Version
        finance_tab = ttk.Frame(self.notebook)
        self.notebook.add(finance_tab, text="Finance Tracker")
        
        # Create notebook for multiple finance sections
        finance_notebook = ttk.Notebook(finance_tab)
        finance_notebook.pack(fill='both', expand=True)
        
        # Transaction Entry Frame
        entry_frame = ttk.Frame(finance_notebook)
        finance_notebook.add(entry_frame, text="Add Transaction")
        
        # Summary Frame
        summary_frame = ttk.Frame(finance_notebook)
        finance_notebook.add(summary_frame, text="Summary")
        
        # Reports Frame
        reports_frame = ttk.Frame(finance_notebook)
        finance_notebook.add(reports_frame, text="Reports")
        
        # ===== Transaction Entry Section =====
        input_frame = ttk.LabelFrame(entry_frame, text="Add Transaction", padding="10")
        input_frame.pack(fill='x', padx=10, pady=5)
        
        # Transaction type
        ttk.Label(input_frame, text="Type:").grid(row=0, column=0, sticky="w")
        self.transaction_type = tk.StringVar(value="expense")
        type_combo = ttk.Combobox(input_frame, textvariable=self.transaction_type, 
                                values=["expense", "income"], state="readonly")
        type_combo.grid(row=0, column=1, sticky="ew")
        type_combo.bind("<<ComboboxSelected>>", self._update_finance_categories)
        
        # Date
        ttk.Label(input_frame, text="Date:").grid(row=1, column=0, sticky="w")
        self.transaction_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(input_frame, textvariable=self.transaction_date).grid(row=1, column=1, sticky="ew")
        
        # Amount
        ttk.Label(input_frame, text="Amount:").grid(row=2, column=0, sticky="w")
        self.amount_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.amount_var).grid(row=2, column=1, sticky="ew")
        
        # Category
        ttk.Label(input_frame, text="Category:").grid(row=3, column=0, sticky="w")
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(input_frame, textvariable=self.category_var)
        self.category_combo.grid(row=3, column=1, sticky="ew")
        
        # Description
        ttk.Label(input_frame, text="Description:").grid(row=4, column=0, sticky="w")
        self.desc_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.desc_var).grid(row=4, column=1, sticky="ew")
        
        # Tax Notes
        ttk.Label(input_frame, text="Tax Notes:").grid(row=5, column=0, sticky="w")
        self.tax_notes_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.tax_notes_var).grid(row=5, column=1, sticky="ew")
        
        # Add button
        ttk.Button(input_frame, text="Add Transaction", command=self.add_transaction).grid(
            row=6, column=0, columnspan=2, pady=5)
        
        # Transaction List
        list_frame = ttk.LabelFrame(entry_frame, text="Recent Transactions", padding="10")
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        columns = ("id", "date", "type", "amount", "category", "description")
        self.transaction_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        for col in columns:
            self.transaction_tree.heading(col, text=col.capitalize())
            self.transaction_tree.column(col, width=100)
        
        self.transaction_tree.pack(side='left', fill='both', expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.transaction_tree.yview)
        scrollbar.pack(side='right', fill='y')
        self.transaction_tree.configure(yscrollcommand=scrollbar.set)
        
        # Context menu for transactions
        self.transaction_menu = tk.Menu(self.root, tearoff=0)
        self.transaction_menu.add_command(label="Edit", command=self.edit_transaction)
        self.transaction_menu.add_command(label="Delete", command=self.delete_transaction)
        self.transaction_tree.bind("<Button-3>", self.show_transaction_context_menu)
        
        # ===== Summary Section =====
        # Summary widgets
        self.summary_text = tk.Text(summary_frame, height=15, state='normal')
        self.summary_text.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Chart frame
        chart_frame = ttk.Frame(summary_frame)
        chart_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Date range controls
        range_frame = ttk.Frame(summary_frame)
        range_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(range_frame, text="From:").pack(side='left')
        self.start_date_var = tk.StringVar(value=(datetime.now() - timedelta(days=30)).strftime('%Y-%m-%d'))
        ttk.Entry(range_frame, textvariable=self.start_date_var, width=10).pack(side='left', padx=5)
        
        ttk.Label(range_frame, text="To:").pack(side='left')
        self.end_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        ttk.Entry(range_frame, textvariable=self.end_date_var, width=10).pack(side='left', padx=5)
        
        ttk.Button(range_frame, text="Update Summary", 
                  command=self.update_finance_summary).pack(side='right', padx=5)
        
        # ===== Reports Section =====
        reports_btn_frame = ttk.Frame(reports_frame)
        reports_btn_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(reports_btn_frame, text="Export to Excel", 
                  command=lambda: self.export_finance('excel')).pack(side='left', padx=5)
        ttk.Button(reports_btn_frame, text="Export to PDF", 
                  command=lambda: self.export_finance('pdf')).pack(side='left', padx=5)
        ttk.Button(reports_btn_frame, text="Generate Tax Report", 
                  command=self.generate_tax_report).pack(side='left', padx=5)
        
        # Initialize categories and summary
        self._update_finance_categories()
        self.update_finance_summary()
        self._load_recent_transactions()

    def create_cloud_tab(self):
        # Cloud Integration Tab
        cloud_tab = ttk.Frame(self.notebook)
        self.notebook.add(cloud_tab, text="Cloud")
        
        # Connection management
        conn_frame = ttk.LabelFrame(cloud_tab, text="Connection Management", padding="10")
        conn_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(conn_frame, text="Connection:").grid(row=0, column=0, sticky="w")
        self.connection_var = tk.StringVar()
        self.connection_combo = ttk.Combobox(conn_frame, textvariable=self.connection_var, state="readonly")
        self.connection_combo.grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(conn_frame, text="Add New", command=self.show_add_connection_dialog).grid(row=0, column=2, padx=5)
        
        btn_frame = ttk.Frame(conn_frame)
        btn_frame.grid(row=1, column=0, columnspan=3, pady=5)
        ttk.Button(btn_frame, text="Connect", command=self.cloud_connect).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Disconnect", command=self.cloud_disconnect).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Get Status", command=self.get_cloud_status).pack(side='left', padx=5)
        
        # File operations
        file_frame = ttk.LabelFrame(cloud_tab, text="File Operations", padding="10")
        file_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(file_frame, text="Upload File", command=self.cloud_upload).pack(side='left', padx=5)
        ttk.Button(file_frame, text="Download File", command=self.cloud_download).pack(side='left', padx=5)
        ttk.Button(file_frame, text="Execute Command", command=self.cloud_execute).pack(side='left', padx=5)
        
        # Status display
        status_frame = ttk.LabelFrame(cloud_tab, text="Status", padding="10")
        status_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.cloud_status_text = tk.Text(status_frame, height=10, state='normal')
        self.cloud_status_text.pack(fill='both', expand=True)
        
        # Initialize connection list
        self.update_connection_list()

    def create_bot_list_tab(self):
        # Bot List Generator Tab
        list_tab = ttk.Frame(self.notebook)
        self.notebook.add(list_tab, text="Bot List")
        
        # Main frames
        input_frame = ttk.LabelFrame(list_tab, text="Add/Edit Profile", padding="10")
        input_frame.pack(fill='x', padx=10, pady=5)
        
        list_frame = ttk.LabelFrame(list_tab, text="Bot Profiles", padding="10")
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Input widgets
        ttk.Label(input_frame, text="Profile Name:").grid(row=0, column=0, sticky="w")
        self.profile_name_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.profile_name_var).grid(row=0, column=1, sticky="ew")
        
        ttk.Label(input_frame, text="Proxy Info:").grid(row=1, column=0, sticky="w")
        self.profile_proxy_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.profile_proxy_var).grid(row=1, column=1, sticky="ew")
        
        ttk.Label(input_frame, text="Notes:").grid(row=2, column=0, sticky="w")
        self.profile_notes_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.profile_notes_var).grid(row=2, column=1, sticky="ew")
        
        # Buttons
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=5)
        ttk.Button(btn_frame, text="Add Profile", command=self.add_bot_profile).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Update Status", command=self.update_bot_status).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Export to CSV", command=self.export_bot_list).pack(side='right', padx=5)
        
        # List of profiles
        columns = ("name", "proxy", "status", "banned", "last_active")
        self.bot_list_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        for col in columns:
            self.bot_list_tree.heading(col, text=col.capitalize().replace("_", " "))
            self.bot_list_tree.column(col, width=120)
        
        self.bot_list_tree.pack(fill='both', expand=True)
        self.update_bot_list()
        
        # Setup context menu
        self.setup_bot_list_context_menu()

    def create_runelite_tab(self):
        """Create tab for Runelite profile management"""
        runelite_tab = ttk.Frame(self.notebook)
        self.notebook.add(runelite_tab, text="Runelite Profiles")
        
        # Main frames
        input_frame = ttk.LabelFrame(runelite_tab, text="Create Profile", padding="10")
        input_frame.pack(fill='x', padx=10, pady=5)
        
        list_frame = ttk.LabelFrame(runelite_tab, text="Existing Profiles", padding="10")
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Input widgets
        ttk.Label(input_frame, text="Profile Name:").grid(row=0, column=0, sticky="w")
        self.rl_profile_name_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.rl_profile_name_var).grid(row=0, column=1, sticky="ew")
        
        ttk.Label(input_frame, text="Username:").grid(row=1, column=0, sticky="w")
        self.rl_username_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.rl_username_var).grid(row=1, column=1, sticky="ew")
        
        ttk.Label(input_frame, text="Password:").grid(row=2, column=0, sticky="w")
        self.rl_password_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.rl_password_var, show="*").grid(row=2, column=1, sticky="ew")
        
        ttk.Label(input_frame, text="Auth Code:").grid(row=3, column=0, sticky="w")
        self.rl_auth_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.rl_auth_var).grid(row=3, column=1, sticky="ew")
        
        # Buttons
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=5)
        ttk.Button(btn_frame, text="Create Profile", command=self.create_runelite_profile).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Refresh List", command=self.update_runelite_profile_list).pack(side='left', padx=5)
        
        # List of profiles
        columns = ("name", "path")
        self.runelite_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        self.runelite_tree.heading("name", text="Profile Name")
        self.runelite_tree.heading("path", text="Location")
        self.runelite_tree.pack(fill='both', expand=True)
        
        # Context menu
        self.runelite_menu = tk.Menu(self.root, tearoff=0)
        self.runelite_menu.add_command(label="Delete Profile", command=self.delete_runelite_profile)
        self.runelite_menu.add_command(label="Copy Path", command=self.copy_runelite_profile_path)
        self.runelite_tree.bind("<Button-3>", self.show_runelite_context_menu)
        
        # Initial list update
        self.update_runelite_profile_list()

    def create_runelite_profile(self):
        """Create a new Runelite profile"""
        profile_name = self.rl_profile_name_var.get()
        username = self.rl_username_var.get()
        password = self.rl_password_var.get()
        auth = self.rl_auth_var.get()
        
        if not profile_name or not username or not password:
            messagebox.showerror("Error", "Profile name, username and password are required")
            return
        
        try:
            credentials = {
                'username': username,
                'password': password,
                'auth': auth if auth else ''
            }
            
            # You can add custom configuration overrides here if needed
            config_overrides = {
                'microbot': {
                    'profile': profile_name,
                    'proxy': self.proxy_var.get() if self.proxy_var.get() else ''
                }
            }
            
            profile_path = self.runelite_profile_manager.create_profile(
                profile_name,
                credentials,
                config_overrides
            )
            
            messagebox.showinfo("Success", f"Profile created at:\n{profile_path}")
            self.update_runelite_profile_list()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create profile: {str(e)}")

    def update_runelite_profile_list(self):
        """Update the list of Runelite profiles"""
        for item in self.runelite_tree.get_children():
            self.runelite_tree.delete(item)
        
        profiles = self.runelite_profile_manager.list_profiles()
        for profile in profiles:
            path = self.runelite_profile_manager.get_profile_path(profile)
            self.runelite_tree.insert('', 'end', values=(profile, path))

    def delete_runelite_profile(self):
        """Delete selected Runelite profile"""
        selected = self.runelite_tree.selection()
        if not selected:
            return
        
        item = selected[0]
        profile_name = self.runelite_tree.item(item, 'values')[0]
        
        if messagebox.askyesno("Confirm", f"Delete profile '{profile_name}'?"):
            if self.runelite_profile_manager.delete_profile(profile_name):
                self.update_runelite_profile_list()
                messagebox.showinfo("Success", "Profile deleted")
            else:
                messagebox.showerror("Error", "Failed to delete profile")

    def copy_runelite_profile_path(self):
        """Copy profile path to clipboard"""
        selected = self.runelite_tree.selection()
        if not selected:
            return
        
        item = selected[0]
        profile_path = self.runelite_tree.item(item, 'values')[1]
        self.root.clipboard_clear()
        self.root.clipboard_append(profile_path)
        self.status_var.set("Profile path copied to clipboard")

    def show_runelite_context_menu(self, event):
        """Show context menu for Runelite profile list"""
        item = self.runelite_tree.identify_row(event.y)
        if item:
            self.runelite_tree.selection_set(item)
            self.runelite_menu.post(event.x_root, event.y_root)

    def _update_finance_categories(self, event=None):
        """Update category dropdown based on transaction type"""
        categories = self.finance_tracker.get_categories(self.transaction_type.get())
        self.category_combo['values'] = categories
        if categories:
            self.category_var.set(categories[0])

    def _load_recent_transactions(self, limit=50):
        """Load recent transactions into the treeview"""
        for item in self.transaction_tree.get_children():
            self.transaction_tree.delete(item)
        
        transactions = self.finance_tracker.get_transactions()
        for i, t in enumerate(sorted(transactions, key=lambda x: x['date'], reverse=True)[:limit]):
            # Use the index as ID if 'id' field is missing
            trans_id = t.get('id', i+1)
            self.transaction_tree.insert('', 'end', values=(
                trans_id,
                t['date'],
                t['type'],
                f"${t['amount']:,.2f}",
                t['category'],
                t['description']
            ))

    def show_transaction_context_menu(self, event):
        """Show context menu for transaction treeview"""
        item = self.transaction_tree.identify_row(event.y)
        if item:
            self.transaction_tree.selection_set(item)
            self.transaction_menu.post(event.x_root, event.y_root)

    def edit_transaction(self):
        """Edit selected transaction"""
        selected = self.transaction_tree.selection()
        if not selected:
            return
        
        item = selected[0]
        trans_id = int(self.transaction_tree.item(item, 'values')[0])
        transaction = self.finance_tracker.get_transaction(trans_id)
        
        if not transaction:
            return
        
        # Create edit dialog
        edit_dialog = tk.Toplevel(self.root)
        edit_dialog.title(f"Edit Transaction #{trans_id}")
        edit_dialog.resizable(False, False)
        
        # Type
        tk.Label(edit_dialog, text="Type:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        type_var = tk.StringVar(value=transaction['type'])
        ttk.Combobox(edit_dialog, textvariable=type_var, 
                    values=["expense", "income"], state="readonly").grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        # Date
        tk.Label(edit_dialog, text="Date:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        date_var = tk.StringVar(value=transaction['date'])
        ttk.Entry(edit_dialog, textvariable=date_var).grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        # Amount
        tk.Label(edit_dialog, text="Amount:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        amount_var = tk.StringVar(value=transaction['amount'])
        ttk.Entry(edit_dialog, textvariable=amount_var).grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        # Category
        tk.Label(edit_dialog, text="Category:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        category_var = tk.StringVar(value=transaction['category'])
        ttk.Combobox(edit_dialog, textvariable=category_var, 
                    values=self.finance_tracker.get_categories(type_var.get())).grid(row=3, column=1, sticky="ew", padx=5, pady=5)
        
        # Description
        tk.Label(edit_dialog, text="Description:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        desc_var = tk.StringVar(value=transaction['description'])
        ttk.Entry(edit_dialog, textvariable=desc_var).grid(row=4, column=1, sticky="ew", padx=5, pady=5)
        
        # Tax Notes
        tk.Label(edit_dialog, text="Tax Notes:").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        tax_notes_var = tk.StringVar(value=transaction.get('tax_notes', ''))
        ttk.Entry(edit_dialog, textvariable=tax_notes_var).grid(row=5, column=1, sticky="ew", padx=5, pady=5)
        
        def save_changes():
            success, message = self.finance_tracker.edit_transaction(
                trans_id,
                type=type_var.get(),
                amount=amount_var.get(),
                category=category_var.get(),
                description=desc_var.get(),
                date=date_var.get(),
                tax_notes=tax_notes_var.get()
            )
            if success:
                self._load_recent_transactions()
                self.update_finance_summary()
                edit_dialog.destroy()
                messagebox.showinfo("Success", message)
            else:
                messagebox.showerror("Error", message)
        
        ttk.Button(edit_dialog, text="Save", command=save_changes).grid(row=6, column=0, columnspan=2, pady=5)

    def delete_transaction(self):
        """Delete selected transaction"""
        selected = self.transaction_tree.selection()
        if not selected:
            return
        
        item = selected[0]
        trans_id = int(self.transaction_tree.item(item, 'values')[0])
        
        if messagebox.askyesno("Confirm", "Delete this transaction?"):
            success, message = self.finance_tracker.delete_transaction(trans_id)
            if success:
                self._load_recent_transactions()
                self.update_finance_summary()
                messagebox.showinfo("Success", message)
            else:
                messagebox.showerror("Error", message)

    def add_transaction(self):
        """Add a new transaction"""
        try:
            amount = float(self.amount_var.get())
            if amount <= 0:
                messagebox.showerror("Error", "Amount must be positive")
                return
            
            success, result = self.finance_tracker.add_transaction(
                self.transaction_type.get(),
                amount,
                self.category_var.get(),
                self.desc_var.get(),
                self.transaction_date.get(),
                self.tax_notes_var.get()
            )
            
            if success:
                messagebox.showinfo("Success", "Transaction added successfully")
                self.amount_var.set("")
                self.desc_var.set("")
                self.tax_notes_var.set("")
                self._load_recent_transactions()
                self.update_finance_summary()
            else:
                messagebox.showerror("Error", result)
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid amount")

    def update_finance_summary(self):
        """Update the finance summary display"""
        try:
            start_date = self.start_date_var.get()
            end_date = self.end_date_var.get()
            
            success, summary = self.finance_tracker.get_summary(start_date, end_date)
            if not success:
                messagebox.showerror("Error", summary)
                return
            
            self.summary_text.config(state='normal')
            self.summary_text.delete(1.0, tk.END)
            
            # Handle empty date ranges safely
            start_date_str = summary['period']['start'] if summary['period']['start'] else "N/A"
            end_date_str = summary['period']['end'] if summary['period']['end'] else "N/A"
            
            text = f"""Financial Summary ({start_date_str} to {end_date_str})
            
Total Income: ${summary['totals']['income']:,.2f}
Total Expenses: ${summary['totals']['expenses']:,.2f}
Net Profit: ${summary['totals']['net_profit']:,.2f}
Transaction Count: {summary['totals']['count']}

Income by Category:
"""
            for cat, amount in summary['by_category']['income'].items():
                text += f"    {cat}: ${amount:,.2f}\n"
            
            text += "\nExpenses by Category:\n"
            for cat, amount in summary['by_category']['expense'].items():
                text += f"    {cat}: ${amount:,.2f}\n"
            
            text += "\nTax Information:\n"
            for tax_cat, amount in summary['by_tax_category']['income'].items():
                text += f"    Income ({tax_cat}): ${amount:,.2f}\n"
            for tax_cat, amount in summary['by_tax_category']['expense'].items():
                text += f"    Expense ({tax_cat}): ${amount:,.2f}\n"
            
            self.summary_text.insert(tk.END, text)
            self.summary_text.config(state='disabled')
            
            # Update charts if they exist
            if hasattr(self, 'finance_chart_frame'):
                self._update_finance_charts(summary)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update finance summary: {str(e)}")

    def export_finance(self, export_type):
        """Export financial data"""
        try:
            if export_type == 'excel':
                success, message = self.finance_tracker.export_to_excel()
            elif export_type == 'pdf':
                success, message = self.finance_tracker.export_to_pdf()
            
            if success:
                messagebox.showinfo("Success", message)
            else:
                messagebox.showerror("Error", message)
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")

    def generate_tax_report(self):
        """Generate tax report for the current year"""
        year = simpledialog.askinteger("Tax Report", "Enter year:", 
                                      parent=self.root,
                                      minvalue=2000, maxvalue=datetime.now().year)
        if year:
            success, message = self.finance_tracker.export_tax_report(year)
            if success:
                messagebox.showinfo("Success", message)
            else:
                messagebox.showerror("Error", message)

    def setup_bot_list_context_menu(self):
        """Create right-click context menu for bot list"""
        self.bot_list_menu = tk.Menu(self.root, tearoff=0)
        self.bot_list_menu.add_command(
            label="Edit Profile", 
            command=self.edit_bot_profile
        )
        self.bot_list_menu.add_command(
            label="Set Status", 
            command=self.show_status_dialog
        )
        self.bot_list_menu.add_separator()
        self.bot_list_menu.add_command(
            label="Export Selected", 
            command=self.export_selected_bots
        )
        
        # Bind right-click to treeview
        self.bot_list_tree.bind("<Button-3>", self.show_bot_list_context_menu)

    def show_bot_list_context_menu(self, event):
        """Show context menu on right-click"""
        item = self.bot_list_tree.identify_row(event.y)
        if item:
            self.bot_list_tree.selection_set(item)
            self.bot_list_menu.post(event.x_root, event.y_root)

    def edit_bot_profile(self):
        """Edit selected bot profile"""
        selected = self.bot_list_tree.selection()
        if not selected:
            return
            
        item = selected[0]
        values = self.bot_list_tree.item(item, 'values')
        profile_data = self.bot_list_generator.profiles.get(values[0], {})
        
        # Create edit dialog
        edit_dialog = tk.Toplevel(self.root)
        edit_dialog.title(f"Edit Profile - {values[0]}")
        edit_dialog.resizable(False, False)
        
        # Name (readonly)
        tk.Label(edit_dialog, text="Profile Name:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        name_var = tk.StringVar(value=values[0])
        tk.Entry(edit_dialog, textvariable=name_var, state='readonly').grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        # Proxy
        tk.Label(edit_dialog, text="Proxy:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        proxy_var = tk.StringVar(value=values[1])
        tk.Entry(edit_dialog, textvariable=proxy_var).grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        # Notes
        tk.Label(edit_dialog, text="Notes:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        notes_var = tk.StringVar(value=profile_data.get("notes", ""))
        notes_entry = tk.Entry(edit_dialog, textvariable=notes_var)
        notes_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        def save_changes():
            # Update profile in the backend
            self.bot_list_generator.profiles[values[0]]["proxy"] = proxy_var.get()
            self.bot_list_generator.profiles[values[0]]["notes"] = notes_var.get()
            self.bot_list_generator._save_profiles()
            
            # Update treeview
            self.update_bot_list()
            edit_dialog.destroy()
        
        tk.Button(edit_dialog, text="Save", command=save_changes).grid(row=3, column=0, columnspan=2, pady=5)

    def export_selected_bots(self):
        """Export selected bots to CSV"""
        selected = self.bot_list_tree.selection()
        if not selected:
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not file_path:
            return
            
        profiles_to_export = []
        for item in selected:
            profile_name = self.bot_list_tree.item(item, 'values')[0]
            profiles_to_export.append(profile_name)
            
        if self.bot_list_generator.export_selected_to_csv(file_path, profiles_to_export):
            messagebox.showinfo("Success", "Selected profiles exported successfully")

    def scan_running_bots(self):
        """Scan for running bot processes and add them to management"""
        jar_name = self.app_config.get("bot_jar_name", "microbot.jar")
        count = self.bot_manager.scan_running_bots(jar_name)
        self.status_var.set(f"Found {count} running bot processes")
        self.update_process_table()

    def cleanup_stopped_processes(self):
        """Cleanup old stopped processes from the list"""
        removed = self.bot_manager.cleanup_stopped_processes()
        self.status_var.set(f"Removed {removed} old processes")
        self.update_process_table()

    def monitor_processes(self):
        """Monitor bot processes using Tkinter's after() instead of threading"""
        try:
            self.bot_manager.monitor_processes()
            self.update_process_table()
        except Exception as e:
            self.status_var.set(f"Monitor error: {str(e)}")
        finally:
            # Schedule the next monitoring cycle
            self.root.after(5000, self.monitor_processes)

    def update_process_table(self):
        """Update the process table with current status"""
        # Clear current items
        for item in self.process_tree.get_children():
            self.process_tree.delete(item)
        
        # Get all processes and their status
        processes = self.bot_manager.get_all_processes()
        
        for pid, status in processes.items():
            if status is None:
                continue
                
            if "error" in status:
                # Error getting status for this process
                self.process_tree.insert('', 'end', values=(
                    pid, "Error", status["error"], "", "", "", "", "", ""
                ), tags=('error',))
                continue
            
            # Determine tag for coloring
            tag = 'running'
            if "crashed" in status["status"].lower():
                tag = 'crashed'
            elif "stopped" in status["status"].lower():
                if "manual" in status["status"].lower():
                    tag = 'manual'
                else:
                    tag = 'stopped'
            
            # Add to treeview
            self.process_tree.insert('', 'end', values=(
                pid,
                status["profile"],
                status["status"],
                status["uptime"],
                status["cpu"],
                status["memory"],
                status["restarts"],
                "Yes" if status["anti_crash"] else "No",
                "Detected" if status.get("detected", False) else "Managed"
            ), tags=(tag,))

    def remove_selected_bots(self):
        """Remove selected bots from monitoring"""
        selected = self.process_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select one or more bots to remove")
            return
            
        for item in selected:
            pid = self.process_tree.item(item, 'values')[0]
            success, result = self.bot_manager.remove_bot(pid)
            
            if not success:
                messagebox.showerror("Remove Failed", 
                    f"Failed to remove bot {pid}: {result}")
            else:
                self.status_var.set(f"Removed bot {pid} from monitoring")
        
        self.update_process_table()

    def start_selected_bots(self):
        """Start selected bots by their BAT files"""
        selected = self.process_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select one or more bots to start")
            return
            
        for item in selected:
            values = self.process_tree.item(item, 'values')
            pid = values[0]
            
            # Check if this is a detected process (can't restart those directly)
            if values[-1] == "Detected":
                messagebox.showwarning("Cannot Start", 
                    f"Bot {pid} was auto-detected and cannot be started directly. "
                    "Please use the BAT Generator to create a starter script.")
                continue
            
            # Get the BAT file path from the process info
            process_info = self.bot_manager.bot_processes.get(pid)
            if not process_info or "bat_file" not in process_info:
                messagebox.showwarning("No BAT File", 
                    f"No BAT file information found for bot {pid}")
                continue
                
            # Start the bot
            success, result = self.bot_manager.start_bot(
                process_info["bat_file"], 
                process_info["profile"])
                
            if not success:
                messagebox.showerror("Start Failed", 
                    f"Failed to start bot {pid}: {result}")
        
        self.update_process_table()

    def stop_selected_bots(self):
        """Stop selected bots"""
        selected = self.process_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select one or more bots to stop")
            return
            
        for item in selected:
            pid = self.process_tree.item(item, 'values')[0]
            success, result = self.bot_manager.stop_bot(pid, manual_stop=True)
            
            if not success:
                messagebox.showerror("Stop Failed", 
                    f"Failed to stop bot {pid}: {result}")
        
        self.update_process_table()

    def restart_selected_bots(self):
        """Restart selected bots"""
        selected = self.process_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select one or more bots to restart")
            return
            
        for item in selected:
            pid = self.process_tree.item(item, 'values')[0]
            success, result = self.bot_manager.restart_bot(pid)
            
            if not success:
                messagebox.showerror("Restart Failed", 
                    f"Failed to restart bot {pid}: {result}")
            else:
                self.status_var.set(f"Restarted bot {pid}")
        
        self.update_process_table()

    def toggle_anti_crash(self):
        """Toggle anti-crash for selected bots"""
        selected = self.process_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select one or more bots to toggle anti-crash")
            return
            
        for item in selected:
            pid = self.process_tree.item(item, 'values')[0]
            success, result = self.bot_manager.toggle_anti_crash(pid)
            
            if not success:
                messagebox.showerror("Toggle Failed", 
                    f"Failed to toggle anti-crash for bot {pid}: {result}")
            else:
                self.status_var.set(f"Anti-crash {result.split()[-1]} for bot {pid}")
        
        self.update_process_table()

    def add_test_proxy(self):
        """Add a proxy to the test list"""
        proxy = self.proxy_test_var.get()
        if proxy:
            # Add to both the treeview and the tester
            self.proxy_tree.insert('', 'end', values=(proxy, "Not Tested", "", "", ""))
            self.proxy_tester.add_proxy(proxy)
            self.proxy_test_var.set("")

    def test_proxies(self):
        """Test all proxies in the list using the ProxyTester component"""
        test_url = self.test_url_var.get()
        if not test_url:
            messagebox.showwarning("Missing URL", "Please enter a test URL")
            return
        
        # Clear previous results
        for item in self.proxy_tree.get_children():
            self.proxy_tree.delete(item)
        
        # Get all proxies from the tester
        proxies_to_test = [proxy["proxy"] for proxy in self.proxy_tester.proxies]
        
        if not proxies_to_test:
            messagebox.showwarning("No Proxies", "No proxies to test")
            return
        
        # Test all proxies
        self.status_var.set("Testing proxies...")
        self.root.update()  # Force UI update
        
        results = self.proxy_tester.test_all_proxies(test_url)
        
        # Update the treeview with results
        for proxy, data in results.items():
            status = data['status'].capitalize()
            response_time = f"{data.get('response_time', 0):.2f} ms" if data.get('response_time') else "N/A"
            success_rate = f"{data.get('success_rate', 0)}%" if data.get('success_rate') is not None else "N/A"
            last_tested = data.get('last_tested', 'Never')
            
            self.proxy_tree.insert('', 'end', values=(
                proxy,
                status,
                response_time,
                success_rate,
                last_tested
            ))
        
        self.status_var.set("Proxy testing completed")
        messagebox.showinfo("Test Complete", f"Finished testing {len(results)} proxies")

    def update_bot_list(self):
        """Refresh the list of bot profiles"""
        for item in self.bot_list_tree.get_children():
            self.bot_list_tree.delete(item)
        
        profiles = self.bot_list_generator.get_all_profiles()
        for name, data in profiles.items():
            self.bot_list_tree.insert('', 'end', values=(
                name,
                data['proxy'],
                data['status'],
                "Yes" if data['banned'] else "No",
                data['last_active'] or "Never"
            ))

    def add_bot_profile(self):
        name = self.profile_name_var.get()
        proxy = self.profile_proxy_var.get()
        notes = self.profile_notes_var.get()
        
        if not name or not proxy:
            messagebox.showwarning("Missing Info", "Profile name and proxy are required")
            return
        
        if self.bot_list_generator.add_profile(name, proxy, notes):
            messagebox.showinfo("Success", "Profile added successfully")
            self.profile_name_var.set("")
            self.profile_proxy_var.set("")
            self.profile_notes_var.set("")
            self.update_bot_list()
        else:
            messagebox.showerror("Error", "Profile already exists")

    def update_bot_status(self):
        selected = self.bot_list_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select a profile first")
            return
        
        profile_name = self.bot_list_tree.item(selected[0], 'values')[0]
        
        # Create a dialog to update status
        status_dialog = tk.Toplevel(self.root)
        status_dialog.title(f"Update Status - {profile_name}")
        
        ttk.Label(status_dialog, text="New Status:").pack(pady=(10, 0))
        status_var = tk.StringVar(value="active")
        ttk.Combobox(status_dialog, textvariable=status_var, 
                    values=["active", "inactive", "building", "error"], state="readonly").pack()
        
        ttk.Label(status_dialog, text="Is Banned?").pack(pady=(10, 0))
        banned_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(status_dialog, variable=banned_var).pack()
        
        ttk.Label(status_dialog, text="Issue (if any):").pack(pady=(10, 0))
        issue_var = tk.StringVar()
        ttk.Entry(status_dialog, textvariable=issue_var).pack()
        
        def submit():
            self.bot_list_generator.update_status(
                profile_name,
                status_var.get(),
                banned_var.get(),
                issue_var.get()
            )
            status_dialog.destroy()
            self.update_bot_list()
            messagebox.showinfo("Success", "Status updated successfully")
        
        ttk.Button(status_dialog, text="Update", command=submit).pack(pady=10)

    def show_status_dialog(self):
        """Wrapper method to show status dialog (used by context menu)"""
        self.update_bot_status()

    def export_bot_list(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            if self.bot_list_generator.export_to_csv(file_path):
                messagebox.showinfo("Success", f"Bot list exported to {file_path}")
            else:
                messagebox.showerror("Error", "Failed to export bot list")

    def browse_jar(self):
        path = filedialog.askopenfilename(filetypes=[("Java Archives", "*.jar")])
        if path:
            self.jar_path_var.set(path)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.save_folder_var.set(folder)

    def add_manual_entry(self):
        proxy = self.proxy_var.get()
        profile = self.profile_var.get()
        if all([proxy, profile]):
            self.entries.append((proxy, profile))
            self.tree.insert('', 'end', values=(proxy, profile))
            self.proxy_var.set("")
            self.profile_var.set("")
        else:
            messagebox.showwarning("Missing Input", "Both proxy and profile are required.")

    def edit_selected(self):
        selected = self.tree.selection()
        if selected:
            item = selected[0]
            proxy, profile = self.tree.item(item, "values")

            new_proxy = self.simple_input("Edit Proxy", "Enter new proxy:", proxy)
            new_profile = self.simple_input("Edit Profile", "Enter new profile name:", profile)

            if new_proxy and new_profile:
                self.tree.item(item, values=(new_proxy, new_profile))
                index = self.tree.index(item)
                self.entries[index] = (new_proxy, new_profile)

    def delete_selected(self):
        selected = self.tree.selection()
        for item in selected:
            index = self.tree.index(item)
            self.tree.delete(item)
            del self.entries[index]

    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path:
            with open(path, newline='') as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    if len(row) >= 2:
                        self.entries.append((row[0], row[1]))
                        self.tree.insert('', 'end', values=(row[0], row[1]))

    def generate_bats(self):
        jar_path = self.jar_path_var.get()
        folder = self.save_folder_var.get()

        if not jar_path or not folder:
            messagebox.showwarning("Missing Paths", "Please set both the jar location and save folder.")
            return

        for i, (proxy, profile) in enumerate(self.entries):
            bat_content = f'@echo off\nstart /MIN javaw -jar "{jar_path}" -proxy={proxy} -proxy-type=socks5 -profile={profile}\n'
            bat_name = f"bot_{i+1}.bat"
            with open(os.path.join(folder, bat_name), 'w') as f:
                f.write(bat_content)

        messagebox.showinfo("Success", "BAT files generated successfully!")

    def simple_input(self, title, prompt, default):
        input_win = tk.Toplevel()
        input_win.title(title)
        input_win.geometry("300x100")
        input_win.transient()
        input_win.grab_set()

        var = tk.StringVar(value=default)
        ttk.Label(input_win, text=prompt).pack(pady=(10, 0))
        entry = ttk.Entry(input_win, textvariable=var)
        entry.pack(pady=5, padx=10, fill="x")
        entry.focus()

        result = []

        def submit():
            result.append(var.get())
            input_win.destroy()

        ttk.Button(input_win, text="OK", command=submit).pack(pady=(0, 10))
        input_win.wait_window()
        return result[0] if result else None

    def browse_target_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.target_folder_var.set(folder)

    def execute_batch_files(self):
        target_folder = self.target_folder_var.get()
        delay = self.delay_var.get()
        
        if not target_folder:
            messagebox.showwarning("Missing Folder", "Please select a target folder containing batch files.")
            return
        
        try:
            delay = int(delay)
        except ValueError:
            messagebox.showwarning("Invalid Delay", "Please enter a valid number for delay.")
            return
        
        self.batch_executor.execute_batch_files(target_folder, delay)
        self.status_var.set(f"Executing batch files from {target_folder} with {delay}s delay")

    def update_connection_list(self):
        connections = [conn['name'] for conn in self.cloud_integration.connections]
        self.connection_combo['values'] = connections
        if connections:
            self.connection_var.set(connections[0])

    def show_add_connection_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add New Connection")
        
        # Connection details
        ttk.Label(dialog, text="Name:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Label(dialog, text="Host:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        host_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=host_var).grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Label(dialog, text="Username:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        user_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=user_var).grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Label(dialog, text="Password:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        pass_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=pass_var, show="*").grid(row=3, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Label(dialog, text="Or SSH Key Path:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        key_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=key_var).grid(row=4, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Label(dialog, text="Port:").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        port_var = tk.StringVar(value="22")
        ttk.Entry(dialog, textvariable=port_var).grid(row=5, column=1, sticky="ew", padx=5, pady=5)
        
        def add_connection():
            try:
                success, message = self.cloud_integration.add_connection(
                    name_var.get(),
                    host_var.get(),
                    user_var.get(),
                    pass_var.get() or None,
                    key_var.get() or None,
                    int(port_var.get())
                )
                if success:
                    self.update_connection_list()
                    dialog.destroy()
                messagebox.showinfo("Result", message)
            except ValueError:
                messagebox.showerror("Error", "Port must be a number")
        
        ttk.Button(dialog, text="Add", command=add_connection).grid(row=6, column=0, columnspan=2, pady=10)

    def cloud_connect(self):
        conn_name = self.connection_var.get()
        if not conn_name:
            messagebox.showwarning("No Selection", "Please select a connection first")
            return
        
        success, message = self.cloud_integration.connect(conn_name)
        if success:
            self.cloud_status_text.insert(tk.END, f"Connected to {conn_name}\n")
        else:
            messagebox.showerror("Connection Failed", message)

    def cloud_disconnect(self):
        success, message = self.cloud_integration.disconnect()
        if success:
            self.cloud_status_text.insert(tk.END, "Disconnected\n")
        else:
            messagebox.showerror("Disconnect Failed", message)

    def get_cloud_status(self):
        success, result = self.cloud_integration.get_server_status()
        self.cloud_status_text.delete(1.0, tk.END)
        self.cloud_status_text.insert(tk.END, result)
        if not success:
            messagebox.showerror("Error", result)

    def cloud_upload(self):
        local_file = filedialog.askopenfilename(title="Select file to upload")
        if not local_file:
            return
        
        remote_file = filedialog.asksaveasfilename(title="Select destination on server")
        if not remote_file:
            return
        
        success, message = self.cloud_integration.upload_file(local_file, remote_file)
        if success:
            self.cloud_status_text.insert(tk.END, f"Upload successful: {message}\n")
        else:
            messagebox.showerror("Upload Failed", message)

    def cloud_download(self):
        remote_file = filedialog.askopenfilename(title="Select file on server")
        if not remote_file:
            return
        
        local_file = filedialog.asksaveasfilename(title="Select local destination")
        if not local_file:
            return
        
        success, message = self.cloud_integration.download_file(remote_file, local_file)
        if success:
            self.cloud_status_text.insert(tk.END, f"Download successful: {message}\n")
        else:
            messagebox.showerror("Download Failed", message)

    def cloud_execute(self):
        command = simpledialog.askstring("Execute Command", "Enter command to execute:")
        if command:
            success, result = self.cloud_integration.execute_command(command)
            self.cloud_status_text.insert(tk.END, f"$ {command}\n{result}\n")
            if not success:
                messagebox.showerror("Command Failed", result)


def main():
    root = tk.Tk()
    try:
        app = MicrobotManagerApp(root)
        root.mainloop()
    except Exception as e:
        with open("error.log", "w") as f:
            f.write(f"Error: {str(e)}\n")
            f.write(traceback.format_exc())
        messagebox.showerror("Critical Error", f"Application crashed. See error.log for details.\n{str(e)}")


if __name__ == "__main__":
    main()
