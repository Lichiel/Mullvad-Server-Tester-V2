# tester_gui.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time
import platform
import os
import csv
import json
import pickle
import subprocess
import logging
from threading import Event
from typing import Optional, List, Dict, Any, Set, Tuple

# --- SV-TTK Import ---
try:
    import sv_ttk
except ImportError:
    sv_ttk = None
    print("WARNING: sv-ttk library not found. Falling back to default ttk theme. Install using: pip install sv-ttk")

# Setup logger for this module
logger = logging.getLogger(__name__)

# --- Import shared modules ---
try:
    from mullvad_api import (load_cached_servers, set_mullvad_location,
                             set_mullvad_protocol, connect_mullvad,
                             disconnect_mullvad, get_mullvad_status, MullvadCLIError)
    # Adapt server_manager imports if needed, maybe just need server fetching
    from server_manager import (get_all_servers, get_servers_by_country,
                                export_to_csv, calculate_latency_color, # Keep color helpers
                                calculate_speed_color) # Keep color helpers
    from config import (load_config, save_config, get_cache_path, get_log_path,
                       get_default_cache_path)
except ImportError as e:
    logger.exception("Failed to import necessary shared modules (mullvad_api, server_manager, config).")
    # Show error immediately if GUI can't even start
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Import Error", f"Failed to import shared modules: {e}\nPlease ensure required files are accessible.")
    exit()

# --- Constants ---
CHECKBOX_UNCHECKED = "☐"
CHECKBOX_CHECKED = "☑"
APP_VERSION = "1.0.0" # Define app version

# --- Helper Functions ---
def get_flag_emoji(country_code: str) -> str:
    """Converts a two-letter country code (ISO 3166-1 alpha-2) to a flag emoji."""
    if not country_code or len(country_code) != 2: return ""
    try:
        offset = 0x1F1E6 - ord('A')
        point1 = chr(ord(country_code[0].upper()) + offset)
        point2 = chr(ord(country_code[1].upper()) + offset)
        return point1 + point2
    except Exception:
        logger.warning(f"Could not generate flag emoji for code: {country_code}")
        return ""

# --- Helper Classes (Reuse LoadingAnimation) ---
class LoadingAnimation:
    """Class to handle loading animation in the status bar."""
    def __init__(self, label_var: tk.StringVar, original_text: str, animation_frames: Optional[List[str]] = None):
        self.label_var = label_var
        self.original_text = original_text
        self.animation_frames = animation_frames or ['⣾', '⣽', '⣻', '⢿', '⡿', '⣟', '⣯', '⣷']
        self.is_running = False
        self.current_frame = 0
        self.after_id: Optional[str] = None
        self.root: Optional[tk.Tk] = None

    def start(self, root: tk.Tk):
        if self.is_running: return
        self.root = root
        self.is_running = True
        self.current_frame = 0
        self.animate()
        logger.debug("Loading animation started.")

    def stop(self):
        if not self.is_running: return
        self.is_running = False
        if self.after_id and self.root:
            try: self.root.after_cancel(self.after_id)
            except Exception: pass
            self.after_id = None
        try: self.label_var.set(self.original_text)
        except Exception: pass
        logger.debug("Loading animation stopped.")

    def animate(self):
        if not self.is_running or not self.root: return
        try:
            frame = self.animation_frames[self.current_frame]
            self.label_var.set(f"{self.original_text} {frame}")
            self.current_frame = (self.current_frame + 1) % len(self.animation_frames)
            self.after_id = self.root.after(150, self.animate)
        except Exception as e:
             logger.error(f"Error during animation cycle: {e}")
             self.stop()

    def update_text(self, new_text: str):
        self.original_text = new_text
        if not self.is_running:
             try: self.label_var.set(self.original_text)
             except Exception: pass

# --- Main Application Class ---
class TesterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Mullvad Automated Tester")
        # self.root.geometry("950x650") # Set in main script

        logger.info("Initializing TesterApp...")
        self._setup_icon() # Reuse icon setup if icon file exists
        self.config = load_config()
        self.theme_colors: Dict[str, str] = {}
        self.main_frame: Optional[ttk.Frame] = None
        self.default_font: Tuple[str, int] = ('Helvetica', 13) # Default fallback

        # --- State Variables ---
        self.server_data: Optional[Dict[str, Any]] = None # Full data from mullvad cache
        self.all_servers_list: List[Dict[str, Any]] = [] # Flattened list of all servers
        self.countries: List[Dict[str, str]] = []
        self.current_operation = tk.StringVar(value="Ready")
        self.test_in_progress = False
        self.sort_column = self.config.get("tester_sort_column", "hostname") # Use different config keys
        self.sort_order = self.config.get("tester_sort_order", "ascending")
        self.selected_server_items: Set[str] = set() # Stores item IDs of checked servers
        self.theme_var = tk.StringVar(value=self.config.get("theme_mode", "system"))

        # --- Thread Control ---
        self.stop_event = Event()
        self.pause_event = Event() # Add pause event
        self.test_thread: Optional[threading.Thread] = None

        # --- UI Elements ---
        self.server_tree: Optional[ttk.Treeview] = None
        self.start_test_button: Optional[ttk.Button] = None
        self.stop_test_button: Optional[ttk.Button] = None
        self.progress_bar: Optional[ttk.Progressbar] = None
        self.progress_var = tk.DoubleVar()
        self.operation_label: Optional[ttk.Label] = None
        self.pause_button: Optional[ttk.Button] = None # Add pause button UI element
        self.control_frame: Optional[ttk.Frame] = None # Frame for pause/stop
        self.status_label: Optional[ttk.Label] = None # For Mullvad connection status

        # --- Other ---
        self.loading_animation = LoadingAnimation(self.current_operation, "Ready")
        self.created_cell_tags: Set[str] = set() # Track dynamic tags for cell colors
        self.country_var = tk.StringVar() # Add country_var initialization here
        self.country_combo: Optional[ttk.Combobox] = None # Add country_combo initialization

        # --- Build UI First ---
        self.create_menu()
        self.create_ui()

        # --- Apply Theme *After* UI is Built ---
        self.apply_theme()

        # --- Post-UI Setup ---
        self.load_server_data() # Initial data load
        # Restore last selected country after loading data and populating combo
        self._restore_last_country()

        logger.info("TesterApp initialization complete.")

    # --- UI Creation Helpers ---
    def _setup_icon(self):
        """Sets the application icon if available (looks for mullvad_icon.*)."""
        # This is identical to the example app's icon setup
        icon_name = 'mullvad_icon'
        icon_path = None
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            # Check in current dir and 'example' subdir if it exists
            potential_dirs = [script_dir, os.path.join(os.path.dirname(script_dir), 'example')]
            for check_dir in potential_dirs:
                 if not os.path.isdir(check_dir): continue
                 potential_paths = [
                     os.path.join(check_dir, f"{icon_name}.ico"),
                     os.path.join(check_dir, f"{icon_name}.png"),
                     os.path.join(check_dir, f"{icon_name}.icns")
                 ]
                 for p in potential_paths:
                     if os.path.exists(p):
                         icon_path = p
                         break
                 if icon_path: break

            if not icon_path:
                 logger.warning("Application icon file 'mullvad_icon.*' not found.")
                 return

            if platform.system() == 'Windows' and icon_path.endswith(".ico"):
                self.root.iconbitmap(icon_path)
                logger.info(f"Set Windows icon: {icon_path}")
            else:
                 try:
                     img = tk.PhotoImage(file=icon_path)
                     self.root.iconphoto(True, img)
                     logger.info(f"Set icon using PhotoImage: {icon_path}")
                 except tk.TclError:
                     logger.warning(f"Could not set icon {icon_path} using PhotoImage.")

        except Exception as e:
            logger.error(f"Error setting application icon: {e}")

    def create_menu(self):
        """Create the application menu bar."""
        menubar = tk.Menu(self.root)

        # File Menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Reload Server Data", command=self.load_server_data)
        file_menu.add_command(label="Export Results to CSV...", command=self.export_results_to_csv, accelerator="Ctrl+E")
        file_menu.add_separator()
        file_menu.add_command(label="Settings", command=self.open_settings)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        # Test Menu
        test_menu = tk.Menu(menubar, tearoff=0)
        test_menu.add_command(label="Start Test", command=self.start_test_thread, accelerator="Ctrl+T")
        test_menu.add_command(label="Stop Test", command=self.stop_tests, accelerator="Ctrl+X")
        menubar.add_cascade(label="Test", menu=test_menu)
        self.root.bind_all("<Control-t>", lambda e: self.start_test_thread())
        self.root.bind_all("<Control-x>", lambda e: self.stop_tests())

        # View Menu (Theme only for now)
        view_menu = tk.Menu(menubar, tearoff=0)
        theme_menu = tk.Menu(view_menu, tearoff=0)
        theme_menu.add_radiobutton(label="System", variable=self.theme_var, value="system", command=self.change_theme)
        theme_menu.add_radiobutton(label="Light", variable=self.theme_var, value="light", command=self.change_theme)
        theme_menu.add_radiobutton(label="Dark", variable=self.theme_var, value="dark", command=self.change_theme)
        view_menu.add_cascade(label="Theme", menu=theme_menu)
        menubar.add_cascade(label="View", menu=view_menu)

        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)
        # Bind accelerator for export
        self.root.bind_all("<Control-e>", lambda e: self.export_results_to_csv())


    def _create_top_frame(self, parent: ttk.Frame):
        """Creates the top frame with controls."""
        top_frame = ttk.Frame(parent)
        # Increased bottom padding and horizontal padding
        top_frame.pack(fill=tk.X, pady=(0, 15), padx=10)

        # Action buttons with increased spacing
        self.start_test_button = ttk.Button(top_frame, text="Start Test", command=self.start_test_thread)
        self.start_test_button.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_test_button = ttk.Button(top_frame, text="Stop Test", command=self.stop_tests, state=tk.DISABLED)
        self.stop_test_button.pack(side=tk.LEFT, padx=(0, 15)) # More space after stop

        # --- Add Country Filter ---
        ttk.Label(top_frame, text="Country:").pack(side=tk.LEFT, padx=(15, 5)) # More space before label
        self.country_var = tk.StringVar()
        self.country_combo = ttk.Combobox(top_frame, textvariable=self.country_var, width=20, state="readonly")
        self.country_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.country_combo.bind("<<ComboboxSelected>>", self.on_country_selected)
        # --- End Add Country Filter ---

        # Status display (aligned right) - Optional, maybe just use operation label
        # status_frame = ttk.Frame(top_frame)
        # status_frame.pack(side=tk.RIGHT, padx=(10, 0))
        # ttk.Label(status_frame, text="Mullvad Status:").pack(side=tk.LEFT)
        # self.status_label = ttk.Label(status_frame, text="Unknown", anchor=tk.W) # Placeholder
        # self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        # self.update_mullvad_status() # Start status polling

    def _create_middle_frame(self, parent: ttk.Frame):
        """Creates the middle frame with the server list Treeview."""
        middle_frame = ttk.Frame(parent)
        # Increased vertical padding and horizontal padding
        middle_frame.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)

        # Define columns including checkbox and new results
        columns = ("selected", "hostname", "city", "country", "status", "ping_ms", "dl_mbps", "ul_mbps")
        self.server_tree = ttk.Treeview(middle_frame, columns=columns, show="headings")

        # Define headings
        self.server_tree.heading("selected", text=CHECKBOX_UNCHECKED, anchor=tk.CENTER,
                                 command=lambda: self._toggle_all_checkboxes())
        self.server_tree.heading("hostname", text="Hostname", anchor=tk.W, command=lambda: self.sort_treeview("hostname"))
        self.server_tree.heading("city", text="City", anchor=tk.W, command=lambda: self.sort_treeview("city"))
        self.server_tree.heading("country", text="Country", anchor=tk.W, command=lambda: self.sort_treeview("country"))
        self.server_tree.heading("status", text="Status", anchor=tk.W, command=lambda: self.sort_treeview("status"))
        self.server_tree.heading("ping_ms", text="Ping (ms)", anchor=tk.E, command=lambda: self.sort_treeview("ping_ms"))
        self.server_tree.heading("dl_mbps", text="DL (Mbps)", anchor=tk.E, command=lambda: self.sort_treeview("dl_mbps"))
        self.server_tree.heading("ul_mbps", text="UL (Mbps)", anchor=tk.E, command=lambda: self.sort_treeview("ul_mbps"))

        # Define column properties
        self.server_tree.column("selected", width=40, minwidth=40, stretch=tk.NO, anchor=tk.CENTER)
        self.server_tree.column("hostname", width=200, stretch=tk.YES, anchor=tk.W)
        self.server_tree.column("city", width=120, stretch=tk.YES, anchor=tk.W)
        self.server_tree.column("country", width=120, stretch=tk.YES, anchor=tk.W)
        self.server_tree.column("status", width=130, stretch=tk.NO, anchor=tk.W)
        self.server_tree.column("ping_ms", width=80, stretch=tk.NO, anchor=tk.E)
        self.server_tree.column("dl_mbps", width=80, stretch=tk.NO, anchor=tk.E)
        self.server_tree.column("ul_mbps", width=80, stretch=tk.NO, anchor=tk.E)

        # Scrollbars
        tree_scroll_y = ttk.Scrollbar(middle_frame, orient=tk.VERTICAL, command=self.server_tree.yview)
        tree_scroll_x = ttk.Scrollbar(middle_frame, orient=tk.HORIZONTAL, command=self.server_tree.xview)
        self.server_tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        # Layout with grid
        middle_frame.grid_rowconfigure(0, weight=1)
        middle_frame.grid_columnconfigure(0, weight=1)
        self.server_tree.grid(row=0, column=0, sticky='nsew')
        tree_scroll_y.grid(row=0, column=1, sticky='ns')
        tree_scroll_x.grid(row=1, column=0, sticky='ew')

        # Configure tags for row colors
        self.server_tree.tag_configure('odd_row', background=self.theme_colors.get("row_odd", "#F8F8F8"))
        self.server_tree.tag_configure('even_row', background=self.theme_colors.get("row_even", "#FFFFFF"))
        self.server_tree.tag_configure('status_error', foreground='red') # Example status tag
        self.server_tree.tag_configure('status_timeout', foreground='orange')
        self.server_tree.tag_configure('status_completed', foreground='green')

        # Bind click event for checkbox toggling
        self.server_tree.bind("<Button-1>", self._on_tree_click)
        # Bind double-click to connect
        self.server_tree.bind("<Double-1>", self._connect_on_double_click)

    def _create_bottom_frame(self, parent: ttk.Frame):
        """Creates the bottom frame with progress bar and status."""
        bottom_frame = ttk.Frame(parent)
        # Increased top padding and horizontal padding
        bottom_frame.pack(fill=tk.X, pady=(15, 0), padx=10)

        # Operation label (with animation) - Expand to take available space
        self.operation_label = ttk.Label(bottom_frame, textvariable=self.current_operation, anchor=tk.W)
        self.operation_label.pack(side=tk.LEFT, padx=(0, 15), fill=tk.X, expand=True) # Increased padding

        # Control buttons (Pause/Stop) - initially hidden
        self.control_frame = ttk.Frame(bottom_frame)
        # Don't pack yet, pack when needed

        self.pause_button = ttk.Button(self.control_frame, text="Pause", command=self.pause_resume_test)
        self.pause_button.pack(side=tk.LEFT, padx=(0, 8)) # Increased padding
        # Note: Stop button is already created in _create_top_frame, move it here?
        # For now, keep stop button in top frame, just add pause here.

        # Progress bar (aligned right)
        self.progress_bar = ttk.Progressbar(bottom_frame, variable=self.progress_var, mode="determinate", length=200)
        self.progress_bar.pack(side=tk.RIGHT, padx=(10, 0))

    def _configure_styles(self):
        """Configure default ttk styles like font."""
        style = ttk.Style()
        # Determine default font based on OS
        if platform.system() == "Darwin": # macOS
            # Try common names for San Francisco font
            # Tkinter might map 'system' or 'SystemFont' correctly too
            # Using '.SF NS Text' is more specific but might fail if name changes
            # Let's try 'system' first as it's often mapped by Tk
            try:
                # Test if 'system' font gives something reasonable
                # You could potentially check ttk.Style().lookup('.', 'font') after setting
                self.default_font = ('system', 13)
                # Or try a known name:
                # self.default_font = ('.SF NS Text', 13)
                style.configure('.', font=self.default_font)
                logger.info(f"Attempting to set macOS system font: {self.default_font}")
            except tk.TclError:
                logger.warning("Failed to set preferred macOS font, falling back.")
                self.default_font = ('Helvetica', 13)
                style.configure('.', font=self.default_font)
        elif platform.system() == "Windows":
            self.default_font = ('Segoe UI', 10) # Standard Windows font
            style.configure('.', font=self.default_font)
        else: # Linux/Other
            self.default_font = ('DejaVu Sans', 10) # Common Linux default
            style.configure('.', font=self.default_font)

        # Configure Treeview heading font (bold)
        try:
            heading_font = (self.default_font[0], self.default_font[1], 'bold')
            style.configure("Treeview.Heading", font=heading_font)
        except Exception as e:
            logger.error(f"Failed to set Treeview heading font: {e}")

        logger.info(f"Default widget font set to: {self.default_font}")


    def create_ui(self):
        """Create the main user interface."""
        self._configure_styles() # Set styles (like font) first
        self.main_frame = ttk.Frame(self.root, padding=15) # Increased main padding
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self._create_top_frame(self.main_frame)
        self._create_middle_frame(self.main_frame)
        self._create_bottom_frame(self.main_frame)

    # --- Event Handlers & UI Logic ---

    def _on_tree_click(self, event):
        """Handle clicks on the Treeview for checkbox toggling."""
        if not self.server_tree: return
        region = self.server_tree.identify_region(event.x, event.y)
        if region == "cell":
            column_id = self.server_tree.identify_column(event.x)
            item_id = self.server_tree.identify_row(event.y)
            if column_id == "#1" and item_id: # Checkbox column
                 self._toggle_checkbox(item_id)

    def _toggle_checkbox(self, item_id: str):
         """Toggles the checkbox state for a given item ID."""
         if not self.server_tree or not self.server_tree.exists(item_id): return
         current_value = self.server_tree.set(item_id, "#1")
         new_value = CHECKBOX_CHECKED if current_value == CHECKBOX_UNCHECKED else CHECKBOX_UNCHECKED
         self.server_tree.set(item_id, "#1", new_value)
         if new_value == CHECKBOX_CHECKED:
             self.selected_server_items.add(item_id)
         else:
             self.selected_server_items.discard(item_id)
         self._update_start_button_text()

    def _toggle_all_checkboxes(self):
        """Toggles all visible checkboxes."""
        if not self.server_tree: return
        all_items = self.server_tree.get_children('')
        if not all_items: return

        # Determine target state based on header or first item
        current_header = self.server_tree.heading("selected", "text")
        target_state = CHECKBOX_CHECKED if current_header == CHECKBOX_UNCHECKED else CHECKBOX_UNCHECKED
        self.server_tree.heading("selected", text=target_state)

        self.selected_server_items.clear()
        for item_id in all_items:
            if not self.server_tree.exists(item_id): continue
            self.server_tree.set(item_id, "#1", target_state)
            if target_state == CHECKBOX_CHECKED:
                self.selected_server_items.add(item_id)
        self._update_start_button_text()

    def _update_start_button_text(self):
        """Updates the 'Start Test' button text based on selections."""
        if not self.start_test_button: return
        num_selected = len(self.selected_server_items)
        if num_selected > 0:
            self.start_test_button.configure(text=f"Start Test ({num_selected} Selected)")
        else:
            self.start_test_button.configure(text="Start Test (All Visible)")

    # --- Data Loading and Display ---

    def load_server_data(self):
        """Load Mullvad server data and populate the list."""
        self.loading_animation.update_text("Loading server data...")
        self.loading_animation.start(self.root)
        self.root.update_idletasks()

        try:
            cache_path = get_cache_path(self.config)
            logger.info(f"Using cache path: {cache_path}")
            self.server_data = load_cached_servers(cache_path)

            if not self.server_data:
                messagebox.showerror("Error", f"Failed to load server data from {cache_path}.", parent=self.root)
                self.loading_animation.update_text("Error loading data")
                self.loading_animation.stop()
                return

            # Extract countries for potential filtering later
            self.countries = [{"code": c.get("code", ""), "name": c.get("name", "Unknown")}
                              for c in self.server_data.get("countries", [])]
            self.countries.sort(key=lambda x: x["name"])

            # Get flattened list of all servers
            # Pass None for protocol to get all initially
            self.all_servers_list = get_all_servers(self.server_data, protocol=None)
            logger.info(f"Loaded {len(self.all_servers_list)} total servers.")

            self.populate_server_list(self.all_servers_list) # Display all servers
                # --- Populate Country Combobox ---
            country_names = ["All Countries"] + sorted([f"{get_flag_emoji(c['code'])} {c['name']}" for c in self.countries])
            if self.country_combo:
                self.country_combo['values'] = country_names
                logger.info(f"Populated country combobox with {len(country_names)} entries.")
            # --- End Populate ---

            self.populate_server_list(self.all_servers_list) # Display all servers initially
            self.loading_animation.update_text(f"{len(self.all_servers_list)} servers loaded")


        except Exception as e:
            logger.exception("An error occurred during server data loading.")
            messagebox.showerror("Error", f"An unexpected error occurred loading server data: {e}", parent=self.root)
            self.loading_animation.update_text("Error loading data")
        finally:
             self.root.after(500, self.loading_animation.stop)

    def populate_server_list(self, servers_to_display: List[Dict[str, Any]]):
        """Populate the Treeview with the given list of servers."""
        if not self.server_tree: return

        logger.info(f"Populating Treeview with {len(servers_to_display)} servers.")
        self.server_tree.delete(*self.server_tree.get_children())
        self.selected_server_items.clear()
        self.server_tree.heading("selected", text=CHECKBOX_UNCHECKED)

        use_alt_colors = self.config.get("alternating_row_colors", True)
        for i, server in enumerate(servers_to_display):
            hostname = server.get("hostname", "N/A")
            city = server.get("city", "N/A")
            country_name = server.get("country", "N/A")
            country_code = server.get("country_code", "") # Assumes get_all_servers adds this
            country_display = f"{get_flag_emoji(country_code)} {country_name}" if country_code else country_name

            tags = ['odd_row' if i % 2 else 'even_row'] if use_alt_colors else []

            # Initial values: checkbox, hostname, city, country, status=Pending, ping="", dl="", ul=""
            self.server_tree.insert("", tk.END, iid=hostname, values=( # Use hostname as item ID
                CHECKBOX_UNCHECKED, hostname, city, country_display, "Pending", "", "", ""
            ), tags=tags)

        # Apply initial sort and update button text
        self.sort_treeview(self.sort_column, force_order=self.sort_order)
        self._update_start_button_text()

    def _restore_last_country(self):
        """Sets the country combobox to the last selected value from config."""
        last_country_code = self.config.get("tester_last_country", "") # Use specific key
        selected_display_name = "All Countries" # Default
        if last_country_code:
            for country in self.countries:
                if country["code"] == last_country_code:
                    selected_display_name = f"{get_flag_emoji(country['code'])} {country['name']}"
                    break
        self.country_var.set(selected_display_name)
        logger.info(f"Restored last selected country: {selected_display_name}")
        # Trigger filtering based on the restored selection
        self.on_country_selected()


    def on_country_selected(self, event=None):
        """Handle country selection change and filter the server list."""
        selected_display_name = self.country_var.get()
        logger.info(f"Country selected: {selected_display_name}")

        servers_to_show = []
        country_code_to_save = "" # Default to empty for "All Countries"

        if selected_display_name == "All Countries":
            servers_to_show = self.all_servers_list # Show all loaded servers
        else:
            # Extract country code from display name (flag + space + name)
            country_name_only = selected_display_name.split(" ", 1)[-1]
            country_code = next((c["code"] for c in self.countries if c["name"] == country_name_only), None)

            if country_code:
                country_code_to_save = country_code
                # Filter the main list
                servers_to_show = [s for s in self.all_servers_list if s.get("country_code") == country_code]
                logger.info(f"Filtering to show {len(servers_to_show)} servers for {country_code}")
            else:
                logger.error(f"Could not find country code for selected name: {country_name_only}")
                servers_to_show = [] # Show empty list on error

        # Save the selected country code (or "" for All)
        self.config["tester_last_country"] = country_code_to_save
        save_config(self.config)

        # Repopulate the treeview with the filtered list
        self.populate_server_list(servers_to_show)


    def sort_treeview(self, column: str, force_order: Optional[str] = None):
        """Sort the treeview by the specified column."""
        if not self.server_tree: return
        logger.debug(f"Sorting Treeview by column '{column}', force_order='{force_order}'")

        # Determine new sort order
        if column == self.sort_column and not force_order:
            self.sort_order = "descending" if self.sort_order == "ascending" else "ascending"
        else:
            self.sort_column = column
            self.sort_order = force_order if force_order else "ascending"

        # Save sort preference with app-specific keys
        self.config["tester_sort_column"] = self.sort_column
        self.config["tester_sort_order"] = self.sort_order
        # save_config(self.config) # Maybe save on exit instead

        items = [(self.server_tree.set(item_id, column), item_id) for item_id in self.server_tree.get_children('')]

        # Define conversion logic
        def get_sort_key(value_str: str) -> Any:
            if column in ['ping_ms', 'dl_mbps', 'ul_mbps']:
                if not value_str or value_str in ["Timeout", "Error", "N/A", ""]:
                    return float('inf')
                try: return float(value_str)
                except ValueError: return float('inf')
            elif column == 'selected':
                 return 0 if value_str == CHECKBOX_CHECKED else 1
            elif column == 'status': # Define order for status
                 order = {"Pending": 0, "Connecting": 1, "Pinging": 2, "Speed Testing": 3, "Disconnecting": 4, "Completed": 5, "Timeout": 6, "Error": 7}
                 return order.get(value_str.split()[0], 99) # Sort by first word of status
            else: # String columns (strip flag for country)
                if column == 'country' and len(value_str.split(" ", 1)) > 1:
                     value_str = value_str.split(" ", 1)[-1]
                return str(value_str).lower()

        try: items.sort(key=lambda x: get_sort_key(x[0]), reverse=(self.sort_order == "descending"))
        except Exception as e:
             logger.exception(f"Error during sorting prep for column {column}: {e}")
             return

        # Rearrange items
        use_alt_colors = self.config.get("alternating_row_colors", True)
        for index, (_, item_id) in enumerate(items):
            if not self.server_tree.exists(item_id): continue
            self.server_tree.move(item_id, '', index)
            if use_alt_colors:
                try:
                    current_tags = list(self.server_tree.item(item_id, "tags"))
                    filtered_tags = [tag for tag in current_tags if not tag.startswith(('odd_row', 'even_row'))]
                    row_tag = 'odd_row' if index % 2 else 'even_row'
                    filtered_tags.append(row_tag)
                    self.server_tree.item(item_id, tags=tuple(filtered_tags))
                except tk.TclError: continue

        logger.debug(f"Treeview sorted by {self.sort_column} {self.sort_order}.")

    # --- Testing Logic ---

    def start_test_thread(self):
        """Prepare server list and start the testing loop in a new thread."""
        if self.test_in_progress:
            messagebox.showwarning("Test in Progress", "A test is already running.", parent=self.root)
            return

        if not self.server_tree: return

        # Determine which servers to test based on checkboxes
        target_item_ids = list(self.selected_server_items)
        if not target_item_ids: # If nothing selected, test all visible
             target_item_ids = list(self.server_tree.get_children(''))
             logger.info("No servers selected via checkbox, testing all visible servers.")
             if not messagebox.askyesno("Confirm Test All", f"No servers selected. Test all {len(target_item_ids)} visible servers?", parent=self.root):
                 return
        else:
             logger.info(f"Testing {len(target_item_ids)} servers selected via checkbox.")

        if not target_item_ids:
            messagebox.showinfo("No Servers", "No servers found to test.", parent=self.root)
            return

        # Get full server details for the target items (use hostname as item ID)
        servers_to_test: List[Dict[str, Any]] = []
        for item_id_hostname in target_item_ids:
            # Find the server details from our loaded list
            server_details = next((s for s in self.all_servers_list if s.get("hostname") == item_id_hostname), None)
            if server_details:
                 servers_to_test.append(server_details.copy()) # Use a copy
            else:
                 logger.warning(f"Could not find server data for hostname '{item_id_hostname}'. Skipping.")

        if not servers_to_test:
            messagebox.showerror("Error", "Could not retrieve details for any servers to test.", parent=self.root)
            return

        # --- Start Test Process ---
        self.test_in_progress = True
        self.stop_event.clear()
        self.pause_event.clear() # Ensure pause is not set initially
        if self.start_test_button: self.start_test_button.configure(state=tk.DISABLED)
        if self.stop_test_button: self.stop_test_button.configure(state=tk.NORMAL)
        # Show and configure control frame (Pause button)
        if self.control_frame and self.pause_button:
            self.control_frame.pack(side=tk.LEFT, padx=(10, 0)) # Pack next to operation label
            self.pause_button.configure(text="Pause", state=tk.NORMAL)
        self.progress_var.set(0)
        self.loading_animation.update_text(f"Starting test on {len(servers_to_test)} servers...")
        self.loading_animation.start(self.root)

        # Clear previous results for the servers being tested
        for server in servers_to_test:
             self.update_server_status(server['hostname'], "Pending", clear_results=True)

        # Launch the test loop in a daemon thread
        self.test_thread = threading.Thread(target=self.run_test_loop, args=(servers_to_test,), daemon=True)
        self.test_thread.start()

    def stop_tests(self):
        """Signal the testing thread to stop."""
        if not self.test_in_progress:
            logger.info("Stop requested but no test running.")
            return

        logger.info("Stop requested. Signaling test thread...")
        self.stop_event.set()
        self.loading_animation.update_text("Stopping tests...")
        if self.stop_test_button: self.stop_test_button.configure(state=tk.DISABLED)
        # The thread's finally block will handle full cleanup

    def _test_cleanup(self):
        """Reset UI state after tests finish or are stopped."""
        logger.debug("Running test cleanup...")
        self.test_in_progress = False
        self.loading_animation.stop()
        if self.start_test_button: self.start_test_button.configure(state=tk.NORMAL)
        if self.stop_test_button: self.stop_test_button.configure(state=tk.DISABLED)
        # Hide control frame (pause button)
        if self.control_frame: self.control_frame.pack_forget()
        self.progress_var.set(0)
        self.test_thread = None

    def pause_resume_test(self):
        """Pause or resume the current test."""
        if not self.test_in_progress: return

        if self.pause_event.is_set():
            self.pause_event.clear()
            if self.pause_button: self.pause_button.configure(text="Pause")
            self.loading_animation.update_text("Resuming test...") # Keep animation running
            logger.info("Test resumed.")
        else:
            self.pause_event.set()
            if self.pause_button: self.pause_button.configure(text="Resume")
            self.loading_animation.update_text("Test paused") # Update text but keep animation
            logger.info("Test paused.")


    def run_test_loop(self, servers_to_run: List[Dict[str, Any]]):
        """The main loop executed in a separate thread to test servers."""
        total_servers = len(servers_to_run)
        logger.info(f"Test thread started for {total_servers} servers.")
        start_time = time.time()

        # --- Get Config ---
        ping_count = self.config.get("ping_count", 3)
        ping_timeout = self.config.get("timeout_seconds", 10) # Timeout for our ping subprocess
        speedtest_timeout = self.config.get("speedtest_timeout_seconds", 90) # Timeout for speedtest CLI

        try:
            for i, server in enumerate(servers_to_run):
                hostname = server.get("hostname", "N/A")
                country_code = server.get("country_code", "")
                city_code = server.get("city_code", "")
                city_name = server.get("city", "Unknown") # Get city name
                country_name = server.get("country", "Unknown") # Get country name
                flag = get_flag_emoji(country_code)
                # Update operation label with detailed info
                op_text_base = f"Testing {i+1}/{total_servers}: {flag} {city_name} ({hostname})"
                self.root.after(0, lambda txt=f"{op_text_base} - Starting...": self.loading_animation.update_text(txt))

                current_progress = (i / total_servers) * 100
                self.root.after(0, lambda p=current_progress: self.progress_var.set(p))

                # --- Pause/Stop Check ---
                while self.pause_event.is_set():
                    if self.stop_event.is_set(): break # Allow stop while paused
                    time.sleep(0.5)
                if self.stop_event.is_set():
                    self.update_server_status(hostname, "Skipped (Stopped)")
                    logger.info(f"Test stopped before server {hostname}")
                    break # Exit loop

                logger.info(f"--- Testing server {i+1}/{total_servers}: {hostname} ---")

                # 1. Connect
                connect_success = False
                connect_error = None
                try:
                    self.root.after(0, lambda txt=f"{op_text_base} - Connecting...": self.loading_animation.update_text(txt))
                    self.update_server_status(hostname, f"Connecting...")
                    if not country_code or not city_code:
                         raise MullvadCLIError(f"Missing country/city code for {hostname}")

                    # Determine protocol (needed for set_location?) - Assume WireGuard default for now
                    # protocol = "wireguard" if "-wg" in hostname.lower() or ".wg." in hostname.lower() else "openvpn"
                    # set_mullvad_protocol(protocol) # Setting protocol might not be needed if location sets it
                    # time.sleep(0.5)

                    set_mullvad_location(country_code, city_code, hostname)
                    time.sleep(1.0) # Give Mullvad time to apply location
                    connect_mullvad() # Run the connect command

                    # --- Verify Connection Status ---
                    connection_verified = False
                    verify_timeout = self.config.get("connection_verify_timeout", 15) # Use config value
                    verify_start_time = time.monotonic()
                    logger.info(f"Verifying connection to {hostname} (timeout: {verify_timeout}s)...")
                    self.root.after(0, lambda txt=f"{op_text_base} - Verifying...": self.loading_animation.update_text(txt))
                    self.update_server_status(hostname, "Verifying...") # Update status before verification loop
                    while time.monotonic() < verify_start_time + verify_timeout:
                        # --- Pause/Stop Check ---
                        while self.pause_event.is_set():
                            if self.stop_event.is_set(): break
                            time.sleep(0.5)
                        if self.stop_event.is_set(): break # Allow stopping during verification

                        try:
                            status_output = get_mullvad_status()
                            # Check if status simply contains "Connected"
                            if "Connected" in status_output:
                                connection_verified = True
                                logger.info(f"Connection verified (Status: '{status_output}').")
                                break
                            # Optional: Check for intermediate states if needed
                        except Exception as status_e:
                            logger.warning(f"Error getting status during verification: {status_e}")
                        time.sleep(0.5) # Check status every half second

                    if not connection_verified and not self.stop_event.is_set():
                        logger.error(f"Connection verification timed out for {hostname} after {verify_timeout}s.")
                        connect_error = "Con. Timeout"
                        # Attempt disconnect just in case it's stuck in a weird state
                        self._safe_disconnect(hostname)
                    elif connection_verified:
                         connect_success = True
                    # --- End Verification ---

                except MullvadCLIError as e:
                    logger.error(f"Connection command for {hostname} failed: {e}")
                    connect_error = f"Connect Error: {e}"
                except Exception as e:
                    logger.exception(f"Unexpected error connecting to {hostname}: {e}")
                    connect_error = f"Connect Error: Unexpected"

                # --- Pause/Stop Check ---
                while self.pause_event.is_set():
                    if self.stop_event.is_set(): break
                    time.sleep(0.5)
                if self.stop_event.is_set():
                    self.update_server_status(hostname, "Stopped during connect")
                    # Attempt disconnect if connection might have partially succeeded
                    if connect_success or "Connected" in get_mullvad_status(): self._safe_disconnect(hostname)
                    break

                # --- Check Connection Success Before Proceeding ---
                if not connect_success:
                    # Update status with connection error and skip to next server
                    final_status = connect_error or "Connect Failed" # Use specific error if available
                    self.update_server_status(hostname, final_status)
                    logger.warning(f"Skipping tests for {hostname} due to connection failure.")
                    # No need to disconnect if connection failed
                    continue # Move to the next server in the loop
                # --- End Check ---

                # 2. Run Tests (only if connected)
                ping_latency = None
                speed_results = {"ping": None, "download": None, "upload": None}
                test_error = None

                if connect_success:
                    # --- Pause/Stop Check ---
                    while self.pause_event.is_set():
                        if self.stop_event.is_set(): break
                        time.sleep(0.5)
                    if self.stop_event.is_set():
                        self.update_server_status(hostname, "Stopped before tests")
                        self._safe_disconnect(hostname)
                        break

                    # 2a. Ping Test (using system ping)
                    try:
                        self.root.after(0, lambda txt=f"{op_text_base} - Pinging...": self.loading_animation.update_text(txt))
                        self.update_server_status(hostname, f"Pinging ({ping_count}x)...")
                        # Ping a reliable external host instead of the server's IP
                        target_ip = "1.1.1.1" # Cloudflare DNS
                        logger.info(f"Pinging external host {target_ip} via {hostname}")
                        # Use -t for overall timeout in seconds on macOS/Linux
                        ping_cmd = ['ping', '-c', str(ping_count), '-t', str(ping_timeout), target_ip]
                        if platform.system() == "Windows":
                            # Windows ping uses -n for count, -w for timeout in ms
                            ping_cmd = ['ping', '-n', str(ping_count), '-w', str(ping_timeout * 1000), target_ip]

                        logger.debug(f"Executing ping command: {' '.join(ping_cmd)}")
                        result = subprocess.run(ping_cmd, capture_output=True, text=True, timeout=ping_timeout + 5, encoding='utf-8', errors='ignore')

                        if result.returncode == 0 and ('avg' in result.stdout or 'Average =' in result.stdout): # Check for Windows/Unix avg indicators
                            # Parse average latency (example for macOS/Linux)
                            try:
                                if platform.system() == "Windows":
                                    # Look for "Average = Xms"
                                    match = [line for line in result.stdout.splitlines() if "Average =" in line]
                                    if match:
                                        ping_latency = float(match[0].split("Average =")[1].split("ms")[0].strip())
                                    else:
                                        raise ValueError("Could not parse Windows ping avg")
                                else: # Linux/macOS
                                    # Look for "round-trip min/avg/max/stddev = x/y/z/a ms"
                                    stats_line = [line for line in result.stdout.splitlines() if 'round-trip' in line or 'rtt' in line]
                                    if stats_line:
                                        ping_latency = float(stats_line[0].split('/')[4]) # avg is the 5th value
                                    else: raise ValueError("Could not parse Linux/macOS ping avg")

                                logger.info(f"Ping {hostname}: {ping_latency:.1f} ms")
                            except (IndexError, ValueError, Exception) as parse_e:
                                logger.error(f"Failed to parse ping output for {hostname}: {parse_e}\nOutput:\n{result.stdout}")
                                test_error = "Ping Parse Error"
                        elif result.returncode != 0 or '100% packet loss' in result.stdout:
                            logger.warning(f"Ping to {target_ip} via {hostname} timed out or failed. RC: {result.returncode}")
                            test_error = "Ping Timeout"
                        else:
                            logger.warning(f"Ping to {target_ip} via {hostname} had non-zero RC or unexpected output. RC: {result.returncode}\nOutput:\n{result.stdout}")
                            test_error = "Ping Error"

                    except subprocess.TimeoutExpired:
                        logger.warning(f"Ping subprocess timed out for {hostname}")
                        test_error = "Ping Timeout"
                    except Exception as e:
                        logger.exception(f"Unexpected error during ping test for {hostname}: {e}")
                        test_error = "Ping Error (Exec)"

                    # --- Pause/Stop Check ---
                    while self.pause_event.is_set():
                        if self.stop_event.is_set(): break
                        time.sleep(0.5)
                    if self.stop_event.is_set():
                        self.update_server_status(hostname, "Stopped after ping")
                        self._safe_disconnect(hostname)
                        break

                    # 2b. Speed Test (Ookla CLI)
                    # 2b. Speed Test (Ookla CLI)
                    # Attempt speed test even if ping failed, but record ping error if it occurred.
                    try:
                        self.root.after(0, lambda txt=f"{op_text_base} - Speed Testing...": self.loading_animation.update_text(txt))
                        self.update_server_status(hostname, f"Speed Testing...")
                        speed_cmd = ['speedtest', '--format=json', '--accept-license', '--accept-gdpr']
                        logger.debug(f"Executing speedtest command: {' '.join(speed_cmd)}")
                        result = subprocess.run(speed_cmd, capture_output=True, text=True, timeout=speedtest_timeout, encoding='utf-8', errors='ignore')

                        if result.returncode == 0 and result.stdout.strip().startswith('{'):
                            try:
                                data = json.loads(result.stdout)
                                speed_results["ping"] = data.get("ping", {}).get("latency")
                                # Convert bps to Mbps
                                speed_results["download"] = data.get("download", {}).get("bandwidth", 0) * 8 / 1_000_000
                                speed_results["upload"] = data.get("upload", {}).get("bandwidth", 0) * 8 / 1_000_000
                                logger.info(f"Speedtest {hostname}: Ping={speed_results['ping']:.1f}ms, DL={speed_results['download']:.1f}Mbps, UL={speed_results['upload']:.1f}Mbps")
                            except json.JSONDecodeError as json_e:
                                logger.error(f"Failed to parse speedtest JSON for {hostname}: {json_e}\nOutput:\n{result.stdout[:500]}...")
                                if not test_error: test_error = "Speedtest Parse Error" # Prioritize earlier errors
                            except Exception as parse_e:
                                logger.error(f"Error processing speedtest data for {hostname}: {parse_e}")
                                if not test_error: test_error = "Speedtest Data Error" # Prioritize earlier errors
                        elif "ERROR:" in result.stdout or "ERROR:" in result.stderr:
                            err_line = result.stdout + result.stderr
                            logger.error(f"Speedtest CLI returned an error for {hostname}. RC: {result.returncode}\nOutput:\n{err_line[:500]}...")
                            # Try to extract specific error type
                            specific_error = "Speedtest Error" # Default
                            if "unable to connect" in err_line.lower(): specific_error = "Speedtest Connect Error"
                            elif "configuration" in err_line.lower(): specific_error = "Speedtest Config Error"
                            if not test_error: test_error = specific_error # Prioritize earlier errors
                        else:
                            logger.error(f"Speedtest CLI failed or gave unexpected output for {hostname}. RC: {result.returncode}\nOutput:\n{result.stdout[:500]}...")
                            if not test_error: test_error = "Speedtest Failed" # Prioritize earlier errors

                    except subprocess.TimeoutExpired:
                        logger.warning(f"Speedtest subprocess timed out for {hostname}")
                        if not test_error: test_error = "Speedtest Timeout" # Prioritize earlier errors
                    except Exception as e:
                        logger.exception(f"Unexpected error during speed test for {hostname}: {e}")
                        if not test_error: test_error = "Speedtest Error (Exec)" # Prioritize earlier errors

                # 3. Disconnect (always attempt if connect was tried, even if it failed)
                # --- Pause/Stop Check ---
                while self.pause_event.is_set():
                    if self.stop_event.is_set(): break
                    time.sleep(0.5)
                if self.stop_event.is_set():
                    self.update_server_status(hostname, "Stopped before disconnect")
                    self._safe_disconnect(hostname) # Still try to disconnect
                    break

                disconnect_error = self._safe_disconnect(hostname)

                # 4. Update Final Status
                final_status = "Error" # Default
                if connect_error:
                    final_status = connect_error
                elif test_error:
                    final_status = test_error
                elif disconnect_error:
                    final_status = f"Completed ({disconnect_error})" # Show completion but note disconnect issue
                elif connect_success:
                    final_status = "Completed"

                self.update_server_status(hostname, final_status, ping_latency, speed_results)

                # Optional delay between servers (check pause/stop during delay)
                delay_end_time = time.monotonic() + self.config.get("delay_between_servers", 1)
                while time.monotonic() < delay_end_time:
                     if self.pause_event.is_set():
                         # If paused during delay, wait until resumed or stopped
                         while self.pause_event.is_set():
                             if self.stop_event.is_set(): break
                             time.sleep(0.5)
                     if self.stop_event.is_set(): break
                     time.sleep(0.1) # Short sleep during active delay check
                if self.stop_event.is_set(): break # Exit outer loop if stopped during delay


            # --- Loop Finished ---
            elapsed = time.time() - start_time
            final_message = "Test Completed" if not self.stop_event.is_set() else "Test Stopped"
            logger.info(f"Test thread finished in {elapsed:.2f}s. Status: {final_message}")
            self.root.after(0, lambda msg=final_message: self.loading_animation.update_text(msg))

        except Exception as e:
            logger.exception("Unhandled error in test loop thread.")
            self.root.after(0, lambda: self.loading_animation.update_text("Test loop error!"))
        finally:
            # Ensure disconnect is attempted if loop breaks unexpectedly while connected
            if "Connected" in get_mullvad_status():
                 logger.warning("Test loop ended unexpectedly while connected, attempting disconnect.")
                 self._safe_disconnect("Unknown")
            # Schedule cleanup on main thread
            self.root.after(500, self._test_cleanup)

    def _safe_disconnect(self, hostname_for_log: str) -> Optional[str]:
        """Attempts to disconnect, logs errors, returns error message string or None."""
        disconnect_error = None
        try:
            # Check status first - don't disconnect if already disconnected
            status = get_mullvad_status()
            if "Disconnected" in status:
                 logger.info(f"Already disconnected before explicit call for {hostname_for_log}.")
                 return None

            # Don't update tree status here, final status update handles it
            # self.update_server_status(hostname_for_log, "Disconnecting...")
            logger.info(f"Disconnecting ({hostname_for_log})...")
            disconnect_mullvad()
            logger.info(f"Disconnect successful ({hostname_for_log}).")
            # Verify?
            time.sleep(0.5)
            status_after = get_mullvad_status()
            if "Disconnected" not in status_after:
                 logger.warning(f"Disconnect command ran for {hostname_for_log}, but status is still: {status_after}")
                 # disconnect_error = "Disconnect Verify Failed" # Maybe too noisy

        except MullvadCLIError as e:
            logger.error(f"Disconnection error ({hostname_for_log}): {e}")
            # Check if already disconnected despite error
            if "Disconnected" in get_mullvad_status():
                 logger.info("Disconnection error occurred, but status is now Disconnected.")
            else:
                 disconnect_error = "Disconnect Error"
        except Exception as e:
            logger.exception(f"Unexpected error disconnecting ({hostname_for_log}): {e}")
            disconnect_error = "Disconnect Error (Exec)"
        return disconnect_error


    # --- UI Update Methods (called via root.after) ---

    def update_server_status(self, hostname: str, status: str, ping: Optional[float] = None, speed: Optional[Dict] = None, clear_results: bool = False):
        """Update the status and results for a specific server in the Treeview."""
        if not self.server_tree: return

        def _update():
            if not self.server_tree or not self.server_tree.exists(hostname):
                # logger.warning(f"Item '{hostname}' no longer exists in tree, cannot update status.")
                return # Item might have been removed/reloaded

            try:
                values = list(self.server_tree.item(hostname, "values"))
                # Update status (column 4)
                values[4] = status

                # Update results (columns 5, 7, 8)
                ping_str = f"{ping:.1f}" if isinstance(ping, (int, float)) else (ping if ping else "") # Handle N/A etc.
                dl_str = f"{speed['download']:.1f}" if speed and isinstance(speed.get('download'), (int, float)) else ""
                ul_str = f"{speed['upload']:.1f}" if speed and isinstance(speed.get('upload'), (int, float)) else ""

                if clear_results:
                     values[5] = ""
                     values[6] = ""
                     values[7] = ""
                else:
                     if ping is not None: values[5] = ping_str
                     if speed is not None:
                         values[6] = dl_str
                         values[7] = ul_str

                # Apply status tags for color
                current_tags = list(self.server_tree.item(hostname, "tags"))
                # Remove old status tags, keep row color tags
                status_tags = ['status_error', 'status_timeout', 'status_completed']
                filtered_tags = [tag for tag in current_tags if not tag.startswith(('odd_row', 'even_row')) and tag not in status_tags]
                # Add back row color tag
                row_index = self.server_tree.index(hostname)
                use_alt_colors = self.config.get("alternating_row_colors", True)
                if use_alt_colors:
                     filtered_tags.append('odd_row' if row_index % 2 else 'even_row')

                # Add new status tag
                if "Error" in status: filtered_tags.append('status_error')
                elif "Timeout" in status: filtered_tags.append('status_timeout')
                elif "Completed" in status: filtered_tags.append('status_completed')

                self.server_tree.item(hostname, values=tuple(values), tags=tuple(filtered_tags))

                # Apply cell coloring for results (optional)
                # self.apply_cell_color(hostname, "ping_ms", ping)
                # self.apply_cell_color(hostname, "dl_mbps", speed.get('download') if speed else None)
                # self.apply_cell_color(hostname, "ul_mbps", speed.get('upload') if speed else None)

            except tk.TclError:
                 logger.warning(f"TCL error updating status for item {hostname} (item might be gone).")
            except Exception as e:
                logger.exception(f"Error updating UI for status of item {hostname}: {e}")

        self.root.after(0, _update)


    # --- Settings ---
    def open_settings(self):
        """Open the settings window."""
        # Simplified settings for now
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Settings")
        settings_window.geometry("450x350")
        settings_window.transient(self.root)
        settings_window.grab_set()
        settings_window.resizable(False, False)

        # Increased padding for the main frame in settings
        frame = ttk.Frame(settings_window, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.grid_columnconfigure(1, weight=1)

        # Cache Path - Increased vertical padding (pady=8)
        ttk.Label(frame, text="Mullvad Cache Path:").grid(row=0, column=0, sticky=tk.W, pady=8)
        cache_path_var = tk.StringVar(value=get_cache_path(self.config))
        cache_entry = ttk.Entry(frame, textvariable=cache_path_var, width=40)
        cache_entry.grid(row=0, column=1, sticky=tk.EW, padx=5)
        def browse_cache():
            initial_dir = os.path.dirname(cache_path_var.get()) or os.path.expanduser("~")
            path = filedialog.askopenfilename(title="Select relays.json", filetypes=[("JSON", "*.json")], initialdir=initial_dir, parent=frame)
            if path: cache_path_var.set(path)
        ttk.Button(frame, text="Browse...", command=browse_cache).grid(row=0, column=2, padx=5)

        # Ping Count - Increased vertical padding
        ttk.Label(frame, text="Ping Count:").grid(row=1, column=0, sticky=tk.W, pady=8)
        ping_count_var = tk.IntVar(value=self.config.get("ping_count", 3))
        ttk.Spinbox(frame, from_=1, to=10, textvariable=ping_count_var, width=5).grid(row=1, column=1, sticky=tk.W, padx=5)

        # Ping Timeout - Increased vertical padding
        ttk.Label(frame, text="Ping Timeout (s):").grid(row=2, column=0, sticky=tk.W, pady=8)
        timeout_var = tk.IntVar(value=self.config.get("timeout_seconds", 10))
        ttk.Spinbox(frame, from_=5, to=60, textvariable=timeout_var, width=5).grid(row=2, column=1, sticky=tk.W, padx=5)

        # Speedtest Timeout - Increased vertical padding
        ttk.Label(frame, text="Speedtest Timeout (s):").grid(row=3, column=0, sticky=tk.W, pady=8)
        speed_timeout_var = tk.IntVar(value=self.config.get("speedtest_timeout_seconds", 90))
        ttk.Spinbox(frame, from_=30, to=180, textvariable=speed_timeout_var, width=5).grid(row=3, column=1, sticky=tk.W, padx=5)

        # Connection Verify Timeout - Increased vertical padding
        ttk.Label(frame, text="Connection Verify Timeout (s):").grid(row=4, column=0, sticky=tk.W, pady=8)
        conn_verify_timeout_var = tk.IntVar(value=self.config.get("connection_verify_timeout", 15))
        ttk.Spinbox(frame, from_=5, to=60, textvariable=conn_verify_timeout_var, width=5).grid(row=4, column=1, sticky=tk.W, padx=5)

        # Delay Between Servers - Increased vertical padding
        ttk.Label(frame, text="Delay Between Servers (s):").grid(row=5, column=0, sticky=tk.W, pady=8)
        delay_var = tk.IntVar(value=self.config.get("delay_between_servers", 1))
        ttk.Spinbox(frame, from_=0, to=10, textvariable=delay_var, width=5).grid(row=5, column=1, sticky=tk.W, padx=5)

        # Alternating Row Colors - Increased vertical padding
        alt_rows_var = tk.BooleanVar(value=self.config.get("alternating_row_colors", True))
        ttk.Checkbutton(frame, text="Use alternating row colors", variable=alt_rows_var).grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=12)

        # --- Save/Cancel Buttons --- Increased top padding
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=7, column=0, columnspan=3, sticky='ew', pady=(20, 0))
        button_frame.grid_columnconfigure(0, weight=1) # Push buttons right
        button_frame.grid_columnconfigure(1, weight=0)
        button_frame.grid_columnconfigure(2, weight=0)


        def save_and_close():
            logger.info("Saving settings...")
            new_config = self.config.copy()
            custom_cache = cache_path_var.get()
            if custom_cache != get_default_cache_path(): new_config["custom_cache_path"] = custom_cache
            else: new_config["custom_cache_path"] = ""
            new_config["ping_count"] = ping_count_var.get()
            new_config["timeout_seconds"] = timeout_var.get()
            new_config["speedtest_timeout_seconds"] = speed_timeout_var.get()
            new_config["connection_verify_timeout"] = conn_verify_timeout_var.get() # Save new setting
            new_config["delay_between_servers"] = delay_var.get()
            new_config["alternating_row_colors"] = alt_rows_var.get()

            if save_config(new_config):
                 self.config = new_config
                 self.apply_theme() # Re-apply theme in case alt colors changed
                 # Reload server data if cache path changed effectively
                 if get_cache_path(self.config) != get_cache_path(load_config()):
                     self.load_server_data()
                 messagebox.showinfo("Settings Saved", "Settings saved successfully.", parent=settings_window)
                 settings_window.destroy()
            else:
                 messagebox.showerror("Save Error", "Failed to save settings. Check logs.", parent=settings_window)

        ttk.Button(button_frame, text="Save", command=save_and_close).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Cancel", command=settings_window.destroy).grid(row=0, column=2)

        settings_window.wait_window()

    # --- Miscellaneous ---
    def show_about(self):
        """Show the about dialog."""
        about_text = (
            f"Mullvad Automated Tester\n\n"
            f"Version: {APP_VERSION}\n\n"
            "Automates connecting to selected Mullvad servers, running ping and speed tests (via Ookla CLI), and disconnecting.\n\n"
            "Dependencies:\n"
            "- Mullvad VPN Client (for 'mullvad' CLI)\n"
            "- Ookla Speedtest CLI ('speedtest')\n"
            "  (Install via `brew install speedtest --force` on macOS)\n\n"
            f"Log File: {get_log_path()}"
        )
        messagebox.showinfo("About Mullvad Automated Tester", about_text, parent=self.root)

    # --- THEMEING (reuse from example/gui.py, simplified) ---
    def apply_theme(self):
        """Applies the selected theme."""
        global sv_ttk
        theme_mode = self.theme_var.get()
        logger.info(f"Applying theme: {theme_mode}")

        if sv_ttk:
            try:
                actual_mode = theme_mode
                if theme_mode == "system":
                    try:
                        # Basic detection, same as before
                        system = platform.system()
                        is_dark = False
                        if system == "Windows":
                            import winreg
                            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                            is_dark = (value == 0)
                            winreg.CloseKey(key)
                        elif system == "Darwin":
                            is_dark = self.root.tk.call("tk", "windowingsystem") == "aqua" and \
                                      self.root.tk.call("::tk::unsupported::MacWindowStyle", "isdark", self.root)
                        actual_mode = "dark" if is_dark else "light"
                    except Exception: actual_mode = "light" # Fallback

                sv_ttk.set_theme(actual_mode)
                self.theme_colors = { # Get colors from sv_ttk
                    "background": sv_ttk.style.colors.bg, "foreground": sv_ttk.style.colors.fg,
                    "row_odd": sv_ttk.style.colors.bg, "row_even": getattr(sv_ttk.style.colors, 'alt_bg', sv_ttk.style.colors.bg),
                    "select_bg": sv_ttk.style.colors.select_bg, "select_fg": sv_ttk.style.colors.select_fg,
                }
                logger.info(f"sv-ttk theme set to '{actual_mode}'.")

                self.root.configure(bg=self.theme_colors["background"])
                if self.main_frame: self.main_frame.configure(style='TFrame')
                if self.server_tree:
                    style = ttk.Style()
                    style.map('Treeview', background=[('selected', self.theme_colors["select_bg"])], foreground=[('selected', self.theme_colors["select_fg"])])
                    self.server_tree.tag_configure('odd_row', background=self.theme_colors["row_odd"], foreground=self.theme_colors["foreground"])
                    self.server_tree.tag_configure('even_row', background=self.theme_colors["row_even"], foreground=self.theme_colors["foreground"])
                    # Re-sort to apply row colors to existing items
                    self.sort_treeview(self.sort_column, force_order=self.sort_order)

                self.root.update_idletasks()
                return # Applied sv-ttk

            except Exception as e:
                logger.error(f"Error applying sv-ttk theme: {e}. Falling back.")
                sv_ttk = None # Disable on error

        # --- Fallback Manual Styling ---
        # Simplified: Just set row colors based on config, rely on default ttk otherwise
        if self.server_tree:
             use_alt = self.config.get("alternating_row_colors", True)
             odd_bg = "#FFFFFF" if not use_alt else "#F5F5F5" # Example light theme colors
             even_bg = "#FFFFFF"
             fg = "#000000"
             self.server_tree.tag_configure('odd_row', background=odd_bg, foreground=fg)
             self.server_tree.tag_configure('even_row', background=even_bg, foreground=fg)
             # Re-sort to apply row colors
             self.sort_treeview(self.sort_column, force_order=self.sort_order)
        self.root.update_idletasks()

    # --- Connection Logic ---
    def _get_server_details_from_item_id(self, item_id: str) -> Optional[Dict[str, Any]]:
         """Finds the full server data dictionary based on a Treeview item ID (hostname)."""
         if not self.server_tree or not self.all_servers_list:
             return None
         try:
             if not self.server_tree.exists(item_id): return None # Check if item exists
             # Assuming item_id is the hostname
             hostname = item_id
             server_details = next((s for s in self.all_servers_list if s.get("hostname") == hostname), None)
             if not server_details:
                  logger.warning(f"Could not find server data for hostname: {hostname}")
             return server_details
         except tk.TclError:
              logger.warning(f"Could not get details for item_id {item_id} (may be invalid).")
              return None

    def _connect_on_double_click(self, event=None):
        """Connect to the server that was double-clicked."""
        if not self.server_tree: return
        item_id = self.server_tree.focus() # Get the currently focused item
        if not item_id: return # No item focused

        server_details = self._get_server_details_from_item_id(item_id)
        if not server_details:
             messagebox.showerror("Error", "Could not retrieve details for the selected server.", parent=self.root)
             return

        hostname = server_details.get("hostname", "N/A")
        country_code = server_details.get("country_code")
        city_code = server_details.get("city_code")

        if not country_code or not city_code:
             messagebox.showerror("Error", f"Missing country/city code for server {hostname}.", parent=self.root)
             return

        # Determine protocol based on hostname convention (or add protocol column later)
        protocol = "wireguard" if "-wg" in hostname.lower() or ".wg." in hostname.lower() else "openvpn"

        logger.info(f"Attempting connection via double-click to: {hostname} ({country_code}/{city_code}) using {protocol}")
        # Run connection in a separate thread
        threading.Thread(
            target=self._connect_to_server_thread, # Use a dedicated thread method
            args=(protocol, country_code, city_code, hostname),
            daemon=True
        ).start()

    def _connect_to_server_thread(self, protocol: str, country_code: str, city_code: str, hostname: str):
        """Internal method to handle connection process in a thread (for double-click)."""
        # Similar to test loop connection, but simpler status updates
        self.root.after(0, lambda: self.loading_animation.update_text(f"Connecting to {hostname}..."))
        self.root.after(0, lambda: self.loading_animation.start(self.root))
        # Disable buttons? Maybe not necessary for a quick connect action.

        try:
            set_mullvad_location(country_code, city_code, hostname)
            time.sleep(1.0)
            connect_mullvad()

            # Verify
            connection_verified = False
            verify_timeout = self.config.get("connection_verify_timeout", 15)
            verify_start_time = time.monotonic()
            while time.monotonic() < verify_start_time + verify_timeout:
                status_output = get_mullvad_status()
                if "Connected" in status_output:
                    connection_verified = True
                    break
                time.sleep(0.5)

            if connection_verified:
                self.root.after(0, lambda h=hostname: self.loading_animation.update_text(f"Connected to {h}"))
                logger.info(f"Connection to {hostname} successful (via double-click).")
                self.root.after(1500, self.loading_animation.stop)
            else:
                 logger.error(f"Connection verification timed out for {hostname} (via double-click).")
                 self.root.after(0, lambda: messagebox.showerror("Connection Failed", f"Connection verification timed out for {hostname}.", parent=self.root))
                 self.root.after(0, lambda: self.loading_animation.update_text("Connection failed"))
                 self.root.after(1500, self.loading_animation.stop)
                 self._safe_disconnect(hostname) # Attempt disconnect

        except MullvadCLIError as e:
            logger.error(f"Mullvad CLI error during double-click connection: {e}")
            self.root.after(0, lambda err=e: messagebox.showerror("Connection Failed", f"Mullvad command failed:\n{err}", parent=self.root))
            self.root.after(0, lambda: self.loading_animation.update_text("Connection failed"))
            self.root.after(1500, self.loading_animation.stop)
        except Exception as e:
            logger.exception("Unexpected error during double-click connection process.")
            self.root.after(0, lambda err=e: messagebox.showerror("Connection Error", f"An unexpected error occurred:\n{err}", parent=self.root))
            self.root.after(0, lambda: self.loading_animation.update_text("Connection error"))
            self.root.after(1500, self.loading_animation.stop)


    # --- Export Function ---
    def export_results_to_csv(self):
        """Export the current server list with results to a CSV file."""
        if not self.server_tree: return
        items = self.server_tree.get_children()
        if not items:
            messagebox.showinfo("Export", "No results to export.", parent=self.root)
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            title="Export Test Results to CSV",
            parent=self.root
        )
        if not file_path: return # User cancelled

        self.loading_animation.update_text("Exporting to CSV...")
        self.loading_animation.start(self.root)
        self.root.update_idletasks()

        headers = ["hostname", "city", "country", "status", "ping_ms", "dl_mbps", "ul_mbps"]
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(headers) # Write header row

                for item_id in items:
                    try:
                        if not self.server_tree.exists(item_id): continue
                        values = self.server_tree.item(item_id, "values")
                        # Extract values corresponding to headers (skip checkbox column #0)
                        row_data = values[1:] # Slice to get values from index 1 onwards
                        writer.writerow(row_data)
                    except (tk.TclError, IndexError):
                        logger.warning(f"Could not get values for item {item_id} during export.")
                        continue # Skip this item

            messagebox.showinfo("Export Successful", f"Results exported to:\n{file_path}", parent=self.root)
            self.loading_animation.update_text("Export successful")

        except IOError as e:
            logger.exception(f"IOError exporting results to CSV {file_path}: {e}")
            messagebox.showerror("Export Failed", f"Failed to write file:\n{e}", parent=self.root)
            self.loading_animation.update_text("Export failed")
        except Exception as e:
            logger.exception(f"Unexpected error exporting results to CSV {file_path}: {e}")
            messagebox.showerror("Export Error", f"An unexpected error occurred:\n{e}", parent=self.root)
            self.loading_animation.update_text("Export error")
        finally:
             self.root.after(1000, self.loading_animation.stop)


    def change_theme(self):
        """Called when the theme radiobutton selection changes."""
        new_theme = self.theme_var.get()
        logger.info(f"Theme selection changed to: {new_theme}")
        if new_theme != self.config.get("theme_mode", "system"):
            self.config["theme_mode"] = new_theme
            save_config(self.config)
        self.apply_theme()

# Example of how to run if this script were executed directly (for testing)
if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    root = tk.Tk()
    # Set theme before creating app if sv_ttk is used
    if sv_ttk:
        sv_ttk.set_theme("light") # Or "dark" for testing
    app = TesterApp(root)
    root.mainloop()
