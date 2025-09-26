#!/usr/bin/env python3
"""
SortAnything Application
A comprehensive GUI tool for manual file and folder sorting with multiple phases.
Dark theme version with optimized code structure.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from PIL import Image, ImageTk
import os
import json
import shutil
import fnmatch
from pathlib import Path
import csv
from typing import List, Any, Optional
import subprocess
import sys
from datetime import datetime
import random


# Constants
DARK_COLORS = {
    'bg': "#2b2b2b", 'fg': "#ffffff", 'select': "#404040", 'tab': "#858585",
    'active': "#505050", 'entry_bg': "#404040", 'success': "#D2ED64", 'danger': "#D2ED64"}

COLOR_PALETTE = [
    '#d6a5c9',  # muted lavender
    '#a6dcef',  # soft cyan
    '#fbc4ab',  # muted peach
    '#b5c9b8',  # muted sage green
    '#f6eac2',  # soft cream
    '#a8a3d1',  # soft periwinkle
    '#dbb0b0',  # muted rose
    '#c5d1d8',  # muted blue-gray
    '#f2c6b3',  # muted coral
    '#c3d9c8'   # soft green
]


ABOUT_TEXT = """SortAnything v1.0
A versatile tool designed to help you organize files and folders efficiently.
• Interactive file selection with advanced filtering options
• Customizable sorting categories (buckets) for personalized organization
• Keyboard shortcuts to speed up manual sorting
• Options to save sorted lists or physically move files
• Session saving and loading for resuming work later
• Detailed reports on sorting activities
GitHub: [Mohsyn](https://github.com/mohsyn)"""

type_map = {
            'image': ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp', '.svg'],
            'text': ['.txt', '.md', '.py', '.js', '.html', '.css', '.json', '.xml', '.b4a','.b4j','.b4i','.log','.nfo','.java', '.c', '.cpp', '.h', '.hpp', '.cs', '.go', '.rb', '.php', '.swift', '.kt', '.kts'],
            'documents': ['.pdf','.doc','.docx','.dot','.ppt','.pptx','.xls','.xlsx','.xlsm','.xlst', '.csv', '.odt', '.ods', '.odp'],
            'compressed': ['.zip', '.rar', '.tar', '.gzip', '.iso', '.7z', '.arc', '.lhz', '.gz'],
            'video': ['.mp4', '.avi', '.mov', '.wmv', '.flv', '.mkv', '.webm'],
            'audio': ['.mp3', '.wav', '.flac', '.aac', '.ogg', '.wma'],
            'executable': ['.exe', '.msi', '.deb', '.rpm', '.dmg', '.app', '.bat', '.sh', '.cmd', '.com', '.run'],
            'font': ['.ttf', '.otf', '.woff', '.woff2'],
            'database': ['.db', '.sqlite', '.sqlite3', '.sql', '.accdb', '.mdb'],
            'ebook': ['.epub', '.mobi', '.azw', '.azw3'],
            'virtualization': ['.vdi', '.vmdk', '.ova', '.ovf'],
            'configuration': ['.ini', '.cfg', '.conf', '.yml', '.yaml', '.toml'],
            'script': ['.ps1', '.vbs', '.sh', '.bash', '.zsh', '.bat', '.cmd'],
            'system': ['.dll', '.sys', '.drv',]
        }


def get_random_color(): return random.choice(COLOR_PALETTE)

def configure_dark_theme(root):
    """Configure dark theme for the application."""
    style = ttk.Style()
    style.theme_use('clam')
    
    # Configure all ttk styles at once
    configs = {
        'TNotebook': {'background': DARK_COLORS['bg'], 'borderwidth': 0},
        'TNotebook.Tab': {'background': DARK_COLORS['bg'], 'foreground': DARK_COLORS['fg'], 'padding': (8,8), 'borderwidth': 1, 'font': ('Arial', 10, 'normal')},
        'TFrame': {'background': DARK_COLORS['bg'], 'borderwidth': 0},
        'TLabelframe': {'background': DARK_COLORS['bg'], 'foreground': DARK_COLORS['fg'], 'borderwidth': 1},
        'TLabelframe.Label': {'background': DARK_COLORS['bg'], 'foreground': DARK_COLORS['fg']},
        'TLabel': {'background': DARK_COLORS['bg'], 'foreground': DARK_COLORS['fg']},
        'TButton': {'background': DARK_COLORS['select'], 'foreground': DARK_COLORS['fg'], 'borderwidth': 1},
        'TEntry': {'background': DARK_COLORS['active'], 'foreground': DARK_COLORS['fg'], 'fieldbackground': DARK_COLORS['entry_bg'], 'borderwidth': 1, 'font': ('Arial', 12)},
        'TCheckbutton': {'background': DARK_COLORS['bg'], 'foreground': DARK_COLORS['fg'], 'borderwidth': 0},
        'TRadiobutton': {'background': DARK_COLORS['bg'], 'foreground': DARK_COLORS['fg'], 'borderwidth': 0},
        'TProgressbar': {'background': DARK_COLORS['success'], 'borderwidth': 0},
        'Treeview': {'background': DARK_COLORS['entry_bg'], 'foreground': DARK_COLORS['fg'], 'fieldbackground': DARK_COLORS['entry_bg'], 'borderwidth': 0, 'font': ('Arial', 12)},
        'Treeview.Heading': {'background': DARK_COLORS['select'], 'foreground': DARK_COLORS['fg'], 'borderwidth': 1},
        'Vertical.TScrollbar': {'background': DARK_COLORS['select'], 'borderwidth': 0, 'arrowcolor': DARK_COLORS['fg'], 'troughcolor': DARK_COLORS['bg']},
        'Horizontal.TScrollbar': {'background': DARK_COLORS['select'], 'borderwidth': 0, 'arrowcolor': DARK_COLORS['fg'], 'troughcolor': DARK_COLORS['bg']} }
    
    for widget, config in configs.items():
        style.configure(widget, **config)
    # Nav buttons
    style.configure('Nav.TButton', font=('Arial', 14, 'bold'), padding=(6, 4), borderwidth=0)
    style.map('TNotebook.Tab',  
              background=[('selected', DARK_COLORS['select']), ('active', DARK_COLORS['select'])],
              foreground=[('selected', DARK_COLORS['fg'])],
              #borderwidth=[('selected', 3),('active', 6)],
              font=[('selected', ("Segoe UI", 12, "bold"))],
              padding=[('selected', (10, 8))])

    style.map('TButton', background=[('active', DARK_COLORS['active'])])
    style.map('Treeview', background=[('selected', DARK_COLORS['active'])])
    # Ensure checkbutton text remains visible on hover/active
    style.map('TCheckbutton', 
              foreground=[('active', DARK_COLORS['fg']), ('selected', DARK_COLORS['fg'])],
              background=[('active', DARK_COLORS['bg']), ('selected', DARK_COLORS['bg'])])
    # Ensure radiobutton text remains visible on hover/active
    style.map('TRadiobutton', 
              foreground=[('active', DARK_COLORS['fg']), ('selected', DARK_COLORS['fg'])],
              background=[('active', DARK_COLORS['bg']), ('selected', DARK_COLORS['bg'])])
    # Ensure treeview headings remain visible on hover/active
    style.map('Treeview.Heading', 
              foreground=[('active', DARK_COLORS['fg']), ('selected', DARK_COLORS['fg'])],
              background=[('active', DARK_COLORS['select']), ('selected', DARK_COLORS['select'])])
    
    root.configure(bg=DARK_COLORS['bg'])

class FileItem:
    def __init__(self, path):
        self.name = path
        self.path = path
        self.is_file = False
        self.size = 0
        self.modified = datetime.now()
        self.extension = ""
        self.selected = False
        self.bucket = None
        self.attributes = {}
        self.skipped = False
        self._populate_metadata()
    
    def _populate_metadata(self):
        """Populate metadata from actual file if path exists."""
        try:
            path = Path(self.path)
            if path.exists():
                self.name = path.name
                self.is_file = path.is_file()
                if self.is_file:
                    self.size = path.stat().st_size
                    self.extension = path.suffix.lower()
                self.modified = datetime.fromtimestamp(path.stat().st_mtime)
        except (OSError, PermissionError, ValueError):
            pass
        
    def __str__(self):
        return f"{self.name} ({self.size} bytes, {self.modified.strftime('%Y-%m-%d %H:%M')})"

class Bucket:
    """Represents a sorting bucket."""
    def __init__(self, number: int, name: str = "", color: Optional[str] = None):
        self.number = number
        self.name = name or f"Bucket {number}"
        self.color = color or get_random_color()
        self.items: List[FileItem] = []
        self.hotkey = str(number) if number < 10 else "0"

class FileSorterApp:
    """Main application class for the SortAnything."""
    def __init__(self, root):
        self.root = root
        self.root.title("SortAnything")
        self.root.geometry("1200x800")
        
        configure_dark_theme(root)
        
        # Application state
        self.current_phase = 1
        self.files: List[FileItem] = []
        self.buckets: List[Bucket] = []
        self.current_file_index = 0
        self.output_mode = "list"
        self.output_directory = ""
        self.session_data = {}
        self.list_mode = False
        self.current_list_items = []
        self.csv_headers = None
        self.current_columns = []
        self.columns_mode = 'folder'
        self._detail_value_labels = []
        self.bucket_button_map = {}
        self.bucket_indicator_map = {}
        self._preview_images = []
        self._resizing_image = False  # Flag to prevent recursive resizes
        self._current_image_dimensions = None

        
        # Create main notebook
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Initialize phases
        self.init_phase1()
        self.init_phase2()
        self.init_phase3()
        self.init_phase4()
        
        # Bind hotkeys
        self.root.bind('<KeyPress>', self.handle_hotkey)
        self.root.focus_set()
        # Focus skip button when entering sorting tab
        self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_change)

    def _on_tab_change(self, event=None):
        try:
            if self.notebook.index(self.notebook.select()) == 2 and hasattr(self, 'skip_btn'):
                self.skip_btn.focus_set()
        except Exception:
            pass

    def create_button_frame(self, parent, buttons):
        """Helper to create standardized button frames."""
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, padx=5, pady=5)
        for i, (text, command, side) in enumerate(buttons):
            if text == "SEPARATOR":
                continue
            btn_config = {'text': text, 'command': command}
            if text in ["Proceed to Configuration", "Start Sorting", "Complete Sorting"]:
                btn = tk.Button(frame, bg=DARK_COLORS['success'], fg="black", 
                               activebackground="#45a049", font=('Arial', 10, 'bold'), **btn_config)
            else:
                btn = ttk.Button(frame, **btn_config)
            btn.pack(side=side, padx=2)
        return frame

    # Replace your helper with this version
    def create_tree_with_scrollbars(self, parent, columns, show='headings',
                                    need_v=True, need_h=True, auto_hide_h=True):
        frame = ttk.Frame(parent)
        tree = ttk.Treeview(frame, columns=columns, show=show)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        v_scroll = None
        h_scroll = None

        if need_v:
            v_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscrollcommand=v_scroll.set)
            v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        if need_h and not auto_hide_h:
            h_scroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
            tree.configure(xscrollcommand=h_scroll.set)
            h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        if need_h and auto_hide_h:
            h_scroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
            # Don't pack initially - only pack when needed
            tree.configure(xscrollcommand=h_scroll.set)

            def toggle_hscroll(event=None):
                try:
                    tree.update_idletasks()
                    total_width = sum(int(tree.column(c, 'width')) for c in columns)
                    needs_scroll = total_width > int(tree.winfo_width())
                except Exception:
                    needs_scroll = False
                mapped = h_scroll.winfo_ismapped()
                if needs_scroll and not mapped:
                    h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
                elif not needs_scroll and mapped:
                    h_scroll.pack_forget()

            tree.bind('<Configure>', toggle_hscroll)
            frame.after(50, toggle_hscroll)

        return frame, tree

    def init_phase1(self):
        """Initialize Phase 1: Input Selection Interface."""
        self.phase1_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.phase1_frame, text="1. Input Selection")
        
        main_frame = ttk.PanedWindow(self.phase1_frame, orient=tk.HORIZONTAL)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left panel
        left_frame = ttk.Frame(main_frame)
        main_frame.add(left_frame, weight=1)
        
        ttk.Label(left_frame, text="Input Mode", font=('Arial', 12, 'bold')).pack(pady=5)
        
        # Mode selection
        mode_frame = ttk.Frame(left_frame)
        mode_frame.pack(fill=tk.X, padx=5, pady=5)
        self.input_mode_var = tk.StringVar(value="folder")
        rb_folder = ttk.Radiobutton(mode_frame, text="Folder Mode", variable=self.input_mode_var,
                           value="folder", command=self.toggle_input_mode)
        rb_folder.pack(side=tk.LEFT, anchor=tk.W, padx=(0, 10))
        self._add_tooltip(rb_folder, "Select if you want to organize files into folders")

        rb_list = ttk.Radiobutton(mode_frame, text="List Mode", variable=self.input_mode_var,
                           value="list", command=self.toggle_input_mode)
        rb_list.pack(side=tk.LEFT, anchor=tk.W, padx=(0, 10))
        self._add_tooltip(rb_list, "Select this if you want to sort list from csv/text/clipboard")
        
        # Folder controls
        self.folder_controls = ttk.Frame(left_frame)
        self.folder_controls.pack(fill=tk.BOTH, expand=True)
        
        btn_frame_1 = self.create_button_frame(self.folder_controls, [])
        add_dir_btn = ttk.Button(btn_frame_1, width=15, text="Add Directory", command=self._add_directory)
        add_dir_btn.pack(side=tk.LEFT, padx=2, pady=20)
        self._add_tooltip(add_dir_btn, "Browse for and add a folder to the list of sources")
        remove_dir_btn = ttk.Button(btn_frame_1,width=20, text="Remove Selected", command=self.remove_directory)
        remove_dir_btn.pack(side=tk.LEFT, padx=20)
        self._add_tooltip(remove_dir_btn, "Removes the selected folder(s) from list below")
        
        # Create listbox with scrollbar
        listbox_frame = ttk.Frame(self.folder_controls)
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.dir_listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE,
                                     bg=DARK_COLORS['entry_bg'], fg='white',
                                     selectbackground=DARK_COLORS['active'], font=('Arial', 12))
        self.dir_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add vertical scrollbar
        listbox_scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.dir_listbox.yview)
        self.dir_listbox.configure(yscrollcommand=listbox_scrollbar.set)
        listbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.dir_listbox.bind('<<ListboxSelect>>', self.on_dir_select)
        
        # List controls
        self.list_controls = ttk.Frame(left_frame)
        list_buttons = [
            ("Import from Clipboard", self.paste_from_clipboard, tk.TOP),
            ("Import CSV", self.import_csv, tk.TOP),
            ("Import Text", self.import_text, tk.TOP)
        ]
        list_btn_frame = ttk.Frame(self.list_controls)
        list_btn_frame.pack(fill=tk.Y, padx=5, pady=5)
        
        btn_clipboard = ttk.Button(list_btn_frame, text="From Clipboard", command=self.paste_from_clipboard)
        btn_clipboard.pack(fill=tk.X, pady=22)
        self._add_tooltip(btn_clipboard, "Import a list of items by pasting from the clipboard")

        btn_csv = ttk.Button(list_btn_frame, text="Import CSV", command=self.import_csv)
        btn_csv.pack(fill=tk.X, pady=12)
        self._add_tooltip(btn_csv, "Import items from a CSV file (First row will be used as headers)")

        btn_text = ttk.Button(list_btn_frame, text="Import Text", command=self.import_text)
        btn_text.pack(fill=tk.X, pady=12)
        self._add_tooltip(btn_text, "Import items from a text file (one item per line)")
        
        # Right panel
        right_frame = ttk.Frame(main_frame)
        main_frame.add(right_frame, weight=2)
        ttk.Label(right_frame, text="Items", font=('Arial', 12, 'bold')).pack(pady=5)
        
        # Filter controls
        filter_frame = ttk.Frame(right_frame)
        filter_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(filter_frame, text="Filter:").pack(side=tk.LEFT)
        self.filter_var = tk.StringVar(value="*")
        self.filter_var.trace('w', self.apply_filter)
        filter_entry = ttk.Entry(filter_frame, textvariable=self.filter_var, font=('Arial', 16), width=16)
        filter_entry.pack(side=tk.LEFT, padx=5)

        btn_select_matching = ttk.Button(filter_frame, text="Select Matching", command=self.select_matching)
        btn_select_matching.pack(side=tk.LEFT, padx=5)
        self._add_tooltip(btn_select_matching, "Select all items that match the filter (wildcards supported)")

        btn_clear_selection = ttk.Button(filter_frame, text="Clear Selection", command=self.clear_selection)
        btn_clear_selection.pack(side=tk.LEFT, padx=5)
        self._add_tooltip(btn_clear_selection, "Deselect all currently selected items.")

        btn_clear_all = ttk.Button(filter_frame, text="Clear All", command=self.clear_all_action)
        btn_clear_all.pack(side=tk.LEFT, padx=5)
        self._add_tooltip(btn_clear_all, "Remove all items from the list below while retaining sources")

        btn_refresh = ttk.Button(filter_frame, text="Refresh", command=self.refresh_button_action)
        btn_refresh.pack(side=tk.LEFT, padx=5)
        self._add_tooltip(btn_refresh, "Refresh list to show changes made to selection")
        
        self.show_selected_only = tk.BooleanVar()
        cb_show_selected = ttk.Checkbutton(filter_frame, text="Show Selected Only",
                       variable=self.show_selected_only, command=self.toggle_show_selected)
        cb_show_selected.pack(side=tk.LEFT, padx=5)
        self._add_tooltip(cb_show_selected, "Toggle to show only the items that are currently selected")
        
        # Item tree
        item_tree_frame, self.item_tree = self.create_tree_with_scrollbars(
            right_frame, ('checkbox', 'Filename', 'Size', 'Modified', 'Type'))
        item_tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Configure tree columns
        col_configs = {
            'checkbox': {'text': '✓', 'width': 40, 'stretch': False, 'anchor': 'center'},
            'Filename': {'text': 'Filename', 'width': 300, 'stretch': True, 'anchor': 'w'},
            'Size': {'text': 'Size', 'width': 100, 'stretch': False},
            'Modified': {'text': 'Modified', 'width': 150, 'stretch': False},
            'Type': {'text': 'Type', 'width': 80, 'stretch': False}
        }
        
        for col, config in col_configs.items():
            self.item_tree.heading(col, text=config['text'], anchor=config.get('anchor', 'w'))
            self.item_tree.column(col, width=config['width'], stretch=config['stretch'])
        
        self.item_tree.column('#0', width=0, stretch=tk.NO)
        self.item_tree.bind('<Button-1>', self.toggle_item_selection)
        
        # Bottom controls
        bottom_frame = ttk.Frame(right_frame)
        bottom_frame.pack(fill=tk.X, padx=5, pady=5)
        self.selection_label = ttk.Label(bottom_frame, text="0 of 0 items selected")
        self.selection_label.pack(side=tk.LEFT)
        btn_proceed1 = tk.Button(bottom_frame, text="Proceed to Configuration", command=self.proceed_to_phase2,
                 bg=DARK_COLORS['success'], fg="black", activebackground="#45a049", 
                 font=('Arial', 10, 'bold'))
        btn_proceed1.pack(side=tk.RIGHT, padx=2)
        self._add_tooltip(btn_proceed1, "Continue to the next step to configure sorting buckets.")
        
        self.list_controls.pack_forget()

    def _add_directory(self):
        """Add a directory to the selection list."""
        directory = filedialog.askdirectory(title="Select Directory to Sort")
        if directory and directory not in self.dir_listbox.get(0, tk.END):
            self.dir_listbox.insert(tk.END, directory)
            self.refresh_items()
    
    def remove_directory(self):
        """Remove selected directories from the list."""
        for index in reversed(self.dir_listbox.curselection()):
            self.dir_listbox.delete(index)
        self.refresh_items()
    
    def configure_item_tree_columns(self, columns, headers_mapping=None):
        """Configure the item tree columns dynamically."""
        self.current_columns = columns
        self.item_tree.configure(columns=columns, show='headings')
        self.item_tree.column('#0', width=0, stretch=tk.NO)
        
        for col in columns:
            if col == 'checkbox':
                self.item_tree.heading('checkbox', text='✓', anchor='center')
                self.item_tree.column('checkbox', width=40, stretch=tk.NO, anchor='center')
            else:
                label = headers_mapping.get(col, col) if headers_mapping else col
                self.item_tree.heading(col, text=label, anchor='w')
                width = 300 if col in ('Filename', 'Item') else 150 if col == 'Modified' else 100
                self.item_tree.column(col, width=width, stretch=tk.YES)
    
    def toggle_input_mode(self):
        """Switch between folder mode and list mode."""
        # Warn if there is existing data
        has_data = bool(self.dir_listbox.size()) or bool(self.files)
        if has_data:
            if not messagebox.askyesno("Switch Mode",
                                       "Switching modes will reset current data and sorting. Continue?"):
                # Revert selection
                self.input_mode_var.set('list' if self.input_mode_var.get() == 'folder' else 'folder')
                return
            # Reset all
            for bucket in self.buckets:
                bucket.items.clear()
            self.buckets.clear()
            self.files.clear()
            self.dir_listbox.delete(0, tk.END)
            self.current_file_index = 0

        if self.input_mode_var.get() == "folder":
            self.folder_controls.pack(fill=tk.BOTH, expand=True)
            self.list_controls.pack_forget()
        else:
            self.folder_controls.pack_forget()
            self.list_controls.pack(fill=tk.BOTH, expand=True)
        self.refresh_display()
    
    def paste_from_clipboard(self):
        """Paste items from clipboard."""
        try:
            content = self.root.clipboard_get()
            lines = [line.strip() for line in content.strip().split('\n') if line.strip()]
            
            self.csv_headers = None
            self.files.clear()
            
            for line in lines:
                file_item = FileItem(line)
                file_item.name = line
                file_item.is_file = False
                self.files.append(file_item)
            
            self.columns_mode = 'list'
            self.configure_item_tree_columns(['checkbox', 'Item'])
            self.refresh_display()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to paste from clipboard: {e}")
    
    def import_csv(self):
        """Import CSV file containing items."""
        filename = filedialog.askopenfilename(
            title="Select CSV File", filetypes=[("CSV files", "*.csv")])
        if not filename:
            return

        try:
            with open(filename, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                headers = [h.strip() for h in next(reader)]
                
                self.csv_headers = headers
                self.files.clear()

                for row in reader:
                    if not row:
                        continue
                    
                    first_val = row[0].strip() if len(row) > 0 else ""
                    file_item = FileItem(first_val)
                    file_item.name = first_val
                    file_item.is_file = False

                    for i, header in enumerate(self.csv_headers):
                        value = row[i].strip() if i < len(row) else ""
                        if header:
                            file_item.attributes[header] = value

                    self.files.append(file_item)

            self.columns_mode = 'list'
            self.configure_item_tree_columns(['checkbox'] + self.csv_headers)
            self.refresh_display()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import CSV: {e}")

    def import_text(self):
        """Import plain text file where each line represents an item."""
        filename = filedialog.askopenfilename(
            title="Select Text File", filetypes=[("Text files", "*.txt *.text")])
        if not filename:
            return

        try:
            with open(filename, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f.readlines() if line.strip()]

            self.csv_headers = None
            self.files.clear()

            for line in lines:
                file_item = FileItem(line)
                file_item.name = line
                file_item.is_file = False
                self.files.append(file_item)

            self.configure_item_tree_columns(['checkbox', 'Item'])
            self.columns_mode = 'list'
            self.refresh_display()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import text file: {e}")
    
    def get_filtered_items(self):
        """Get items matching the current filter."""
        filter_pattern = self.filter_var.get().strip()
        if not filter_pattern:
            return self.files
        return [f for f in self.files if fnmatch.fnmatch(f.name.lower(), filter_pattern.lower())]
    
    def refresh_display(self):
        """Refresh the item tree display."""
        for item_id in self.item_tree.get_children():
            self.item_tree.delete(item_id)

        filtered_items = self.get_filtered_items()

        # Configure columns based on current mode
        if self.columns_mode == 'list':
            expected_cols = ['checkbox'] + (self.csv_headers if self.csv_headers else ['Item'])
        else:
            expected_cols = ['checkbox', 'Filename', 'Size', 'Modified', 'Type']
        
        if getattr(self, 'current_columns', []) != expected_cols:
            self.configure_item_tree_columns(expected_cols, {'checkbox': '✓'} if self.columns_mode != 'list' else None)

        # Populate rows
        for file_item in filtered_items:
            if self.show_selected_only.get() and not file_item.selected:
                continue

            checkbox_symbol = "✓" if file_item.selected else ""
            
            if self.input_mode_var.get() == 'list':
                if self.csv_headers:
                    row_values = [checkbox_symbol] + [file_item.attributes.get(h, "") for h in self.csv_headers]
                else:
                    row_values = [checkbox_symbol, file_item.name]
            else:
                size_str = f"{file_item.size:,} bytes" if file_item.is_file else "N/A"
                modified_str = file_item.modified.strftime("%Y-%m-%d %H:%M")
                type_str = self._get_file_type(file_item)
                row_values = [checkbox_symbol, file_item.name, size_str, modified_str, type_str]

            self.item_tree.insert('', 'end', values=row_values,
                                  tags=('selected' if file_item.selected else 'unselected'))

        self.item_tree.tag_configure('selected', background=DARK_COLORS['active'])

        # Update selection counter
        selected_count = sum(1 for f in self.files if f.selected)
        total_count = len(filtered_items)
        self.selection_label.config(text=f"{selected_count} of {total_count} items selected")

    def _get_file_type(self, file_item):
        """Determine file type display string."""
        if not file_item.is_file:
            return "Folder"
        
        
        
        for type_name, extensions in type_map.items():
            if file_item.extension in extensions:
                return type_name.title()
        
        return file_item.extension.upper().replace('.', '') + " File" if file_item.extension else "File"
    
    def on_dir_select(self, event=None):
        """Handle directory selection change."""
        self.refresh_items()
    
    def toggle_item_selection(self, event):
        """Toggle item selection when clicked."""
        item = self.item_tree.selection()[0] if self.item_tree.selection() else None
        if not item:
            return
        
        values = self.item_tree.item(item, 'values')
        if not values or len(values) < 2:
            return
        
        key_value = values[1]
        
        if self.input_mode_var.get() == 'folder':
            match_fn = lambda fi: fi.name == key_value
        else:
            if self.csv_headers:
                first_col = self.csv_headers[0]
                match_fn = lambda fi: fi.attributes.get(first_col, fi.name) == key_value
            else:
                match_fn = lambda fi: fi.name == key_value
        
        for file_item in self.files:
            if match_fn(file_item):
                file_item.selected = not file_item.selected
                break
        
        self.refresh_display()
    
    def proceed_to_phase2(self):
        """Move to Phase 2: Bucket Configuration."""
        selected_items = [f for f in self.files if f.selected]
        if not selected_items:
            messagebox.showwarning("No Selection", "Please select at least one item to sort.")
            return

        self.output_mode_var.set('folder' if self.input_mode_var.get() == 'folder' else 'list')
        self.on_output_mode_change()
        self.notebook.select(1)
    
    def refresh_items(self):
        """Refresh the item list based on selected directories."""
        # Preserve current selections by path
        previously_selected_paths = {str(f.path) for f in self.files if getattr(f, 'selected', False)}
        self.files.clear()
        for directory in self.dir_listbox.get(0, tk.END):
            try:
                path = Path(directory)
                if path.exists() and path.is_dir():
                    for item in path.iterdir():
                        if not item.name.startswith('.'):
                            fi = FileItem(str(item))
                            if str(fi.path) in previously_selected_paths:
                                fi.selected = True
                            self.files.append(fi)
            except PermissionError:
                messagebox.showwarning("Permission Error", f"Cannot access directory: {directory}")
        
        self.columns_mode = 'folder'
        self.configure_item_tree_columns(['checkbox', 'Filename', 'Size', 'Modified', 'Type'],
                                         headers_mapping={'checkbox': '✓'})
        self.refresh_display()
    
    def refresh_button_action(self):
        """Refresh action depending on current mode."""
        if self.input_mode_var.get() == 'folder':
            self.refresh_items()
        else:
            self.refresh_display()
    
    def clear_all_action(self):
        """Clear list items and selections, reset filter."""
        for f in self.files:
            f.selected = False
        self.files.clear()
        self.filter_var.set("*")
        self.show_selected_only.set(False)
        self.csv_headers = None
        if self.input_mode_var.get() == 'list':
            self.configure_item_tree_columns(['checkbox', 'Item'])
        self.refresh_display()
    
    def select_matching(self):
        """Select all files matching the current filter."""
        filter_pattern = self.filter_var.get().strip()
        if filter_pattern:
            for file_item in self.files:
                if fnmatch.fnmatch(file_item.name.lower(), filter_pattern.lower()):
                    file_item.selected = True
            self.refresh_display()
    
    def clear_selection(self):
        """Clear all file selections."""
        for file_item in self.files:
            file_item.selected = False
        self.refresh_display()
    
    def toggle_show_selected(self):
        """Toggle showing only selected items."""
        self.refresh_display()
    
    def apply_filter(self, *args):
        """Apply the current filter to the item display."""
        self.refresh_display()

    def init_phase2(self):
        """Initialize Phase 2: Bucket Configuration."""
        self.phase2_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.phase2_frame, text="2. Bucket Configuration")
        
        main_frame = ttk.Frame(self.phase2_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(main_frame, text="Customize your Buckets (Collections / Categories / Types / Folders)", font=('Arial', 14, 'bold')).pack(pady=10)
        
        # Header with add button
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=5)
        btn_add_bucket = ttk.Button(header_frame, text="Add New Bucket", command=self.add_new_bucket)
        btn_add_bucket.pack(side=tk.LEFT, padx=10)
        self._add_tooltip(btn_add_bucket, "Add a new category/folder to sort items into (max 10).")
        btn_reset_buckets = ttk.Button(header_frame, text="Reset All", command=self.reset_all_buckets)
        btn_reset_buckets.pack(side=tk.LEFT, padx=(10, 0))
        self._add_tooltip(btn_reset_buckets, "Remove all custom buckets and restore the defaults.")
        
        # Scrollable bucket config
        canvas = tk.Canvas(main_frame, bg=DARK_COLORS['bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        self.bucket_config_frame = ttk.Frame(canvas)
        
        self.bucket_config_frame.bind("<Configure>", 
                                      lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.bucket_config_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Output mode selection
        mode_frame = ttk.Frame(main_frame, width=50)
        mode_frame.pack(fill=tk.X, pady=10)
        ttk.Label(mode_frame, text="Output Mode:").pack(side=tk.LEFT)
        self.output_mode_var = tk.StringVar(value="list")
        rb_list_out = ttk.Radiobutton(mode_frame, text="List Mode", variable=self.output_mode_var,
                           value="list", command=self.on_output_mode_change)
        rb_list_out.pack(side=tk.LEFT, padx=10)
        self._add_tooltip(rb_list_out, "Output the sorting results as a file (e.g., CSV, JSON).")
        rb_folder_out = ttk.Radiobutton(mode_frame, text="Folder Mode (Move Files)", variable=self.output_mode_var,
                           value="folder", command=self.on_output_mode_change)
        rb_folder_out.pack(side=tk.LEFT, padx=10)
        self._add_tooltip(rb_folder_out, "Use this mode if you want to work with files and move them into folders")
        
        # Output directory selection (for folder mode)
        self.output_frame = ttk.Frame(main_frame)
        self.output_label = ttk.Label(self.output_frame, text="Output Directory:")
        self.output_label.pack(side=tk.LEFT)
        self.output_dir_var = tk.StringVar()
        self.output_entry = ttk.Entry(self.output_frame, textvariable=self.output_dir_var, width=50)
        self.output_entry.pack(side=tk.LEFT, padx=5)
        self.output_browse_btn = ttk.Button(self.output_frame, text="Browse", command=self.browse_output_directory)
        self.output_browse_btn.pack(side=tk.LEFT, padx=2)
        self._add_tooltip(self.output_browse_btn, "Select Output folder where files will be moved (create folders yourself)")
        
        # Always pack the frame to maintain consistent layout
        self.output_frame.pack(fill=tk.X, pady=10)
        
        self.update_bucket_config()
        self.add_new_bucket()  # Add second default bucket

        # Bottom buttons (keep a handle to ensure order relative to output_frame)
        self.phase2_buttons_frame = ttk.Frame(main_frame)
        self.phase2_buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Set initial state based on output mode
        self.on_output_mode_change()
        btn_back1 = ttk.Button(self.phase2_buttons_frame, text="Back to selection", command=self.back_to_phase1)
        btn_back1.pack(side=tk.LEFT, padx=2)
        self._add_tooltip(btn_back1, "Return to the previous step to change item selection.")
        btn_start_sorting = tk.Button(self.phase2_buttons_frame, text="Start Sorting", command=self.proceed_to_phase3, bg=DARK_COLORS['success'], fg="black", activebackground="#45a049", font=('Arial', 10, 'bold'))
        btn_start_sorting.pack(side=tk.RIGHT, padx=2)
        self._add_tooltip(btn_start_sorting, "Proceed to the next step to begin sorting the selected items.")

    def init_phase3(self):
        """Initialize Phase 3: Interactive Sorting."""
        self.phase3_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.phase3_frame, text="3. Interactive Sorting")
        
        main_frame = ttk.Frame(self.phase3_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.progress_var = tk.DoubleVar()

        item_display_frame = ttk.Frame(main_frame)
        item_display_frame.pack(fill=tk.BOTH, expand=True, pady=15)
        
        self.current_filename_label = ttk.Label(item_display_frame, text="",
                                                font=('Arial', 20, 'bold'),
                                                anchor='center', wraplength=1000)
        self.current_filename_label.pack(pady=12, ipady=8, fill=tk.X)
        
        self.side_by_side_frame = tk.PanedWindow(item_display_frame, orient=tk.HORIZONTAL,sashwidth=2, sashrelief='flat')
        self.side_by_side_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.info_frame = ttk.LabelFrame(self.side_by_side_frame, text="   File Information")
        self.side_by_side_frame.add(self.info_frame, width=400) 

        # Details Treeview
        self.details_tree = ttk.Treeview(self.info_frame, columns=('Property', 'Value'), show='headings')
        self.details_tree.heading('Property', text='Property')
        self.details_tree.heading('Value', text='Value')
        self.details_tree.column('Property', width=100, stretch=False)
        self.details_tree.column('Value', width=120, stretch=True)
        self.details_tree.pack(pady=4, padx=4, fill=tk.BOTH, expand=True)

        self.preview_frame = ttk.LabelFrame(self.side_by_side_frame, text="   Preview (Click to Open)")
        self.side_by_side_frame.add(self.preview_frame)
        
        self.preview_label = ttk.Label(self.preview_frame, text="No preview available", anchor='center' , cursor="hand2")
        self.preview_label.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)
        self.preview_label.bind("<Button-1>", self.on_preview_click)
        self.preview_frame.pack_propagate(False)
        
        self.bucket_buttons_frame = ttk.Frame(main_frame)
        self.bucket_buttons_frame.pack(pady=10)
        
        nav_frame = ttk.Frame(main_frame)
        nav_frame.pack(fill=tk.X, pady=5) 

        self.progress_bar = ttk.Progressbar(nav_frame, variable=self.progress_var, maximum=100, length=300)
        self.progress_label = ttk.Label(nav_frame, text="Ready to sort...")
       
        # Navigation with icon-like labels and tooltips
        self.first_btn = ttk.Button(nav_frame, text="⏮", width=4, style='Nav.TButton', command=lambda: self._jump_to_item(0))
        self.prev_btn = ttk.Button(nav_frame, text="⏪", width=4, style='Nav.TButton', command=self.previous_file)
        self.skip_btn = ttk.Button(nav_frame, text=">>", width=4, style='Nav.TButton', command=self.skip_file)
        self.next_btn = ttk.Button(nav_frame, text="⏩", width=4, style='Nav.TButton', command=self.next_file)
        self.last_btn = ttk.Button(nav_frame, text="⏭", width=4, style='Nav.TButton', command=lambda: self._jump_to_item(-1))
        self.save_btn = ttk.Button(nav_frame, text="💾", width=4, style='Nav.TButton', command=self.pause_sorting)
        self.complete_btn = tk.Button(nav_frame, text="Complete Sorting", command=self.proceed_to_phase4,
                                      bg=DARK_COLORS['success'], fg="black", activebackground="#45a049", font=('Arial', 10, 'bold'))

        for btn in [self.first_btn, self.prev_btn, self.skip_btn, self.next_btn, self.last_btn]:
            btn.pack(side=tk.LEFT, padx=3)
        self.save_btn.pack(side=tk.LEFT, padx=3, pady=1)

        self.complete_btn.pack(side=tk.RIGHT, padx=5)
        self._add_tooltip(self.complete_btn, "Finish sorting and proceed to the final review step.")

        # Skip sorted checkbox
        self.skip_sorted_var = tk.BooleanVar(value=False)
        skip_sorted_cb = ttk.Checkbutton(nav_frame, text="Skip Sorted",
                                        variable=self.skip_sorted_var)
        skip_sorted_cb.pack(side=tk.LEFT, pady=5, padx=8)
        self._add_tooltip(skip_sorted_cb, "Skip items already sorted or skipped")
        
        # Place progress bar and label in nav_frame
        self.progress_bar.pack(side=tk.LEFT, padx=5)
        self.progress_label.pack(side=tk.LEFT, padx=8)
        
        # Progress bar and label already packed earlier
        discard_btn = tk.Button(nav_frame, text="Discard Sorting", bg=DARK_COLORS['danger'], fg="black",
                                activebackground="#d32f2f", font=('Arial', 10, 'bold'), command=self.discard_sorting)
        discard_btn.pack(side=tk.RIGHT, padx=5)
        self._add_tooltip(discard_btn, "Reset all sorting progress for the current session.")
    
        self.add_new_bucket()

        # Tooltips
        self._add_tooltip(self.first_btn, "First item")
        self._add_tooltip(self.prev_btn, "Previous item")
        self._add_tooltip(self.skip_btn, "Skip: Mark item as skipped")
        self._add_tooltip(self.next_btn, "Next item")
        self._add_tooltip(self.last_btn, "Last item")
        self._add_tooltip(self.save_btn, "Save current progress")

    def init_phase4(self):
        """Initialize Phase 4: Review and Finalization."""
        self.phase4_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.phase4_frame, text="4. Review & Finalize")
        
        main_frame = ttk.Frame(self.phase4_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ttk.Label(main_frame, text="Review and Finalization", font=('Arial', 14, 'bold')).pack(pady=10)
        
        # Bucket review area
        self.review_notebook = ttk.Notebook(main_frame)
        self.review_notebook.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Statistics frame
        stats_frame = ttk.LabelFrame(main_frame, text="Statistics")
        stats_frame.pack(fill=tk.X, pady=10)
        self.stats_label = ttk.Label(stats_frame, text="Sorting not over")
        self.stats_label.pack(pady=10)
        
        # Final action buttons
        self.phase4_action_frame = ttk.Frame(main_frame)
        self.phase4_action_frame.pack(fill=tk.X, pady=10)
        
        btn_export = ttk.Button(self.phase4_action_frame, text="Export Results", command=self.export_results)
        btn_export.pack(side=tk.LEFT, padx=5)
        self._add_tooltip(btn_export, "Save the sorted lists to a file (JSON, CSV, TXT, HTML).")

        btn_save_session = ttk.Button(self.phase4_action_frame, text="Save Session", command=self.save_session)
        btn_save_session.pack(side=tk.LEFT, padx=5)
        self._add_tooltip(btn_save_session, "Save the entire session (selections, buckets, progress) to a file to resume later.")

        btn_refresh_review = ttk.Button(self.phase4_action_frame, text="Refresh", command=self.refresh_review)
        btn_refresh_review.pack(side=tk.LEFT, padx=5)
        self._add_tooltip(btn_refresh_review, "Update the review tabs with the latest sorting status.")
        
        # Add Discard Sorting button
        self.discard_phase4_btn = tk.Button(self.phase4_action_frame, text="Discard Sorting", bg=DARK_COLORS['danger'], fg="black",
                                activebackground="#d32f2f", font=('Arial', 10, 'bold'), command=self.discard_sorting)
        self.discard_phase4_btn.pack(side=tk.RIGHT, padx=10)
        self._add_tooltip(self.discard_phase4_btn, "Reset all sorting progress for the current session.")
        
        # Always create Move Files button; control visibility/state later (match style with discard)
        self.move_files_btn = tk.Button(self.phase4_action_frame, text="Move Files", command=self.execute_file_moves,
                                        bg=DARK_COLORS['success'], fg="black", activebackground="#45a049", font=('Arial', 10, 'bold'))
        self.move_files_btn.pack(side=tk.RIGHT, padx=5)
        self._add_tooltip(self.move_files_btn, "Physically move all sorted files to their corresponding bucket folders.")

    def handle_hotkey(self, event):
        """Handle hotkey presses for bucket selection and navigation."""
        if self.notebook.index(self.notebook.select()) != 2:  # Only in sorting phase
            return
            
        key_actions = {
            'Left': self.previous_file, 'Right': self.next_file, 
            'Up': lambda: self._jump_to_item(0),
            'Down': lambda: self._jump_to_item(-1),
            '\x1b': self.previous_file, '\x08': self.previous_file, '\x7f': self.previous_file,  # Esc, Backspace, Delete
            ' ': self.skip_file  # Space
        }
        
        if event.keysym in key_actions:
            key_actions[event.keysym]()
        elif event.char.isdigit():
            bucket_num = int(event.char) if event.char != '0' else 10
            bucket = next((b for b in self.buckets if b.number == bucket_num), None)
            if bucket:
                self.sort_to_bucket(bucket)

    def _jump_to_item(self, index):
        """Jump to specific item index."""
        selected_files = [f for f in self.files if f.selected]
        if selected_files:
            self.current_file_index = index if index >= 0 else len(selected_files) - 1
            self.show_current_file()

    def previous_file(self):
        """Go back to the previous file."""
        if self.current_file_index > 0:
            self.current_file_index -= 1
            self.show_current_file()
    
    def next_file(self):
        """Go to next file without changing any association."""
        selected_files = [f for f in self.files if f.selected]
        if self.current_file_index < len(selected_files) - 1:
            self.current_file_index += 1
        else:
            self.current_file_index = len(selected_files)
        self.show_current_file()
    
    def skip_file(self):
        """Skip the current file without sorting."""
        selected_files = [f for f in self.files if f.selected]
        if self.current_file_index < len(selected_files):
            current = selected_files[self.current_file_index]
            # Remove any existing bucket association
            if current.bucket:
                try:
                    current.bucket.items.remove(current)
                except ValueError:
                    pass
                current.bucket = None
            current.skipped = True
        self.current_file_index += 1
        self.show_current_file()

    def _add_tooltip(self, widget, text):
        """Simple tooltip for a widget with 600 delay."""
        tip = {'tw': None, 'after_id': None}
        
        def show_tip(event=None):
            # Cancel any existing scheduled tooltip
            if tip['after_id'] is not None:
                widget.after_cancel(tip['after_id'])
            
            tip['after_id'] = widget.after(600, create_tooltip)
        
        def create_tooltip():
            if tip['tw'] is not None:
                return
            x = widget.winfo_rootx() + 25
            y = widget.winfo_rooty() + widget.winfo_height() + 15
            tw = tk.Toplevel(widget)
            tw.wm_overrideredirect(True)
            tw.wm_geometry(f"+{x}+{y}")
            lbl = tk.Label(tw, text=text, fg="#333333", bg="#ffffff", relief=tk.SOLID, borderwidth=1, padx=6, pady=3)
            lbl.pack()
            tip['tw'] = tw
            tip['after_id'] = None
        
        def hide_tip(event=None):
            # Cancel scheduled tooltip if mouse leaves before delay
            if tip['after_id'] is not None:
                widget.after_cancel(tip['after_id'])
                tip['after_id'] = None
            
            # Hide existing tooltip
            if tip['tw'] is not None:
                tip['tw'].destroy()
                tip['tw'] = None
        
        widget.bind("<Enter>", show_tip)
        widget.bind("<Leave>", hide_tip)    
    
    def pause_sorting(self):
        """Pause sorting and save current state."""
        if messagebox.askyesno("Pause Sorting", "Do you want to save the current progress and pause?"):
            self.save_session(resume_to_last_item=True)

    def on_preview_click(self, event):
        """Handle click on preview to open file in associated application."""
        selected_files = [f for f in self.files if f.selected]
        if self.current_file_index < len(selected_files):
            current_file = selected_files[self.current_file_index]
            self._open_file(current_file.path)

    def _open_file(self, file_path):
        """Open file with system default application."""
        try:
            if os.path.exists(file_path):
                if os.name == 'nt':  # Windows
                    os.startfile(file_path)
                elif sys.platform == 'darwin':  # macOS
                    subprocess.call(['open', file_path])
                else:  # Linux
                    subprocess.call(['xdg-open', file_path])
            else:
                messagebox.showwarning("File Not Found", f"Cannot open file: {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file: {str(e)}")

    def update_preview(self, file_item):
        """Update the preview area for the current file."""
        # Clear existing preview
        for widget in self.preview_frame.winfo_children():
            if widget != self.preview_label:
                widget.destroy()
        
        # Clear the preview label completely
        self.preview_label.config(image='', text='', compound=tk.TOP)
        
        preview_handlers = {
            'text': ['.txt', '.py', '.md'],
            'image': ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp']
        }
        
        if file_item.is_file and file_item.extension in preview_handlers['text']:
            self._current_image_dimensions = None
            self._preview_text_file(file_item.path)
        elif file_item.is_file and file_item.extension in preview_handlers['image']:
            self._preview_image_file(file_item.path)
        else:
            self._current_image_dimensions = None
            self.preview_label.config(image='', anchor='center' , text="No preview available", foreground="#cccccc")

    def _preview_text_file(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read(1500)
                self.preview_label.config(text=f"{content}...", foreground="#ffffff")
        except Exception:
            self.preview_label.config(text="Cannot preview this file", anchor='center' , foreground="#cccccc")

    def _preview_image_file(self, file_path):
        try:
            self.preview_label.unbind("<Configure>")
            # Initial display
            self._resize_image(file_path)
            # Bind to configure event to handle resizing
            self.preview_label.bind("<Configure>", lambda e: self._resize_image(file_path))
        except Exception as e:
            self.preview_label.config(
                image='',
                anchor='center' ,
                text=f"Failed to load image: {str(e)}",
                foreground="#cccccc"
            )

    def _resize_image(self, file_path):
        """Resize image to fit preview label, without upscaling."""
        if self._resizing_image:
            return
        try:
            self._resizing_image = True
            self.preview_frame.update_idletasks()
            frame_width = self.preview_frame.winfo_width()
            frame_height = self.preview_frame.winfo_height()

            if frame_width <= 1 or frame_height <= 1:
                self._current_image_dimensions = None
                return

            img = Image.open(file_path)
            img_width, img_height = img.width, img.height
            self._current_image_dimensions = (img_width, img_height)
            if img_height == 0: return # Avoid division by zero
            aspect = img_width / img_height

            max_width = max(frame_width - 8, 10)
            max_height = max(frame_height - 8, 10)

            # Calculate the size to fit the frame
            if (max_width / max_height) > aspect:
                fit_height = max_height
                fit_width = int(fit_height * aspect)
            else:
                fit_width = max_width
                fit_height = int(fit_width / aspect)

            # Final dimensions are the smaller of the "fit" size and the original size
            final_width = min(fit_width, img_width)
            final_height = min(fit_height, img_height)

            # Only resize if the final dimensions are different from the original
            if final_width != img_width or final_height != img_height:
                display_img = img.resize((final_width, final_height), Image.LANCZOS)
            else:
                display_img = img

            photo = ImageTk.PhotoImage(display_img)
            self._preview_images.clear()
            self._preview_images.append(photo)
            self.preview_label.config(image=photo, text="", compound=tk.CENTER)

        except Exception as e:
            self.preview_label.config(
                image='',
                anchor='center',
                text=f"Failed to load image: {str(e)}",
                foreground="#cccccc"
            )
        finally:
            self._resizing_image = False


    def show_current_file(self):
        """Display the current file for sorting."""
        selected_files = [f for f in self.files if f.selected]
        
        # Hide preview frame in list mode; show in folder mode
        if self.columns_mode == 'list':
            try:
                self.side_by_side_frame.remove(self.preview_frame)
            except Exception:
                pass
        else:
            try:
                if self.preview_frame not in self.side_by_side_frame.panes():
                    self.side_by_side_frame.add(self.preview_frame, weight=2)
            except Exception:
                pass
        
        # Skip already sorted/skipped items if option is checked
        if self.skip_sorted_var.get():
            while self.current_file_index < len(selected_files):
                current_file = selected_files[self.current_file_index]
                if not current_file.bucket and not current_file.skipped:
                    break
                self.current_file_index += 1
        
        if self.current_file_index >= len(selected_files):
            self._show_completion()
            return
            
        current_file = selected_files[self.current_file_index]
        
        # Update display - use full name and let it wrap
        self.current_filename_label.config(text=current_file.name)
        self.info_frame.configure(text="Item Information" if self.columns_mode == 'list' else "File Information")
        
        # Update preview first to get image dimensions
        self.update_preview(current_file)

        # Rebuild details
        self._rebuild_details(current_file, selected_files)
        
        # Update progress
        progress = (self.current_file_index / len(selected_files)) * 100
        self.progress_var.set(progress)
        self.progress_label.config(text=f"{self.current_file_index} / {len(selected_files)}")
        
        # Update bucket indicators
        self._update_bucket_indicators(current_file.bucket)

    def _show_completion(self):
        """Show completion state."""
        total = len([f for f in self.files if f.selected])
        self.progress_var.set(100)
        self.progress_label.config(text=f"{total} / {total}")
        self.current_filename_label.config(text="All items sorted!", anchor='center' )
        
        for w in self.item_details_container.winfo_children():
            w.destroy()
        ttk.Label(self.item_details_container, 
                 text="Sorting is complete.\nPlease proceed to Review & Finalization.",
                 font=('Arial', 10)).pack(anchor='w')
        self.preview_label.config(image='', anchor='center' , text="Task Completed", foreground="#cccccc")
        self.complete_sorting()

    def _truncate_middle(self, s: str, max_len: int) -> str:
        """Truncate string with ellipses in middle."""
        if len(s) <= max_len:
            return s
        keep = max_len - 3
        left = keep // 2
        right = keep - left
        return s[:left] + '...' + s[-right:]

    def _rebuild_details(self, current_file, selected_files):
        """Rebuild the details Treeview."""
        # Clear existing items
        for item in self.details_tree.get_children():
            self.details_tree.delete(item)

        def add_kv(key, value):
            self.details_tree.insert('', 'end', values=(key, value))

        # Add details based on mode
        label_key = "Item" if self.columns_mode == 'list' else "Full Name"
        add_kv(label_key, current_file.name)
        
        if self.columns_mode == 'list':
            if self.csv_headers:
                for h in self.csv_headers:
                    add_kv(h, current_file.attributes.get(h, ""))
        else:
            add_kv("Path", current_file.path)
            size_str = f"{current_file.size:,} bytes" if current_file.is_file else "N/A"
            add_kv("Size", size_str)
            add_kv("Modified", current_file.modified.strftime('%Y-%m-%d %H:%M:%S'))
            add_kv("Type", self._get_file_type(current_file))

            if self._current_image_dimensions:
                try:
                    width, height = self._current_image_dimensions
                    megapixels = (width * height) / 1_000_000
                    add_kv("Resolution", f"{width} x {height}")
                    add_kv("Megapixels", f"{megapixels:.2f} MP")
                except Exception:
                    pass # Failsafe
        
        add_kv("Item", f"{self.current_file_index + 1} of {len(selected_files)}")
        add_kv("Bucket", current_file.bucket.name if current_file.bucket else "")

    def _update_bucket_indicators(self, active_bucket: Optional[Bucket]):
        """Update bucket selection indicators with thicker bars."""
        if not hasattr(self, 'bucket_indicator_map'):
            return
            
        active_color = DARK_COLORS['success']  # Use the same green as buttons
        inactive_color = DARK_COLORS['entry_bg']
        
        for bucket, indicator in self.bucket_indicator_map.items():
            color = active_color if bucket is active_bucket else inactive_color
            indicator.configure(bg=color)

    def sort_to_bucket(self, bucket):
        """Sort current file to the specified bucket."""
        selected_files = [f for f in self.files if f.selected]
        if self.current_file_index < len(selected_files):
            current_file = selected_files[self.current_file_index]
            # Remove from previous bucket
            if current_file.bucket:
                try:
                    current_file.bucket.items.remove(current_file)
                except ValueError:
                    pass
            # Add to new bucket
            bucket.items.append(current_file)
            current_file.bucket = bucket
            current_file.skipped = False
            self.current_file_index += 1
            self.show_current_file()

    def setup_sorting_interface(self):
        """Setup the sorting interface with bucket buttons."""
        for widget in self.bucket_buttons_frame.winfo_children():
            widget.destroy()
        
        self.bucket_button_map.clear()
        self.bucket_indicator_map.clear()
            
        # Create bucket buttons with thicker indicator bars
        for bucket in self.buckets:
            container = tk.Frame(self.bucket_buttons_frame, bg=DARK_COLORS['bg'])
            container.pack(side=tk.LEFT, padx=5)
            
            # Thicker indicator bar (increased from height=3 to height=8)
            indicator = tk.Frame(container, height=8, bg=DARK_COLORS['entry_bg'], width=100)
            indicator.pack(fill=tk.X, side=tk.TOP)
            
            btn = tk.Button(container, text=f"{bucket.hotkey}\n{bucket.name}",
                           background=bucket.color, foreground="#424242",
                           font=('Arial', 12, 'bold'), width=10, height=2, relief=tk.RAISED,
                           command=lambda b=bucket: self.sort_to_bucket(b))
            btn.pack(side=tk.TOP)
            
            self.bucket_button_map[bucket] = btn
            self.bucket_indicator_map[bucket] = indicator

    def proceed_to_phase3(self):
        """Move to Phase 3: Interactive Sorting."""
        selected_items = [f for f in self.files if f.selected]
        if not selected_items:
            messagebox.showwarning("No Items", "No items selected for sorting.")
            return
            
        if self.output_mode_var.get() == "folder" and not self.output_dir_var.get():
            messagebox.showwarning("No Output Directory", "Please select an output directory for folder mode.")
            return
            
        self.output_mode = self.output_mode_var.get()
        self.output_directory = self.output_dir_var.get()
        self.current_file_index = 0
        
        if not self.buckets:
            self.add_new_bucket()
            
        self.setup_sorting_interface()
        self.notebook.select(2)
        self.show_current_file()
        try:
            self.skip_btn.focus_set()
        except Exception:
            pass
    
    def back_to_phase1(self):
        """Return to Phase 1."""
        self.notebook.select(0)
    
    def on_output_mode_change(self):
        """Handle change in output mode."""
        self.output_mode = self.output_mode_var.get()

        if self.output_mode_var.get() == "folder":
            # Show output directory widgets in folder mode
            self.output_label.pack(side=tk.LEFT)
            self.output_entry.pack(side=tk.LEFT, padx=5)
            self.output_browse_btn.pack(side=tk.LEFT, padx=2)
            self.load_buckets_from_subfolders()
        else:
            # Hide output directory widgets in list mode but keep frame for consistent layout
            self.output_label.pack_forget()
            self.output_entry.pack_forget()
            self.output_browse_btn.pack_forget()
            self.update_bucket_config()
        # Also immediately toggle Move Files button visibility in Phase 4
        try:
            if self.output_mode_var.get() == 'folder':
                self.move_files_btn.configure(state=tk.NORMAL, text='Move Files', bg=DARK_COLORS['success'], fg='black')
                self.move_files_btn.pack(side=tk.RIGHT, padx=5)
            else:
                self.move_files_btn.pack_forget()
        except Exception:
            pass
    
    def load_buckets_from_subfolders(self):
        """Load subfolders as buckets when in folder mode."""
        output_dir = self.output_dir_var.get()
        if not output_dir:
            return
        
        try:
            subfolders = [f.name for f in os.scandir(output_dir) if f.is_dir()][:10]
            if len(subfolders) < 2:
                messagebox.showwarning("Not Enough Subfolders",
                    "Folder mode requires at least two subfolders in the output directory.")
                self.update_bucket_config()
                return
                
            self.buckets = [Bucket(idx + 1, name=name) for idx, name in enumerate(subfolders)]
            self.update_bucket_config()
            if len(subfolders) == 10:
                messagebox.showinfo("Subfolder Limit", "Only the first 10 subfolders will be used as buckets.")
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read output directory: {e}")
    
    def add_new_bucket(self):
        """Add a new bucket."""
        if len(self.buckets) >= 10:
            messagebox.showwarning("Limit Reached", "Maximum of 10 buckets allowed.")
            return
        
        next_number = len(self.buckets) + 1
        used_colors = {b.color.lower() for b in self.buckets}
        available_colors = [c for c in COLOR_PALETTE if c.lower() not in used_colors]
        color = available_colors[0] if available_colors else get_random_color()
        
        self.buckets.append(Bucket(next_number, f"Bucket {next_number}", color))
        self.update_bucket_config()
    
    def reset_all_buckets(self):
        """Remove all buckets and create two default Buckets"""
        if messagebox.askyesno("Reset All Buckets",
                               "This will remove all existing buckets. Continue?"):
            # Clear all existing buckets
            self.buckets.clear()

            # Create default Bucket 1 and 2
            self.buckets.append(Bucket(1, "Bucket 1", get_random_color()))
            self.buckets.append(Bucket(2, "Bucket 2", get_random_color()))

            # Update the configuration display
            self.update_bucket_config()
    
    def delete_bucket(self, bucket):
        """Delete a bucket."""
        if len(self.buckets) <= 2:
            messagebox.showwarning("Minimum Required", "At least two buckets are required.")
            return
        self.buckets.remove(bucket)
        self.update_bucket_config()
    
    def update_bucket_config(self):
        """Update the bucket configuration interface."""
        for widget in self.bucket_config_frame.winfo_children():
            widget.destroy()
            
        for bucket in self.buckets:
            bucket_frame = tk.LabelFrame(self.bucket_config_frame,
                                        text=f"Bucket {bucket.number} (Hotkey: {bucket.hotkey})", 
                                        bg=bucket.color, fg="#424242")
            bucket_frame.pack(fill=tk.X, padx=10, pady=5)
            
            name_frame = tk.Frame(bucket_frame, bg=bucket.color)
            name_frame.pack(fill=tk.X, padx=15, pady=8)
            tk.Label(name_frame, text="Name:", bg=bucket.color, fg="#424242").pack(side=tk.LEFT)
            
            name_var = tk.StringVar(value=bucket.name)
            name_var.trace('w', lambda *args, b=bucket, v=name_var: setattr(b, 'name', v.get()))
            name_entry = tk.Entry(name_frame, textvariable=name_var, width=40, 
                                 bg=DARK_COLORS['entry_bg'], fg='white', insertbackground='white')
            name_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            
            def choose_color(b=bucket, frame=bucket_frame, namef=name_frame):
                color_code = colorchooser.askcolor(title="Choose Bucket Color", initialcolor=b.color)
                if color_code[1]:
                    b.color = color_code[1]
                    frame.config(bg=b.color)
                    namef.config(bg=b.color)
                    for widget in namef.winfo_children():
                        if isinstance(widget, tk.Label):
                            widget.config(bg=b.color, fg="#424242")
            
            btn_color = tk.Button(name_frame, text="Color", command=choose_color, 
                     bg=DARK_COLORS['success'], fg="black")
            btn_color.pack(side=tk.RIGHT, padx=2)
            self._add_tooltip(btn_color, "Choose a custom color for this bucket.")
            btn_delete = tk.Button(name_frame, text="X", width=2, bg=DARK_COLORS['danger'], fg="black",
                     command=lambda b=bucket: self.delete_bucket(b))
            btn_delete.pack(side=tk.RIGHT, padx=2)
            self._add_tooltip(btn_delete, "Delete this bucket.")
            
        self.root.update_idletasks()
    
    def browse_output_directory(self):
        """Browse for output directory."""
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir_var.set(directory)
            if self.output_mode_var.get() == "folder":
                self.load_buckets_from_subfolders()
    
    def complete_sorting(self):
        """Complete the sorting process."""
        self.proceed_to_phase4()
        

    def discard_sorting(self):
        """Reset all sorting progress and refresh UI before returning to Phase 3."""
        if messagebox.askyesno(
            "Confirm Discard",
            "Are you sure you want to discard all sorting progress? This cannot be undone."
        ):
            # Reset item states
            for file_item in self.files:
                file_item.bucket = None
                file_item.skipped = False

            # Clear bucket contents
            for bucket in self.buckets:
                bucket.items.clear()

            # If currently on Phase 4, refresh the review UI so it reflects the cleared state
            try:
                if self.notebook.index(self.notebook.select()) == 3:
                    self.setup_review_interface()
            except Exception:
                # Safe fallback: do nothing if notebook state cannot be read
                pass

            # Reset progress and return to Phase 3
            self.current_file_index = 0
            self.setup_sorting_interface()
            self.notebook.select(2)
            self.show_current_file()
            messagebox.showinfo("Sorting Reset", "Sorting progress has been discarded. You can start sorting again.")

    def proceed_to_phase4(self):
        """Move to Phase 4: Review and Finalization."""
        self.setup_review_interface()
        self.notebook.select(3)

    def setup_review_interface(self):
        """Setup the review interface showing all buckets and skipped items."""
        for tab_id in self.review_notebook.tabs():
            self.review_notebook.forget(tab_id)
            
        # Create tabs for buckets with items
        for bucket in self.buckets:
            if bucket.items:
                self._create_bucket_tab(bucket)
        
        # Create skipped items tab
        skipped_items = [f for f in self.files if f.selected and f.skipped]
        if skipped_items:
            self._create_skipped_tab(skipped_items)
                
        self.update_statistics()
        
        # Control Move Files button visibility/state
        try:
            if self.output_mode == 'folder':
                self.move_files_btn.configure(state=tk.NORMAL, text='Move Files', bg=DARK_COLORS['success'], fg='black')
                self.move_files_btn.pack(side=tk.RIGHT, padx=5)
            else:
                self.move_files_btn.pack_forget()
        except Exception:
            pass
        
    def _review_columns_and_row_builder(self):
        """Return (columns, row_builder) based on current input mode."""
        input_mode = self.output_mode_var.get()

        if input_mode == 'list':
            if getattr(self, 'csv_headers', None):
                columns = tuple(self.csv_headers)
                def row_builder(item):
                    return [item.attributes.get(h, '') for h in self.csv_headers]
            else:
                columns = ('Item',)
                def row_builder(item):
                    return [item.name]
        else:
            columns = ('Filename', 'Path', 'Size', 'Modified', 'Type')
            def row_builder(item):
                type_str = self._get_file_type(item)
                size_str = f"{item.size:,} bytes" if getattr(item, 'size', 0) else "N/A"
                modified_str = item.modified.strftime("%Y-%m-%d %H:%M") if getattr(item, 'modified', None) else "N/A"
                return [item.name, str(item.path), size_str, modified_str, type_str]

        return columns, row_builder


    def _create_review_tab(self, title, items, context='bucket', bucket=None):
        """
        Build a single review tab for either a bucket or skipped items.
        context: 'bucket' | 'skipped'
        """
        tab_frame = ttk.Frame(self.review_notebook)
        self.review_notebook.add(tab_frame, text=title)

        columns, row_builder = self._review_columns_and_row_builder()

        # Use only vertical scroll - no horizontal scrollbar
        tree_frame, tree = self.create_tree_with_scrollbars(
            tab_frame, columns, show='headings', need_v=True, need_h=False, auto_hide_h=False
        )
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # Headings and column sizing policy - make columns fit within available width
        for col in columns:
            tree.heading(col, text=col)
            if col in ('Filename', 'Item'):
                tree.column(col, width=300, stretch=True)     # primary stretch column
            elif col == 'Path':
                tree.column(col, width=400, stretch=True)     # also stretch to use available space
            elif col in ('Type', 'Size', 'Modified'):
                tree.column(col, width=120, stretch=False)
            else:
                tree.column(col, width=100, stretch=False)

        # Populate rows
        for item in items:
            tree.insert('', 'end', values=row_builder(item))

        # Context menus
        if context == 'bucket':
            self._add_context_menu(tree, bucket)
        elif context == 'pending':
            self._add_pending_context_menu(tree)
        else:
            self._add_skipped_context_menu(tree)
        
        # Add double-click functionality
        tree.bind("<Double-1>", lambda event: self.open_file_from_tree(tree))

        return tab_frame


    def setup_review_interface(self):
        """Setup the review interface showing all buckets and skipped items."""
        for tab_id in self.review_notebook.tabs():
            self.review_notebook.forget(tab_id)

        # Create tabs for buckets with items
        for bucket in self.buckets:
            if bucket.items:
                self._create_review_tab(f"{bucket.name} ({len(bucket.items)})", bucket.items, context='bucket', bucket=bucket)

        # Create pending items tab (unsorted items)
        pending_items = [f for f in self.files if f.selected and not f.bucket and not f.skipped]
        if pending_items:
            self._create_review_tab(f"Pending ({len(pending_items)})", pending_items, context='pending')

        # Create skipped items tab
        skipped_items = [f for f in self.files if f.selected and f.skipped]
        if skipped_items:
            self._create_review_tab(f"Skipped ({len(skipped_items)})", skipped_items, context='skipped')
        
        self.update_statistics()

    def _add_context_menu(self, tree_widget, bucket):
        """Add context menu to bucket tree."""
        def on_right_click(event):
            item = tree_widget.selection()
            if item:
                menu = tk.Menu(self.root, tearoff=0, bg=DARK_COLORS['entry_bg'], fg='white')
                # Only show "Open File" in folder mode, not in list mode
                if self.output_mode_var.get() == 'folder':
                    menu.add_command(label="Open File", command=lambda: self.open_file_from_tree(tree_widget))
                menu.add_command(label="Remove from Bucket", 
                               command=lambda: self.remove_from_bucket(tree_widget, bucket))
                menu.tk_popup(event.x_root, event.y_root)
        
        tree_widget.bind("<Button-3>", on_right_click)  # Windows/Linux
        tree_widget.bind("<Button-2>", on_right_click)  # macOS

    def _add_skipped_context_menu(self, tree_widget):
        """Add context menu to skipped items tree."""
        def on_right_click(event):
            item = tree_widget.selection()
            if item:
                menu = tk.Menu(self.root, tearoff=0, bg=DARK_COLORS['entry_bg'], fg='white')
                # Only show "Open File" in folder mode, not in list mode
                if self.output_mode_var.get() == 'folder':
                    menu.add_command(label="Open File", command=lambda: self.open_file_from_tree(tree_widget))
                menu.add_command(label="Un-skip Item", command=lambda: self.unskip_item(tree_widget))
                menu.tk_popup(event.x_root, event.y_root)
        
        tree_widget.bind("<Button-3>", on_right_click)
        tree_widget.bind("<Button-2>", on_right_click)

    def _add_pending_context_menu(self, tree_widget):
        """Add context menu to pending items tree."""
        def on_right_click(event):
            item = tree_widget.selection()
            if item:
                menu = tk.Menu(self.root, tearoff=0, bg=DARK_COLORS['entry_bg'], fg='white')
                # Only show "Open File" in folder mode, not in list mode
                if self.output_mode_var.get() == 'folder':
                    menu.add_command(label="Open File", command=lambda: self.open_file_from_tree(tree_widget))
                menu.add_command(label="Skip Item", command=lambda: self.skip_pending_item(tree_widget))
                # No "Remove from Bucket" for pending items
                menu.tk_popup(event.x_root, event.y_root)

        tree_widget.bind("<Button-3>", on_right_click)
        tree_widget.bind("<Button-2>", on_right_click)

    def skip_pending_item(self, tree_widget):
        """Skip selected item from the pending list."""
        selection = tree_widget.selection()
        if not selection:
            return
        
        values = tree_widget.item(selection[0], 'values')
        if values and len(values) > 0:
            first_value = values[0]
            for file_item in self.files:
                if file_item.selected and not file_item.bucket and not file_item.skipped and file_item.name == first_value:
                    file_item.skipped = True
                    break
            self.setup_review_interface()

    def unskip_item(self, tree_widget):
        """Un-skip selected item from the skipped list."""
        selection = tree_widget.selection()
        if not selection:
            return
        
        values = tree_widget.item(selection[0], 'values')
        if values and len(values) > 0:
            file_path = values[0]
            for file_item in self.files:
                if str(file_item.path) == file_path:
                    file_item.skipped = False
                    break
            self.setup_review_interface()

    def update_statistics(self):
        """Update the statistics display."""
        total_items = sum(len(bucket.items) for bucket in self.buckets)
        total_skipped = sum(1 for f in self.files if f.selected and f.skipped)
        total_pending = sum(1 for f in self.files if f.selected and not f.bucket and not f.skipped)
        total_buckets = len([bucket for bucket in self.buckets if bucket.items])
        
        stats_text = f"Total Items Sorted:  {total_items}\n"
        stats_text += f"Total Pending Items: {total_pending}\n"
        stats_text += f"Total Skipped Items: {total_skipped}\n"
        stats_text += f"Buckets Used:        {total_buckets}\n"
        
        for bucket in self.buckets:
            if bucket.items:
                stats_text += f"    {bucket.name}: {len(bucket.items)} items\n"
        
        self.stats_label.config(text=stats_text)
    
    def open_file_from_tree(self, tree_widget):
        """Open file from tree selection."""
        selection = tree_widget.selection()
        if selection:
            values = tree_widget.item(selection[0], 'values')
            if values and len(values) > 0:
                # Try to find the actual file item from the tree selection
                file_item = None
                
                # Get the current tab to determine context
                current_tab = self.review_notebook.tab(self.review_notebook.select(), 'text')
                
                # Find the file item by matching the first column value
                first_value = values[0]
                
                # Search through all buckets and skipped items
                for bucket in self.buckets:
                    for item in bucket.items:
                        if item.name == first_value:
                            file_item = item
                            break
                    if file_item:
                        break
                
                # If not found in buckets, check skipped items
                if not file_item:
                    for item in self.files:
                        if item.selected and item.skipped and item.name == first_value:
                            file_item = item
                            break
                
                # Use the file item's path if found, otherwise fall back to the old method
                if file_item and hasattr(file_item, 'path'):
                    self._open_file(str(file_item.path))
                else:
                    # Fallback to old method
                    if self.output_mode_var.get() == 'list':
                        # In list mode, we can't open files as they might not be actual file paths
                        messagebox.showwarning("Cannot Open", "This item cannot be opened as a file.")
                    else:
                        # In folder mode, second column is Path
                        path_value = values[1] if len(values) > 1 else values[0]
                        self._open_file(path_value)
    
    def remove_from_bucket(self, tree_widget, bucket):
        """Remove selected item from bucket and move to pending."""
        selection = tree_widget.selection()
        if not selection:
            return
        
        values = tree_widget.item(selection[0], 'values')
        if values and len(values) > 0:
            first_value = values[0]
            for file_item in bucket.items[:]:
                if file_item.name == first_value:
                    bucket.items.remove(file_item)
                    file_item.bucket = None
                    break
            self.setup_review_interface()
    
    def export_results(self):
        """Export sorting results to various formats."""
        if not any(bucket.items for bucket in self.buckets) and not any(f.skipped for f in self.files):
            messagebox.showwarning("No Data", "No sorted or skipped items to export.")
            return
        
        # Export format dialog
        export_dialog = tk.Toplevel(self.root)
        export_dialog.title("Export Results")
        export_dialog.geometry("400x300")
        export_dialog.transient(self.root)
        export_dialog.grab_set()
        export_dialog.configure(bg=DARK_COLORS['bg'])
        
        # Center dialog
        export_dialog.update_idletasks()
        w, h = export_dialog.winfo_width(), export_dialog.winfo_height()
        x = (export_dialog.winfo_screenwidth() - w) // 2
        y = (export_dialog.winfo_screenheight() - h) // 2
        export_dialog.geometry(f'{w}x{h}+{x}+{y}')
        
        tk.Label(export_dialog, text="Select Export Format:", font=('Arial', 12, 'bold'),
                bg=DARK_COLORS['bg'], fg='white').pack(pady=10)
                 
        export_format = tk.StringVar(value="json")
        formats = [("JSON", "json"), ("CSV", "csv"), ("Text Report", "txt"), ("HTML Report", "html")]
        for text, value in formats:
            tk.Radiobutton(export_dialog, text=text, variable=export_format, value=value,
                          bg=DARK_COLORS['bg'], fg='white', selectcolor=DARK_COLORS['entry_bg']).pack(anchor=tk.W, padx=20, pady=5)
        
        def do_export():
            format_type = export_format.get()
            file_types = {
                "json": [("JSON files", "*.json")], "csv": [("CSV files", "*.csv")],
                "txt": [("Text files", "*.txt")], "html": [("HTML files", "*.html")]
            }
            filename = filedialog.asksaveasfilename(
                title="Save Export File", filetypes=file_types[format_type],
                defaultextension=f".{format_type}")
            if filename:
                try:
                    export_methods = {
                        "json": self.export_json, "csv": self.export_csv,
                        "txt": self.export_txt, "html": self.export_html
                    }
                    export_methods[format_type](filename)
                    messagebox.showinfo("Export Complete", f"Results exported to {filename}")
                    export_dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Export Error", f"Failed to export: {str(e)}")
        
        tk.Button(export_dialog, text="Export", command=do_export, 
                 bg=DARK_COLORS['success'], fg="white").pack(pady=20)
        tk.Button(export_dialog, text="Cancel", command=export_dialog.destroy, 
                 bg=DARK_COLORS['danger'], fg="white").pack()
    
    def refresh_review(self):
        """Refresh the review interface."""
        # Preserve current tab selection
        try:
            current_tab = self.review_notebook.tab(self.review_notebook.select(), 'text')
        except Exception:
            current_tab = None
        self.setup_review_interface()
        if current_tab:
            for tab_id in self.review_notebook.tabs():
                if self.review_notebook.tab(tab_id, 'text') == current_tab:
                    self.review_notebook.select(tab_id)
                    break
    
    def save_session(self, resume_to_last_item: bool = False):
        """Save the current session state."""
        session_data = {
            "version": "1.0", "save_date": datetime.now().isoformat(),
            "current_phase": self.current_phase, "directories": list(self.dir_listbox.get(0, tk.END)),
            "files": [], "buckets": [], "current_file_index": self.current_file_index,
            "output_mode": self.output_mode, "output_directory": self.output_directory,
            "resume_to_last_item": bool(resume_to_last_item)
        }
        
        # Save selected file data
        for file_item in self.files:
            if file_item.selected:
                file_data = {
                    "path": str(file_item.path), "selected": file_item.selected,
                    "bucket_number": file_item.bucket.number if file_item.bucket else None,
                    "skipped": file_item.skipped
                }
                if file_item.attributes:
                    file_data["attributes"] = file_item.attributes
                session_data["files"].append(file_data)
        
        # Save bucket data
        for bucket in self.buckets:
            session_data["buckets"].append({
                "number": bucket.number, "name": bucket.name, "color": bucket.color
            })
            
        filename = filedialog.asksaveasfilename(
            title="Save Session", filetypes=[("JSON files", "*.json")], defaultextension=".json")
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(session_data, f, indent=2)
                messagebox.showinfo("Session Saved", f"Session saved to {filename}")
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save session: {str(e)}")
    
    def load_session(self):
        """Load a saved session."""
        filename = filedialog.askopenfilename(
            title="Load Session", filetypes=[("JSON files", "*.json")])
        if not filename:
            return
        
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                session_data = json.load(f)
            
            # Restore state
            self.current_phase = session_data.get("current_phase", 1)
            self.current_file_index = session_data.get("current_file_index", 0)
            self.output_mode = session_data.get("output_mode", "list")
            self.output_directory = session_data.get("output_directory", "")
            
            # Restore directories
            self.dir_listbox.delete(0, tk.END)
            for directory in session_data.get("directories", []):
                self.dir_listbox.insert(tk.END, directory)
                
            # Restore files
            self.files.clear()
            for file_data in session_data.get("files", []):
                if os.path.exists(file_data["path"]):
                    file_item = FileItem(file_data["path"])
                    file_item.selected = file_data.get("selected", False)
                    file_item.skipped = file_data.get("skipped", False)
                    if "attributes" in file_data:
                        file_item.attributes = file_data["attributes"]
                    self.files.append(file_item)
                    
            # Restore buckets
            bucket_data_list = session_data.get("buckets", [])
            if bucket_data_list:
                self.buckets = [Bucket(bd["number"], bd["name"], bd["color"]) for bd in bucket_data_list]
                
                # Restore file-bucket assignments
                for file_data in session_data.get("files", []):
                    if file_data.get("bucket_number"):
                        for file_item in self.files:
                            if str(file_item.path) == file_data["path"]:
                                bucket = next((b for b in self.buckets 
                                             if b.number == file_data["bucket_number"]), None)
                                if bucket:
                                    file_item.bucket = bucket
                                    bucket.items.append(file_item)
                                    
            # Update UI
            self.output_mode_var.set(self.output_mode)
            self.output_dir_var.set(self.output_directory)
            self.refresh_display()
            self.update_bucket_config()
            
            # Resume to last item if requested
            if session_data.get("resume_to_last_item"):
                self.setup_sorting_interface()
                self.notebook.select(2)
                self.show_current_file()
            
         #   messagebox.showinfo("Session Loaded", f"Session loaded from {filename}")
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load session: {str(e)}")

    def export_json(self, filename):
        """Export results to JSON format."""
        export_data = {
            "export_date": datetime.now().isoformat(),
            "buckets": [], "skipped_items": []
        }
        
        for bucket in self.buckets:
            bucket_data = {"name": bucket.name, "color": bucket.color, "items": []}
            for item in bucket.items:
                item_data = {
                    "name": item.name, "path": str(item.path), "size": item.size,
                    "modified": item.modified.isoformat(), "is_file": item.is_file
                }
                if item.attributes:
                    item_data["attributes"] = item.attributes
                bucket_data["items"].append(item_data)
            export_data["buckets"].append(bucket_data)
        
        # Export skipped items
        for item in [f for f in self.files if f.selected and f.skipped]:
            item_data = {
                "name": item.name, "path": str(item.path), "size": item.size,
                "modified": item.modified.isoformat(), "is_file": item.is_file
            }
            if item.attributes:
                item_data["attributes"] = item.attributes
            export_data["skipped_items"].append(item_data)
        
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(export_data, f, indent=2)
    
    def export_csv(self, filename):
        """Export results to CSV format."""
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Category', 'Bucket', 'Item Name', 'Full Path', 'Size', 'Modified', 'Type'])
            
            for bucket in self.buckets:
                for item in bucket.items:
                    size_str = f"{item.size:,}" if item.is_file else "N/A"
                    modified_str = item.modified.strftime("%Y-%m-%d %H:%M:%S")
                    type_str = "File" if item.is_file else "Folder"
                    writer.writerow(['Sorted', bucket.name, item.name, str(item.path), 
                                   size_str, modified_str, type_str])
            
            for item in [f for f in self.files if f.selected and f.skipped]:
                size_str = f"{item.size:,}" if item.is_file else "N/A"
                modified_str = item.modified.strftime("%Y-%m-%d %H:%M:%S")
                type_str = "File" if item.is_file else "Folder"
                writer.writerow(['Skipped', 'N/A', item.name, str(item.path), 
                               size_str, modified_str, type_str])
    
    def export_txt(self, filename):
        """Export results to text format."""
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(f"SortAnything Results\n")
            f.write(f"Export Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 50 + "\n\n")
            
            for bucket in self.buckets:
                if bucket.items:
                    f.write(f"BUCKET: {bucket.name}\n" + "-" * 30 + "\n")
                    for item in bucket.items:
                        f.write(f"  {item.name}\n    Path: {item.path}\n")
                        if item.is_file:
                            f.write(f"    Size: {item.size:,} bytes\n")
                        f.write(f"    Modified: {item.modified.strftime('%Y-%m-%d %H:%M:%S')}\n")
                        if item.attributes:
                            for key, value in item.attributes.items():
                                f.write(f"    {key}: {value}\n")
                        f.write("\n")
                    f.write("\n")
            
            skipped_items = [f for f in self.files if f.selected and f.skipped]
            if skipped_items:
                f.write("SKIPPED ITEMS\n" + "-" * 30 + "\n")
                for item in skipped_items:
                    f.write(f"  {item.name}\n    Path: {item.path}\n")
                    if item.is_file:
                        f.write(f"    Size: {item.size:,} bytes\n")
                    f.write(f"    Modified: {item.modified.strftime('%Y-%m-%d %H:%M:%S')}\n")
                    if item.attributes:
                        for key, value in item.attributes.items():
                            f.write(f"    {key}: {value}\n")
                    f.write("\n")
    
    def export_html(self, filename):
        """Export results to HTML format."""
        html_template = f"""<!DOCTYPE html>
<html>
<head>
    <title>SortAnything Results</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; background-color: {DARK_COLORS['bg']}; color: {DARK_COLORS['fg']}; }}
        .bucket {{ margin-bottom: 30px; border: 1px solid #555; padding: 15px; }}
        .bucket-header {{ background-color: {DARK_COLORS['entry_bg']}; padding: 10px; margin: -15px -15px 15px -15px; }}
        .item {{ margin-bottom: 10px; padding: 10px; background-color: {DARK_COLORS['entry_bg']}; }}
        .item-name {{ font-weight: bold; }}
        .item-details {{ color: #ccc; font-size: 0.9em; }}
        .skipped-section {{ margin-top: 30px; }}
    </style>
</head>
<body>
    <h1>SortAnything Results</h1>
    <p>Export Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
"""
        
        for bucket in self.buckets:
            if bucket.items:
                html_template += f"""
    <div class="bucket">
        <div class="bucket-header" style="background-color: {bucket.color};">
            <h2>{bucket.name} ({len(bucket.items)} items)</h2>
        </div>
"""
                for item in bucket.items:
                    size_str = f"{item.size:,} bytes" if item.is_file else "Folder"
                    html_template += f"""
        <div class="item">
            <div class="item-name">{item.name}</div>
            <div class="item-details">
                Path: {item.path}<br>
                Size: {size_str}<br>
                Modified: {item.modified.strftime('%Y-%m-%d %H:%M:%S')}
"""
                    if item.attributes:
                        for key, value in item.attributes.items():
                            html_template += f"<br>{key}: {value}"
                    html_template += "            </div>\n        </div>\n"
                html_template += "    </div>\n"
        
        skipped_items = [f for f in self.files if f.selected and f.skipped]
        if skipped_items:
            html_template += f"""
    <div class="skipped-section">
        <div class="bucket">
            <div class="bucket-header">
                <h2>Skipped Items ({len(skipped_items)} items)</h2>
            </div>
"""
            for item in skipped_items:
                size_str = f"{item.size:,} bytes" if item.is_file else "Folder"
                html_template += f"""
            <div class="item">
                <div class="item-name">{item.name}</div>
                <div class="item-details">
                    Path: {item.path}<br>
                    Size: {size_str}<br>
                    Modified: {item.modified.strftime('%Y-%m-%d %H:%M:%S')}
"""
                if item.attributes:
                    for key, value in item.attributes.items():
                        html_template += f"<br>{key}: {value}"
                html_template += "                </div>\n            </div>\n"
            html_template += "        </div>\n    </div>\n"
        
        html_template += "</body>\n</html>"
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(html_template)
    
    def execute_file_moves(self):
        """Execute the actual file moves to bucket folders."""
        if self.output_mode != "folder":
            messagebox.showwarning("Invalid Mode", "File moves are only available in folder mode.")
            return
        
        if not self.output_directory:
            messagebox.showwarning("No Output Directory", "Please select an output directory.")
            return
        
        total_moves = sum(len(bucket.items) for bucket in self.buckets)
        if not messagebox.askyesno("Confirm File Moves",
            f"This will move {total_moves} files to their respective bucket folders. "
            "This action cannot be easily undone. Continue?"):
            return
        
        try:
            moved_count = 0
            for bucket in self.buckets:
                if not bucket.items:
                    continue
                
                bucket_path = os.path.join(self.output_directory, bucket.name)
                os.makedirs(bucket_path, exist_ok=True)
                
                for item in bucket.items:
                    if item.is_file and os.path.exists(item.path):
                        dest_path = os.path.join(bucket_path, item.name)
                        counter = 1
                        original_dest = dest_path
                        while os.path.exists(dest_path):
                            name_part, ext_part = os.path.splitext(original_dest)
                            dest_path = f"{name_part}_{counter}{ext_part}"
                            counter += 1
                        
                        shutil.move(item.path, dest_path)
                        moved_count += 1
            
            messagebox.showinfo("Move Complete", f"Successfully moved {moved_count} files.")
            # Update button state
            try:
                self.move_files_btn.configure(text='Files Moved!', state=tk.DISABLED)
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Move Error", f"Error during file moves: {str(e)}")
            try:
                self.move_files_btn.configure(text='Error Moving Files', state=tk.DISABLED)
            except Exception:
                pass


def create_menu_bar(root, app):
    """Create consistent menu bar for all windows."""
    menubar = tk.Menu(root, bg=DARK_COLORS['entry_bg'], fg='white', font=('Arial', 14), tearoff=0)
    root.config(menu=menubar)



    # File menu
    file_menu = tk.Menu(menubar, tearoff=40, bg=DARK_COLORS['entry_bg'], fg='white')
    menubar.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="New Session", command=lambda: open_new_session(root))
    file_menu.add_command(label="Load Session", command=app.load_session)
    file_menu.add_command(label="Save Session", command=app.save_session)
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=root.quit)

    # About
    menubar.add_command(label="About", command=lambda: messagebox.showinfo("About", ABOUT_TEXT))

def open_new_session(parent_root):
    """Open a new session window with consistent menus."""
    new_win = tk.Toplevel(parent_root)
    new_win.title("SortAnything")
    new_win.configure(bg=DARK_COLORS['bg'])
    app = FileSorterApp(new_win)
    create_menu_bar(new_win, app)

def main():
    """Main function to run the application."""
    root = tk.Tk()
    root.title("SortAnything")
    root.configure(bg=DARK_COLORS['bg'])
    
    configure_dark_theme(root)
    
    app = FileSorterApp(root)
    create_menu_bar(root, app)
    
    root.mainloop()

if __name__ == "__main__":
    main()
