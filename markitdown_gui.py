#!/usr/bin/env python3
"""
MarkItDown GUI - Universal Document to Markdown Converter
Powered by Microsoft MarkItDown
"""

import os
import sys
import threading
import subprocess
import traceback
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Ensure local package can be imported whether run from source or compiled
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(CURRENT_DIR, "packages", "markitdown", "src")
if os.path.isdir(SRC_DIR) and SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

# Import MarkItDown backend
try:
    from markitdown import MarkItDown
except ImportError:
    MarkItDown = None

# Attempt to load TkinterDnD for drag-and-drop support
HAS_DND = False
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except Exception:
    HAS_DND = False
    from tkinter import Tk as TkinterDnD


# Supported formats mapping
SUPPORTED_EXTENSIONS = {
    # Documents
    ".pdf": ("PDF Document", "Documents"),
    ".docx": ("Word Document", "Documents"),
    ".doc": ("Legacy Word Document", "Documents"),
    ".pptx": ("PowerPoint Presentation", "Documents"),
    ".ppt": ("Legacy PowerPoint Presentation", "Documents"),
    ".xlsx": ("Excel Spreadsheet", "Spreadsheets"),
    ".xls": ("Excel 97-2004 Spreadsheet", "Spreadsheets"),
    # Data & Code
    ".csv": ("Comma-Separated Values", "Data"),
    ".tsv": ("Tab-Separated Values", "Data"),
    ".json": ("JSON Document", "Data"),
    ".xml": ("XML Document", "Data"),
    ".rss": ("RSS Feed", "Data"),
    ".atom": ("Atom Feed", "Data"),
    ".ipynb": ("Jupyter Notebook", "Code"),
    # Web & Text
    ".html": ("HTML Document", "Web & Text"),
    ".htm": ("HTML Document", "Web & Text"),
    ".txt": ("Plain Text", "Web & Text"),
    ".text": ("Plain Text", "Web & Text"),
    ".log": ("Log File", "Web & Text"),
    ".md": ("Markdown Document", "Web & Text"),
    ".rtf": ("Rich Text Format", "Web & Text"),
    # Media & Audio
    ".mp3": ("MP3 Audio", "Media & Audio"),
    ".wav": ("WAV Audio", "Media & Audio"),
    ".m4a": ("M4A Audio", "Media & Audio"),
    ".flac": ("FLAC Audio", "Media & Audio"),
    ".ogg": ("OGG Audio", "Media & Audio"),
    ".wma": ("WMA Audio", "Media & Audio"),
    ".aac": ("AAC Audio", "Media & Audio"),
    # Images
    ".jpg": ("JPEG Image", "Images"),
    ".jpeg": ("JPEG Image", "Images"),
    ".png": ("PNG Image", "Images"),
    ".bmp": ("Bitmap Image", "Images"),
    ".tiff": ("TIFF Image", "Images"),
    ".webp": ("WebP Image", "Images"),
    # Books & Archives
    ".epub": ("EPUB E-Book", "Archives & Books"),
    ".zip": ("ZIP Archive (Multi-file)", "Archives & Books"),
    ".msg": ("Outlook Email Message", "Email"),
}


def format_size(bytes_num):
    """Format bytes into human readable format."""
    for unit in ["B", "KB", "MB", "GB"]:
        if bytes_num < 1024.0:
            return f"{bytes_num:.1f} {unit}"
        bytes_num /= 1024.0
    return f"{bytes_num:.1f} TB"


class MarkItDownApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MarkItDown - Universal Document to Markdown Converter")
        self.root.geometry("880x820")
        self.root.minsize(720, 650)

        # Initialize MarkItDown instance
        self.md_instance = None
        self._init_backend()

        # State tracking
        self.file_queue = []  # List of dicts representing queued items
        self.item_counter = 0
        self.is_converting = False
        self.cancel_requested = False
        self.active_preview_path = None
        self.auto_convert_var = tk.BooleanVar(value=True)
        self.output_same_dir_var = tk.BooleanVar(value=True)
        self.custom_output_dir_var = tk.StringVar(value="")
        self.word_wrap_var = tk.BooleanVar(value=True)
        self.show_formats_var = tk.BooleanVar(value=True)

        # Configure styles
        self._setup_styles()

        # Build UI layout
        self._build_ui()

        # Setup Drag and Drop
        self._setup_dnd()

    def _init_backend(self):
        """Initialize MarkItDown engine safely in background or main thread."""
        try:
            if MarkItDown:
                self.md_instance = MarkItDown(enable_builtins=True)
        except Exception as e:
            print(f"Warning: Error initializing MarkItDown: {e}")
            self.md_instance = None

    def _setup_styles(self):
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        # Progress bar style
        self.style.configure(
            "Custom.Horizontal.TProgressbar",
            background="#2563eb",
            troughcolor="#e2e8f0",
            bordercolor="#cbd5e1",
            lightcolor="#3b82f6",
            darkcolor="#1d4ed8",
        )

        # Treeview styling
        self.style.configure(
            "Queue.Treeview",
            font=("Segoe UI", 9),
            rowheight=26,
            background="#ffffff",
            fieldbackground="#ffffff",
        )
        self.style.configure(
            "Queue.Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            background="#f1f5f9",
            foreground="#1e293b",
        )
        self.style.map(
            "Queue.Treeview",
            background=[("selected", "#dbeafe")],
            foreground=[("selected", "#1e3a8a")],
        )

    def _build_ui(self):
        # Configure grid expansion
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(3, weight=1)  # Queue frame expands
        self.root.rowconfigure(5, weight=2)  # Preview frame expands

        # =====================================================================
        # 1. Header Frame
        # =====================================================================
        header_frame = tk.Frame(self.root, bg="#0f172a", padx=16, pady=12)
        header_frame.grid(row=0, column=0, sticky="ew")
        header_frame.columnconfigure(0, weight=1)

        title_label = tk.Label(
            header_frame,
            text="MarkItDown Converter",
            font=("Segoe UI", 15, "bold"),
            fg="#f8fafc",
            bg="#0f172a",
        )
        title_label.grid(row=0, column=0, sticky="w")

        subtitle_label = tk.Label(
            header_frame,
            text="Universal Document & Media to Markdown • Drag & drop files to convert instantly",
            font=("Segoe UI", 9),
            fg="#94a3b8",
            bg="#0f172a",
        )
        subtitle_label.grid(row=1, column=0, sticky="w")

        # =====================================================================
        # 2. Supported Formats Frame (Collapsible)
        # =====================================================================
        self.formats_container = ttk.LabelFrame(
            self.root,
            text="📋 Supported File Formats (20+ Formats)",
            padding="8 6 8 8",
        )
        self.formats_container.grid(row=1, column=0, sticky="ew", padx=12, pady=(8, 4))
        self.formats_container.columnconfigure((0, 1, 2, 3), weight=1)

        # Column 1: Documents
        f1 = ttk.Frame(self.formats_container)
        f1.grid(row=0, column=0, sticky="nsew", padx=4)
        ttk.Label(f1, text="📄 Office & Docs", font=("Segoe UI", 9, "bold"), foreground="#0369a1").pack(anchor="w")
        ttk.Label(f1, text="• PDF (.pdf)\n• Word (.docx, .doc)\n• PowerPoint (.pptx, .ppt)\n• Excel (.xlsx, .xls)", font=("Segoe UI", 8)).pack(anchor="w")

        # Column 2: Data & Code
        f2 = ttk.Frame(self.formats_container)
        f2.grid(row=0, column=1, sticky="nsew", padx=4)
        ttk.Label(f2, text="📊 Data & Tables", font=("Segoe UI", 9, "bold"), foreground="#047857").pack(anchor="w")
        ttk.Label(f2, text="• CSV & TSV (.csv, .tsv)\n• JSON (.json)\n• XML & RSS (.xml, .rss)\n• Jupyter (.ipynb)", font=("Segoe UI", 8)).pack(anchor="w")

        # Column 3: Web & Text
        f3 = ttk.Frame(self.formats_container)
        f3.grid(row=0, column=2, sticky="nsew", padx=4)
        ttk.Label(f3, text="🌐 Web & Text", font=("Segoe UI", 9, "bold"), foreground="#b45309").pack(anchor="w")
        ttk.Label(f3, text="• HTML (.html, .htm)\n• Plain Text (.txt, .text)\n• Log Files (.log)\n• Markdown & RTF", font=("Segoe UI", 8)).pack(anchor="w")

        # Column 4: Media & Archives
        f4 = ttk.Frame(self.formats_container)
        f4.grid(row=0, column=3, sticky="nsew", padx=4)
        ttk.Label(f4, text="🎵 Media & Archives", font=("Segoe UI", 9, "bold"), foreground="#6d28d9").pack(anchor="w")
        ttk.Label(f4, text="• Audio (.mp3, .wav, .m4a)\n• Images (.jpg, .png, .webp)\n• ZIP Archives (.zip)\n• EPUB (.epub) & Outlook (.msg)", font=("Segoe UI", 8)).pack(anchor="w")

        # =====================================================================
        # 3. Drop Zone & Source Selection
        # =====================================================================
        drop_container = ttk.Frame(self.root, padding="12 4 12 4")
        drop_container.grid(row=2, column=0, sticky="ew")
        drop_container.columnconfigure(0, weight=1)

        # Interactive drop zone box
        self.drop_box = tk.Frame(
            drop_container,
            bg="#f8fafc",
            highlightthickness=2,
            highlightbackground="#94a3b8",
            cursor="hand2",
            padx=20,
            pady=16,
        )
        self.drop_box.grid(row=0, column=0, sticky="ew")
        self.drop_box.columnconfigure(0, weight=1)

        drop_icon_text = "📥 Drag & Drop Files or Folders Here"
        self.drop_title = tk.Label(
            self.drop_box,
            text=drop_icon_text,
            font=("Segoe UI", 12, "bold"),
            fg="#1e293b",
            bg="#f8fafc",
        )
        self.drop_title.pack()

        self.drop_subtitle = tk.Label(
            self.drop_box,
            text="Drop any PDF, Office doc, audio, image, table, archive, or folder to convert to Markdown",
            font=("Segoe UI", 9),
            fg="#64748b",
            bg="#f8fafc",
        )
        self.drop_subtitle.pack(pady=(2, 8))

        # Buttons inside drop zone
        btn_bar = tk.Frame(self.drop_box, bg="#f8fafc")
        btn_bar.pack()

        self.browse_files_btn = ttk.Button(
            btn_bar,
            text="📂 Browse File(s)...",
            command=self.browse_files,
        )
        self.browse_files_btn.pack(side="left", padx=4)

        self.browse_folder_btn = ttk.Button(
            btn_bar,
            text="📁 Browse Folder...",
            command=self.browse_folder,
        )
        self.browse_folder_btn.pack(side="left", padx=4)

        # Bind click on drop box to browse files
        for widget in (self.drop_box, self.drop_title, self.drop_subtitle):
            widget.bind("<Button-1>", lambda e: self.browse_files())

        # Options Row
        options_frame = ttk.Frame(drop_container, padding="0 6 0 0")
        options_frame.grid(row=1, column=0, sticky="ew")
        options_frame.columnconfigure(2, weight=1)

        ttk.Checkbutton(
            options_frame,
            text="Auto-convert immediately when files are dropped",
            variable=self.auto_convert_var,
        ).grid(row=0, column=0, sticky="w", padx=(0, 16))

        ttk.Checkbutton(
            options_frame,
            text="Save .md next to original file",
            variable=self.output_same_dir_var,
            command=self._on_output_toggle,
        ).grid(row=0, column=1, sticky="w")

        # Custom directory row
        self.custom_dir_frame = ttk.Frame(drop_container, padding="0 2 0 0")
        self.custom_dir_frame.grid(row=2, column=0, sticky="ew")
        self.custom_dir_frame.columnconfigure(1, weight=1)

        ttk.Label(self.custom_dir_frame, text="Custom Output Dir:").grid(row=0, column=0, sticky="w", padx=(0, 4))
        self.custom_dir_entry = ttk.Entry(
            self.custom_dir_frame,
            textvariable=self.custom_output_dir_var,
            state="disabled",
        )
        self.custom_dir_entry.grid(row=0, column=1, sticky="ew", padx=4)

        self.custom_dir_btn = ttk.Button(
            self.custom_dir_frame,
            text="Choose...",
            command=self.browse_output_dir,
            state="disabled",
        )
        self.custom_dir_btn.grid(row=0, column=2, padx=(4, 0))

        # =====================================================================
        # 4. Queue / File List Frame
        # =====================================================================
        queue_frame = ttk.LabelFrame(
            self.root,
            text="Conversion Queue",
            padding="6 6 6 6",
        )
        queue_frame.grid(row=3, column=0, sticky="nsew", padx=12, pady=4)
        queue_frame.columnconfigure(0, weight=1)
        queue_frame.rowconfigure(1, weight=1)

        # Queue toolbar
        q_toolbar = ttk.Frame(queue_frame)
        q_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 4))

        self.convert_all_btn = ttk.Button(
            q_toolbar,
            text="⚡ Convert All",
            command=self.start_convert_all_thread,
        )
        self.convert_all_btn.pack(side="left", padx=2)

        self.convert_selected_btn = ttk.Button(
            q_toolbar,
            text="▶ Convert Selected",
            command=self.start_convert_selected_thread,
        )
        self.convert_selected_btn.pack(side="left", padx=2)

        self.remove_btn = ttk.Button(
            q_toolbar,
            text="✖ Remove Selected",
            command=self.remove_selected,
        )
        self.remove_btn.pack(side="left", padx=2)

        self.clear_btn = ttk.Button(
            q_toolbar,
            text="🗑 Clear Queue",
            command=self.clear_queue,
        )
        self.clear_btn.pack(side="left", padx=2)

        self.queue_count_label = ttk.Label(
            q_toolbar,
            text="0 files in queue",
            foreground="#64748b",
            font=("Segoe UI", 9, "italic"),
        )
        self.queue_count_label.pack(side="right", padx=6)

        # Treeview for Queue
        tree_container = ttk.Frame(queue_frame)
        tree_container.grid(row=1, column=0, sticky="nsew")
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        columns = ("status", "filename", "type", "size", "output")
        self.tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show="headings",
            selectmode="extended",
            style="Queue.Treeview",
        )

        self.tree.heading("status", text="Status", anchor="center")
        self.tree.heading("filename", text="File Name", anchor="w")
        self.tree.heading("type", text="Type", anchor="w")
        self.tree.heading("size", text="Size", anchor="e")
        self.tree.heading("output", text="Output File (.md)", anchor="w")

        self.tree.column("status", width=95, minwidth=80, anchor="center")
        self.tree.column("filename", width=260, minwidth=180, anchor="w")
        self.tree.column("type", width=140, minwidth=100, anchor="w")
        self.tree.column("size", width=85, minwidth=60, anchor="e")
        self.tree.column("output", width=240, minwidth=140, anchor="w")

        tree_scroll_y = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        # =====================================================================
        # 5. Progress Bar & Status Line
        # =====================================================================
        status_frame = ttk.Frame(self.root, padding="12 2 12 2")
        status_frame.grid(row=4, column=0, sticky="ew")
        status_frame.columnconfigure(0, weight=1)

        self.progress_bar = ttk.Progressbar(
            status_frame,
            orient="horizontal",
            mode="determinate",
            style="Custom.Horizontal.TProgressbar",
        )
        self.progress_bar.grid(row=0, column=0, sticky="ew", pady=(2, 4))

        self.status_var = tk.StringVar(value="Ready. Drop files above or browse to start converting.")
        self.status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            foreground="#334155",
        )
        self.status_label.grid(row=1, column=0, sticky="w", pady=(0, 2))

        # =====================================================================
        # 6. Markdown Preview Pane
        # =====================================================================
        preview_frame = ttk.LabelFrame(
            self.root,
            text="📝 Markdown Preview",
            padding="6 6 6 6",
        )
        preview_frame.grid(row=5, column=0, sticky="nsew", padx=12, pady=(4, 10))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(1, weight=1)

        # Preview toolbar
        p_toolbar = ttk.Frame(preview_frame)
        p_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 4))

        self.preview_title_var = tk.StringVar(value="No file selected")
        ttk.Label(
            p_toolbar,
            textvariable=self.preview_title_var,
            font=("Segoe UI", 9, "bold"),
            foreground="#1e293b",
        ).pack(side="left", padx=4)

        self.copy_btn = ttk.Button(
            p_toolbar,
            text="📋 Copy Markdown",
            command=self.copy_preview_to_clipboard,
        )
        self.copy_btn.pack(side="right", padx=2)

        self.save_as_btn = ttk.Button(
            p_toolbar,
            text="💾 Save As...",
            command=self.save_preview_as,
        )
        self.save_as_btn.pack(side="right", padx=2)

        self.open_file_btn = ttk.Button(
            p_toolbar,
            text="📂 Open File",
            command=self.open_active_output_file,
        )
        self.open_file_btn.pack(side="right", padx=2)

        self.open_folder_btn = ttk.Button(
            p_toolbar,
            text="📁 Open Folder",
            command=self.open_active_output_folder,
        )
        self.open_folder_btn.pack(side="right", padx=2)

        ttk.Checkbutton(
            p_toolbar,
            text="Wrap",
            variable=self.word_wrap_var,
            command=self._on_wrap_toggle,
        ).pack(side="right", padx=6)

        # Text area for preview
        text_container = ttk.Frame(preview_frame)
        text_container.grid(row=1, column=0, sticky="nsew")
        text_container.columnconfigure(0, weight=1)
        text_container.rowconfigure(0, weight=1)

        self.preview_text = tk.Text(
            text_container,
            wrap="word",
            font=("Consolas", 10),
            bg="#f8fafc",
            fg="#0f172a",
            insertbackground="#0f172a",
            selectbackground="#cbd5e1",
            relief="solid",
            borderwidth=1,
            highlightthickness=0,
            padx=8,
            pady=8,
        )
        self.preview_text.grid(row=0, column=0, sticky="nsew")

        text_scroll_y = ttk.Scrollbar(text_container, orient="vertical", command=self.preview_text.yview)
        text_scroll_y.grid(row=0, column=1, sticky="ns")
        self.preview_text.configure(yscrollcommand=text_scroll_y.set)

    def _setup_dnd(self):
        """Configure Drag & Drop on the drop box, treeview, and root."""
        if not HAS_DND:
            self.drop_title.configure(text="📂 Select Files or Folders to Convert")
            self.drop_subtitle.configure(text="Click 'Browse File(s)' or 'Browse Folder' below to add files")
            return

        try:
            # Register drop target on drop zone
            self.drop_box.drop_target_register(DND_FILES)
            self.drop_box.dnd_bind("<<Drop>>", self._on_dnd_drop)
            self.drop_box.dnd_bind("<<DropEnter>>", self._on_dnd_enter)
            self.drop_box.dnd_bind("<<DropLeave>>", self._on_dnd_leave)

            # Register on treeview too
            self.tree.drop_target_register(DND_FILES)
            self.tree.dnd_bind("<<Drop>>", self._on_dnd_drop)

            # Register on main window
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_dnd_drop)
        except Exception as e:
            print(f"Error binding DND events: {e}")

    def _on_dnd_enter(self, event):
        self.drop_box.configure(bg="#e0f2fe", highlightbackground="#0284c7")
        self.drop_title.configure(bg="#e0f2fe", fg="#0369a1")
        self.drop_subtitle.configure(bg="#e0f2fe")

    def _on_dnd_leave(self, event):
        self.drop_box.configure(bg="#f8fafc", highlightbackground="#94a3b8")
        self.drop_title.configure(bg="#f8fafc", fg="#1e293b")
        self.drop_subtitle.configure(bg="#f8fafc")

    def _on_dnd_drop(self, event):
        self._on_dnd_leave(event)
        raw_data = event.data
        if not raw_data:
            return

        # Parse paths safely
        try:
            paths = self.root.splitlist(raw_data)
        except Exception:
            paths = [raw_data]

        self.add_paths(paths)

    def _on_output_toggle(self):
        is_same = self.output_same_dir_var.get()
        state = "disabled" if is_same else "normal"
        self.custom_dir_entry.configure(state=state)
        self.custom_dir_btn.configure(state=state)

    def _on_wrap_toggle(self):
        wrap_mode = "word" if self.word_wrap_var.get() else "none"
        self.preview_text.configure(wrap=wrap_mode)

    def browse_output_dir(self):
        chosen = filedialog.askdirectory(title="Select Output Directory for Markdown Files")
        if chosen:
            self.custom_output_dir_var.set(chosen)

    def browse_files(self):
        # Build file dialog extensions filter
        ext_list = ";*".join(SUPPORTED_EXTENSIONS.keys())
        filetypes = [
            ("Supported Documents & Media", f"*{ext_list}"),
            ("PDF Documents (*.pdf)", "*.pdf"),
            ("Word Documents (*.docx;*.doc)", "*.docx;*.doc"),
            ("PowerPoint (*.pptx;*.ppt)", "*.pptx;*.ppt"),
            ("Excel Spreadsheets (*.xlsx;*.xls)", "*.xlsx;*.xls"),
            ("Data Files (*.csv;*.tsv;*.json;*.xml;*.rss)", "*.csv;*.tsv;*.json;*.xml;*.rss"),
            ("Web & Text (*.html;*.htm;*.txt;*.md;*.ipynb)", "*.html;*.htm;*.txt;*.md;*.ipynb"),
            ("Audio Files (*.mp3;*.wav;*.m4a;*.flac;*.ogg)", "*.mp3;*.wav;*.m4a;*.flac;*.ogg"),
            ("Image Files (*.jpg;*.jpeg;*.png;*.webp)", "*.jpg;*.jpeg;*.png;*.webp"),
            ("Archives & Books (*.zip;*.epub;*.msg)", "*.zip;*.epub;*.msg"),
            ("All Files (*.*)", "*.*"),
        ]
        files = filedialog.askopenfilenames(
            title="Select Files to Convert to Markdown",
            filetypes=filetypes,
        )
        if files:
            self.add_paths(files)

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Folder to Scan for Supported Documents")
        if folder:
            self.add_paths([folder])

    def add_paths(self, paths):
        """Process incoming files or directories, adding them to the conversion queue."""
        added_count = 0
        all_files = []

        for p in paths:
            clean_p = p.strip('"{}\'')
            if not os.path.exists(clean_p):
                continue

            if os.path.isdir(clean_p):
                # Recursively discover supported files in directory
                for root_dir, _, filenames in os.walk(clean_p):
                    for fn in filenames:
                        _, ext = os.path.splitext(fn)
                        if ext.lower() in SUPPORTED_EXTENSIONS:
                            all_files.append(os.path.join(root_dir, fn))
            elif os.path.isfile(clean_p):
                all_files.append(clean_p)

        # Filter and queue files
        for fpath in all_files:
            # Check if already in queue
            if any(item["path"] == fpath for item in self.file_queue):
                continue

            ext = os.path.splitext(fpath)[1].lower()
            type_desc, _ = SUPPORTED_EXTENSIONS.get(ext, ("Unknown Document", "Other"))

            try:
                fsize = os.path.getsize(fpath)
                fsize_str = format_size(fsize)
            except Exception:
                fsize = 0
                fsize_str = "Unknown"

            item_id = f"item_{self.item_counter}"
            self.item_counter += 1

            target_out = self._calculate_output_path(fpath)

            queue_item = {
                "id": item_id,
                "path": fpath,
                "filename": os.path.basename(fpath),
                "ext": ext,
                "type": type_desc,
                "size": fsize,
                "size_str": fsize_str,
                "status": "⏳ Queued",
                "output_path": target_out,
                "result_markdown": None,
                "error_msg": None,
            }

            self.file_queue.append(queue_item)

            self.tree.insert(
                "",
                "end",
                iid=item_id,
                values=(
                    queue_item["status"],
                    queue_item["filename"],
                    queue_item["type"],
                    queue_item["size_str"],
                    os.path.basename(target_out),
                ),
            )
            added_count += 1

        self._update_queue_label()

        if added_count > 0:
            self.status_var.set(f"Added {added_count} file(s) to queue. Total: {len(self.file_queue)}")
            # Auto-convert if option checked
            if self.auto_convert_var.get() and not self.is_converting:
                self.start_convert_all_thread()

    def _calculate_output_path(self, src_path):
        """Determine destination .md path for a given source path."""
        base_name = os.path.splitext(os.path.basename(src_path))[0]
        out_filename = f"{base_name}.md"

        if self.output_same_dir_var.get():
            parent_dir = os.path.dirname(src_path)
            # Avoid overwriting original if source was already .md
            if src_path.lower().endswith(".md"):
                out_filename = f"{base_name}_converted.md"
            return os.path.join(parent_dir, out_filename)
        else:
            custom_dir = self.custom_output_dir_var.get().strip()
            if not custom_dir or not os.path.isdir(custom_dir):
                custom_dir = os.path.dirname(src_path)
            return os.path.join(custom_dir, out_filename)

    def _update_queue_label(self):
        total = len(self.file_queue)
        done = sum(1 for item in self.file_queue if item["status"].startswith("✅"))
        self.queue_count_label.configure(text=f"{done}/{total} completed • {total} in queue")

    def remove_selected(self):
        selected_ids = self.tree.selection()
        if not selected_ids:
            return
        for iid in selected_ids:
            self.tree.delete(iid)
            self.file_queue = [item for item in self.file_queue if item["id"] != iid]
        self._update_queue_label()
        self.status_var.set(f"Removed {len(selected_ids)} item(s) from queue.")

    def clear_queue(self):
        if self.is_converting:
            messagebox.showwarning("Busy", "Conversion is currently running. Please wait for it to finish.")
            return
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.file_queue.clear()
        self._update_queue_label()
        self.preview_text.delete("1.0", "end")
        self.preview_title_var.set("No file selected")
        self.progress_bar["value"] = 0
        self.status_var.set("Queue cleared. Ready.")

    def _on_tree_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        iid = selected[-1]
        item = next((x for x in self.file_queue if x["id"] == iid), None)
        if item:
            self._display_preview(item)

    def _on_tree_double_click(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        iid = selected[0]
        item = next((x for x in self.file_queue if x["id"] == iid), None)
        if item and item["result_markdown"] is not None:
            self.open_active_output_file()
        elif item:
            # Convert this single item
            threading.Thread(target=self._convert_items_worker, args=([item],), daemon=True).start()

    def _display_preview(self, item):
        self.active_preview_path = item["output_path"]
        self.preview_title_var.set(f"File: {item['filename']} -> {os.path.basename(item['output_path'])}")

        self.preview_text.delete("1.0", "end")
        if item["result_markdown"] is not None:
            self.preview_text.insert("1.0", item["result_markdown"])
        elif item["error_msg"] is not None:
            self.preview_text.insert("1.0", f"--- ERROR CONVERTING FILE ---\n\n{item['error_msg']}")
        elif os.path.exists(item["output_path"]):
            try:
                with open(item["output_path"], "r", encoding="utf-8", errors="replace") as f:
                    content = f.read()
                self.preview_text.insert("1.0", content)
            except Exception as e:
                self.preview_text.insert("1.0", f"Error reading output file: {e}")
        else:
            self.preview_text.insert("1.0", f"[File is queued for conversion]\n\nSource: {item['path']}\nDestination: {item['output_path']}\nType: {item['type']} ({item['size_str']})\n\nClick 'Convert' to convert.")

    def start_convert_all_thread(self):
        if self.is_converting:
            return
        items_to_convert = [item for item in self.file_queue if not item["status"].startswith("✅")]
        if not items_to_convert:
            if not self.file_queue:
                self.status_var.set("Queue is empty. Drop or browse files to convert.")
            else:
                self.status_var.set("All files in queue have already been converted.")
            return

        threading.Thread(target=self._convert_items_worker, args=(items_to_convert,), daemon=True).start()

    def start_convert_selected_thread(self):
        if self.is_converting:
            return
        selected_ids = self.tree.selection()
        if not selected_ids:
            self.status_var.set("No files selected in queue. Please select files first.")
            return

        items = [item for item in self.file_queue if item["id"] in selected_ids]
        threading.Thread(target=self._convert_items_worker, args=(items,), daemon=True).start()

    def _convert_items_worker(self, items):
        self.is_converting = True
        self.cancel_requested = False
        total = len(items)

        def set_ui_busy(busy):
            state = "disabled" if busy else "normal"
            self.convert_all_btn.configure(state=state)
            self.convert_selected_btn.configure(state=state)
            self.browse_files_btn.configure(state=state)
            self.browse_folder_btn.configure(state=state)

        self.root.after(0, set_ui_busy, True)

        # Ensure MarkItDown engine is ready
        if self.md_instance is None:
            try:
                self.md_instance = MarkItDown(enable_builtins=True)
            except Exception as e:
                err_str = f"Failed to initialize MarkItDown backend: {e}"
                self.root.after(0, lambda: messagebox.showerror("Backend Error", err_str))
                self.is_converting = False
                self.root.after(0, set_ui_busy, False)
                return

        converted_count = 0
        error_count = 0

        for idx, item in enumerate(items):
            if self.cancel_requested:
                break

            current_num = idx + 1
            progress_pct = int(((idx) / total) * 100)

            # Update status to converting
            def update_item_start(it=item, cnum=current_num, pct=progress_pct):
                it["status"] = "⚙️ Converting..."
                self.tree.item(it["id"], values=(
                    it["status"],
                    it["filename"],
                    it["type"],
                    it["size_str"],
                    os.path.basename(it["output_path"]),
                ))
                self.progress_bar["value"] = pct
                self.status_var.set(f"Converting ({cnum}/{total}): {it['filename']}...")

            self.root.after(0, update_item_start)

            # Perform actual conversion using MarkItDown
            src_path = item["path"]
            out_path = self._calculate_output_path(src_path)
            item["output_path"] = out_path

            success = False
            markdown_content = ""
            err_details = None

            try:
                res = self.md_instance.convert(src_path)
                markdown_content = res.text_content or ""

                # Write markdown to file
                out_dir = os.path.dirname(out_path)
                if out_dir and not os.path.exists(out_dir):
                    os.makedirs(out_dir, exist_ok=True)

                with open(out_path, "w", encoding="utf-8") as out_f:
                    out_f.write(markdown_content)

                success = True
                converted_count += 1
            except Exception as e:
                err_details = f"{type(e).__name__}: {str(e)}\n\n{traceback.format_exc()}"
                error_count += 1

            # Update item result
            def update_item_end(it=item, sc=success, md=markdown_content, err=err_details):
                if sc:
                    it["status"] = "✅ Done"
                    it["result_markdown"] = md
                    it["error_msg"] = None
                else:
                    it["status"] = "❌ Error"
                    it["result_markdown"] = None
                    it["error_msg"] = err

                self.tree.item(it["id"], values=(
                    it["status"],
                    it["filename"],
                    it["type"],
                    it["size_str"],
                    os.path.basename(it["output_path"]),
                ))
                self._update_queue_label()

                # If active preview is this item, update preview
                selected_ids = self.tree.selection()
                if not selected_ids or it["id"] in selected_ids:
                    self._display_preview(it)

            self.root.after(0, update_item_end)

        # Finished all items in batch
        def finish_batch(c_count=converted_count, e_count=error_count, tot=total):
            self.progress_bar["value"] = 100
            if e_count == 0:
                self.status_var.set(f"✅ Successfully converted all {c_count} file(s)!")
            else:
                self.status_var.set(f"Completed with issues: {c_count} converted, {e_count} failed. Click item to view error.")
            set_ui_busy(False)
            self.is_converting = False

        self.root.after(0, finish_batch)

    def copy_preview_to_clipboard(self):
        content = self.preview_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showinfo("Clipboard", "Preview area is empty.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.status_var.set("📋 Markdown copied to clipboard!")

    def save_preview_as(self):
        content = self.preview_text.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showinfo("Save As", "No markdown content to save.")
            return
        init_name = "converted.md"
        if self.active_preview_path:
            init_name = os.path.basename(self.active_preview_path)
        dest = filedialog.asksaveasfilename(
            title="Save Markdown As",
            initialfile=init_name,
            defaultextension=".md",
            filetypes=[("Markdown Document", "*.md"), ("Text Document", "*.txt"), ("All Files", "*.*")],
        )
        if dest:
            try:
                with open(dest, "w", encoding="utf-8") as f:
                    f.write(content)
                self.status_var.set(f"Saved to {os.path.basename(dest)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")

    def open_active_output_file(self):
        if self.active_preview_path and os.path.exists(self.active_preview_path):
            try:
                os.startfile(self.active_preview_path)
            except Exception as e:
                messagebox.showerror("Error", f"Could not open file: {e}")
        else:
            messagebox.showinfo("Open File", "Output file has not been created yet.")

    def open_active_output_folder(self):
        folder = None
        if self.active_preview_path:
            folder = os.path.dirname(self.active_preview_path)
        elif self.custom_output_dir_var.get():
            folder = self.custom_output_dir_var.get()

        if folder and os.path.isdir(folder):
            try:
                os.startfile(folder)
            except Exception as e:
                messagebox.showerror("Error", f"Could not open folder: {e}")
        else:
            messagebox.showinfo("Open Folder", "No valid folder found for selected item.")


def main():
    # Fix high DPI blur on Windows
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = TkinterDnD.Tk()
    app = MarkItDownApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
