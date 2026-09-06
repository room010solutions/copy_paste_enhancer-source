import customtkinter as ctk
import pyperclip
import threading
import time
import os
import json
import random
import platform
from datetime import datetime
from PIL import Image, ImageTk
from pathlib import Path


class ClipboardManager:
    ROULETTE_MESSAGES = [
        "🎲 The clipboard gods have chosen item #{n}!",
        "🎲 Rolling the dice... item #{n} it is!",
        "🎲 Fate has selected item #{n}!",
        "🎲 Behold, a blast from the past — item #{n}!",
        "🎲 Random.org would be proud: #{n}!",
    ]

    PLAYABLE_EXTENSIONS = {
        '.mp3', '.wav', '.ogg', '.flac', '.m4a', '.aac',      # audio
        '.mp4', '.avi', '.mkv', '.mov', '.wmv', '.webm',      # video
    }

    def __init__(self):
        # Configure appearance
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Main window
        self.window = ctk.CTk()
        self.window.title("Clipboard Manager Pro")
        self.window.geometry("900x700")
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

        # Maximum items to store
        self.max_items = 24
        self.clipboard_history = []       # List to maintain chronological order
        self.clipboard_items_type = []    # Store type: 'text' or 'file'
        self.clipboard_pinned = []        # Bool per item — pinned items skip eviction
        self.clipboard_timestamps = []    # ISO timestamp per item
        self.last_clipboard_content = ""

        # All-time fun stats (survive clears/restarts)
        self.stats = {"total_items": 0, "total_characters": 0}

        # Search
        self.search_query = ""

        # Persistence
        self.history_file = Path.home() / ".clipboard_manager_history.json"

        # Track selected items
        self.selected_items = {}  # Dictionary to store checkbox variables
        self.checkboxes = {}      # Store checkbox widgets

        # Selection mode (True = multiple, False = single)
        self.multiple_selection_mode = False

        # Toast stacking
        # Toast stacking
        self.active_toasts = []

        # Warn about missing pywin32 only once per session
        self._pywin32_warning_shown = False

        # Track scheduled marquee `after()` callbacks so we can cancel them
        # when the preview panel refreshes — otherwise they keep firing
        # against widgets that no longer exist
        self._marquee_jobs = []

        # Timestamp until which the monitor thread ignores clipboard changes —
        # prevents our own programmatic copies from being re-detected as new items
        self._suppress_monitor_until = 0.0

        # Load persisted history before building UI so it renders immediately
        self.load_history()

        # Create UI
        self.create_widgets()
        self.update_display()

        # Start clipboard monitoring
        self.monitoring = True
        self.monitor_thread = threading.Thread(target=self.monitor_clipboard, daemon=True)
        self.monitor_thread.start()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def create_widgets(self):
        # Main container
        self.main_container = ctk.CTkFrame(self.window)
        self.main_container.pack(fill="both", expand=True, padx=10, pady=10)

        # Title
        self.title_label = ctk.CTkLabel(
            self.main_container,
            text="Clipboard Manager Pro",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.title_label.pack(pady=5)

        # Top controls frame
        self.top_controls = ctk.CTkFrame(self.main_container)
        self.top_controls.pack(fill="x", pady=5, padx=10)

        # Topmost toggle (left side)
        self.topmost_var = ctk.BooleanVar(value=False)
        self.topmost_toggle = ctk.CTkSwitch(
            self.top_controls,
            text="Keep on Top",
            variable=self.topmost_var,
            command=self.toggle_topmost,
            font=ctk.CTkFont(size=11)
        )
        self.topmost_toggle.pack(side="left", padx=5)

        # Selection mode toggle (middle)
        self.mode_var = ctk.BooleanVar(value=False)
        self.mode_toggle = ctk.CTkSwitch(
            self.top_controls,
            text="Single-Select",
            variable=self.mode_var,
            command=self.toggle_selection_mode,
            font=ctk.CTkFont(size=11)
        )
        self.mode_toggle.pack(side="left", padx=20)

        # Copy files button (right side)
        self.copy_files_button = ctk.CTkButton(
            self.top_controls,
            text="📁 Add Files",
            command=self.add_files_manually,
            width=100,
            height=30,
            font=ctk.CTkFont(size=12)
        )
        self.copy_files_button.pack(side="right", padx=5)

        # Instructions
        self.instructions = ctk.CTkLabel(
            self.main_container,
            text=(
                f"Copy text or files (Ctrl+C)\n"
                f"Supports text, images, documents, music, and more\n"
                f"Multi-Select: Check multiple items to combine | Single-Select: one at a time\n"
                f"📌 Pin items to protect them from eviction | Ctrl+F to search (max {self.max_items} items)"
            ),
            font=ctk.CTkFont(size=11),
            text_color="gray",
            justify="center"
        )
        self.instructions.pack(pady=5)

        # Selection info
        self.selection_label = ctk.CTkLabel(
            self.main_container,
            text="No items selected",
            font=ctk.CTkFont(size=11),
            text_color="yellow"
        )
        self.selection_label.pack(pady=2)

        # Main content frame (left: list, right: preview)
        self.content_frame = ctk.CTkFrame(self.main_container)
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Left frame for clipboard items
        self.left_frame = ctk.CTkFrame(self.content_frame)
        self.left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        # Search bar
        self.search_frame = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.search_frame.pack(fill="x", padx=5, pady=(5, 0))

        self.search_entry = ctk.CTkEntry(
            self.search_frame,
            placeholder_text="🔍 Search clipboard history..."
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.search_entry.bind("<KeyRelease>", self.on_search_changed)

        self.clear_search_button = ctk.CTkButton(
            self.search_frame,
            text="✕",
            width=30,
            command=self.clear_search
        )
        self.clear_search_button.pack(side="right")

        # Item count label
        self.count_label = ctk.CTkLabel(
            self.left_frame,
            text="",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        self.count_label.pack(anchor="e", padx=10)

        # Scrollable frame for clipboard items
        self.scrollable_frame = ctk.CTkScrollableFrame(
            self.left_frame,
            width=450,
            height=380
        )
        self.scrollable_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Right frame for preview
        self.right_frame = ctk.CTkFrame(self.content_frame, width=300)
        self.right_frame.pack(side="right", fill="both", padx=(5, 0))
        self.right_frame.pack_propagate(False)

        # Preview title
        self.preview_title = ctk.CTkLabel(
            self.right_frame,
            text="Preview",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.preview_title.pack(pady=10)

        # Preview content area
        self.preview_frame = ctk.CTkScrollableFrame(
            self.right_frame,
            width=280,
            height=350
        )
        self.preview_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Selection buttons row
        self.button_frame = ctk.CTkFrame(self.main_container)
        self.button_frame.pack(pady=(10, 5), padx=20, fill="x")

        self.select_all_button = ctk.CTkButton(
            self.button_frame, text="Select All", command=self.select_all_items,
            width=100, font=ctk.CTkFont(size=12)
        )
        self.select_all_button.pack(side="left", padx=5)

        self.deselect_all_button = ctk.CTkButton(
            self.button_frame, text="Deselect All", command=self.deselect_all_items,
            width=100, font=ctk.CTkFont(size=12)
        )
        self.deselect_all_button.pack(side="left", padx=5)

        self.roulette_button = ctk.CTkButton(
            self.button_frame, text="🎲 Surprise Me", command=self.clipboard_roulette,
            width=110, fg_color="#8e44ad", hover_color="#6c3483", font=ctk.CTkFont(size=12)
        )
        # self.roulette_button.pack(side="left", padx=5)

        self.copy_selected_button = ctk.CTkButton(
            self.button_frame, text="📋 Copy Selected", command=self.copy_selected_items,
            width=130, fg_color="#2e7d32", hover_color="#1b5e20", font=ctk.CTkFont(size=12)
        )
        self.copy_selected_button.pack(side="left", padx=5)

        self.clear_button = ctk.CTkButton(
            self.button_frame, text="Clear All", command=self.clear_history,
            width=80, fg_color="red", hover_color="darkred", font=ctk.CTkFont(size=12)
        )
        self.clear_button.pack(side="right", padx=5)

        self.max_items_button = ctk.CTkButton(
            self.button_frame, text="Max Items", command=self.open_max_items_dialog,
            width=80, font=ctk.CTkFont(size=12)
        )
        self.max_items_button.pack(side="right", padx=5)

        self.stats_button = ctk.CTkButton(
            self.button_frame, text="📊 Stats", command=self.open_stats_dialog,
            width=80, font=ctk.CTkFont(size=12)
        )
        self.stats_button.pack(side="right", padx=5)

        # Keyboard shortcuts
        self.window.bind("<Control-f>", lambda e: self.search_entry.focus_set())
        self.window.bind("<Control-F>", lambda e: self.search_entry.focus_set())
        self.window.bind("<Escape>", lambda e: self.clear_search())

    # ------------------------------------------------------------------
    # Persistence
    # ------------------------------------------------------------------

    def load_history(self):
        if not self.history_file.exists():
            return
        try:
            with open(self.history_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.clipboard_history = data.get("history", [])
            self.clipboard_items_type = data.get("types", [])
            n = len(self.clipboard_history)
            self.clipboard_pinned = data.get("pinned", [False] * n)
            self.clipboard_timestamps = data.get("timestamps", [""] * n)
            self.max_items = data.get("max_items", self.max_items)
            self.stats = data.get("stats", self.stats)
            if self.clipboard_history:
                last = self.clipboard_history[-1]
                self.last_clipboard_content = last if isinstance(last, str) else ""
        except Exception:
            pass  # Corrupt or unreadable history file — start fresh rather than crash

    def save_history(self):
        try:
            data = {
                "history": self.clipboard_history,
                "types": self.clipboard_items_type,
                "pinned": self.clipboard_pinned,
                "timestamps": self.clipboard_timestamps,
                "max_items": self.max_items,
                "stats": self.stats,
            }
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(data, f)
        except Exception:
            pass  # Saving is best-effort — never let it interrupt the user's workflow

    def on_close(self):
        self.save_history()
        self.stop_monitoring()
        self.window.destroy()

    # ------------------------------------------------------------------
    # Toggles
    # ------------------------------------------------------------------

    def toggle_topmost(self):
        """Toggle window always-on-top mode"""
        self.window.attributes('-topmost', self.topmost_var.get())
        if self.topmost_var.get():
            self.show_toast("Window will stay on top", "green")
        else:
            self.show_toast("Topmost disabled", "gray")

    def toggle_selection_mode(self):
        """Toggle between single and multiple selection mode"""
        self.multiple_selection_mode = self.mode_var.get()

        if not self.multiple_selection_mode:
            self.mode_toggle.configure(text="Single-Select")
            self.enforce_single_selection()
            self.show_toast("Single-select mode enabled", "blue")
        else:
            self.mode_toggle.configure(text="Multi-Select")
            self.show_toast("Multi-select mode enabled", "green")

    def enforce_single_selection(self):
        """Ensure only one item is selected (for single-select mode)"""
        if not self.multiple_selection_mode:
            selected_indices = [i for i in self.selected_items if self.selected_items[i].get()]

            if len(selected_indices) > 1:
                last_selected = selected_indices[-1]
                for i in selected_indices:
                    if i != last_selected:
                        self.selected_items[i].set(False)
                        if i in self.checkboxes:
                            self.checkboxes[i].deselect()

                if last_selected in self.selected_items:
                    self.copy_item_to_clipboard(last_selected)

    # ------------------------------------------------------------------
    # Clipboard monitoring
    # ------------------------------------------------------------------

    def monitor_clipboard(self):
        """Monitor clipboard for changes in background thread"""
        while self.monitoring:
            if time.time() < self._suppress_monitor_until:
                # We just wrote to the clipboard ourselves (via a Copy action) —
                # skip this poll entirely so we don't re-detect our own write as new
                time.sleep(0.2)
                continue
            try:
                current_content = pyperclip.paste()
                file_paths = self.get_clipboard_files()

                if file_paths:
                    if file_paths != self.last_clipboard_content:
                        self.last_clipboard_content = file_paths
                        self.window.after(0, lambda fp=file_paths: self.add_new_item(fp, 'file'))
                elif current_content != self.last_clipboard_content and current_content.strip():
                    self.last_clipboard_content = current_content
                    if current_content not in self.clipboard_history:
                        self.window.after(0, lambda c=current_content: self.add_new_item(c, 'text'))
            except Exception:
                pass
            time.sleep(0.5)

    def get_clipboard_files(self):
        """Get file paths from clipboard (Windows, macOS, Linux best-effort)"""
        system = platform.system()
        if system == "Windows":
            return self._get_clipboard_files_windows()
        try:
            import subprocess
            if system == "Darwin":
                check = subprocess.run(
                    ['osascript', '-e', 'clipboard info'],
                    capture_output=True, text=True, timeout=2
                )
                if 'furl' in check.stdout:
                    result = subprocess.run(
                        ['osascript', '-e', 'POSIX path of (the clipboard as «class furl»)'],
                        capture_output=True, text=True, timeout=2
                    )
                    path = result.stdout.strip()
                    if path:
                        return [path]
            elif system == "Linux":
                targets = subprocess.run(
                    ['xclip', '-selection', 'clipboard', '-t', 'TARGETS', '-o'],
                    capture_output=True, text=True, timeout=2
                )
                if 'text/uri-list' in targets.stdout:
                    result = subprocess.run(
                        ['xclip', '-selection', 'clipboard', '-t', 'text/uri-list', '-o'],
                        capture_output=True, text=True, timeout=2
                    )
                    files = [
                        f.strip().replace('file://', '')
                        for f in result.stdout.split('\n') if f.strip()
                    ]
                    return files if files else None
        except (FileNotFoundError, Exception):
            pass
        return None

    def _get_clipboard_files_windows(self):
        """Read CF_HDROP directly via the Win32 API. Far faster and more reliable
        than shelling out to PowerShell on every 500ms poll, and avoids parsing
        PowerShell's formatted table output entirely."""
        try:
            import win32clipboard
        except ImportError:
            if not self._pywin32_warning_shown:
                self._pywin32_warning_shown = True
                self.window.after(0, lambda: self.show_toast(
                    "File-copy detection needs pywin32: pip install pywin32", "orange"
                ))
            return None

        for _ in range(3):
            try:
                win32clipboard.OpenClipboard()
                try:
                    if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_HDROP):
                        files = win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
                        files = [f for f in files if os.path.exists(f)]
                        return files if files else None
                    return None
                finally:
                    win32clipboard.CloseClipboard()
            except Exception:
                # Clipboard transiently locked by another process (common — retry briefly)
                time.sleep(0.05)
        return None

    def add_files_manually(self):
        """Open file dialog to add files"""
        from tkinter import filedialog
        files = filedialog.askopenfilenames(
            title="Select files to add",
            filetypes=[
                ("All files", "*.*"),
                ("Images", "*.png *.jpg *.jpeg *.gif *.bmp"),
                ("Documents", "*.txt *.pdf *.doc *.docx"),
                ("Music", "*.mp3 *.wav *.ogg"),
                ("Videos", "*.mp4 *.avi *.mkv")
            ]
        )
        if files:
            self.add_new_item(list(files), 'file')
            self.show_toast(f"Added {len(files)} file(s)", "green")

    # ------------------------------------------------------------------
    # Core history management
    # ------------------------------------------------------------------

    def add_new_item(self, content, item_type='text'):
        """Add new item to the end of the list (maintains chronological order)"""
        self.clipboard_history.append(content)
        self.clipboard_items_type.append(item_type)
        self.clipboard_pinned.append(False)
        self.clipboard_timestamps.append(datetime.now().isoformat())

        self.stats["total_items"] += 1
        if item_type == 'text':
            self.stats["total_characters"] += len(content)

        if len(self.clipboard_history) > self.max_items:
            oldest_unpinned = next(
                (i for i, pinned in enumerate(self.clipboard_pinned) if not pinned), None
            )
            if oldest_unpinned is not None:
                self.clipboard_history.pop(oldest_unpinned)
                self.clipboard_items_type.pop(oldest_unpinned)
                self.clipboard_pinned.pop(oldest_unpinned)
                self.clipboard_timestamps.pop(oldest_unpinned)
                self.show_toast(f"Removed oldest unpinned item to maintain {self.max_items} max", "orange")
            else:
                self.show_toast("All items pinned — unpin some or raise Max Items", "orange")

        self.save_history()
        self.update_display()

    def item_matches_query(self, index, query):
        item_type = self.clipboard_items_type[index]
        content = self.clipboard_history[index]
        if item_type == 'file':
            names = content if isinstance(content, list) else [content]
            return any(query in os.path.basename(n).lower() for n in names)
        return query in str(content).lower()

    def update_display(self):
        """Update the scrollable frame with clipboard items"""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        self.checkboxes.clear()

        old_selected = {i: var.get() for i, var in self.selected_items.items()}
        self.selected_items.clear()

        query = self.search_query.strip().lower()
        visible_count = 0
        for i, item in enumerate(self.clipboard_history):
            item_type = self.clipboard_items_type[i] if i < len(self.clipboard_items_type) else 'text'
            if query and not self.item_matches_query(i, query):
                continue
            visible_count += 1
            self.create_clipboard_item(i, item, old_selected.get(i, False), item_type)

        if query and visible_count == 0:
            no_results = ctk.CTkLabel(
                self.scrollable_frame,
                text=f'No items match "{self.search_query}"',
                font=ctk.CTkFont(size=12),
                text_color="gray"
            )
            no_results.pack(pady=20)

        self.count_label.configure(text=f"{len(self.clipboard_history)}/{self.max_items} items")
        self.update_selection_label()
        self.update_preview()

    def create_clipboard_item(self, index, content, was_selected=False, item_type='text'):
        """Create a UI element for each clipboard item"""
        is_pinned = index < len(self.clipboard_pinned) and self.clipboard_pinned[index]

        if is_pinned:
            item_frame = ctk.CTkFrame(self.scrollable_frame, border_width=2, border_color="#DAA520")
        else:
            item_frame = ctk.CTkFrame(self.scrollable_frame)
        item_frame.pack(fill="x", pady=5, padx=5)

        # Checkbox for selection
        checkbox_var = ctk.BooleanVar(value=was_selected)
        checkbox = ctk.CTkCheckBox(
            item_frame,
            text="",
            variable=checkbox_var,
            width=20,
            command=lambda i=index, v=checkbox_var: self.on_checkbox_toggle(i, v)
        )
        checkbox.pack(side="left", padx=5)

        self.selected_items[index] = checkbox_var
        self.checkboxes[index] = checkbox

        # Determine display text and icon
        if item_type == 'file':
            if isinstance(content, list):
                if len(content) == 1:
                    display_text = os.path.basename(content[0])
                    icon = "📄"
                else:
                    display_text = f"{len(content)} files selected"
                    icon = "📁"
            else:
                display_text = os.path.basename(content)
                icon = "📄"
        else:
            display_text = content[:50] + "..." if len(content) > 50 else content
            display_text = display_text.replace('\n', ' ')
            icon = "📝"

        icon_label = ctk.CTkLabel(item_frame, text=icon, width=30, font=ctk.CTkFont(size=16))
        icon_label.pack(side="left", padx=2)

        index_label = ctk.CTkLabel(
            item_frame, text=f"#{index + 1}", width=40, font=ctk.CTkFont(size=14, weight="bold")
        )
        index_label.pack(side="left", padx=5)

        time_text = ""
        if index < len(self.clipboard_timestamps):
            time_text = self.format_relative_time(self.clipboard_timestamps[index])
        time_label = ctk.CTkLabel(
            item_frame, text=time_text, width=50, font=ctk.CTkFont(size=10), text_color="gray"
        )
        time_label.pack(side="left", padx=2)

        content_label = ctk.CTkLabel(item_frame, text=display_text, anchor="w", cursor="hand2")
        content_label.pack(side="left", fill="x", expand=True, padx=5)
        content_label.bind("<Button-1>", lambda e, i=index: self.toggle_checkbox_from_label(i))
        icon_label.bind("<Button-1>", lambda e, i=index: self.toggle_checkbox_from_label(i))

        # Delete button
        delete_button = ctk.CTkButton(
            item_frame, text="×", width=30, height=30, font=ctk.CTkFont(size=14),
            fg_color="red", hover_color="darkred",
            command=lambda i=index: self.delete_item(i)
        )
        delete_button.pack(side="right", padx=2)

        # Copy button
        copy_button = ctk.CTkButton(
            item_frame, text="Copy", width=60, height=30, font=ctk.CTkFont(size=12),
            command=lambda i=index: self.copy_item_to_clipboard(i)
        )
        copy_button.pack(side="right", padx=5)

        # Pin button
        pin_icon = "📌" if is_pinned else "📍"
        pin_button = ctk.CTkButton(
            item_frame, text=pin_icon, width=30, height=30, font=ctk.CTkFont(size=13),
            fg_color="transparent", hover_color=("gray75", "gray25"),
            command=lambda i=index: self.toggle_pin(i)
        )
        pin_button.pack(side="right", padx=2)

    def toggle_checkbox_from_label(self, index):
        """Let clicking an item's text toggle its selection, same as the checkbox"""
        if index in self.selected_items:
            var = self.selected_items[index]
            var.set(not var.get())
            if var.get():
                self.checkboxes[index].select()
            else:
                self.checkboxes[index].deselect()
            self.on_checkbox_toggle(index, var)

    def toggle_pin(self, index):
        """Pin/unpin an item — pinned items are protected from eviction"""
        if 0 <= index < len(self.clipboard_pinned):
            self.clipboard_pinned[index] = not self.clipboard_pinned[index]
            state = "pinned 📌" if self.clipboard_pinned[index] else "unpinned"
            self.show_toast(f"Item #{index + 1} {state}", "blue")
            self.save_history()
            self.update_display()

    def format_relative_time(self, iso_timestamp):
        try:
            ts = datetime.fromisoformat(iso_timestamp)
        except (ValueError, TypeError):
            return ""
        seconds = (datetime.now() - ts).total_seconds()
        if seconds < 60:
            return "just now"
        minutes = seconds / 60
        if minutes < 60:
            return f"{int(minutes)}m ago"
        hours = minutes / 60
        if hours < 24:
            return f"{int(hours)}h ago"
        return f"{int(hours / 24)}d ago"

    def on_checkbox_toggle(self, index, var):
        """Handle checkbox toggle event.

        Checking/unchecking a box is purely an internal selection action — it
        builds up the working set shown in the preview panel. It does NOT push
        anything to the system clipboard. In single-select mode we still copy
        immediately, since there checking a box unambiguously means "use this
        one item now." In multi-select mode, pushing the combined selection to
        the clipboard is a separate, explicit action (the Copy Selected button).
        """
        if not self.multiple_selection_mode:
            if var.get():
                for i in self.selected_items:
                    if i != index:
                        self.selected_items[i].set(False)
                        if i in self.checkboxes:
                            self.checkboxes[i].deselect()

                self.copy_item_to_clipboard(index)
                self.show_toast(f"Item #{index + 1} selected and copied", "green")

        self.update_selection_label()
        self.update_preview()

    def copy_item_to_clipboard(self, index):
        """Copy a specific item to clipboard"""
        if 0 <= index < len(self.clipboard_history):
            content = self.clipboard_history[index]
            item_type = self.clipboard_items_type[index]

            if item_type == 'file':
                self.copy_files_to_clipboard(content)
            else:
                pyperclip.copy(content)
                self.last_clipboard_content = content
                self._suppress_monitor_until = time.time() + 1.0
                self.show_toast(f"Item #{index + 1} copied", "green")

    def copy_files_to_clipboard(self, files):
        """Copy files to system clipboard (Windows/macOS/Linux, with text fallback)"""
        system = platform.system()
        file_list = files if isinstance(files, list) else [files]
        try:
            import subprocess
            if system == "Windows":
                path_args = ','.join(f'"{f}"' for f in file_list)
                subprocess.run(
                    ['powershell', '-Command', f"Set-Clipboard -Path {path_args}"],
                    timeout=2, check=True
                )
            elif system == "Darwin":
                if len(file_list) == 1:
                    ascript = f'set the clipboard to (POSIX file "{file_list[0]}")'
                else:
                    posix_list = ", ".join(f'POSIX file "{f}"' for f in file_list)
                    ascript = f'set the clipboard to {{{posix_list}}}'
                subprocess.run(['osascript', '-e', ascript], timeout=2, check=True)
            elif system == "Linux":
                uri_list = "\n".join(Path(f).resolve().as_uri() for f in file_list)
                subprocess.run(
                    ['xclip', '-selection', 'clipboard', '-t', 'text/uri-list'],
                    input=uri_list, text=True, timeout=2, check=True
                )
            else:
                raise OSError(f"Unsupported platform: {system}")
            self.last_clipboard_content = file_list
            self._suppress_monitor_until = time.time() + 1.0
            self.show_toast("Files copied to clipboard", "green")
        except FileNotFoundError:
            pyperclip.copy("\n".join(file_list))
            self.last_clipboard_content = "\n".join(file_list)
            self._suppress_monitor_until = time.time() + 1.0
            self.show_toast("Native file copy unavailable — copied paths as text instead", "orange")
        except Exception as e:
            self.show_toast(f"Error copying files: {str(e)}", "red")

    def copy_selected_items(self):
        """Copy all selected items to clipboard (combined if multiple).

        This is now the single explicit trigger for pushing the checked
        selection to the system clipboard — checking boxes alone no longer
        does this in multi-select mode.
        """
        selected_indices = [i for i in self.selected_items if self.selected_items[i].get()]

        if len(selected_indices) == 0:
            self.show_toast("Check one or more items first", "gray")
            return

        if len(selected_indices) == 1:
            self.copy_item_to_clipboard(selected_indices[0])
        else:
            text_contents = []
            file_contents = []

            for i in selected_indices:
                if self.clipboard_items_type[i] == 'text':
                    text_contents.append(self.clipboard_history[i])
                else:
                    if isinstance(self.clipboard_history[i], list):
                        file_contents.extend(self.clipboard_history[i])
                    else:
                        file_contents.append(self.clipboard_history[i])

            if text_contents and not file_contents:
                combined = "\n".join(text_contents)
                pyperclip.copy(combined)
                self.last_clipboard_content = combined
                self._suppress_monitor_until = time.time() + 1.0
                self.show_toast(f"{len(text_contents)} text items combined", "blue")
            elif file_contents and not text_contents:
                self.copy_files_to_clipboard(file_contents)
                self.show_toast(f"{len(file_contents)} files copied", "blue")
            else:
                if text_contents:
                    combined = "\n".join(text_contents)
                    pyperclip.copy(combined)
                    self.last_clipboard_content = combined
                    self._suppress_monitor_until = time.time() + 1.0
                if file_contents:
                    self.show_toast(f"Text + {len(file_contents)} files (files copied separately)", "orange")

    def update_selection_label(self):
        """Update the selection count label"""
        selected_indices = [i for i in self.selected_items if self.selected_items[i].get()]
        selected_count = len(selected_indices)

        if selected_count == 0:
            mode_text = "Single-Select" if not self.multiple_selection_mode else "Multi-Select"
            self.selection_label.configure(
                text=f"No items selected ({mode_text} mode)",
                text_color="yellow"
            )
        elif selected_count == 1:
            self.selection_label.configure(
                text=f"Item #{selected_indices[0] + 1} selected - ready to paste",
                text_color="green"
            )
        else:
            selected_nums = [i + 1 for i in selected_indices]
            self.selection_label.configure(
                text=f"Items {selected_nums} selected",
                text_color="orange"
            )

    def update_preview(self):
        """Update the preview panel with selected item details"""
        for job in self._marquee_jobs:
            try:
                self.window.after_cancel(job)
            except Exception:
                pass
        self._marquee_jobs.clear()

        for widget in self.preview_frame.winfo_children():
            widget.destroy()

        selected_indices = [i for i in self.selected_items if self.selected_items[i].get()]

        if not selected_indices:
            label = ctk.CTkLabel(
                self.preview_frame, text="Select an item\nto preview",
                font=ctk.CTkFont(size=14), text_color="gray"
            )
            label.pack(pady=50)
            return

        for idx in selected_indices:
            content = self.clipboard_history[idx]
            item_type = self.clipboard_items_type[idx]
            pin_note = " 📌" if idx < len(self.clipboard_pinned) and self.clipboard_pinned[idx] else ""
            time_note = self.format_relative_time(self.clipboard_timestamps[idx]) if idx < len(self.clipboard_timestamps) else ""

            num_label = ctk.CTkLabel(
                self.preview_frame, text=f"Item #{idx + 1}{pin_note}",
                font=ctk.CTkFont(size=14, weight="bold"), text_color="blue"
            )
            num_label.pack(pady=(5, 0))

            if time_note:
                time_label = ctk.CTkLabel(
                    self.preview_frame, text=time_note, font=ctk.CTkFont(size=10), text_color="gray"
                )
                time_label.pack(pady=(0, 5))

            if item_type == 'file':
                if isinstance(content, list):
                    for file_path in content[:5]:
                        self.create_file_preview(file_path)
                    if len(content) > 5:
                        more_label = ctk.CTkLabel(
                            self.preview_frame, text=f"... and {len(content) - 5} more files",
                            font=ctk.CTkFont(size=11), text_color="gray"
                        )
                        more_label.pack(pady=2)
                else:
                    self.create_file_preview(content)
            else:
                preview_text = content[:500] + "..." if len(content) > 500 else content
                text_label = ctk.CTkLabel(
                    self.preview_frame, text=preview_text, font=ctk.CTkFont(size=12),
                    wraplength=250, justify="left"
                )
                text_label.pack(pady=5, padx=10)

            if len(selected_indices) > 1 and idx != selected_indices[-1]:
                separator = ctk.CTkFrame(self.preview_frame, height=2, fg_color="gray")
                separator.pack(fill="x", pady=5, padx=10)

    def create_file_preview(self, file_path):
        """Create preview for a file"""
        if not os.path.exists(file_path):
            return

        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)

        file_frame = ctk.CTkFrame(self.preview_frame)
        file_frame.pack(fill="x", pady=2, padx=5)

        ext = os.path.splitext(file_name)[1].lower()
        is_playable = ext in self.PLAYABLE_EXTENSIONS
        if ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp']:
            icon = "🖼️"
            try:
                img = Image.open(file_path)
                img.thumbnail((200, 200))
                photo = ImageTk.PhotoImage(img)
                img_label = ctk.CTkLabel(self.preview_frame, image=photo, text="")
                img_label.image = photo
                img_label.pack(pady=5)
            except Exception:
                pass
        elif ext in ['.mp3', '.wav', '.ogg', '.flac', '.m4a', '.aac']:
            icon = "🎵"
        elif ext in ['.mp4', '.avi', '.mkv', '.mov', '.wmv', '.webm']:
            icon = "🎬"
        elif ext in ['.pdf', '.doc', '.docx', '.txt']:
            icon = "📄"
        else:
            icon = "📁"

        icon_label = ctk.CTkLabel(file_frame, text=icon, font=ctk.CTkFont(size=13), width=20)
        icon_label.pack(side="left", padx=(5, 0))

        # Play button packs right after the icon — before the name — so it's
        # never squeezed out by a long filename in the fixed-width frame
        if is_playable:
            play_button = ctk.CTkButton(
                file_frame, text="▶", width=28, height=22,
                font=ctk.CTkFont(size=11), fg_color="#1565c0", hover_color="#0d47a1",
                command=lambda p=file_path: self.play_file(p)
            )
            play_button.pack(side="left", padx=(2, 5))

        size_label = ctk.CTkLabel(
            file_frame, text=f"({self.format_size(file_size)})",
            font=ctk.CTkFont(size=10), text_color="gray"
        )
        size_label.pack(side="right", padx=5)

        # Name gets whatever space is left. If it's too long to fit, it
        # scrolls (marquee-style) instead of pushing other widgets off-frame.
        name_label = ctk.CTkLabel(
            file_frame, text=file_name, font=ctk.CTkFont(size=11),
            anchor="w", width=130
        )
        name_label.pack(side="left", fill="x", expand=True, padx=5)
        self.start_marquee(name_label, file_name, visible_chars=20)

    def start_marquee(self, label, full_text, visible_chars=20):
        """Slowly scroll `full_text` through `label` if it's too long to fit.
        Short text is shown as-is with no animation."""
        if len(full_text) <= visible_chars:
            label.configure(text=full_text)
            return

        padded = full_text + "   •   "
        position = {"i": 0}

        def tick():
            if not label.winfo_exists():
                return
            i = position["i"]
            window_text = (padded + padded)[i:i + visible_chars]
            label.configure(text=window_text)
            position["i"] = (i + 1) % len(padded)
            job = self.window.after(280, tick)
            self._marquee_jobs.append(job)

        tick()

    def play_file(self, file_path):
        """Open a media file with the OS default player. This just launches
        whatever the system has associated with the extension (e.g. Windows
        Media Player, QuickTime, VLC) — it doesn't embed playback in-app."""
        if not os.path.exists(file_path):
            self.show_toast("File no longer exists at that path", "red")
            return
        system = platform.system()
        try:
            if system == "Windows":
                os.startfile(file_path)
            elif system == "Darwin":
                import subprocess
                subprocess.run(['open', file_path], timeout=2)
            else:
                import subprocess
                subprocess.run(['xdg-open', file_path], timeout=2)
            self.show_toast(f"Playing {os.path.basename(file_path)}", "blue")
        except Exception as e:
            self.show_toast(f"Couldn't open file: {str(e)}", "red")

    def format_size(self, size):
        """Format file size"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"

    # ------------------------------------------------------------------
    # Search
    # ------------------------------------------------------------------

    def on_search_changed(self, event=None):
        self.search_query = self.search_entry.get()
        self.update_display()

    def clear_search(self):
        self.search_entry.delete(0, 'end')
        self.search_query = ""
        self.update_display()

    # ------------------------------------------------------------------
    # Bulk actions
    # ------------------------------------------------------------------

    def select_all_items(self):
        """Select all items in the list"""
        if not self.multiple_selection_mode:
            if len(self.clipboard_history) > 0:
                last_index = len(self.clipboard_history) - 1
                for i in self.selected_items:
                    self.selected_items[i].set(False)
                    if i in self.checkboxes:
                        self.checkboxes[i].deselect()
                if last_index in self.selected_items:
                    self.selected_items[last_index].set(True)
                    if last_index in self.checkboxes:
                        self.checkboxes[last_index].select()
                    self.copy_item_to_clipboard(last_index)
                    self.show_toast(f"Item #{last_index + 1} selected (single mode)", "green")
        else:
            for i in range(len(self.clipboard_history)):
                if i in self.selected_items:
                    self.selected_items[i].set(True)
                    if i in self.checkboxes:
                        self.checkboxes[i].select()

            self.copy_selected_items()
            if len(self.clipboard_history) > 0:
                self.show_toast(f"All {len(self.clipboard_history)} items selected", "green")

        self.update_selection_label()
        self.update_preview()

    def deselect_all_items(self):
        """Deselect all items in the list"""
        for i in range(len(self.clipboard_history)):
            if i in self.selected_items:
                self.selected_items[i].set(False)
                if i in self.checkboxes:
                    self.checkboxes[i].deselect()

        self.update_selection_label()
        self.update_preview()
        self.show_toast("All items deselected", "gray")

    def delete_item(self, index):
        """Delete specific item from history (maintains order of remaining items)"""
        if 0 <= index < len(self.clipboard_history):
            del self.clipboard_history[index]
            del self.clipboard_items_type[index]
            del self.clipboard_pinned[index]
            del self.clipboard_timestamps[index]

            self.selected_items.clear()
            self.checkboxes.clear()

            self.save_history()
            self.update_display()
            self.show_toast(f"Item #{index + 1} deleted", "red")

    def clear_history(self):
        """Clear all clipboard history (with confirmation)"""
        if not self.clipboard_history:
            self.show_toast("History is already empty", "gray")
            return

        from tkinter import messagebox
        if not messagebox.askyesno(
            "Clear All",
            "This will permanently delete all clipboard history, including pinned items. Continue?"
        ):
            return

        self.clipboard_history.clear()
        self.clipboard_items_type.clear()
        self.clipboard_pinned.clear()
        self.clipboard_timestamps.clear()
        self.selected_items.clear()
        self.checkboxes.clear()
        self.save_history()
        self.update_display()
        self.show_toast("History cleared!", "red")

    # ------------------------------------------------------------------
    # Fun stuff
    # ------------------------------------------------------------------

    def clipboard_roulette(self):
        """Randomly re-copy a past clipboard item, just for fun"""
        if not self.clipboard_history:
            self.show_toast("Nothing to roll — clipboard's empty!", "gray")
            return
        index = random.randrange(len(self.clipboard_history))
        self.copy_item_to_clipboard(index)
        message = random.choice(self.ROULETTE_MESSAGES).format(n=index + 1)
        self.show_toast(message, "purple")

    def open_stats_dialog(self):
        """Show lifetime clipboard stats, with a silly novel/tweet comparison"""
        dialog = ctk.CTkToplevel(self.window)
        dialog.title("Clipboard Stats")
        dialog.geometry("340x340")
        dialog.transient(self.window)
        dialog.grab_set()
        if self.topmost_var.get():
            dialog.attributes('-topmost', True)

        title = ctk.CTkLabel(
            dialog, text="📊 Your Clipboard Stats", font=ctk.CTkFont(size=16, weight="bold")
        )
        title.pack(pady=(15, 10))

        total_items = self.stats["total_items"]
        total_chars = self.stats["total_characters"]
        novel_pct = (total_chars / 500_000) * 100  # ~500k chars ≈ an average novel
        tweets = total_chars // 280
        pinned_count = sum(1 for p in self.clipboard_pinned if p)

        stats_text = (
            f"Items copied all-time: {total_items:,}\n"
            f"Characters copied all-time: {total_chars:,}\n\n"
            f"That's {novel_pct:.3f}% of an average novel,\n"
            f"or roughly {tweets:,} tweets worth of text.\n\n"
            f"Current session: {len(self.clipboard_history)} item(s)\n"
            f"Pinned right now: {pinned_count}"
        )
        label = ctk.CTkLabel(
            dialog, text=stats_text, font=ctk.CTkFont(size=12), justify="left"
        )
        label.pack(pady=10, padx=20)

        close_button = ctk.CTkButton(dialog, text="Nice!", command=dialog.destroy)
        close_button.pack(pady=10)

    # ------------------------------------------------------------------
    # Toasts
    # ------------------------------------------------------------------

    def show_toast(self, message, color="green"):
        """Show a temporary toast message, stacking with any already visible"""
        slot = len(self.active_toasts)
        toast = ctk.CTkLabel(
            self.window, text=message, font=ctk.CTkFont(size=12),
            fg_color=color, corner_radius=10, width=350, height=30
        )
        toast.place(relx=0.5, rely=0.9 - (slot * 0.07), anchor="center")
        self.active_toasts.append(toast)

        def remove():
            if toast in self.active_toasts:
                self.active_toasts.remove(toast)
            toast.destroy()
            for i, t in enumerate(self.active_toasts):
                t.place(relx=0.5, rely=0.9 - (i * 0.07), anchor="center")

        self.window.after(1800, remove)

    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------

    def open_max_items_dialog(self):
        """Open dialog to set maximum items"""
        dialog = ctk.CTkToplevel(self.window)
        dialog.title("Set Max Items")
        dialog.geometry("300x180")
        dialog.transient(self.window)
        dialog.grab_set()

        if self.topmost_var.get():
            dialog.attributes('-topmost', True)

        label = ctk.CTkLabel(dialog, text="Set maximum items (1-50):", font=ctk.CTkFont(size=14))
        label.pack(pady=10)

        entry = ctk.CTkEntry(dialog, width=100, placeholder_text=str(self.max_items))
        entry.pack(pady=10)

        def set_max_items():
            try:
                new_max = int(entry.get())
                if 1 <= new_max <= 50:
                    self.max_items = new_max

                    if len(self.clipboard_history) > new_max:
                        pinned_idx = [i for i, p in enumerate(self.clipboard_pinned) if p]
                        unpinned_idx = [i for i, p in enumerate(self.clipboard_pinned) if not p]
                        keep_count = max(0, new_max - len(pinned_idx))
                        keep_unpinned = set(unpinned_idx[-keep_count:]) if keep_count > 0 else set()
                        keep_indices = sorted(set(pinned_idx) | keep_unpinned)

                        self.clipboard_history = [self.clipboard_history[i] for i in keep_indices]
                        self.clipboard_items_type = [self.clipboard_items_type[i] for i in keep_indices]
                        self.clipboard_pinned = [self.clipboard_pinned[i] for i in keep_indices]
                        self.clipboard_timestamps = [self.clipboard_timestamps[i] for i in keep_indices]

                        self.show_toast(f"Kept {len(keep_indices)} items (pinned items protected)", "orange")

                    self.selected_items.clear()
                    self.checkboxes.clear()

                    self.save_history()
                    dialog.destroy()
                    self.update_display()
                    self.show_toast(f"Max items set to {new_max}", "green")
                else:
                    entry.delete(0, 'end')
                    entry.insert(0, "Invalid!")
            except ValueError:
                entry.delete(0, 'end')
                entry.insert(0, "Enter number!")

        button = ctk.CTkButton(dialog, text="Set", command=set_max_items)
        button.pack(pady=10)

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    def run(self):
        """Start the application"""
        self.window.mainloop()

    def stop_monitoring(self):
        """Stop clipboard monitoring"""
        self.monitoring = False


if __name__ == "__main__":
    app = ClipboardManager()
    try:
        app.run()
    finally:
        app.stop_monitoring()