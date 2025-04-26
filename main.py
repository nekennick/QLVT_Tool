import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pyperclip
import os
import sys
import json
import re
from tkinter import font
from functools import lru_cache

class QLVTApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QLVT Tool V2")
        self.root.geometry("600x500")
        self.root.minsize(500, 400)
        
        # Biến theo dõi trạng thái ghim
        self.is_pinned = False
        
        # Data storage
        self.items = []
        self.filtered_items = []
        self.bookmarked_items = []
        self.status_message = ""
        self.status_timer = None
        self.search_timer = None
        self.last_query = ""
        
        # Search index for faster lookups
        self.search_index = {}
        
        # Create custom styles
        self.setup_styles()
        
        # Create the UI
        self.create_ui()
        
        # Load existing data if available
        self.load_data()
    
    def create_ui(self):
        # Create main frame with modern style
        main_frame = ttk.Frame(self.root, style="Main.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Top frame for buttons and search
        top_frame = ttk.Frame(main_frame, style="Main.TFrame")
        top_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Left frame for buttons
        button_frame = ttk.Frame(top_frame, style="Main.TFrame")
        button_frame.pack(side=tk.LEFT)
        
        # Import button with accent color
        import_btn = ttk.Button(button_frame, 
                              text="📂 Import Excel",
                              style="Accent.TButton",
                              command=self.import_excel)
        import_btn.pack(side=tk.LEFT, padx=(0, 8))
        
        # Pin button with special style
        self.pin_btn = ttk.Button(button_frame,
                                text="📌 Ghim",
                                style="Pin.TButton",
                                command=self.toggle_pin)
        self.pin_btn.pack(side=tk.LEFT, padx=(0, 8))
        
        # Search frame with modern style
        search_frame = ttk.Frame(top_frame, style="Main.TFrame")
        search_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        
        # Search label with icon
        search_lbl = ttk.Label(search_frame,
                             text="🔍 Tìm kiếm:",
                             background="#ffffff",
                             font=("Segoe UI", 9))
        search_lbl.pack(side=tk.LEFT, padx=(0, 8))
        
        # Search variable and entry with modern style
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.on_search_input)
        search_entry = ttk.Entry(search_frame,
                               textvariable=self.search_var,
                               style="Search.TEntry")
        search_entry.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(0, 5))
        
        # Create a frame for the list with scrollbar
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Frame for bookmarked items
        self.bookmarked_frame = ttk.Frame(list_frame)
        self.bookmarked_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Separator between bookmarked and normal items
        self.separator = ttk.Separator(list_frame, orient='horizontal')
        
        # Frame for normal items with scrollbar
        normal_items_frame = ttk.Frame(list_frame)
        normal_items_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(normal_items_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Canvas for scrollable content
        self.canvas = tk.Canvas(normal_items_frame, yscrollcommand=scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.canvas.yview)
        
        # Frame inside canvas for items
        self.items_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.items_frame, anchor="nw")
        
        # Configure canvas to scroll with mousewheel
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.items_frame.bind("<Configure>", self.on_frame_configure)
        self.root.bind_all("<MouseWheel>", self.on_mousewheel)
        
        # Status bar at the bottom
        self.status_bar = ttk.Label(main_frame, text="", anchor=tk.W)
        self.status_bar.pack(fill=tk.X, pady=(10, 0))
        
        # Bind the drag and drop events
        self.drag_data = {"widget": None, "index": -1, "y_pos": 0}
        
    def setup_styles(self):
        """Set up custom styles for the UI elements"""
        style = ttk.Style()
        
        # Configure colors
        style.configure(".", font=("Segoe UI", 9))
        style.configure("Item.TFrame", background="#ffffff")
        style.configure("DragActive.TFrame", background="#e3f2fd")
        style.configure("Bookmarked.TFrame", background="#fff8e1")
        
        # Buttons
        style.configure("Accent.TButton", 
                      padding=2,
                      background="#f0f0f0",
                      foreground="black")
        style.map("Accent.TButton",
                 background=[("active", "#e0e0e0")])
        
        style.configure("Secondary.TButton",
                      padding=2,
                      background="#f5f5f5")
        style.map("Secondary.TButton",
                 background=[("active", "#e0e0e0")])
        
        # Pin button
        style.configure("Pin.TButton",
                      padding=2,
                      background="#f0f0f0")
        
        # Search entry
        style.configure("Search.TEntry",
                      padding=5,
                      fieldbackground="#f5f5f5")
        
        # Configure root window
        self.root.configure(bg="#ffffff")
        
        # Configure main frame padding
        style.configure("Main.TFrame", background="#ffffff", padding=10)
        style.configure("List.TFrame", background="#ffffff", padding=2)
    
    def toggle_pin(self):
        """Toggle window pin state"""
        self.is_pinned = not self.is_pinned
        if self.is_pinned:
            self.root.attributes('-topmost', True)
            self.pin_btn.configure(text="📌 Bỏ ghim", style="Pinned.TButton")
            self.show_status_message("✅ Đã ghim cửa sổ")
        else:
            self.root.attributes('-topmost', False)
            self.pin_btn.configure(text="📌 Ghim", style="")
            self.show_status_message("✅ Đã bỏ ghim cửa sổ")
    
    def on_canvas_configure(self, event):
        # Ensure items frame width matches canvas width
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        
    def on_frame_configure(self, event):
        # Reset the scroll region to encompass the inner frame
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def on_mousewheel(self, event):
        # Windows uses 'delta' with different values
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def toggle_bookmark(self, item):
        """Toggle bookmark status for an item"""
        # Check if item is already bookmarked
        is_bookmarked = False
        for i, bookmarked in enumerate(self.bookmarked_items):
            if bookmarked["code"] == item["code"]:
                # Remove from bookmarks
                self.bookmarked_items.pop(i)
                is_bookmarked = True
                self.show_status_message("✅ Đã bỏ đánh dấu")
                break
        
        if not is_bookmarked:
            # Add to bookmarks
            self.bookmarked_items.append(item)
            self.show_status_message("✅ Đã đánh dấu")
        
        # Save changes
        self.save_data()
        
        # Refresh display
        self.display_items()
    
    def display_items(self):
        # Clear existing items
        for widget in self.bookmarked_frame.winfo_children():
            widget.destroy()
        for widget in self.items_frame.winfo_children():
            widget.destroy()
        
        # Get items to display
        items_to_display = self.filtered_items if self.filtered_items else self.items
        
        # Display bookmarked items first
        if self.bookmarked_items and not self.filtered_items:
            for index, item in enumerate(self.bookmarked_items):
                self.create_item_widget(item, index, self.bookmarked_frame)
            self.separator.pack(fill=tk.X, pady=5)
        else:
            self.separator.pack_forget()
        
        # Display other items
        for index, item in enumerate(items_to_display):
            if not self.filtered_items and any(b['code'] == item['code'] for b in self.bookmarked_items):
                continue  # Skip if item is bookmarked and we're not filtering
            self.create_item_widget(item, index)
        
        # Update the canvas scrollregion
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def create_item_widget(self, item, index, parent_frame=None):
        # Create a frame for each item with modern style
        item_frame = ttk.Frame(parent_frame if parent_frame else self.items_frame)
        item_frame.pack(fill=tk.X, pady=1, padx=2)
        
        # Make the frame draggable if it's in bookmarked items
        if parent_frame == self.bookmarked_frame:
            item_frame.bind("<ButtonPress-1>", lambda e, idx=index: self.on_drag_start(e, idx, item_frame))
            item_frame.bind("<B1-Motion>", self.on_drag_motion)
            item_frame.bind("<ButtonRelease-1>", self.on_drag_release)
        
        # Set a distinctive background with hover effect
        is_bookmarked = any(b['code'] == item['code'] for b in self.bookmarked_items)
        item_frame.configure(style="Bookmarked.TFrame" if is_bookmarked else "Item.TFrame")
        
        # Display item code and name (truncated if too long)
        name_text = item["name"]
        if len(name_text) > 40:
            name_text = name_text[:40] + "..."
        
        # Create a label with code and truncated name
        item_label = ttk.Label(item_frame,
                             text=f"{item['code']} - {name_text}",
                             background="#ffffff" if not is_bookmarked else "#fff8e1",
                             font=("Segoe UI", 9))
        item_label.pack(side=tk.LEFT, anchor=tk.W, padx=(4, 0), pady=0)
        
        # Double-click to edit
        item_label.bind("<Double-1>", lambda e, i=item, idx=index: self.edit_item(i, idx))
        
        # Button frame for multiple buttons
        btn_frame = ttk.Frame(item_frame, style="Main.TFrame")
        btn_frame.pack(side=tk.RIGHT, padx=2)
        
        # Bookmark button with star icon
        bookmark_text = "⭐" if is_bookmarked else "☆"
        bookmark_btn = ttk.Button(btn_frame,
                                text=bookmark_text,
                                style="Secondary.TButton",
                                width=3,
                                command=lambda i=item: self.toggle_bookmark(i))
        bookmark_btn.pack(side=tk.RIGHT, padx=2)
        
        # Copy button with modern style
        copy_btn = ttk.Button(btn_frame,
                            text="📋 Copy",
                            style="Secondary.TButton",
                            width=8,
                            command=lambda i=item: self.copy_item_code(i))
        copy_btn.pack(side=tk.RIGHT, padx=2)
    
    def copy_item_code(self, item):
        pyperclip.copy(item["code"])
        self.show_status_message("✅ Đã copy")
    
    def show_status_message(self, message, duration=2000):
        # Clear any existing timer
        if self.status_timer:
            self.root.after_cancel(self.status_timer)
        
        # Update the status message
        self.status_bar.config(text=message)
        
        # Set a timer to clear the message
        self.status_timer = self.root.after(duration, lambda: self.status_bar.config(text=""))
    
    def import_excel(self):
        file_path = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=[("Excel files", "*.xlsx")],
        )
        
        if not file_path:
            return
        
        try:
            # Read Excel file
            df = pd.read_excel(file_path)
            
            # Check if required columns exist
            if "Mã VT" not in df.columns or "Tên VT" not in df.columns:
                messagebox.showerror(
                    "Lỗi",
                    "File Excel không đúng định dạng. Cần có cột 'Mã VT' và 'Tên VT'."
                )
                return
            
            # Extract relevant columns and convert to list of dictionaries
            self.items = [
                {"code": str(row["Mã VT"]).strip(), "name": str(row["Tên VT"]).strip(),
                 "code_lower": str(row["Mã VT"]).strip().lower(), "name_lower": str(row["Tên VT"]).strip().lower()}
                for _, row in df.iterrows()
                if pd.notna(row["Mã VT"]) and pd.notna(row["Tên VT"])
            ]
            
            # Build search index
            self.build_search_index()
            
            # Reset search and display items
            self.search_var.set("")
            self.filtered_items = []
            self.display_items()
            
            # Save data
            self.save_data()
            
            # Show success message
            self.show_status_message(f"✅ Đã import {len(self.items)} vật tư")
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đọc file Excel: {str(e)}")
    
    def on_search_input(self, *args):
        # Cancel any existing timer
        if self.search_timer:
            self.root.after_cancel(self.search_timer)
        
        # Start a new timer (300ms delay)
        self.search_timer = self.root.after(300, self.perform_search)
    
    def perform_search(self):
        # Get search query
        query = self.search_var.get().lower()
        
        # If query is same as last search, skip processing
        if query == self.last_query:
            return
            
        self.last_query = query
        
        if not query:
            # If search is empty, show all items
            self.filtered_items = []
        else:
            # Use optimized search
            self.filtered_items = self.search_items(query)
        
        # Update display
        self.display_items()
    
    def build_search_index(self):
        """Build an index to speed up searches"""
        self.search_index = {}
        
        # Add words from both codes and names to the index
        for i, item in enumerate(self.items):
            # Add whole code
            code = item["code_lower"]
            if code not in self.search_index:
                self.search_index[code] = []
            self.search_index[code].append(i)
            
            # Add parts of the code (e.g., for 'VT001', index 'VT' and '001')
            code_parts = re.findall(r'[a-z]+|\d+', code)
            for part in code_parts:
                if part not in self.search_index:
                    self.search_index[part] = []
                if i not in self.search_index[part]:
                    self.search_index[part].append(i)
            
            # Add words from name
            name = item["name_lower"]
            words = re.findall(r'\w+', name)
            for word in words:
                if word not in self.search_index:
                    self.search_index[word] = []
                if i not in self.search_index[word]:
                    self.search_index[word].append(i)
    
    @lru_cache(maxsize=128)
    def get_matching_words(self, query):
        """Get all words from the index that contain the query"""
        return [word for word in self.search_index.keys() if query in word]
    
    def search_items(self, query):
        """Optimized search using the index"""
        # For very short queries, fall back to direct search
        if len(query) < 2:
            return [item for item in self.items 
                   if query in item["code_lower"] or query in item["name_lower"]]
        
        # Get all words that contain the query
        matching_words = self.get_matching_words(query)
        
        # Get item indices for these words
        matching_indices = set()
        for word in matching_words:
            matching_indices.update(self.search_index[word])
        
        # Direct match for query in code or name
        for i, item in enumerate(self.items):
            if query in item["code_lower"] or query in item["name_lower"]:
                matching_indices.add(i)
        
        # Return items at these indices
        return [self.items[i] for i in matching_indices]
    
    def edit_item(self, item, index):
        # Create a dialog for editing
        dialog = tk.Toplevel(self.root)
        dialog.title("Sửa vật tư")
        dialog.geometry("400x150")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.geometry("+%d+%d" % (
            self.root.winfo_rootx() + (self.root.winfo_width() // 2) - (400 // 2),
            self.root.winfo_rooty() + (self.root.winfo_height() // 2) - (150 // 2)
        ))
        
        # Create a frame for the form
        form_frame = ttk.Frame(dialog, padding="10")
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Code label and entry
        code_lbl = ttk.Label(form_frame, text="Mã vật tư:")
        code_lbl.grid(row=0, column=0, sticky=tk.W, pady=5)
        
        code_var = tk.StringVar(value=item["code"])
        code_entry = ttk.Entry(form_frame, textvariable=code_var, width=30)
        code_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # Name label and entry
        name_lbl = ttk.Label(form_frame, text="Tên vật tư:")
        name_lbl.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        name_var = tk.StringVar(value=item["name"])
        name_entry = ttk.Entry(form_frame, textvariable=name_var, width=30)
        name_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Buttons
        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        save_btn = ttk.Button(
            btn_frame,
            text="Lưu",
            command=lambda: self.save_edited_item(index, code_var.get(), name_var.get(), dialog)
        )
        save_btn.pack(side=tk.LEFT, padx=5)
        
        cancel_btn = ttk.Button(
            btn_frame,
            text="Hủy",
            command=dialog.destroy
        )
        cancel_btn.pack(side=tk.LEFT, padx=5)
    
    def save_edited_item(self, index, code, name, dialog):
        """Save edited item and close dialog"""
        if not code or not name:
            messagebox.showerror("Lỗi", "Mã và tên vật tư không được để trống")
            return
        
        # Update item in the list
        if index < len(self.items):
            old_code = self.items[index]["code"]
            self.items[index]["code"] = code
            self.items[index]["name"] = name
            self.items[index]["code_lower"] = code.lower()
            self.items[index]["name_lower"] = name.lower()
            
            # Find and update the item widget
            for widget in self.items_frame.winfo_children():
                if isinstance(widget, ttk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Label) and old_code in child["text"]:
                            child.configure(text=f"{code} - {name}")
                            break
        
        # Update bookmarked items if needed
        for i, item in enumerate(self.bookmarked_items):
            if item["code"] == old_code:
                self.bookmarked_items[i]["code"] = code
                self.bookmarked_items[i]["name"] = name
                self.bookmarked_items[i]["code_lower"] = code.lower()
                self.bookmarked_items[i]["name_lower"] = name.lower()
                # Find and update the bookmarked item widget
                for widget in self.bookmarked_frame.winfo_children():
                    if isinstance(widget, ttk.Frame):
                        for child in widget.winfo_children():
                            if isinstance(child, ttk.Label) and old_code in child["text"]:
                                child.configure(text=f"{code} - {name}")
                                break
        
        # Close dialog
        dialog.destroy()
        
        # Save data
        self.save_data()
        
        # Show success message
        self.show_status_message("✅ Đã lưu thay đổi")
    def on_drag_start(self, event, index, frame):
        # Record the widget and its starting position
        self.drag_data["widget"] = frame
        self.drag_data["index"] = index
        self.drag_data["y_pos"] = event.y_root
        
        # Visual feedback - change background color
        frame.configure(style="DragActive.TFrame")
        
        # Store whether we're dragging a bookmarked item
        self.drag_data["is_bookmark"] = frame.master == self.bookmarked_frame
    
    def on_drag_motion(self, event):
        if self.drag_data["widget"]:
            # Get the dragged widget
            widget = self.drag_data["widget"]
            
            # Get mouse coordinates relative to the canvas
            y = event.y_root
            
            # Auto-scroll if near the edges
            canvas_height = self.canvas.winfo_height()
            canvas_y = y - self.canvas.winfo_rooty()
            
            if canvas_y < 20 and self.canvas.yview()[0] > 0:
                # Scroll up
                self.canvas.yview_scroll(-1, "units")
            elif canvas_y > canvas_height - 20 and self.canvas.yview()[1] < 1:
                # Scroll down
                self.canvas.yview_scroll(1, "units")
    
    def on_drag_release(self, event):
        if self.drag_data["widget"] and self.drag_data["index"] >= 0:
            # Only handle drag and drop for bookmarked items
            if not self.drag_data.get("is_bookmark"):
                self.drag_data["widget"].configure(style="Item.TFrame")
                return
                
            # Reset visual style
            is_bookmarked = any(b['code'] == self.bookmarked_items[self.drag_data["index"]]['code'] 
                               for b in self.bookmarked_items)
            self.drag_data["widget"].configure(
                style="Bookmarked.TFrame" if is_bookmarked else "Item.TFrame")
            
            # Get all bookmark frames
            frames = [w for w in self.bookmarked_frame.winfo_children() if isinstance(w, ttk.Frame)]
            
            # Find the closest frame to drop position
            drop_y = event.y_root
            
            closest_idx = -1
            min_distance = float('inf')
            
            for i, frame in enumerate(frames):
                frame_y = frame.winfo_rooty() + frame.winfo_height() // 2
                distance = abs(drop_y - frame_y)
                
                if distance < min_distance:
                    min_distance = distance
                    closest_idx = i
            
            # If we have a valid drop target and it's different from source
            if closest_idx >= 0 and closest_idx != self.drag_data["index"]:
                # Move item in bookmarked list
                moved_item = self.bookmarked_items.pop(self.drag_data["index"])
                self.bookmarked_items.insert(closest_idx, moved_item)
                
                # Save the updated order
                self.save_data()
                
                # Refresh the display
                self.display_items()
                
                # Show status message
                self.show_status_message("✅ Đã thay đổi thứ tự bookmark")
            
            # Reset drag data
            self.drag_data = {"widget": None, "index": -1, "y_pos": 0, "is_bookmark": False}
    
    def save_data(self):
        """Save items data to a JSON file"""
        try:
            # Remove preprocessing fields before saving
            save_data = {
                "items": [
                    {"code": item["code"], "name": item["name"]}
                    for item in self.items
                ],
                "bookmarks": [
                    {"code": item["code"], "name": item["name"]}
                    for item in self.bookmarked_items
                ]
            }
                
            # Get the directory containing the executable in PyInstaller bundle
            if getattr(sys, 'frozen', False):
                application_path = os.path.dirname(sys.executable)
            else:
                application_path = os.path.dirname(os.path.abspath(__file__))
            
            data_file = os.path.join(application_path, "data.json")
            with open(data_file, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.show_status_message(f"❌ Lỗi lưu dữ liệu: {str(e)}")
    
    def load_data(self):
        """Load items data from JSON file if it exists"""
        try:
            # Get the directory containing the executable in PyInstaller bundle
            if getattr(sys, 'frozen', False):
                application_path = os.path.dirname(sys.executable)
            else:
                application_path = os.path.dirname(os.path.abspath(__file__))
            
            data_file = os.path.join(application_path, "data.json")
            if os.path.exists(data_file):
                with open(data_file, 'r', encoding='utf-8') as f:
                    loaded_data = json.load(f)
                
                # Handle both old and new format
                if isinstance(loaded_data, list):
                    loaded_items = loaded_data
                    loaded_bookmarks = []
                else:
                    loaded_items = loaded_data.get("items", [])
                    loaded_bookmarks = loaded_data.get("bookmarks", [])
                
                # Add preprocessing fields
                self.items = []
                for item in loaded_items:
                    self.items.append({
                        "code": item["code"],
                        "name": item["name"],
                        "code_lower": item["code"].lower(),
                        "name_lower": item["name"].lower()
                    })
                
                # Load bookmarks
                self.bookmarked_items = []
                for item in loaded_bookmarks:
                    self.bookmarked_items.append({
                        "code": item["code"],
                        "name": item["name"],
                        "code_lower": item["code"].lower(),
                        "name_lower": item["name"].lower()
                    })
                
                # Build search index
                self.build_search_index()
                
                self.display_items()
                self.show_status_message(f"✅ Đã tải {len(self.items)} vật tư")
        except Exception as e:
            self.show_status_message(f"❌ Lỗi tải dữ liệu: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = QLVTApp(root)
    root.mainloop()
