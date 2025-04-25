import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pyperclip
import os
import json
import re
from tkinter import font
from functools import lru_cache

class QLVTApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QLVT Tool V2")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # Data storage
        self.items = []
        self.filtered_items = []
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
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Top frame for buttons and search
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Import button
        import_btn = ttk.Button(top_frame, text="Import Excel", command=self.import_excel)
        import_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Search frame
        search_frame = ttk.Frame(top_frame)
        search_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        
        # Search label
        search_lbl = ttk.Label(search_frame, text="Tìm kiếm:")
        search_lbl.pack(side=tk.LEFT, padx=(0, 5))
        
        # Search variable and entry
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.on_search_input)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        
        # Create a frame for the list with scrollbar
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Canvas for scrollable content
        self.canvas = tk.Canvas(list_frame, yscrollcommand=scrollbar.set)
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
        style.configure("Item.TFrame", background="#f0f0f0")
        style.configure("DragActive.TFrame", background="#d0d0ff")
    
    def on_canvas_configure(self, event):
        # Ensure items frame width matches canvas width
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        
    def on_frame_configure(self, event):
        # Reset the scroll region to encompass the inner frame
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def on_mousewheel(self, event):
        # Windows uses 'delta' with different values
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def display_items(self):
        # Clear the existing items
        for widget in self.items_frame.winfo_children():
            widget.destroy()
        
        # Display the items (filtered or all)
        items_to_display = self.filtered_items if self.filtered_items else self.items
        
        for index, item in enumerate(items_to_display):
            self.create_item_widget(item, index)
        
        # Update the canvas scrollregion
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def create_item_widget(self, item, index):
        # Create a frame for each item
        item_frame = ttk.Frame(self.items_frame)
        item_frame.pack(fill=tk.X, pady=2)
        
        # Make the frame draggable
        item_frame.bind("<ButtonPress-1>", lambda e, idx=index: self.on_drag_start(e, idx, item_frame))
        item_frame.bind("<B1-Motion>", self.on_drag_motion)
        item_frame.bind("<ButtonRelease-1>", self.on_drag_release)
        
        # Set a distinctive background for better visibility during drag
        item_frame.configure(style="Item.TFrame")
        
        # Display item code and name (truncated if too long)
        name_text = item["name"]
        if len(name_text) > 60:
            name_text = name_text[:60] + "..."
        
        # Create a label with code and truncated name
        item_label = ttk.Label(item_frame, text=f"{item['code']} - {name_text}")
        item_label.pack(side=tk.LEFT, anchor=tk.W, padx=(5, 0))
        
        # Double-click to edit
        item_label.bind("<Double-1>", lambda e, i=item, idx=index: self.edit_item(i, idx))
        
        # Copy button
        copy_btn = ttk.Button(item_frame, text="Copy", width=8,
                              command=lambda i=item: self.copy_item_code(i))
        copy_btn.pack(side=tk.RIGHT, padx=5)
    
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
        if not code or not name:
            messagebox.showerror("Lỗi", "Mã và tên vật tư không được để trống")
            return
        
        # Update the item
        items_list = self.filtered_items if self.filtered_items else self.items
        item_index = index if not self.filtered_items else self.items.index(items_list[index])
        
        self.items[item_index]["code"] = code
        self.items[item_index]["name"] = name
        self.items[item_index]["code_lower"] = code.lower()
        self.items[item_index]["name_lower"] = name.lower()
        
        # Rebuild search index
        self.build_search_index()
        
        # Close the dialog
        dialog.destroy()
        
        # Refresh the display
        self.on_search_change()
        
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
            # Reset visual style
            self.drag_data["widget"].configure(style="Item.TFrame")
            
            # Determine new position based on mouse position
            items_list = self.filtered_items if self.filtered_items else self.items
            
            # Get all item frames
            frames = [w for w in self.items_frame.winfo_children() if isinstance(w, ttk.Frame)]
            
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
                if self.filtered_items:
                    # If we're working with filtered items, we need to update both lists
                    moved_item = self.filtered_items.pop(self.drag_data["index"])
                    self.filtered_items.insert(closest_idx, moved_item)
                    
                    # Find the original indices in the full list
                    orig_source_idx = self.items.index(moved_item)
                    orig_item = self.items.pop(orig_source_idx)
                    
                    # Find where to insert in the full list
                    target_item = self.filtered_items[closest_idx]
                    orig_target_idx = self.items.index(target_item)
                    
                    self.items.insert(orig_target_idx, orig_item)
                else:
                    # Direct movement in the main list
                    moved_item = self.items.pop(self.drag_data["index"])
                    self.items.insert(closest_idx, moved_item)
                
                # Save the updated order
                self.save_data()
                
                # Refresh the display
                self.display_items()
                
                # Show status message
                self.show_status_message("✅ Đã thay đổi thứ tự")
            
            # Reset drag data
            self.drag_data = {"widget": None, "index": -1, "y_pos": 0}
    
    def save_data(self):
        """Save items data to a JSON file"""
        try:
            # Remove preprocessing fields before saving
            save_items = []
            for item in self.items:
                save_item = {"code": item["code"], "name": item["name"]}
                save_items.append(save_item)
                
            data_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")
            with open(data_file, 'w', encoding='utf-8') as f:
                json.dump(save_items, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.show_status_message(f"❌ Lỗi lưu dữ liệu: {str(e)}")
    
    def load_data(self):
        """Load items data from JSON file if it exists"""
        try:
            data_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.json")
            if os.path.exists(data_file):
                with open(data_file, 'r', encoding='utf-8') as f:
                    loaded_items = json.load(f)
                
                # Add preprocessing fields
                self.items = []
                for item in loaded_items:
                    self.items.append({
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
