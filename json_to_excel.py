"""
JSON to Excel Converter
A GUI application that converts JSON arrays to Excel files.

Author: Michael Dehne
License: MIT License
Version: 1.0.0

Copyright (c) 2024 Michael Dehne
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import pandas as pd
import os
import subprocess
import platform
from pathlib import Path
import re

class JsonToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("JSON to Excel Converter")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Variables
        self.json_file_path = tk.StringVar()
        self.selected_array_key = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_filename = tk.StringVar()
        self.json_data = None
        self.array_keys = []
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="JSON to Excel Converter", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="Step 1: Select JSON File", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="JSON File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(file_frame, textvariable=self.json_file_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(file_frame, text="Browse", command=self.browse_json_file).grid(row=0, column=2)
        
        # Array selection section
        array_frame = ttk.LabelFrame(main_frame, text="Step 2: Select Array to Convert", padding="10")
        array_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        array_frame.columnconfigure(0, weight=1)
        
        ttk.Label(array_frame, text="Available Arrays:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # Combobox for array selection
        self.array_combobox = ttk.Combobox(array_frame, textvariable=self.selected_array_key, 
                                          state="readonly", width=50)
        self.array_combobox.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.array_combobox.bind('<<ComboboxSelected>>', self.on_array_selected)
        
        # Array info label
        self.array_info_label = ttk.Label(array_frame, text="No JSON file loaded")
        self.array_info_label.grid(row=2, column=0, sticky=tk.W)
        
        # Output settings section
        output_frame = ttk.LabelFrame(main_frame, text="Step 3: Set Output Location", padding="10")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(1, weight=1)
        
        # Output folder selection
        ttk.Label(output_frame, text="Save Folder:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(output_frame, textvariable=self.output_folder, width=40).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(output_frame, text="Browse", command=self.browse_output_folder).grid(row=0, column=2)
        
        # Output filename
        ttk.Label(output_frame, text="File Name:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        ttk.Entry(output_frame, textvariable=self.output_filename, width=40).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(10, 0))
        ttk.Label(output_frame, text=".xlsx").grid(row=1, column=2, sticky=tk.W, pady=(10, 0))
        
        # Process button
        self.process_button = ttk.Button(main_frame, text="Convert to Excel", 
                                       command=self.convert_to_excel, state="disabled")
        self.process_button.grid(row=4, column=0, columnspan=3, pady=20)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready to convert JSON to Excel")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=(10, 0))
        
    def browse_json_file(self):
        """Open file dialog to select JSON file"""
        file_path = filedialog.askopenfilename(
            title="Select JSON File",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            self.json_file_path.set(file_path)
            self.load_json_file()
    
    def browse_output_folder(self):
        """Open folder dialog to select output directory"""
        folder_path = filedialog.askdirectory(
            title="Select Output Folder"
        )
        
        if folder_path:
            self.output_folder.set(folder_path)
            # Auto-generate filename based on selected array
            if self.selected_array_key.get():
                suggested_name = f"{self.selected_array_key.get()}_data"
                self.output_filename.set(suggested_name)
    
    def load_json_file(self):
        """Load and parse the selected JSON file"""
        try:
            with open(self.json_file_path.get(), 'r', encoding='utf-8') as file:
                self.json_data = json.load(file)
            
            # Find all arrays in the JSON
            self.array_keys = self.find_arrays(self.json_data)
            
            if self.array_keys:
                self.array_combobox['values'] = self.array_keys
                self.array_combobox.set(self.array_keys[0])
                self.selected_array_key.set(self.array_keys[0])
                
                # Update info label
                selected_key = self.array_keys[0]
                array_data = self.get_array_data(self.json_data, selected_key)
                self.update_array_info(selected_key, array_data)
                
                self.process_button['state'] = 'normal'
                self.status_label['text'] = f"Found {len(self.array_keys)} array(s) in JSON file"
            else:
                self.array_combobox['values'] = []
                self.array_info_label['text'] = "No arrays found in JSON file"
                self.process_button['state'] = 'disabled'
                self.status_label['text'] = "No arrays found in JSON file"
                
        except json.JSONDecodeError as e:
            messagebox.showerror("Error", f"Invalid JSON file: {str(e)}")
            self.status_label['text'] = "Error: Invalid JSON file"
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")
            self.status_label['text'] = "Error loading file"
    
    def find_arrays(self, data, path=""):
        """Recursively find all arrays in the JSON data"""
        arrays = []
        
        if isinstance(data, dict):
            for key, value in data.items():
                current_path = f"{path}.{key}" if path else key
                if isinstance(value, list) and value:
                    arrays.append(current_path)
                elif isinstance(value, (dict, list)):
                    arrays.extend(self.find_arrays(value, current_path))
        
        elif isinstance(data, list) and data:
            # If the root is an array, add it
            if not path:
                arrays.append("root")
            else:
                arrays.append(path)
        
        return arrays
    
    def get_array_data(self, data, array_path):
        """Get the array data for a given path"""
        if array_path == "root":
            return data
        
        keys = array_path.split('.')
        current = data
        
        for key in keys:
            if isinstance(current, dict) and key in current:
                current = current[key]
            else:
                return None
        
        return current if isinstance(current, list) else None
    
    def update_array_info(self, array_key, array_data):
        """Update the info label with array details"""
        if array_data and len(array_data) > 0:
            sample_item = array_data[0]
            if isinstance(sample_item, dict):
                columns = list(sample_item.keys())
                self.array_info_label['text'] = f"Array '{array_key}': {len(array_data)} items, {len(columns)} columns"
            else:
                self.array_info_label['text'] = f"Array '{array_key}': {len(array_data)} items (non-object array)"
        else:
            self.array_info_label['text'] = f"Array '{array_key}': Empty array"
    
    def on_array_selected(self, event=None):
        """Handle array selection change"""
        selected_key = self.selected_array_key.get()
        if selected_key and self.json_data:
            array_data = self.get_array_data(self.json_data, selected_key)
            self.update_array_info(selected_key, array_data)
            
            # Auto-generate filename if output folder is set
            if self.output_folder.get():
                suggested_name = f"{selected_key}_data"
                self.output_filename.set(suggested_name)
    
    def convert_to_excel(self):
        """Convert the selected array to Excel"""
        if not self.selected_array_key.get():
            messagebox.showerror("Error", "Please select an array to convert")
            return
        
        if not self.output_folder.get():
            messagebox.showerror("Error", "Please select an output folder")
            return
        
        if not self.output_filename.get():
            messagebox.showerror("Error", "Please enter a filename")
            return
        
        # Get the array data
        array_data = self.get_array_data(self.json_data, self.selected_array_key.get())
        
        if not array_data:
            messagebox.showerror("Error", "Selected array is empty or invalid")
            return
        
        # Check if array contains objects
        if not array_data or not isinstance(array_data[0], dict):
            messagebox.showerror("Error", "Array must contain objects with key-value pairs")
            return
        
        # Create DataFrame
        try:
            df = pd.DataFrame(array_data)
            
            # Build the full file path
            filename = self.output_filename.get().strip()
            if not filename.endswith('.xlsx'):
                filename += '.xlsx'
            
            file_path = os.path.join(self.output_folder.get(), filename)
            
            # Check if file already exists
            if os.path.exists(file_path):
                result = messagebox.askyesno("File Exists", 
                                          f"File '{filename}' already exists in the selected folder.\n"
                                          "Do you want to overwrite it?")
                if not result:
                    return
            
            # Save to Excel
            df.to_excel(file_path, index=False)
            
            messagebox.showinfo("Success", 
                             f"Excel file saved successfully!\n"
                             f"File: {file_path}\n"
                             f"Rows: {len(df)}\n"
                             f"Columns: {len(df.columns)}")
            
            self.status_label['text'] = f"Excel file saved: {filename}"
            
            # Open the Excel file automatically
            self.open_excel_file(file_path)
             
        except Exception as e:
            messagebox.showerror("Error", f"Error creating Excel file: {str(e)}")
            self.status_label['text'] = "Error creating Excel file"
    
    def open_excel_file(self, file_path):
         """Open the Excel file with the default application"""
         try:
             system = platform.system()
             
             if system == "Windows":
                 os.startfile(file_path)
             elif system == "Darwin":  # macOS
                 subprocess.run(["open", file_path], check=True)
             else:  # Linux
                 subprocess.run(["xdg-open", file_path], check=True)
                 
             self.status_label['text'] = f"Excel file opened: {os.path.basename(file_path)}"
             
         except Exception as e:
             # If opening fails, just show a message but don't crash
             messagebox.showwarning("Warning", 
                                 f"Excel file saved successfully, but could not open automatically.\n"
                                 f"File location: {file_path}\n"
                                 f"Error: {str(e)}")
             self.status_label['text'] = f"Excel file saved (could not open): {os.path.basename(file_path)}"

def main():
    root = tk.Tk()
    app = JsonToExcelConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()
