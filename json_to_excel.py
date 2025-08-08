"""
JSON Data Converter
A GUI application that converts JSON arrays to multiple formats (Excel, CSV, JSON).

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
        self.root.title("JSON Data Converter")
        self.root.geometry("800x700")
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
        title_label = ttk.Label(main_frame, text="JSON Data Converter", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="Step 1: Select JSON Source", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # File input option
        ttk.Label(file_frame, text="JSON File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(file_frame, textvariable=self.json_file_path, width=40).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(file_frame, text="Browse", command=self.browse_json_file).grid(row=0, column=2)
        
        # Direct JSON input option
        ttk.Label(file_frame, text="OR Paste JSON:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        ttk.Button(file_frame, text="Open JSON Input", command=self.open_json_input_window).grid(row=1, column=1, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        
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
        
        # Output format selection
        ttk.Label(output_frame, text="Output Format:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.output_format = tk.StringVar(value="excel")
        format_combobox = ttk.Combobox(output_frame, textvariable=self.output_format, 
                                      values=["Excel (.xlsx)", "CSV (.csv)", "JSON (.json)"], 
                                      state="readonly", width=15)
        format_combobox.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        format_combobox.bind('<<ComboboxSelected>>', self.on_format_selected)
        
        # Output folder selection
        ttk.Label(output_frame, text="Save Folder:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        ttk.Entry(output_frame, textvariable=self.output_folder, width=40).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(10, 0))
        ttk.Button(output_frame, text="Browse", command=self.browse_output_folder).grid(row=1, column=2, pady=(10, 0))
        
        # Output filename
        ttk.Label(output_frame, text="File Name:").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        ttk.Entry(output_frame, textvariable=self.output_filename, width=40).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(10, 0))
        self.extension_label = ttk.Label(output_frame, text=".xlsx")
        self.extension_label.grid(row=2, column=2, sticky=tk.W, pady=(10, 0))
        
        # Process buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=20)
        
        # Process button
        self.process_button = ttk.Button(button_frame, text="Convert to File", 
                                       command=self.convert_to_file, state="disabled")
        self.process_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Clear button
        self.clear_button = ttk.Button(button_frame, text="Clear All", 
                                     command=self.clear_all)
        self.clear_button.pack(side=tk.LEFT)
        
        # Status section
        status_frame = ttk.LabelFrame(main_frame, text="Status & Events", padding="10")
        status_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        status_frame.columnconfigure(0, weight=1)
        
        # Status label with better styling
        self.status_label = ttk.Label(status_frame, text="Ready to convert JSON data", 
                                     font=("Arial", 10), foreground="green")
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
    def clear_all(self):
        """Clear all inputs and reset the application state"""
        # Clear file path
        self.json_file_path.set("")
        
        # Clear array selection
        self.selected_array_key.set("")
        self.array_combobox['values'] = []
        self.array_combobox.set("")
        
        # Clear output settings
        self.output_folder.set("")
        self.output_filename.set("")
        
        # Reset data
        self.json_data = None
        self.array_keys = []
        
        # Update UI
        self.array_info_label['text'] = "No JSON file loaded"
        self.process_button['state'] = 'disabled'
        self.status_label['text'] = "Ready to convert JSON data"
        self.status_label['foreground'] = "green"
        
        # Reset format to default
        self.output_format.set("Excel (.xlsx)")
        self.extension_label['text'] = ".xlsx"
        self.process_button['text'] = "Convert to File"
    
    def browse_json_file(self):
        """Open file dialog to select JSON file"""
        file_path = filedialog.askopenfilename(
            title="Select JSON File",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            self.json_file_path.set(file_path)
            self.load_json_file()
    
    def open_json_input_window(self):
        """Open a window for direct JSON input"""
        self.json_input_window = tk.Toplevel(self.root)
        self.json_input_window.title("Paste JSON Data")
        self.json_input_window.geometry("600x400")
        self.json_input_window.resizable(True, True)
        
        # Configure grid weights
        self.json_input_window.columnconfigure(0, weight=1)
        self.json_input_window.rowconfigure(1, weight=1)
        
        # Instructions
        instruction_label = ttk.Label(self.json_input_window, 
                                    text="Paste your JSON data below (e.g., from Postman, API response, etc.):",
                                    font=("Arial", 10, "bold"))
        instruction_label.grid(row=0, column=0, sticky=tk.W, padx=10, pady=(10, 5))
        
        # Text area for JSON input
        self.json_text = tk.Text(self.json_input_window, wrap=tk.WORD, font=("Consolas", 10))
        self.json_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=(0, 10))
        
        # Scrollbar for text area
        scrollbar = ttk.Scrollbar(self.json_input_window, orient=tk.VERTICAL, command=self.json_text.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.json_text.configure(yscrollcommand=scrollbar.set)
        
        # Buttons frame
        button_frame = ttk.Frame(self.json_input_window)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(0, 10))
        
        ttk.Button(button_frame, text="Load JSON", command=self.load_json_from_text).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Clear", command=self.clear_json_text).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Cancel", command=self.json_input_window.destroy).pack(side=tk.LEFT)
        
        # Focus on text area
        self.json_text.focus()
    
    def load_json_from_text(self):
        """Load JSON from the text input"""
        json_text = self.json_text.get("1.0", tk.END).strip()
        
        if not json_text:
            messagebox.showerror("Error", "Please enter JSON data")
            return
        
        try:
            # Parse JSON from text
            self.json_data = json.loads(json_text)
            
            # Clear file path since we're using direct input
            self.json_file_path.set("")
            
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
                self.status_label['text'] = f"✓ Found {len(self.array_keys)} array(s) in JSON data"
                self.status_label['foreground'] = "green"
                
                # Close the input window
                self.json_input_window.destroy()
                
            else:
                messagebox.showerror("Error", "No arrays found in the JSON data")
                
        except json.JSONDecodeError as e:
            messagebox.showerror("Error", f"Invalid JSON format: {str(e)}")
        except Exception as e:
            messagebox.showerror("Error", f"Error parsing JSON: {str(e)}")
    
    def clear_json_text(self):
        """Clear the JSON text area"""
        self.json_text.delete("1.0", tk.END)
    
    def on_format_selected(self, event=None):
        """Handle output format selection"""
        format_text = self.output_format.get()
        
        if "Excel" in format_text:
            self.extension_label['text'] = ".xlsx"
        elif "CSV" in format_text:
            self.extension_label['text'] = ".csv"
        elif "JSON" in format_text:
            self.extension_label['text'] = ".json"
        
        # Update button text
        format_name = format_text.split(" ")[0]
        self.process_button['text'] = f"Convert to {format_name}"
    
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
                self.status_label['text'] = f"✓ Found {len(self.array_keys)} array(s) in JSON file"
                self.status_label['foreground'] = "green"
            else:
                self.array_combobox['values'] = []
                self.array_info_label['text'] = "No arrays found in JSON file"
                self.process_button['state'] = 'disabled'
                self.status_label['text'] = "⚠ No arrays found in JSON file"
                self.status_label['foreground'] = "orange"
                
        except json.JSONDecodeError as e:
            messagebox.showerror("Error", f"Invalid JSON file: {str(e)}")
            self.status_label['text'] = "❌ Error: Invalid JSON file"
            self.status_label['foreground'] = "red"
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")
            self.status_label['text'] = "❌ Error loading file"
            self.status_label['foreground'] = "red"
    
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
    
    def convert_to_file(self):
        """Convert the selected array to the chosen format"""
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
            
            # Get format and extension
            format_text = self.output_format.get()
            if "Excel" in format_text:
                extension = ".xlsx"
                format_type = "Excel"
            elif "CSV" in format_text:
                extension = ".csv"
                format_type = "CSV"
            elif "JSON" in format_text:
                extension = ".json"
                format_type = "JSON"
            else:
                extension = ".xlsx"
                format_type = "Excel"
            
            # Build the full file path
            filename = self.output_filename.get().strip()
            if not filename.endswith(extension):
                filename += extension
            
            file_path = os.path.join(self.output_folder.get(), filename)
            
            # Check if file already exists
            if os.path.exists(file_path):
                result = messagebox.askyesno("File Exists", 
                                          f"File '{filename}' already exists in the selected folder.\n"
                                          "Do you want to overwrite it?")
                if not result:
                    return
            
            # Save based on format
            if format_type == "Excel":
                df.to_excel(file_path, index=False)
            elif format_type == "CSV":
                df.to_csv(file_path, index=False)
            elif format_type == "JSON":
                df.to_json(file_path, orient='records', indent=2)
            
            messagebox.showinfo("Success", 
                             f"{format_type} file saved successfully!\n"
                             f"File: {file_path}\n"
                             f"Rows: {len(df)}\n"
                             f"Columns: {len(df.columns)}")
            
            self.status_label['text'] = f"✓ {format_type} file saved: {filename}"
            self.status_label['foreground'] = "green"
            
            # Open the file automatically
            if format_type == "Excel":
                self.open_excel_file(file_path)
            elif format_type == "CSV":
                self.open_csv_file(file_path)
            elif format_type == "JSON":
                self.open_json_file(file_path)
            else:
                # For other formats, just show the file location
                self.open_file_location(file_path)
             
        except Exception as e:
            messagebox.showerror("Error", f"Error creating {format_type} file: {str(e)}")
            self.status_label['text'] = f"❌ Error creating {format_type} file"
            self.status_label['foreground'] = "red"
    
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
                
            self.status_label['text'] = f"✓ Excel file opened: {os.path.basename(file_path)}"
            self.status_label['foreground'] = "green"
            
        except Exception as e:
            # If opening fails, just show a message but don't crash
            messagebox.showwarning("Warning", 
                                f"Excel file saved successfully, but could not open automatically.\n"
                                f"File location: {file_path}\n"
                                f"Error: {str(e)}")
            self.status_label['text'] = f"⚠ Excel file saved (could not open): {os.path.basename(file_path)}"
            self.status_label['foreground'] = "orange"
    
    def open_csv_file(self, file_path):
        """Open the CSV file with the default application"""
        try:
            system = platform.system()
            
            if system == "Windows":
                os.startfile(file_path)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", file_path], check=True)
            else:  # Linux
                subprocess.run(["xdg-open", file_path], check=True)
                
            self.status_label['text'] = f"✓ CSV file opened: {os.path.basename(file_path)}"
            self.status_label['foreground'] = "green"
            
        except Exception as e:
            # If opening fails, just show a message but don't crash
            messagebox.showwarning("Warning", 
                                f"CSV file saved successfully, but could not open automatically.\n"
                                f"File location: {file_path}\n"
                                f"Error: {str(e)}")
            self.status_label['text'] = f"⚠ CSV file saved (could not open): {os.path.basename(file_path)}"
            self.status_label['foreground'] = "orange"
    
    def open_json_file(self, file_path):
        """Open the JSON file with the default application"""
        try:
            system = platform.system()
            
            if system == "Windows":
                os.startfile(file_path)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", file_path], check=True)
            else:  # Linux
                subprocess.run(["xdg-open", file_path], check=True)
                
            self.status_label['text'] = f"✓ JSON file opened: {os.path.basename(file_path)}"
            self.status_label['foreground'] = "green"
            
        except Exception as e:
            # If opening fails, just show a message but don't crash
            messagebox.showwarning("Warning", 
                                f"JSON file saved successfully, but could not open automatically.\n"
                                f"File location: {file_path}\n"
                                f"Error: {str(e)}")
            self.status_label['text'] = f"⚠ JSON file saved (could not open): {os.path.basename(file_path)}"
            self.status_label['foreground'] = "orange"
    
    def open_file_location(self, file_path):
        """Open the file location in file explorer"""
        try:
            system = platform.system()
            
            if system == "Windows":
                # Open folder and select the file
                subprocess.run(["explorer", "/select,", file_path], check=True)
            elif system == "Darwin":  # macOS
                # Show in Finder
                subprocess.run(["open", "-R", file_path], check=True)
            else:  # Linux
                # Open folder containing the file
                folder_path = os.path.dirname(file_path)
                subprocess.run(["xdg-open", folder_path], check=True)
                
            self.status_label['text'] = f"✓ File location opened: {os.path.basename(file_path)}"
            self.status_label['foreground'] = "green"
            
        except Exception as e:
            # If opening fails, just show a message but don't crash
            messagebox.showinfo("File Saved", 
                             f"File saved successfully!\n"
                             f"Location: {file_path}")
            self.status_label['text'] = f"✓ File saved: {os.path.basename(file_path)}"
            self.status_label['foreground'] = "green"

def main():
    root = tk.Tk()
    app = JsonToExcelConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()
