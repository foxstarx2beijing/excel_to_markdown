#!/usr/bin/env python3
"""
Excel/CSV to Markdown Converter - GUI Version

This script provides a graphical user interface for converting Excel (.xlsx, .xls) 
or CSV files to Markdown tables.
"""

import sys
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from excel_to_md import read_file, dataframe_to_markdown
import pandas as pd
from pathlib import Path


class ExcelToMarkdownApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/CSV to Markdown Converter")
        self.root.geometry("800x600")
        self.root.minsize(600, 500)
        
        # Set style
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6)
        self.style.configure("TLabel", font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 12, "bold"))
        
        # File paths
        self.input_file = None
        self.output_file = None
        self.df = None
        self.sheets = []
        self.current_sheet = None
        
        # Create the UI
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20 10 20 10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel/CSV to Markdown Converter", style="Header.TLabel")
        title_label.pack(pady=(0, 10))
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10 10 10 10")
        file_frame.pack(fill=tk.X, expand=False, pady=(0, 10))
        
        # Input file
        input_frame = ttk.Frame(file_frame)
        input_frame.pack(fill=tk.X, expand=True, pady=(0, 5))
        
        ttk.Label(input_frame, text="Input File:").pack(side=tk.LEFT, padx=(0, 5))
        self.input_entry = ttk.Entry(input_frame)
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        input_button = ttk.Button(input_frame, text="Browse", command=self.select_input_file)
        input_button.pack(side=tk.LEFT)
        
        # Output file
        output_frame = ttk.Frame(file_frame)
        output_frame.pack(fill=tk.X, expand=True, pady=(0, 5))
        
        ttk.Label(output_frame, text="Output File:").pack(side=tk.LEFT, padx=(0, 5))
        self.output_entry = ttk.Entry(output_frame)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        output_button = ttk.Button(output_frame, text="Browse", command=self.select_output_file)
        output_button.pack(side=tk.LEFT)
        
        # Sheet selection
        sheet_frame = ttk.Frame(file_frame)
        sheet_frame.pack(fill=tk.X, expand=True, pady=(0, 5))
        
        ttk.Label(sheet_frame, text="Sheet:").pack(side=tk.LEFT, padx=(0, 5))
        self.sheet_combobox = ttk.Combobox(sheet_frame, state="readonly")
        self.sheet_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.sheet_combobox.bind("<<ComboboxSelected>>", self.on_sheet_selected)
        
        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10 10 10 10")
        options_frame.pack(fill=tk.X, expand=False, pady=(0, 10))
        
        # Include headers option
        self.include_headers_var = tk.BooleanVar(value=True)
        headers_check = ttk.Checkbutton(options_frame, text="Include Headers", variable=self.include_headers_var, command=self.update_preview)
        headers_check.pack(anchor=tk.W)
        
        # Preview frame
        preview_frame = ttk.LabelFrame(main_frame, text="Preview", padding="10 10 10 10")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Preview text area
        self.preview_text = scrolledtext.ScrolledText(preview_frame, wrap=tk.WORD, font=("Courier New", 10))
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, expand=False, pady=(0, 10))
        
        # Copy to clipboard button
        clipboard_button = ttk.Button(buttons_frame, text="Copy to Clipboard", command=self.copy_to_clipboard)
        clipboard_button.pack(side=tk.LEFT, padx=(0, 5))
        
        # Convert button
        convert_button = ttk.Button(buttons_frame, text="Convert & Save", command=self.convert_and_save)
        convert_button.pack(side=tk.RIGHT)
    
    def select_input_file(self):
        filetypes = [
            ('Excel/CSV Files', '*.xlsx *.xls *.csv'),
            ('Excel Files', '*.xlsx *.xls'),
            ('CSV Files', '*.csv'),
            ('All Files', '*.*')
        ]
        
        filename = filedialog.askopenfilename(
            title="Select Input File",
            filetypes=filetypes
        )
        
        if filename:
            self.input_file = filename
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, filename)
            
            # Suggest output filename
            output_path = str(Path(filename).with_suffix('.md'))
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, output_path)
            
            self.load_file()
    
    def select_output_file(self):
        filetypes = [
            ('Markdown Files', '*.md'),
            ('Text Files', '*.txt'),
            ('All Files', '*.*')
        ]
        
        filename = filedialog.asksaveasfilename(
            title="Select Output File",
            filetypes=filetypes,
            defaultextension=".md"
        )
        
        if filename:
            self.output_file = filename
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, filename)
    
    def load_file(self):
        try:
            file_ext = os.path.splitext(self.input_file)[1].lower()
            
            # For Excel files, get sheet names
            if file_ext in ['.xlsx', '.xls']:
                excel_file = pd.ExcelFile(self.input_file)
                self.sheets = excel_file.sheet_names
                
                # Update the sheet combobox
                self.sheet_combobox['values'] = ['All Sheets'] + self.sheets
                self.sheet_combobox.current(0)  # Set to "All Sheets" by default
                self.current_sheet = None
                
                # Load the first sheet for preview
                self.df = pd.read_excel(self.input_file, sheet_name=self.sheets[0])
            else:
                # For CSV files
                self.df = read_file(self.input_file)
                self.sheets = []
                self.sheet_combobox['values'] = ['']
                self.sheet_combobox.current(0)
                self.current_sheet = None
            
            self.update_preview()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")
    
    def on_sheet_selected(self, event):
        selected = self.sheet_combobox.get()
        
        if selected == 'All Sheets':
            self.current_sheet = None
            if self.sheets:
                self.df = pd.read_excel(self.input_file, sheet_name=self.sheets[0])
        else:
            self.current_sheet = selected
            self.df = pd.read_excel(self.input_file, sheet_name=selected)
        
        self.update_preview()
    
    def update_preview(self):
        if self.df is not None:
            include_headers = self.include_headers_var.get()
            markdown = dataframe_to_markdown(self.df, headers=include_headers)
            
            # Update preview
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(tk.END, markdown)
    
    def copy_to_clipboard(self):
        if self.df is not None:
            # Get the markdown from the preview
            markdown = self.preview_text.get(1.0, tk.END)
            
            # Copy to clipboard
            self.root.clipboard_clear()
            self.root.clipboard_append(markdown)
            
            messagebox.showinfo("Success", "Markdown copied to clipboard!")
    
    def convert_and_save(self):
        if not self.input_file:
            messagebox.showerror("Error", "Please select an input file first.")
            return
        
        output_file = self.output_entry.get()
        if not output_file:
            messagebox.showerror("Error", "Please specify an output file.")
            return
        
        try:
            file_ext = os.path.splitext(self.input_file)[1].lower()
            include_headers = self.include_headers_var.get()
            all_markdown = []
            
            # Process all sheets or just the selected one
            if file_ext in ['.xlsx', '.xls'] and self.current_sheet is None:
                # Convert all sheets
                for sheet in self.sheets:
                    df = pd.read_excel(self.input_file, sheet_name=sheet)
                    markdown = f"## Sheet: {sheet}\n\n" + dataframe_to_markdown(df, headers=include_headers)
                    all_markdown.append(markdown)
            else:
                # Convert single sheet or CSV
                markdown = dataframe_to_markdown(self.df, headers=include_headers)
                all_markdown.append(markdown)
            
            # Write to output file
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write("\n\n".join(all_markdown))
            
            messagebox.showinfo("Success", f"Successfully converted to Markdown and saved to {output_file}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")


def main():
    root = tk.Tk()
    app = ExcelToMarkdownApp(root)
    root.mainloop()


if __name__ == "__main__":
    main() 