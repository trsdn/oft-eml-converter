#!/usr/bin/env python3
"""
OFT to EML Converter - GUI Version (Clean, Single Window)

A simple GUI application for converting Outlook Template (.oft) files 
to standard EML format. Supports batch processing and remembers output directory.
This version uses only standard tkinter to avoid window conflicts.

Requirements:
- Python 3.7+
- tkinter (usually comes with Python)
- extract-msg (for OFT file parsing)

Usage:
    python oft_to_eml_gui_fixed.py
"""

import sys
import os
import json
import threading
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError as e:
    print(f"Error: tkinter not available: {e}")
    print("Please install tkinter (usually comes with Python)")
    sys.exit(1)

from oft_to_eml_converter import convert_oft_to_eml


class OFTtoEMLGUI:
    def __init__(self):
        # Create a single, clean root window
        self.root = tk.Tk()
        self.root.title("OFT to EML Converter")
        self.root.geometry("600x500")
        self.root.minsize(500, 400)
        
        # Configuration file for settings
        self.config_file = "converter_config.json"
        self.load_config()
        
        # Variables
        self.output_dir = tk.StringVar(value=self.config.get('last_output_dir', os.getcwd()))
        self.is_converting = False
        self.files_to_convert = []
        self.conversion_results = []
        
        self.setup_ui()
        
    def load_config(self):
        """Load configuration from file."""
        try:
            with open(self.config_file, 'r') as f:
                self.config = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self.config = {}
    
    def save_config(self):
        """Save configuration to file."""
        self.config['last_output_dir'] = self.output_dir.get()
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=2)
        except Exception:
            pass  # Silently ignore config save errors
    
    def setup_ui(self):
        """Set up the user interface."""
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="OFT to EML Converter", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # File selection frame
        drop_frame = ttk.LabelFrame(main_frame, text="Select Files", padding="10")
        drop_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        drop_frame.columnconfigure(0, weight=1)
        drop_frame.rowconfigure(0, weight=1)
        
        # File selection area
        self.drop_label = tk.Label(drop_frame, 
                                  text="📁 Click to browse and select OFT files\n\n🗂️ Multiple files supported",
                                  font=('Arial', 14),
                                  anchor=tk.CENTER,
                                  relief=tk.SOLID,
                                  bd=2,
                                  background='#f0f8ff',
                                  fg='#333333',
                                  cursor='hand2',
                                  height=6,
                                  wraplength=400)
        self.drop_label.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5, ipady=20)
        
        # Make clickable
        self.drop_label.bind('<Button-1>', lambda e: self.browse_files())
        self.drop_label.bind('<Enter>', lambda e: self.drop_label.config(background='#e6f3ff', relief=tk.RAISED))
        self.drop_label.bind('<Leave>', lambda e: self.drop_label.config(background='#f0f8ff', relief=tk.SOLID))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, pady=10)
        
        # Browse button
        self.browse_btn = ttk.Button(button_frame, text="Browse Files", 
                                    command=self.browse_files)
        self.browse_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Clear files button
        self.clear_files_btn = ttk.Button(button_frame, text="Clear Files", 
                                         command=self.clear_files)
        self.clear_files_btn.pack(side=tk.LEFT)
        
        # Output directory frame
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=10)
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="Output Directory:").grid(row=0, column=0, padx=(0, 10))
        
        output_entry = ttk.Entry(output_frame, textvariable=self.output_dir, state='readonly')
        output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(output_frame, text="Browse", 
                  command=self.browse_output_dir).grid(row=0, column=2)
        
        # Progress frame
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=10)
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_label = ttk.Label(progress_frame, text="Ready to convert files")
        self.progress_label.grid(row=0, column=0, sticky=tk.W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="Conversion Results", padding="10")
        results_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Results text area with scrollbar
        text_frame = ttk.Frame(results_frame)
        text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.results_text = tk.Text(text_frame, height=8, wrap=tk.WORD, state=tk.DISABLED)
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.results_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        # Action buttons frame
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=6, column=0, pady=10)
        
        self.convert_btn = ttk.Button(action_frame, text="Convert Files", 
                                     command=self.start_conversion, state=tk.DISABLED)
        self.convert_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.open_folder_btn = ttk.Button(action_frame, text="Open Output Folder", 
                                         command=self.open_output_folder, state=tk.DISABLED)
        self.open_folder_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.clear_btn = ttk.Button(action_frame, text="Clear All", 
                                   command=self.clear_all)
        self.clear_btn.pack(side=tk.LEFT)
    
    def browse_files(self):
        """Browse for OFT files."""
        files = filedialog.askopenfilenames(
            title="Select OFT Files",
            filetypes=[("Outlook Template files", "*.oft"), ("All files", "*.*")]
        )
        
        if files:
            self.files_to_convert = list(files)
            self.update_drop_label()
            self.convert_btn.config(state=tk.NORMAL)
    
    def clear_files(self):
        """Clear selected files."""
        self.files_to_convert = []
        self.update_drop_label()
        self.convert_btn.config(state=tk.DISABLED)
    
    def browse_output_dir(self):
        """Browse for output directory."""
        directory = filedialog.askdirectory(
            title="Select Output Directory",
            initialdir=self.output_dir.get()
        )
        
        if directory:
            self.output_dir.set(directory)
            self.save_config()
    
    def update_drop_label(self):
        """Update the file selection label with file count."""
        if self.files_to_convert:
            count = len(self.files_to_convert)
            file_names = [os.path.basename(f) for f in self.files_to_convert[:3]]
            text = f"✅ {count} OFT file{'s' if count != 1 else ''} ready to convert:\n\n"
            text += "\n".join(f"• {name}" for name in file_names)
            if count > 3:
                text += f"\n• ... and {count - 3} more files"
            self.drop_label.config(text=text, background='#e6ffe6', fg='#2d5a2d')
        else:
            self.drop_label.config(text="📁 Click to browse and select OFT files\n\n🗂️ Multiple files supported", 
                                  background='#f0f8ff', fg='#333333')
    
    def update_results(self, message, success=True):
        """Update the results text area."""
        self.results_text.config(state=tk.NORMAL)
        icon = "✓" if success else "✗"
        self.results_text.insert(tk.END, f"{icon} {message}\n")
        self.results_text.see(tk.END)
        self.results_text.config(state=tk.DISABLED)
        self.root.update()
    
    def start_conversion(self):
        """Start the conversion process in a separate thread."""
        if self.is_converting or not self.files_to_convert:
            return
        
        # Clear previous results
        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete(1.0, tk.END)
        self.results_text.config(state=tk.DISABLED)
        
        # Start conversion in thread to avoid blocking UI
        thread = threading.Thread(target=self.convert_files)
        thread.daemon = True
        thread.start()
    
    def convert_files(self):
        """Convert all selected files."""
        self.is_converting = True
        self.convert_btn.config(state=tk.DISABLED)
        self.browse_btn.config(state=tk.DISABLED)
        
        total_files = len(self.files_to_convert)
        self.progress_bar.config(maximum=total_files)
        successful_conversions = 0
        
        for i, oft_file in enumerate(self.files_to_convert):
            try:
                # Update progress
                self.progress_label.config(text=f"Converting file {i+1} of {total_files}...")
                self.progress_bar.config(value=i)
                self.root.update()
                
                # Generate output path
                base_name = Path(oft_file).stem
                output_path = os.path.join(self.output_dir.get(), f"{base_name}.eml")
                
                # Convert file
                convert_oft_to_eml(oft_file, output_path)
                
                # Update results
                file_name = os.path.basename(oft_file)
                self.update_results(f"{file_name} → {os.path.basename(output_path)}")
                successful_conversions += 1
                
            except Exception as e:
                file_name = os.path.basename(oft_file)
                self.update_results(f"{file_name} - Error: {str(e)}", success=False)
        
        # Final progress update
        self.progress_bar.config(value=total_files)
        self.progress_label.config(text=f"Conversion complete: {successful_conversions}/{total_files} files converted")
        
        if successful_conversions > 0:
            self.open_folder_btn.config(state=tk.NORMAL)
        
        # Re-enable buttons
        self.is_converting = False
        self.convert_btn.config(state=tk.NORMAL)
        self.browse_btn.config(state=tk.NORMAL)
        
        # Show completion message
        if successful_conversions == total_files:
            messagebox.showinfo("Conversion Complete", 
                              f"All {total_files} files converted successfully!")
        elif successful_conversions > 0:
            messagebox.showwarning("Partial Success", 
                                 f"{successful_conversions} of {total_files} files converted successfully.")
        else:
            messagebox.showerror("Conversion Failed", "No files were converted successfully.")
    
    def open_output_folder(self):
        """Open the output folder in file manager."""
        output_path = self.output_dir.get()
        if os.path.exists(output_path):
            if os.name == 'nt':  # Windows
                os.startfile(output_path)
            elif os.name == 'posix':  # macOS and Linux
                os.system(f'open "{output_path}"' if sys.platform == 'darwin' else f'xdg-open "{output_path}"')
    
    def clear_all(self):
        """Clear all files and results."""
        self.files_to_convert = []
        self.update_drop_label()
        self.convert_btn.config(state=tk.DISABLED)
        self.open_folder_btn.config(state=tk.DISABLED)
        
        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete(1.0, tk.END)
        self.results_text.config(state=tk.DISABLED)
        
        self.progress_bar.config(value=0)
        self.progress_label.config(text="Ready to convert files")
    
    def run(self):
        """Start the GUI application."""
        # Handle window closing
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            pass
    
    def on_closing(self):
        """Handle application closing."""
        if self.is_converting:
            if messagebox.askokcancel("Quit", "Conversion in progress. Do you want to quit?"):
                self.save_config()
                self.root.destroy()
        else:
            self.save_config()
            self.root.destroy()


def main():
    """Main entry point for the GUI application."""
    try:
        app = OFTtoEMLGUI()
        app.run()
    except Exception as e:
        try:
            messagebox.showerror("Error", f"Failed to start application: {str(e)}")
        except:
            print(f"Failed to start application: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()