import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from extractData import InventoryManager


class ExcelLoaderGUI:
    """GUI for loading and submitting Excel files."""
    
    def __init__(self, root):
        """Initialize the Excel Loader GUI.
        
        Args:
            root: The tkinter root window
        """
        self.root = root
        self.root.title("Excel File Loader")
        self.root.geometry("360x300")
        self.selected_file_path = None
        self.manager = InventoryManager()  # Keep a reference to the manager for recompute
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Create and layout the GUI widgets."""
        ttk.Label(self.root, text="Entry Type:").pack(pady=5)
        
        # Create a frame for radio buttons
        type_frame = ttk.Frame(self.root)
        type_frame.pack(pady=5)
        
        self.entry_type = tk.StringVar(value="sale")
        ttk.Radiobutton(type_frame, text="Sale", variable=self.entry_type, value="sale").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="Restock", variable=self.entry_type, value="restock").pack(side=tk.LEFT, padx=5)
        
        ttk.Label(self.root, text="Sale Number:").pack(pady=5)
        self.sale_entry = ttk.Entry(self.root, width=30)
        self.sale_entry.pack(pady=5)
        
        ttk.Button(self.root, text="Browse Excel File", command=self.browse_file).pack(pady=10)
        self.file_label = ttk.Label(self.root, text="No file selected")
        self.file_label.pack(pady=5)
        
        # Create a frame for buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Submit", command=self.submit).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Recompute", command=self.recompute).pack(side=tk.LEFT, padx=5)
    
    def browse_file(self):
        """Open file dialog and update the file label."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.selected_file_path = file_path
            self.file_label.config(text=f"Selected: {os.path.basename(file_path)}")
    
    def submit(self):
        """Validate and submit the form."""
        entry_type = self.entry_type.get()
        sale_number = self.sale_entry.get()
        file_path = self.selected_file_path
        
        if entry_type == "sale" and not sale_number:
            messagebox.showwarning("Input Error", "Please enter a sale number")
            return
        if not file_path:
            messagebox.showwarning("Input Error", "Please select an Excel file")
            return
        
        try:
            # Process the inventory data
            
            if entry_type == "sale":
                output_file = self.manager.process_inventory(
                    input_file=file_path,
                    sale_number=sale_number,
                    label_column='Label',
                    stock_column='Stock'
                )
            else:  # restock
                output_file = self.manager.process_restock(
                    input_file=file_path,
                    label_column='Label',
                    stock_column='Stock'
                )
            
            messagebox.showinfo("Success", f"Inventory processed successfully!\n\nOutput file: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process inventory:\n{str(e)}")
    
    def recompute(self):
        """Recompute all analysis sheets without importing new data."""
        try:
            # Recompute all analysis sheets
            self.manager.update_sales_differences()
            self.manager.update_average_use()
            self.manager.update_predictions()
            
            messagebox.showinfo("Success", "Analysis sheets recomputed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to recompute analysis:\n{str(e)}")