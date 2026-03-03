import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from PIL import Image, ImageTk
from extractData import InventoryManager


class InstructionsWindow:
    """Window to display instructions and help information with images."""
    
    def __init__(self, parent):
        """Initialize the instructions window.
        
        Args:
            parent: The parent window
        """
        self.window = tk.Toplevel(parent)
        self.window.title("Shipping Inventory Program - Instructions")
        self.window.geometry("900x800")
        self.window.resizable(True, True)
        
        # Create a main frame with scrollbar
        main_frame = ttk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create canvas with scrollbar
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Title
        title_label = ttk.Label(scrollable_frame, text="Inventory Update Procedure", 
                               font=("Helvetica", 16, "bold"))
        title_label.pack(pady=10)
        
        # Section 1
        section1_label = ttk.Label(scrollable_frame, text="1. Refresh and Download Shipping Saint Box Inventory", 
                                   font=("Helvetica", 12, "bold"))
        section1_label.pack(pady=(10, 5), anchor="w", padx=10)
        
        step1a = ttk.Label(scrollable_frame, text="a. Go to box Inventory:", font=("Helvetica", 10))
        step1a.pack(anchor="w", padx=20)
        self._load_and_display_image(scrollable_frame, "picture1.png", 700)
        
        step1b = ttk.Label(scrollable_frame, text="b. Hit the printer icon (green) to update the export, then hit the link icon to download the excel file from ShippingSaint:", 
                          font=("Helvetica", 10), wraplength=850)
        step1b.pack(anchor="w", padx=20, pady=(10, 5))
        self._load_and_display_image(scrollable_frame, "picture2.png", 700)
        
        # Section 2
        section2_label = ttk.Label(scrollable_frame, text="2. Open Inventory Analysis Program", 
                                   font=("Helvetica", 12, "bold"))
        section2_label.pack(pady=(20, 5), anchor="w", padx=10)
        
        step2a = ttk.Label(scrollable_frame, text="a. Go to the 'Shipping Box Inventory' folder in the 'Company Files (f:)' shared drive:", 
                          font=("Helvetica", 10), wraplength=850)
        step2a.pack(anchor="w", padx=20)
        self._load_and_display_image(scrollable_frame, "picture3.png", 700)
        
        step2b = ttk.Label(scrollable_frame, text="b. Launch the BDBox_Inventory_Manager program with a double-click:", 
                          font=("Helvetica", 10), wraplength=850)
        step2b.pack(anchor="w", padx=20, pady=(10, 5))
        self._load_and_display_image(scrollable_frame, "picture4.png", 700)
        
        # Section 3
        step3_label = ttk.Label(scrollable_frame, text="3. Click the 'Browse Excel File' button in the program to select the excel file which was downloaded from ShippingSaint. Typically named 'build' and located in your computers 'Downloads' folder:", 
                               font=("Helvetica", 10), wraplength=850)
        step3_label.pack(anchor="w", padx=20, pady=(20, 5))
        self._load_and_display_image(scrollable_frame, "picture5.png", 700)
        
        # Section 4
        step4_label = ttk.Label(scrollable_frame, text="4. Enter sale number (reference b) or select the restock check box (reference a), then click 'Submit' to process the inventory. NOTE: Do NOT have the 'Inventory_Analysis' excel file open when you submit the ShippingSaint Inventory.", 
                               font=("Helvetica", 10), wraplength=850)
        step4_label.pack(anchor="w", padx=20, pady=(20, 5))
        
        step4a = ttk.Label(scrollable_frame, text="   a. Check the restock box when you are submitting inventory to be analyzed directly after updating the inventory in ShippingSaint (regardless of reason such as a Uline order or inventory audit).", 
                          font=("Helvetica", 9), wraplength=850)
        step4a.pack(anchor="w", padx=40, pady=2)
        
        step4b = ttk.Label(scrollable_frame, text="   b. Check the sale box and enter a sale number when submitting the inventory after the conclusion of shipping for a given sale.", 
                          font=("Helvetica", 9), wraplength=850)
        step4b.pack(anchor="w", padx=40, pady=2)
        
        # Section 5
        step5_label = ttk.Label(scrollable_frame, text="5. Open the 'Inventory_Analysis' excel file to review the updated inventory and projections", 
                               font=("Helvetica", 10), wraplength=850)
        step5_label.pack(anchor="w", padx=20, pady=(20, 5))
        
        step5a = ttk.Label(scrollable_frame, text="   a. The Current Inventory & Predictions sheet shows the Current Stock, a quarterly prediction for how much stock is needed for the next 6 sales, and an automatic status", 
                          font=("Helvetica", 9), wraplength=850)
        step5a.pack(anchor="w", padx=40, pady=2)
        
        step5ai = ttk.Label(scrollable_frame, text="      i. The automatically displayed status will show adequate stock for items which have more than the quarterly prediction. Any item which has less current stock than the quarterly prediction, the quantity of items you are short is displayed. Ie., current stock is 20, the prediction is 35, the status will show 15", 
                           font=("Helvetica", 9), wraplength=850)
        step5ai.pack(anchor="w", padx=60, pady=2)
        
        step5b = ttk.Label(scrollable_frame, text="   b. There is also an inventory history sheet available to log the history of the inventory, and the remaining 2 sheets are used to hold sale difference and average use stats.", 
                          font=("Helvetica", 9), wraplength=850)
        step5b.pack(anchor="w", padx=40, pady=2)
        
        # Helpful Tips
        tips_label = ttk.Label(scrollable_frame, text="Helpful Tips:", 
                              font=("Helvetica", 12, "bold"))
        tips_label.pack(pady=(20, 5), anchor="w", padx=10)
        
        tip1 = ttk.Label(scrollable_frame, text="• When making a Uline order, the number displayed in the status column of the Current Inventory & Predictions sheet is the number of items you want to target in your order.", 
                        font=("Helvetica", 9), wraplength=850)
        tip1.pack(anchor="w", padx=20, pady=5)
        
        # Close button
        close_btn = ttk.Button(scrollable_frame, text="Close", command=self.window.destroy)
        close_btn.pack(pady=15)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def _load_and_display_image(self, parent_frame, image_name, max_width=700):
        """Load and display an image from the images folder.
        
        Args:
            parent_frame: The frame to add the image to
            image_name: The name of the image file (e.g., 'picture1.png')
            max_width: Maximum width for the image
        """
        try:
            # Get the images folder path relative to this script
            script_dir = os.path.dirname(os.path.abspath(__file__))
            image_path = os.path.join(script_dir, "images", image_name)
            
            if os.path.exists(image_path):
                img = Image.open(image_path)
                
                # Resize image if needed
                if img.width > max_width:
                    ratio = max_width / img.width
                    new_height = int(img.height * ratio)
                    img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
                
                # Convert to PhotoImage
                photo = ImageTk.PhotoImage(img)
                
                # Create label with image
                img_label = ttk.Label(parent_frame, image=photo)
                img_label.image = photo  # Keep a reference to prevent garbage collection
                img_label.pack(pady=(5, 10))
            else:
                # Show placeholder if image not found
                error_label = ttk.Label(parent_frame, text=f"[Image not found: {image_name}]", 
                                       foreground="red", font=("Helvetica", 9))
                error_label.pack(pady=5)
        except Exception as e:
            error_label = ttk.Label(parent_frame, text=f"[Error loading {image_name}: {str(e)}]", 
                                   foreground="red", font=("Helvetica", 9))
            error_label.pack(pady=5)


class ExcelLoaderGUI:
    """GUI for loading and submitting Excel files."""
    
    def __init__(self, root):
        """Initialize the Excel Loader GUI.
        
        Args:
            root: The tkinter root window
        """
        self.root = root
        self.root.title("Excel File Loader")
        self.root.geometry("380x300")
        self.selected_file_path = None
        self.manager = InventoryManager()  # Keep a reference to the manager for recompute
        
        self._create_widgets()
        
        # Show instructions window on startup
        self.root.after(100, self._show_instructions)
    
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
        
        ttk.Button(button_frame, text="Submit new sale or restock", command=self.submit).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Recompute analysis based on existing inventory data", command=self.recompute).pack(side=tk.LEFT, padx=5)
        
        # Add help button
        ttk.Button(self.root, text="Help", command=self._show_instructions).pack(pady=5)
    
    def _show_instructions(self):
        """Display the instructions window."""
        InstructionsWindow(self.root)
    
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