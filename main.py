import tkinter as tk
from loadExcel import ExcelLoaderGUI


def main():
    """Main function to initialize and run the tkinter application."""
    root = tk.Tk()
    
    # Initialize the Excel Loader GUI
    app = ExcelLoaderGUI(root)
    
    root.mainloop()


if __name__ == "__main__":
    main()