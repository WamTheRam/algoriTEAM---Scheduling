import pandas as pd
import tkinter as tk
from tkinter import Menu, ttk
import os

class ExcelViewerApp:
    def __init__(self, root, file_name="TestFile.xlsx"):
        self.root = root
        self.root.title("Excel Viewer")

        # Menu
        self.menu = Menu(self.root)
        self.root.config(menu=self.menu)
        
        # File menu
        self.file_menu = Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="File", menu=self.file_menu)
        self.file_menu.add_command(label="Open", command=self.load_file)

        # Schedule menu
        self.schedule_menu = Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Schedule", menu=self.schedule_menu)
        self.schedule_menu.add_command(label="Add Schedule", command=self.add_schedule)
        self.schedule_menu.add_command(label="Edit Schedule", command=self.edit_schedule)
        self.schedule_menu.add_command(label="Suggest Merge", command=self.suggest_merge)

        # Frame for displaying the table
        self.frame = tk.Frame(self.root)
        self.frame.pack(fill=tk.BOTH, expand=1)

        self.tree = None

        # Scrollbars
        self.tree_scroll_y = tk.Scrollbar(self.frame, orient=tk.VERTICAL)
        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_scroll_x = tk.Scrollbar(self.frame, orient=tk.HORIZONTAL)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Load the file automatically on start
        self.file_path = os.path.join(os.path.dirname(__file__), file_name)
        self.load_file()

    def load_file(self):
        try:
            df = pd.read_excel(self.file_path, header=None)
            self.show_table(df)
        except Exception as e:
            print(f"Error loading file: {e}")

    def show_table(self, df):
        if self.tree is not None:
            self.tree.destroy()

        self.tree = ttk.Treeview(self.frame, yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)
        self.tree.pack(fill=tk.BOTH, expand=1)

        # Define columns based on row 4 (index 3 in zero-indexed DataFrame)
        columns = df.iloc[3].dropna().tolist()
        self.tree["columns"] = columns
        self.tree["show"] = "headings"

        # Define headings
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        # Process and insert data
        college_name = None
        for index, row in df.iterrows():
            if index <= 3:  # Skip the first 4 rows
                continue

            # Identify college header
            non_na_values = row.dropna().tolist()
            if len(non_na_values) < 3 and pd.notna(row[0]):
                college_name = row[0]
            else:
                row_data = row.dropna().tolist()
                if college_name:
                    row_data.insert(0, college_name)  # Insert college name as the first column
                self.tree.insert("", "end", values=row_data)

    def add_schedule(self):
        pass  # Placeholder for adding schedule functionality

    def edit_schedule(self):
        pass  # Placeholder for editing schedule functionality

    def suggest_merge(self):
        pass  # Placeholder for suggesting merge functionality


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.geometry("800x600")
    root.mainloop()
