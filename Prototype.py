import pandas as pd
import tkinter as tk
from tkinter import Menu, ttk, simpledialog, messagebox
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
        self.file_menu.add_command(label="Save", command=self.save_file)

        # Schedule menu
        self.schedule_menu = Menu(self.menu, tearoff=0)
        self.menu.add_cascade(label="Schedule", menu=self.schedule_menu)
        self.schedule_menu.add_command(label="Add Schedule", command=self.add_schedule)
        self.schedule_menu.add_command(label="Edit Schedule", command=self.edit_schedule)
        self.schedule_menu.add_command(label="Suggest Merge", command=self.suggest_merge)
        self.schedule_menu.add_command(label="Find Conflict", command=self.find_conflict)

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
            self.df = pd.read_excel(self.file_path, header=None)
            self.df.fillna('', inplace=True)  # Replace NaN with empty string
            self.show_table(self.df)
        except Exception as e:
            print(f"Error loading file: {e}")

    def save_file(self):
        try:
            self.df.to_excel(self.file_path, index=False, header=False)
            messagebox.showinfo("Save", "File saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving file: {e}")

    def show_table(self, df):
        if self.tree is not None:
            self.tree.destroy()

        self.tree = ttk.Treeview(self.frame, yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)
        self.tree.pack(fill=tk.BOTH, expand=1)

        # Define columns based on row 4 (index 3 in zero-indexed DataFrame), shifting right by one
        columns = [''] + df.iloc[3].tolist()  # Prepend an empty string
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
            non_na_values = [value for value in row if value]  # Get non-empty values
            if len(non_na_values) < 3 and row[0]:  # Check for college header
                college_name = row[0]
            else:
                row_data = row.tolist()
                if college_name:
                    row_data.insert(0, college_name)  # Insert college name as the first column
                row_data = [''] + row_data[1:]  # Shift to the right, skip the already shifted college name
                self.tree.insert("", "end", values=row_data)

    def add_schedule(self):
        # Ensure main window focus
        self.root.lift()
        self.root.focus_force()
        
        new_schedule = [''] * len(self.df.columns)  # Create an empty row with the right length

        for col in range(len(new_schedule)):
            column_name = self.df.iloc[3, col]  # Get the column name

            # Create a Toplevel window for input
            input_window = tk.Toplevel(self.root)
            input_window.title("Input Schedule Info")

            # Create a label for the column name
            label = tk.Label(input_window, text=f"Enter {column_name}:")
            label.pack(pady=5)

            # Create an entry field for input
            entry = tk.Entry(input_window)
            entry.pack(pady=5)

            # Function to handle skip button
            def on_skip():
                new_schedule[col] = ""  # Leave empty for this column
                input_window.destroy()

            # Function to handle submit
            def on_submit():
                value = entry.get()
                new_schedule[col] = value if value else ""  # Set value or leave empty
                input_window.destroy()

            # Function to handle cancel
            def on_cancel():
                input_window.destroy()
                self.cancel_add = True  # Set a flag to indicate cancellation

            # Create buttons for skip, submit, and cancel
            button_frame = tk.Frame(input_window)
            button_frame.pack(pady=5)

            skip_button = tk.Button(button_frame, text="Skip", command=on_skip)
            skip_button.pack(side=tk.LEFT, padx=5)

            submit_button = tk.Button(button_frame, text="Submit", command=on_submit)
            submit_button.pack(side=tk.LEFT, padx=5)

            cancel_button = tk.Button(button_frame, text="Cancel", command=on_cancel)
            cancel_button.pack(side=tk.LEFT, padx=5)

            # Wait for the input window to close
            self.root.wait_window(input_window)

            # Check if the process was cancelled
            if hasattr(self, 'cancel_add') and self.cancel_add:
                return  # Cancel was pressed, terminate the process

        # Insert the new schedule in the DataFrame
        second_column_value = new_schedule[1]  # Assuming the second column is the one to match
        insert_index = len(self.df)

        for index, row in self.df.iterrows():
            if row[1] == second_column_value:
                insert_index = index + 1
                break

        self.df.loc[insert_index:insert_index] = [new_schedule]
        self.df.reset_index(drop=True, inplace=True)
        self.show_table(self.df)


    def edit_schedule(self):
        pass  # Placeholder for editing schedule functionality

    def suggest_merge(self):
        pass  # Placeholder for suggesting merge functionality

    def find_conflict(self):
        pass  # Placeholder for finding conflict functionality


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.geometry("800x600")
    root.mainloop()
