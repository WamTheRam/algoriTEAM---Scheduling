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
        self.schedule_menu.add_command(label="Merge Schedule", command=self.merge_schedules)
        self.schedule_menu.add_command(label="Find Conflict", command=self.find_conflict)
        self.schedule_menu.add_command(label="Delete Schedule", command=self.delete_schedule)

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

        # Create a new window for schedule entry
        input_window = tk.Toplevel(self.root)
        input_window.title("Add New Schedule")

        # Retrieve column names from the table (row 4 in zero-indexed DataFrame)
        column_names = self.df.iloc[3].tolist()

        # Store new schedule data
        new_schedule = [''] * len(column_names)  # Create an empty row with the right length
        entries = []  # To keep track of entry widgets

        def on_submit():
            # Function to handle submit
            for i, entry in enumerate(entries):
                value = entry.get()
                new_schedule[i] = value if value else ""  # Set value or leave empty

            # Determine the position to insert the new schedule
            course_code = new_schedule[1]  # Assuming course code is in the second column
            insert_index = len(self.df)  # Default to the end of the DataFrame

            # Iterate through DataFrame to find the insertion point
            for idx, row in self.df.iterrows():
                if row[1] == course_code:
                    insert_index = idx + 1  # Insert below the last occurrence of the same course code
                elif idx > insert_index:
                    break  # Stop when we've gone past the insertion point

            # Insert new schedule into DataFrame
            new_row = pd.DataFrame([new_schedule], columns=self.df.columns)
            self.df = pd.concat([self.df.iloc[:insert_index], new_row, self.df.iloc[insert_index:]]).reset_index(drop=True)

            # Refresh the displayed table
            self.show_table(self.df)
            input_window.destroy()

        def on_cancel():
            # Function to handle cancel
            input_window.destroy()

        # Create labels and entry fields for each column
        for col, column_name in enumerate(column_names):
            # Create a label for the column name
            label = tk.Label(input_window, text=f"Enter {column_name}:")  # Use the actual column name
            label.grid(row=col, column=0, padx=10, pady=5, sticky='e')

            # Create an entry field for input
            entry = tk.Entry(input_window)
            entry.grid(row=col, column=1, padx=10, pady=5, sticky='w')

            # Append entry to the list
            entries.append(entry)

        # Create buttons for submit and cancel
        button_frame = tk.Frame(input_window)
        button_frame.grid(row=len(column_names), columnspan=2, pady=10)

        submit_button = tk.Button(button_frame, text="Submit", command=on_submit)
        submit_button.pack(side=tk.LEFT, padx=5)

        cancel_button = tk.Button(button_frame, text="Cancel", command=on_cancel)
        cancel_button.pack(side=tk.LEFT, padx=5)






    def edit_schedule(self):
        selected_item = self.tree.selection()

        if not selected_item:
            messagebox.showwarning("Warning", "Please select a schedule to edit.")
            return

        # Get the selected item's values
        item_values = self.tree.item(selected_item)['values'][1:]  # Skip the first empty string

        # Retrieve column names from the table (row 4 in zero-indexed DataFrame)
        column_names = self.df.iloc[3].tolist()

        # Create a Toplevel window for input
        input_window = tk.Toplevel(self.root)
        input_window.title("Edit Schedule Info")

        # Create labels and entry fields for each column
        entries = []
        for col, column_name in enumerate(column_names):
            label = tk.Label(input_window, text=f"Edit {column_name}:")
            label.grid(row=col, column=0, padx=10, pady=5, sticky='e')

            entry = tk.Entry(input_window)
            entry.insert(0, item_values[col])
            entry.grid(row=col, column=1, padx=10, pady=5, sticky='w')
            entries.append(entry)

        # Create buttons for submit and cancel
        def on_submit():
            new_values = [entry.get() for entry in entries]
            
            selected_index = self.tree.index(selected_item)
            adjusted_index = selected_index + 6

            if adjusted_index < 0 or adjusted_index >= len(self.df):
                messagebox.showwarning("Warning", "Adjusted index is out of range.")
                input_window.destroy()
                return

            # Update the DataFrame at the adjusted index
            self.df.iloc[adjusted_index] = new_values
            self.df.reset_index(drop=True, inplace=True)
            self.show_table(self.df)
            input_window.destroy()

        def on_cancel():
            input_window.destroy()

        button_frame = tk.Frame(input_window)
        button_frame.grid(row=len(entries), columnspan=2, pady=10)
        
        submit_button = tk.Button(button_frame, text="Submit", command=on_submit)
        submit_button.pack(side=tk.LEFT, padx=5)
        
        cancel_button = tk.Button(button_frame, text="Cancel", command=on_cancel)
        cancel_button.pack(side=tk.LEFT, padx=5)











    def suggest_merge(self):
        # Ask the user for the enrollment threshold
        threshold = simpledialog.askinteger("Enrollment Threshold", "Enter the student threshold for merging:")

        if threshold is None:
            return  # User canceled the input

        # Convert the Enrl Cap column to numeric, coercing errors to NaN
        self.df[14] = pd.to_numeric(self.df[14], errors='coerce')

        # Find rows with Enrl Cap below the threshold
        below_threshold = self.df[self.df[14] < threshold]

        if below_threshold.empty:
            messagebox.showinfo("No Merges Suggested", "No schedules below the specified threshold found.")
            return

        # Create a Toplevel window to show suggestions
        suggestion_window = tk.Toplevel(self.root)
        suggestion_window.title("Merge Suggestions")

        # Create a Canvas widget for scrolling
        canvas = tk.Canvas(suggestion_window)
        scrollbar = tk.Scrollbar(suggestion_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack the Canvas and Scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Populate the scrollable frame with suggestions
        label = tk.Label(scrollable_frame, text=f"Schedules with Enrollment Capacity below {threshold}:")
        label.pack(pady=5)

        for index, row in below_threshold.iterrows():
            row_text = ', '.join(map(str, row))
            row_label = tk.Label(scrollable_frame, text=row_text, anchor="w", padx=5)
            row_label.pack(pady=2, fill="x")

        # Adjust the window size based on content
        suggestion_window.update_idletasks()
        suggestion_window.geometry(f"600x{min(400, suggestion_window.winfo_height())}")





    def merge_schedules(self):
        selected_items = self.tree.selection()

        if len(selected_items) != 2:
            messagebox.showwarning("Merge Error", "Please select exactly two schedules to merge.")
            return

        # Retrieve the values of the selected items
        schedules = [self.tree.item(item)['values'] for item in selected_items]  # Retrieve all values, including empty columns

        # Check if the course codes are the same
        if schedules[0][2] != schedules[1][2]:  # Assuming course code is the third column
            messagebox.showwarning("Merge Error", "Cannot merge schedules with different course codes.")
            return

        # Construct a new merged schedule
        merged_schedule = [
            '',  # First column (e.g., ID) remains empty or you can decide how to handle it
            f"{schedules[0][1]} + {schedules[1][1]}",  # Takers, assuming it's the second column
            schedules[0][2],  # Course Code (assuming it's the third column)
            schedules[0][3],  # Course Title (assuming it's the fourth column)
            schedules[0][4],  # Offered To (assuming it's the fifth column)
            '',  # Leave empty for this column
            '',  # Leave empty for this column
            '',  # Leave empty for this column
            '',  # Leave empty for this column
            '',  # Leave empty for this column
            '',  # Leave empty for this column
            '',  # Leave empty for this column
            '',  # Leave empty for this column
            '',  # Leave empty for this column
            '',  # Leave empty for this column
            int(schedules[0][15]) + int(schedules[1][15]),  # Enrl Cap (adding integer values), assuming it's the 16th column
            '',  # Leave empty for this column
            '',  # Leave empty for this column
        ]

        # Ensure that the merged_schedule length matches the number of columns in the DataFrame
        if len(merged_schedule) != len(self.df.columns):
            messagebox.showwarning("Merge Error", "Merged schedule does not match the DataFrame columns.")
            return

        # Add the merged schedule to the DataFrame
        self.df.loc[len(self.df)] = merged_schedule

        # Delete the original schedules from the DataFrame
        for item in selected_items:
            index = self.tree.index(item)
            self.df.drop(index, inplace=True)

        # Reset index to keep it clean
        self.df.reset_index(drop=True, inplace=True)

        # Refresh the displayed table
        self.show_table(self.df)

        # Inform the user of the successful merge
        messagebox.showinfo("Success", "Schedules merged successfully.")






    def find_conflict(self):
        conflicts = []  # To hold all conflicts

        for index1, row1 in self.df.iterrows():
            for index2, row2 in self.df.iterrows():
                if index1 >= index2:  # Skip duplicate checks and self-checks
                    continue

                # Check for non-empty values in room and days
                if (row1[5] and row2[5]) and (row1[0] and row2[0]):  # Room1 is not empty
                    try:
                        # Check Day1, Begin1, End1, Room1 for conflicts
                        if (row1[6] == row2[6] and row1[5] == row2[5] and  # Same Room
                            not (int(row1[7]) >= int(row2[8]) or int(row2[7]) >= int(row1[8]))):  # Time overlap
                            conflicts.append((row1.tolist(), row2.tolist()))
                    except ValueError:
                        print(f"Invalid time format for {row1[7]}, {row1[8]}, {row2[7]}, {row2[8]}")

                if (row1[11] and row2[11]) and (row1[0] and row2[0]):  # Room2 is not empty
                    try:
                        # Check Day2, Begin2, End2, Room2 for conflicts
                        if (row1[12] == row2[12] and row1[11] == row2[11] and  # Same Room
                            not (int(row1[13]) >= int(row2[14]) or int(row2[13]) >= int(row1[14]))):  # Time overlap
                            conflicts.append((row1.tolist(), row2.tolist()))
                    except ValueError:
                        print(f"Invalid time format for {row1[13]}, {row1[14]}, {row2[13]}, {row2[14]}")

        # Display conflicts
        if conflicts:
            # Create a Toplevel window for conflict display
            conflict_window = tk.Toplevel(self.root)
            conflict_window.title("Conflict Groups")

            # Create a Canvas widget for scrolling
            canvas = tk.Canvas(conflict_window)
            canvas.pack(side="left", fill="both", expand=True)

            # Create a Scrollbar widget
            scrollbar_y = tk.Scrollbar(conflict_window, orient="vertical", command=canvas.yview)
            scrollbar_x = tk.Scrollbar(conflict_window, orient="horizontal", command=canvas.xview)
            scrollbar_y.pack(side="right", fill="y")
            scrollbar_x.pack(side="bottom", fill="x")

            # Create a Frame to hold the Treeview inside the Canvas
            scrollable_frame = tk.Frame(canvas)
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

            # Bind the Frame to the Canvas' scroll region
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            # Configure the Canvas' scrollbars
            canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

            # Create a Notebook widget for tabbed pages
            notebook = ttk.Notebook(scrollable_frame)
            notebook.pack(fill="both", expand=True)

            # Define column names and widths
            column_names = self.df.iloc[3].tolist()
            column_names = [str(col) for col in column_names]  # Convert to strings
            column_widths = [150] * len(column_names)  # Adjust column widths as needed

            # Create tabs for each conflict group
            for i, conflict_pair in enumerate(conflicts):
                tab = ttk.Frame(notebook)
                notebook.add(tab, text=f"Conflict {i+1}")

                # Create Treeview widget for tabular display
                tree = ttk.Treeview(tab, columns=column_names, show="headings")
                tree.pack(fill="both", expand=True)

                # Define column headings
                for col, col_name in enumerate(column_names):
                    tree.heading(col, text=col_name)
                    tree.column(col, width=column_widths[col])

                # Add rows to Treeview
                for row in conflict_pair:
                    tree.insert("", "end", values=row)

            # Ensure the canvas is updated with the new scroll region
            scrollable_frame.update_idletasks()
            canvas.config(scrollregion=canvas.bbox("all"))

        else:
            messagebox.showinfo("No Conflicts", "No scheduling conflicts found.")












    def delete_schedule(self):
        selected_item = self.tree.selection()
        
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a schedule to delete.")
            return
        
        # Confirm deletion
        confirm = messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete the selected schedule?")
        if not confirm:
            return

        # Get the selected item's values
        item_values = self.tree.item(selected_item)['values'][1:]  # Skip the first empty string
        
        # Debug output
        print("Selected Item Values:", item_values)

        # Find the original row in the DataFrame that matches the selected item's values
        found_match = False
        for index, row in self.df.iterrows():
            print(f"Checking row {index}: {row.tolist()}")  # Debug output
            if len(row) == len(item_values):  # Ensure lengths match
                # Convert both to strings and compare
                if all(str(row[i]).strip() == str(item_values[i]).strip() for i in range(len(item_values))):
                    print(f"Deleting row {index}")  # Indicate which row will be deleted
                    self.df.drop(index, inplace=True)
                    found_match = True
                    break

        if not found_match:
            print("No matching row found for deletion.")  # Indicate no match found

        # Reset the DataFrame index
        self.df.reset_index(drop=True, inplace=True)

        # Refresh the displayed table
        self.show_table(self.df)



if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.geometry("800x600")
    root.mainloop()
