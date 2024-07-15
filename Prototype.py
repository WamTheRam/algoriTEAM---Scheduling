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

        # Append the new schedule to the DataFrame
        self.df.loc[len(self.df)] = new_schedule  # Append to the end
        self.df.reset_index(drop=True, inplace=True)  # Reset index to keep it clean

        # Refresh the displayed table
        self.show_table(self.df)






    def edit_schedule(self):
        selected_item = self.tree.selection()

        if not selected_item:
            messagebox.showwarning("Warning", "Please select a schedule to edit.")
            return

        # Get the selected item's values
        item_values = self.tree.item(selected_item)['values'][1:]  # Skip the first empty string

        # Create a Toplevel window for input
        self.root.lift()
        self.root.focus_force()
        new_schedule = list(item_values)  # Start with existing values

        for col in range(len(new_schedule)):
            column_name = self.df.iloc[3, col]  # Get the column name

            # Create a Toplevel window for input
            input_window = tk.Toplevel(self.root)
            input_window.title("Edit Schedule Info")

            # Create a label for the column name
            label = tk.Label(input_window, text=f"Edit {column_name}:")
            label.pack(pady=5)

            # Create an entry field for input
            entry = tk.Entry(input_window)
            entry.insert(0, new_schedule[col])  # Populate with existing value
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
                self.cancel_edit = True  # Set a flag to indicate cancellation

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
            if hasattr(self, 'cancel_edit') and self.cancel_edit:
                return  # Cancel was pressed, terminate the process

        # Delete the old schedule from the DataFrame
        second_column_value = item_values[1]  # Assuming the second column is the one to match

        found_match = False
        for index, row in self.df.iterrows():
            if row[1] == second_column_value and all(row[i] == item_values[i] for i in range(len(item_values))):
                self.df.drop(index, inplace=True)  # Delete the old entry
                found_match = True
                break

        if not found_match:
            print("No matching row found for deletion.")

        # Insert the new schedule in the DataFrame
        insert_index = len(self.df)
        
        # Find the correct insert index based on the second column value
        for index, row in self.df.iterrows():
            if row[1] == second_column_value:
                insert_index = index + 1
                break

        self.df.loc[insert_index:insert_index] = [new_schedule]
        self.df.reset_index(drop=True, inplace=True)

        # Refresh the displayed table
        self.show_table(self.df)




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

        label = tk.Label(suggestion_window, text=f"Schedules with Enrollment Capacity below {threshold}:")
        label.pack(pady=5)

        for index, row in below_threshold.iterrows():
            row_text = ', '.join(map(str, row))
            row_label = tk.Label(suggestion_window, text=row_text)
            row_label.pack(pady=2)

        # Option to confirm merging
        confirm_button = tk.Button(suggestion_window, text="Confirm Merge", command=lambda: self.confirm_merge(below_threshold))
        confirm_button.pack(pady=10)





    def merge_schedules(self):
        selected_items = self.tree.selection()

        if len(selected_items) != 2:
            messagebox.showwarning("Merge Error", "Please select exactly two schedules to merge.")
            return

        # Retrieve the values of the selected items
        schedules = [self.tree.item(item)['values'][1:] for item in selected_items]  # Skip the first empty string

        # Check if the course codes are the same
        if schedules[0][1] != schedules[1][1]:  # Course code is in the second column
            messagebox.showwarning("Merge Error", "Cannot merge schedules with different course codes.")
            return

        # Create a new merged schedule
        merged_schedule = [
            '',  # First column remains empty
            f"{schedules[0][0]} + {schedules[1][0]}",  # Takers
            schedules[0][1],  # Course Code
            schedules[0][2],  # Course Title (assuming same)
            schedules[0][3],  # Offered To (assuming same)
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
            int(schedules[0][14]) + int(schedules[1][14]),  # Enrl Cap (adding integer values)
            '',  # Leave empty for this column
            '',  # Leave empty for this column
        ]

        # Add the merged schedule to the DataFrame without shifting
        self.df.loc[len(self.df)] = merged_schedule

        # Delete the original schedules from the DataFrame
        for item in selected_items:
            index = self.tree.index(item)
            self.df.drop(index, inplace=True)

        self.df.reset_index(drop=True, inplace=True)  # Reset index to keep it clean

        # Refresh the displayed table
        self.show_table(self.df)

        # Inform the user of the successful merge
        messagebox.showinfo("Success", "Schedules merged successfully.")





    def find_conflict(self):
        conflicts = []  # To hold all conflicts
        conflict_page = []  # Current page of conflicts

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
                            conflict_page.append((row1.tolist(), row2.tolist()))
                    except ValueError:
                        print(f"Invalid time format for {row1[7]}, {row1[8]}, {row2[7]}, {row2[8]}")

                if (row1[11] and row2[11]) and (row1[0] and row2[0]):  # Room2 is not empty
                    try:
                        # Check Day2, Begin2, End2, Room2 for conflicts
                        if (row1[12] == row2[12] and row1[11] == row2[11] and  # Same Room
                            not (int(row1[13]) >= int(row2[14]) or int(row2[13]) >= int(row1[14]))):  # Time overlap
                            conflict_page.append((row1.tolist(), row2.tolist()))
                    except ValueError:
                        print(f"Invalid time format for {row1[13]}, {row1[14]}, {row2[13]}, {row2[14]}")

            if conflict_page:
                conflicts.append(conflict_page)
                conflict_page = []  # Reset for the next set of conflicts

        # Display conflicts
        if conflicts:
            for idx, conflict_group in enumerate(conflicts):
                conflict_window = tk.Toplevel(self.root)
                conflict_window.title(f"Conflict Group {idx + 1}")

                label = tk.Label(conflict_window, text=f"Conflicts in Group {idx + 1}:")
                label.pack(pady=5)

                for row1, row2 in conflict_group:
                    row_label = tk.Label(conflict_window, text=f"{row1} <-> {row2}")
                    row_label.pack(pady=2)

                close_button = tk.Button(conflict_window, text="Close", command=conflict_window.destroy)
                close_button.pack(pady=10)
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
