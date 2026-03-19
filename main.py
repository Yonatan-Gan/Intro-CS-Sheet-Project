import tkinter as tk
from tkinter import simpledialog
from tkinter.font import Font
from sheet import *
from table import Table
import os
import argparse

MAIN_LABEL_FONT_SIZE = 18
BUTTON_FONT_SIZE = 14


def validate_cell(cell, sheet):
    """
    Validate if the cell is in the correct format and within the sheet's range
    """
    if len(cell) < 2 or not cell[0].isalpha() or not cell[1:].isdigit():
        return False

    col, row = cell[:-1], int(cell[1:])
    max_col = convert_column_index_to_letters(sheet.get_width())
    return 'A' <= col <= max_col and 1 <= row <= sheet.get_length()


class SheetProgramGUI:
    """
    Represents the main GUI of the Sheet Program.
    """
    def __init__(self):
        """
        Initializes the SheetGUI class.
        """
        self.root = tk.Tk()
        self.root.title("Sheet Program")
        self.root.geometry("300x200")
        self.create_main_window()

    def create_main_window(self):
        """
        Creates the main window of the GUI with buttons for creating a new sheet and loading an existing sheet.
        """
        main_label = tk.Label(self.root, text="Welcome to Sheet Program", font=("Helvetica", MAIN_LABEL_FONT_SIZE))
        main_label.pack(pady=10)

        button_create = tk.Button(self.root, text="Create a New Sheet", font=("Helvetica", BUTTON_FONT_SIZE),
                                  command=self.create_sheet_window)
        button_create.pack(expand=True)

        button_load = tk.Button(self.root, text="Load an Existing Sheet from JSON file",
                                command=self.import_sheet_window)
        button_load.pack(expand=True)

    def create_sheet_window(self):
        """
        Opens a new window for creating a new sheet.
        """
        CreateSheetWindow(self.root)

    def import_sheet_window(self):
        """
        Opens a new window for importing a sheet from a JSON file.
        """
        ImportSheetWindow(self.root)

    def run(self):
        """
        Runs the main event loop of the GUI.
        """
        self.root.mainloop()


class CreateSheetWindow:
    """
    Represents the window for creating a new sheet.
    """
    def __init__(self, parent):
        """
        Initializes the CreateSheetWindow class.
        """
        self.width_entry = None
        self.length_entry = None
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("Create a New Sheet")
        self.create_widgets()

    def create_widgets(self):
        """
        Creates widgets for entering the dimensions of the new sheet.
        """
        main_label = tk.Label(self.window, text="Enter width and length of the new sheet:", font=("Helvetica", 14))
        main_label.pack(pady=5)

        length_label = tk.Label(self.window, text="Length:", font=("Helvetica", 12))
        length_label.pack(side=tk.LEFT, padx=5)

        self.length_entry = tk.Entry(self.window, width=5)
        self.length_entry.pack(side=tk.LEFT)

        width_label = tk.Label(self.window, text="Width:", font=("Helvetica", 12))
        width_label.pack(side=tk.LEFT, padx=5)

        self.width_entry = tk.Entry(self.window, width=5)
        self.width_entry.pack(side=tk.LEFT)

        self.window.bind('<Return>', self.create_sheet)

        create_button = tk.Button(self.window, text="Create Sheet", command=self.create_sheet)
        create_button.pack(pady=5)

    def create_sheet(self, event=None):
        """
        Creates a new sheet based on the entered dimensions.
        """
        error_label = tk.Label(self.window, text="",
                               fg="red")  # Initialize the error label outside the try-except block

        try:
            length = int(self.length_entry.get())
            width = int(self.width_entry.get())
            new_sheet = Sheet(length, width)
            self.parent.destroy()
            SheetProcessor(new_sheet)
        except ValueError:
            # Remove any existing error labels from the window
            for widget in self.window.winfo_children():
                if isinstance(widget, tk.Label) and widget['fg'] == "red":
                    widget.pack_forget()

            error_label.config(text="Invalid input. Please enter valid numbers.")  # Update the text of the error label
            error_label.pack()  # Pack the error label after updating its text


class ImportSheetWindow:
    """
    Represents the window for importing a sheet from a JSON file.
    """
    def __init__(self, parent):
        """
        Initializes the ImportSheetWindow class.
        """
        self.json_entry = None
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("Import Sheet from JSON")
        self.create_widgets()

    def create_widgets(self):
        """
        Creates widgets for entering the path of the JSON file.
        """
        main_label = tk.Label(self.window, text="Enter the path of the JSON file:", font=("Helvetica", 12))
        main_label.pack(pady=5)

        self.json_entry = tk.Entry(self.window)
        self.json_entry.pack(pady=5)

        import_button = tk.Button(self.window, text="Load Sheet from JSON", command=self.import_sheet_from_json)
        import_button.pack(pady=5)

    def import_sheet_from_json(self):
        """
        Creates a new sheet from the JSON file and runs the sheetProcessor on it
        """
        file_path = self.json_entry.get()
        if file_path:
            if os.path.isfile(file_path):  # Check if the file exists
                json_data = read_from_json(file_path)
                if json_data:
                    new_sheet = Sheet()
                    populate_sheet_from_json(json_data, new_sheet)
                    self.parent.destroy()
                    sheet_processor = SheetProcessor(new_sheet)
                    sheet_processor.sheet_window.update_table()
                else:
                    for widget in self.window.winfo_children():
                        if isinstance(widget, tk.Label) and widget['fg'] == "red":
                            widget.pack_forget()

                    error_label = tk.Label(self.window, text="Error: Invalid JSON file.", fg="red")
                    error_label.pack()
            else:
                for widget in self.window.winfo_children():
                    if isinstance(widget, tk.Label) and widget['fg'] == "red":
                        widget.pack_forget()
                error_label = tk.Label(self.window, text="Error: Invalid JSON file.", fg="red")
                error_label.pack()
        else:
            for widget in self.window.winfo_children():
                if isinstance(widget, tk.Label) and widget['fg'] == "red":
                    widget.pack_forget()

            error_label = tk.Label(self.window, text="Error: Please enter a file path.", fg="red")
            error_label.pack()


class SheetProcessor:
    """
    Represents the processor for handling operations on a sheet.
    """
    def __init__(self, sheet):
        """
        Initializes the SheetProcessor class.
        """
        self.sheet_window = None  # to be set later
        self.sheet = sheet
        self.create_control_window()

    def set_function_cell(self, func, result_label):
        """
        Sets a cell as a function cell based on user input.
        """
        cell_to_set = simpledialog.askstring("Input", "Enter the cell to set as a function cell (e.g., A1):")
        if cell_to_set and validate_cell(cell_to_set, self.sheet):
            cell_range = simpledialog.askstring("Input",
                                                "Enter the cells to set as the range of data for function cell "
                                                "(e.g., A1,B2,C3):")

            if cell_range:
                cell_range = cell_range.split(',')
                for cell_name in cell_range:
                    if not validate_cell(cell_name, self.sheet):
                        result_label.config(text=f"Invalid cell name: {cell_name}", fg="red")
                        return False
                happened = self.sheet.set_function_cell(cell_to_set, func, cell_range)
                if happened:
                    self.sheet_window.update_table()
                else:
                    result_label.config(text=f"Cells in range of Function cell must all be numbers", fg="red")

        else:
            result_label.config(text=f"Invalid cell name: {cell_to_set}", fg="red")

    def give_cell_value(self, cell_name, result_label):
        """
        Calculates the value of the cell that was entered in the command line and displays the result.
        """
        if validate_cell(cell_name, self.sheet):
            cell_value = self.sheet.get_value_in_cell(cell_name)  # Calculate the value of the cell
            result_label.config(text=f"Value of {cell_name}: {cell_value}", fg="white")  # Update the result label
        else:
            result_label.config(text="Error: Invalid argument", fg="red")  # Update the result label

    def calculate_expression_value(self, cell_data, result_label):
        """
        Calculates the value of the expression that was entered in the command line and displays the result.
        """
        cell_data = cell_data.split()
        if len(cell_data) == 3:  # Check if input contains cell name, operator, and number
            cell_name, operator, number = cell_data
            if operator in ['+', '-', '*', '/']:
                if validate_cell(cell_name, self.sheet):
                    cell_value = self.sheet.get_value_in_cell(cell_name)  # Get value of the cell
                    if self.sheet.get_type_of_value_in_cell(cell_name) is not str:
                        if operator == '/' and number == '0':
                            result_label.config(text="Error: Division by zero", fg="red")  # Update the result label
                        else:
                            result = eval(f"{cell_value} {operator} {number}")  # Perform calculation
                            result_label.config(text=f"Result: {result}", fg="green")  # Update the result label
                    else:
                        result_label.config(text="Error: Cell contains a string", fg="red")  # Update the result label
                else:
                    result_label.config(text="Error: Invalid cell name", fg="red")  # Update the result label
            else:
                result_label.config(text="Error: Invalid operator", fg="red")  # Update the result label
        elif len(cell_data) == 1:
            self.give_cell_value(cell_data[0], result_label)
        else:
            result_label.config(text="Error: Invalid argument", fg="red")  # Update the result label

    def save_to_json(self, file_path_entry, result_label, sheet_table):
        """
        Saves the sheet to a JSON file.
        """
        file_path = file_path_entry.get()
        if file_path:
            save_sheet_to_json(self.sheet, file_path)
            # Provide feedback to the user
            result_label.config(text=f"Sheet saved to {file_path}.json", fg="purple")
        else:
            result_label.config(text="Error: Please enter a file path.", fg="red")

    def create_control_window(self):
        """
        Creates the control window for interacting with the sheet.
        """
        def on_enter(event):
            """
            Handle the Enter key press event.
            """
            self.sheet_window.enter("")

        def exit_program():
            """
            Exits the program.
            """
            control_window.destroy()

        def back_to_main():
            """
            Returns to the main window.
            """
            control_window.destroy()
            app.root = tk.Tk()  # Recreate the root window
            app.root.title("Sheet Program")
            app.root.geometry("300x200")
            app.create_main_window()

        control_window = tk.Tk()
        control_window.title("Control Board")

        # Create frames to organize widgets
        top_frame = tk.Frame(control_window, padx=10, pady=10)
        top_frame.grid(row=0, column=0, sticky="ew")

        mid_frame = tk.Frame(control_window, padx=10, pady=10)
        mid_frame.grid(row=1, column=0, sticky="ew")

        bottom_frame = tk.Frame(control_window, padx=10, pady=10)
        bottom_frame.grid(row=2, column=0, sticky="ew")

        # Table window
        self.sheet_window = Table(tk.Toplevel(), self.sheet)
        self.sheet_window.enter("")

        # Command line entry and label
        command_line_label = tk.Label(mid_frame, text="This is your calculator line:")
        command_line_label.grid(row=0, column=0, sticky="w", padx=5)

        command_line_entry = tk.Entry(mid_frame)
        command_line_entry.grid(row=0, column=1, padx=5)

        # Bind Enter key to command line entry
        control_window.bind('<Return>', on_enter)

        # Instruction labels for command line
        instruction_label1 = tk.Label(mid_frame, text="To calculate: Enter cell operator number (e.g., A1 + 5)")
        instruction_label1.grid(row=1, column=0, columnspan=3, pady=(5, 0), sticky="w")
        instruction_label2 = tk.Label(mid_frame, text="Operators: +, -, *, /")
        instruction_label2.grid(row=2, column=0, columnspan=3, sticky="w")

        # Calculate button
        calculate_button = tk.Button(mid_frame, text="Calculate", command=lambda: self.calculate_expression_value(
            command_line_entry.get(), console_label))
        calculate_button.grid(row=0, column=2, padx=5, pady=5)

        # Console label
        console_label = tk.Label(mid_frame, text="", fg="red")
        console_label.grid(row=0, column=3, padx=5, pady=5)

        # Buttons
        max_button = tk.Button(mid_frame, text="MAX", command=lambda: self.set_function_cell("MAX", console_label))
        max_button.grid(row=3, column=0, padx=5, pady=5)

        min_button = tk.Button(mid_frame, text="MIN", command=lambda: self.set_function_cell("MIN", console_label))
        min_button.grid(row=3, column=1, padx=5, pady=5)

        sum_button = tk.Button(mid_frame, text="SUM", command=lambda: self.set_function_cell("SUM", console_label))
        sum_button.grid(row=3, column=2, padx=5, pady=5)

        avg_button = tk.Button(mid_frame, text="AVG", command=lambda: self.set_function_cell("AVG", console_label))
        avg_button.grid(row=3, column=3, padx=5, pady=5)

        # Save to JSON
        json_label = tk.Label(mid_frame, text="Enter path for JSON file:")
        json_label.grid(row=4, column=0, sticky="e", padx=5, pady=5)

        file_path_entry = tk.Entry(mid_frame)
        file_path_entry.grid(row=4, column=1, padx=5, pady=5)

        save_to_json_button = tk.Button(mid_frame, text="Save to JSON",
                                        command=lambda: self.save_to_json(file_path_entry, console_label,
                                                                          self.sheet_window))
        save_to_json_button.grid(row=4, column=2, padx=5, pady=5, sticky="w")

        # Bold label for informing the user about Enter key
        bold_font = Font(weight="bold", size=18)
        enter_info_label = tk.Label(mid_frame, text="Note: The sheet updates only "
                                                    "when <Enter> is pressed.",
                                    font=bold_font, bg="white", fg="blue")
        enter_info_label.grid(row=5, column=0, columnspan=4, pady=(10, 0), sticky="w")

        # Add Row and Add Column buttons
        add_row_button = tk.Button(bottom_frame, text="Add Row", command=lambda: self.sheet_window.add_row())
        add_row_button.grid(row=0, column=0, padx=5, pady=5)

        add_column_button = tk.Button(bottom_frame, text="Add Column", command=lambda: self.sheet_window.add_col())
        add_column_button.grid(row=0, column=1, padx=5, pady=5)

        # Exit and back to main buttons
        exit_button = tk.Button(bottom_frame, text="Exit Program", command=exit_program)
        exit_button.grid(row=0, column=2, pady=5)

        back_to_main_button = tk.Button(bottom_frame, text="Back to Main Menu", command=back_to_main)
        back_to_main_button.grid(row=0, column=3, pady=5)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=
                                     '''Sheet Program that works with a GUI.
                                        Just run the program and choose to create a new sheet.
                                        If you have already created one, you can use the "load from Json" window 
                                        to load it. write: python3 main.py to run the program'''
                                     )

    args = parser.parse_args()
    app = SheetProgramGUI()
    app.run()
