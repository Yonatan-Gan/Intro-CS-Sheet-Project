from tkinter import *
from sheet import *


class Table:
    """
    This class will represent a GUI window for a single sheet
    """
    def __init__(self, root_1, sheet_1):
        self.label_cols = None  # to be defined later
        self.label_rows = None  # to be defined later
        self.root = root_1
        self.__sheet = sheet_1
        self.entry_dict = {}
        self.create_table()

    def create_table(self):
        """
        Create the table GUI
        """
        max_widths = {}  # Dictionary to store the maximum width needed for each column
        for i in range(1, self.__sheet.get_length() + 1):
            label_rows = Label(self.root, text=i, width=1)
            label_rows.grid(row=i + 1, column=0)
            for j in range(1, self.__sheet.get_width() + 1):
                if i == 1:
                    label_cols = Label(self.root, text=convert_column_index_to_letters(j), width=5)
                    label_cols.grid(row=i, column=j)
                cell_name = convert_column_index_to_letters(j) + str(i)
                entry = Entry(self.root, justify=CENTER)
                entry.grid(row=i + 1, column=j)
                entry.insert(END, self.__sheet.get_value_in_cell(cell_name))
                entry.bind('<Return>', self.enter)
                self.entry_dict[cell_name] = entry
                if isinstance(self.__sheet.get_dict()[cell_name], FunctionCell):
                    entry.config(bg='green')

                # Update the maximum width for the column
                content_width = len(str(self.__sheet.get_value_in_cell(cell_name)))
                max_widths[j] = max(max_widths.get(j, 0), content_width)

        # Set the width of each entry based on the maximum width needed for each column
        for j in range(1, self.__sheet.get_width() + 1):
            width = max_widths.get(j, 5) + 2  # Add some padding
            for i in range(1, self.__sheet.get_length() + 1):
                cell_name = convert_column_index_to_letters(j) + str(i)
                self.entry_dict[cell_name].config(width=width)

    def add_row(self):
        """
        Adds a new row to the sheet and updates the GUI
        """
        self.__sheet.add_row()
        self.label_rows = Label(self.root, text=self.__sheet.get_length(), width=1)
        self.label_rows.grid(row=self.__sheet.get_length() + 1, column=0)
        # Create entry widgets for the new row
        new_row_index = self.__sheet.get_length()
        for j in range(1, self.__sheet.get_width() + 1):
            cell_name = convert_column_index_to_letters(j) + str(new_row_index)
            entry = Entry(self.root, width=5, justify=CENTER)
            entry.grid(row=new_row_index + 1, column=j)
            entry.insert(END, self.__sheet.get_value_in_cell(cell_name))
            entry.bind('<Return>', self.enter)
            self.entry_dict[cell_name] = entry
        self.update_table()

    def add_col(self):
        """
        Adds a new column to the sheet and updates the GUI
        """
        self.__sheet.add_col()
        self.label_cols = Label(self.root, text=convert_column_index_to_letters(self.__sheet.get_width()), width=5)
        self.label_cols.grid(row=1, column=self.__sheet.get_width())
        # Create entry widgets for the new column
        new_col_index = self.__sheet.get_width()
        new_col_letter = convert_column_index_to_letters(new_col_index)
        for i in range(1, self.__sheet.get_length() + 1):
            cell_name = new_col_letter + str(i)
            entry = Entry(self.root, width=5, justify=CENTER)
            entry.grid(row=i + 1, column=new_col_index)
            entry.insert(END, self.__sheet.get_value_in_cell(cell_name))
            entry.bind('<Return>', self.enter)
            self.entry_dict[cell_name] = entry
        self.update_table()

    def enter(self, event):
        """
        Handle cell value changes
        """
        for cell_name, entry_widget in self.entry_dict.items():
            input_value = entry_widget.get()
            cell_type = self.__sheet.get_type_of_value_in_cell(cell_name)

            # Convert input value to the appropriate data type based on the cell's type
            if cell_type not in CELL_FUNCTION_TYPES:
                try:
                    input_value = float(input_value)
                except ValueError:
                    input_value = input_value
            # Update the cell value in te sheet
            self.__sheet.set_value_in_cell(cell_name, input_value)
        # Update the GUI with the new cell values
        self.update_table()

    def update_table(self):
        """
        Update the table with new cell values and adjust entry widths dynamically
        """
        max_widths = {}  # Dictionary to store the maximum width needed for each column

        # Update the cell values and calculate the maximum width for each column
        for cell_name, entry_widget in self.entry_dict.items():
            # Highlight function cells with a different background color
            if isinstance(self.__sheet.get_dict()[cell_name], FunctionCell):
                entry_widget.config(bg='green')
            else:
                entry_widget.config(bg='grey')
            value = self.__sheet.get_value_in_cell(cell_name)
            content_width = len(str(value))
            max_widths[cell_name[0]] = max(max_widths.get(cell_name[0], 0), content_width)
            entry_widget.delete(0, END)
            entry_widget.insert(END, value)

        # Adjust the width of entry widgets based on the maximum width needed for each column
        for column, max_width in max_widths.items():
            width = max_width + 2  # Add some padding
            for i in range(1, self.__sheet.get_length() + 1):
                cell_name = column + str(i)
                self.entry_dict[cell_name].config(width=width)
