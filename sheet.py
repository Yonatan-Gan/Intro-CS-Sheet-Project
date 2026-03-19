from cell import *
import json

CELL_FUNCTION_TYPES = ["MIN", "MAX", "AVG", "SUM"]
NUMBER_OF_LETTERS_A_TO_Z = 26
A_ASCII_VALUE = 65


def convert_column_index_to_letters(num):
    """
    Change number index of columns to letters
    """
    letters = ''
    while num:
        mod = (num - 1) % NUMBER_OF_LETTERS_A_TO_Z
        letters += chr(mod + A_ASCII_VALUE)
        num = (num - 1) // NUMBER_OF_LETTERS_A_TO_Z
    return ''.join(reversed(letters))


class Sheet:
    """
    Represents a sheet with cells organized in a dictionary
    """

    def __init__(self, length=10, width=10):
        """
        Initializes a new sheet with the given dimensions.
        """
        self.__dict = {}
        self.__length = length
        self.__width = width
        self.__function_cell_list = []
        # Populate the sheet with cells
        for i in range(length):
            for j in range(width):
                new_cell = Cell()
                name = convert_column_index_to_letters(j + 1)
                name += str(i + 1)
                self.__dict[name] = new_cell

    def update_sheet(self):
        """
        Updates all the function cells in the sheet
        """
        # Iterate through the function cell list
        for cell_key in self.__function_cell_list:
            # Get the cell object
            cell_obj = self.__dict[cell_key]
            all_float = True
            # Iterate through the cells in the range of the function cell
            for cell in cell_obj.get_range():
                # Get the type of the cell
                cell_type = cell.get_type()
                # Check if the cell type is not float
                if cell_type != float:
                    all_float = False
                    break
            # If any cell in the range is not float, reset the function cell to a regular cell
            if not all_float:
                self.__dict[cell_key] = Cell()
                self.__function_cell_list.remove(cell_key)
        # Call the calculate_value method for all function cells
        for cell_key in self.__function_cell_list:
            cell_obj = self.__dict[cell_key]
            if isinstance(cell_obj, FunctionCell):
                cell_obj.calculate_value()

    def get_dict(self):
        """
        returns the dictionary of the cells in the sheet
        """
        return self.__dict.copy()

    def get_length(self):
        """
        returns the length of the sheet
        """
        return self.__length

    def get_width(self):
        """
        returns the width of the sheet
        """
        return self.__width

    def column_range(self):
        """
        Returns the range of column identifiers
        """
        return [convert_column_index_to_letters(i) for i in range(1, self.get_width() + 1)]

    def set_value_in_cell(self, name, value):
        """
        Sets the value of a cell and updates the sheet.
        """
        if isinstance(value, (float, int)):
            value = float(value)
        self.__dict[name].set_value(value)
        self.update_sheet()  # Update the sheet after setting the value

    def get_value_in_cell(self, name):
        """
        Returns the value in a cell
        """
        return self.__dict[name].get_value()

    def get_function_cell_list(self):
        """
        Returns the list of function cells
        """
        return self.__function_cell_list

    def get_type_of_value_in_cell(self, name):
        """
        returns the type of the cell
        """
        return self.__dict[name].get_type()

    def get_cell(self, name):
        """
        Returns the cell object
        """
        return self.__dict[name]

    def __str__(self):
        """
        Returns a string representation of the sheet.
        """
        sheet_str = ""
        for j in range(self.__width):
            sheet_str += " " + convert_column_index_to_letters(j + 1)
        sheet_str += "\n"
        for i in range(self.__length):
            sheet_str += str(i + 1) + ' '
            for j in range(self.__width):
                cur_name = convert_column_index_to_letters(j + 1) + str(i + 1)
                sheet_str += "|" + str(self.__dict[cur_name].get_value()) + "| "
            sheet_str += "\n"
        return sheet_str

    def add_row(self):
        """
        Adds a new row to the sheet
        """
        new_row_index = self.__length + 1
        for j in range(self.__width):
            new_cell = Cell()
            name = convert_column_index_to_letters(j + 1) + str(new_row_index)
            self.__dict[name] = new_cell
        self.__length += 1

    def add_col(self):
        """
        Adds a new column to the sheet
        """
        new_col_index = self.__width + 1
        new_col_letter = convert_column_index_to_letters(new_col_index)
        for i in range(self.__length):
            new_cell = Cell()
            name = new_col_letter + str(i + 1)
            self.__dict[name] = new_cell
        self.__width += 1

    def set_function_cell(self, name, function_type, cell_range):
        """
        Sets a cell to be a function cell.
        """
        if function_type not in CELL_FUNCTION_TYPES:
            return False
        if name not in self.__dict.keys():
            return False
        for cell in cell_range:
            if cell not in self.__dict.keys():
                return False
            if self.__dict[cell].get_type() == str:
                return False
        cell_list_for_function = []
        for cell in cell_range:
            cell_list_for_function.append(self.__dict[cell])
        new_function_cell = FunctionCell(cell_list_for_function, function_type)
        self.__dict[name] = new_function_cell
        self.__function_cell_list.append(name)
        new_function_cell.calculate_value()
        return True

    def get_function_type_of_cell(self, name):
        """
        Returns the function type of a cell
        """
        cell = self.__dict[name]
        if isinstance(cell, FunctionCell):
            return cell.get_type()
        else:
            return None

    def get_range_of_function_cell(self, name):
        """
        Returns the range of a function cell
        """
        cell = self.__dict[name]
        if isinstance(cell, FunctionCell):
            cell_range_objects = cell.get_range()
            cell_range_names = [cell_name for cell_name, cell_object in self.__dict.items()
                                if cell_object in cell_range_objects]
            return cell_range_names
        else:
            return None

    def get_all_cells_data(self):
        """
        Returns data from all cells in the sheet as a list of dictionaries
        """
        all_cells_data = []
        # iterate through each row
        for i in range(self.__length):
            row_data = {}
            # iterate through each column
            for j in range(self.__width):
                cell_name = convert_column_index_to_letters(j + 1) + str(i + 1)
                cell_value = self.__dict[cell_name].get_value()
                row_data[self.__dict[cell_name].get_value()] = cell_value
            all_cells_data.append(row_data)
        return all_cells_data


def read_from_json(filename: str):
    """
    Read data from a JSON file and populate the sheet
    """
    try:
        with open(filename, 'r') as file:
            data = json.load(file)
        return data
    except Exception as e:
        print("An error occurred while reading from JSON file:", e)


def save_sheet_to_json(sheet, file_path):
    """
    Save data from a Sheet object to a JSON file.
    """
    data = []
    for row in range(sheet.get_length()):
        row_data = {}
        for col in range(sheet.get_width()):
            cell_name = convert_column_index_to_letters(col + 1) + str(row + 1)
            cell_value = sheet.get_value_in_cell(cell_name)
            cell_type = type(cell_value).__name__
            cell_data = {"value": cell_value, "type": cell_type}
            if cell_name in sheet.get_function_cell_list():
                cell_data["function_type"] = str(sheet.get_function_type_of_cell(cell_name))
                cell_data["cell_range"] = str(sheet.get_range_of_function_cell(cell_name))
            row_data[cell_name] = cell_data
        data.append(row_data)

    if not file_path.endswith('.json'):
        file_path += '.json'

    with open(file_path, 'w') as json_file:
        json.dump(data, json_file, indent=4)


def populate_sheet_from_json(json_data, sheet):
    """
    populate the sheet with data from a JSON file 
    """
    # Clear existing data in the sheet
    sheet.__init__(len(json_data), len(json_data[0]))
    # Populate sheet with JSON data
    for row_index, row_data in enumerate(json_data):
        for cell_name, cell_data in row_data.items():
            if 'function_type' in cell_data and 'cell_range' in cell_data:
                # Create a function cell
                cell_range = eval(cell_data['cell_range'])
                function_type = cell_data['function_type']
                sheet.set_function_cell(cell_name, function_type, cell_range)
            else:
                # Set regular cell value
                cell_value = cell_data['value']
                cell_type = cell_data['type']
                if cell_type == 'float':
                    cell_value = float(cell_value)
                elif cell_type == 'int':
                    cell_value = int(cell_value)
                sheet.set_value_in_cell(cell_name, cell_value)
