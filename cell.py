import math
CELL_FUNCTION_TYPES = ["MIN", "MAX", "AVG", "SUM"]


class Cell:
    """ This class represents a cell in the sheet we will work with, a cell can be a "Normal cell" with a numerical or
    a string value, or it can be a "Function cell" with a range of cells and a functions type, its value is determined
    by the values of the cells in its range and the functions type"""
    def __init__(self, value=None):
        """
        Initializes a cell with a given value. If no value is provided, initializes a regular cell with a default
        value of 0.0.
        """
        if not value:
            self.__type = None
            self.__value = 0.0
            self.__range = []
        else:
            self.set_value(value)

    def set_value(self, value):
        """
        Sets the value of the cell.
        legal value(float, int, str)
        """
        # Validate input value
        if not isinstance(value, (float, int, str)):
            print("Invalid value type. Value can only be numbers or string.")
            return

        # Convert value to float if it's numeric
        if isinstance(value, (float, int)):
            value = float(value)

        # Set the value and type of the cell
        self.__value = value
        self.__type = type(value)

    def get_range(self):
        """
        Returns a range of the cells of a function cell
        """
        return self.__range

    def get_value(self):
        """
        Returns the value in the cell
        """
        return self.__value

    def get_type(self):
        """
        Returns the type of the cell
        """
        return self.__type

    def __str__(self):
        """
        Returns a string representation the value in the cell
        """
        return str(self.get_value())


class FunctionCell(Cell):
    """
    Represents a function cell, which calculates its value based on a range of cells and a function type.
    """

    def __init__(self, list_of_cells=None, function_type=None):
        """
        Initializes a function cell with a range of cells and a function type.
        """
        super().__init__()
        if list_of_cells is not None and function_type in CELL_FUNCTION_TYPES:
            self.__range = list_of_cells
            self.__type = function_type
            self.calculate_value()

    def get_type(self):
        """
        Returns the function type of the cell
        """
        return self.__type

    def get_range(self):
        """
        Returns a range of the cells of a function cell
        """
        return self.__range

    def calculate_value(self):
        """
        Calculates the value of the function cell based on the function type and the values of the cells in its range.
        """
        if self.__type == "MIN":
            minimum = math.inf
            for cell in self.__range:
                cell_value = cell.get_value()
                if isinstance(cell_value, (int, float)):
                    if cell_value < minimum:
                        minimum = cell_value
                else:
                    # If the cell value is not numeric set the cell back to normal cell
                    self.set_value(0.0)
                    self.__class__ = Cell
                    return
            self.set_value(minimum)
        elif self.__type == "MAX":
            maximum = -math.inf
            for cell in self.__range:
                cell_value = cell.get_value()
                if isinstance(cell_value, (int, float)):
                    if cell_value > maximum:
                        maximum = cell_value
                else:
                    # If the cell value is not numeric set cell back to normal cell
                    self.set_value(0.0)
                    self.__class__ = Cell
                    return
            self.set_value(maximum)
        elif self.__type == "AVG" or self.__type == "SUM":
            sum_of_values = 0
            num_valid_cells = 0
            for cell in self.__range:
                cell_value = cell.get_value()
                if isinstance(cell_value, (int, float)):
                    sum_of_values += cell_value
                    num_valid_cells += 1
                else:
                    # If the cell value is not numeric, set the cell back to normal cell
                    self.__class__ = Cell
                    return
            if num_valid_cells > 0:
                if self.__type == "AVG":
                    self.set_value(sum_of_values / num_valid_cells)
                else:
                    self.set_value(sum_of_values)
            else:
                # If no valid cells found, set the cell back to normal cell
                self.__class__ = Cell
