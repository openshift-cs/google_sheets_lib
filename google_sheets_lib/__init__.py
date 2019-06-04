from logging import getLogger, Formatter, StreamHandler
from typing import Dict, List, Union, Pattern
from sys import stdout
import pygsheets
import re


class GoogleSheets:
    sheet: 'pygsheets.Spreadsheet' = None
    ws: 'pygsheets.Worksheet' = None
    ws_range_format: Pattern = re.compile(r'(?P<worksheet>[a-zA-Z0-9_]+)!(?P<start_range>[A-Z0-9]+):(?P<end_range>[A-Z0-9]+)')

    def __init__(self, drive_folder_id=None, logging_level: str='INFO', service_account_file: str=None) -> None:
        """Google Sheets class initializer"""
        # Disable sub-logging from the `googleapiclient` discovery.py
        getLogger('googleapiclient.discovery').setLevel('WARNING')
        self.log = getLogger(__name__)
        self.log.setLevel(logging_level)
        formatter = Formatter('{message:<80} {asctime}:{levelname:<9} - {module}:{funcName}:{lineno}', "%H:%M:%S", '{')
        stdout_handler = StreamHandler(stdout)
        stdout_handler.setFormatter(formatter)
        self.log.addHandler(stdout_handler)
        self.folder = drive_folder_id
        scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        if service_account_file:
            self.client = pygsheets.authorize(service_account_file=service_account_file, scopes=scopes)
        else:
            self.client = pygsheets.authorize(credentials_directory=None, scopes=scopes)

    @staticmethod
    def format_addr(addr):
        return pygsheets.utils.format_addr(addr)

    def list_sheets(self) -> List[str]:
        """Lists the available Google Spreadsheets under `self.folder`

        Returns:
            A list of Google Spreadsheet titles
        """
        if self.folder:
            return self.client.spreadsheet_titles(query=f'"{self.folder}" in parents')
        else:
            return self.client.spreadsheet_titles()

    def create_sheet(self, title: str) -> 'GoogleSheets':
        """Creates and activates a Google Spreadsheet within `drive_folder_id`
                https://pygsheets.readthedocs.io/en/stable/reference.html#pygsheets.Client.create

        Args:
            title (str): Title of the spreadsheet to create

        Returns:
            A copy of the current object; this allows call chaining
        """
        self.sheet = self.client.create(title, folder=self.folder)
        return self

    def delete_sheet(self, *, title: str=None, key: str=None, url: str=None, ignore_errors: bool=False) -> 'GoogleSheets':
        """Deletes the specified Google Spreadsheet by ID only

        Args:
            title (str, optional): The title of the spreadsheet to delete
            key (str, optional): The key ID of the spreadsheet to delete
            url (str, optional): The URL of the spreadsheet to delete
            ignore_errors (bool, optional): Can optionally ignore any errors like missing spreadsheets

        Returns:
            A copy of the current object; this allows call chaining

        Raises:
            KeyError: If the specified Spreadsheet does not exist
        """
        prev_sheet = self.sheet
        try:
            self.set_sheet(title=title, key=key, url=url)
            self.sheet.delete()
            self.sheet = prev_sheet
        except (pygsheets.SpreadsheetNotFound, KeyError) as e:
            self.log.info(f'Spreadsheet not found by {title if title else key if key else url}: {e}')
            if not ignore_errors:
                raise(KeyError('Spreadsheet not found'))
        return self

    def set_sheet(self, *, title: str=None, key: str=None, url: str=None) -> 'GoogleSheets':
        """Activates the current Google Spreadsheet that subsequent actions will be performed on
                 At least one parameter must be specified. If more that one is specified, then only one will be used
                 https://pygsheets.readthedocs.io/en/stable/reference.html#pygsheets.Client.open
                 Sets the worksheet to index: 0 by default

        Args:
            title (str, optional): The title of the spreadsheet to activate
            key (str, optional): The key ID of the spreadsheet to activate
            url (str, optional): The URL of the spreadsheet to activate

        Returns:
            A copy of the current object; this allows call chaining

        Raises:
            TypeError: If no keyword arguments are specified
            KeyError: If the specified Spreadsheet does not exist
        """
        if not any([title is not None, key is not None, url is not None]):
            raise(ValueError('At least one of `title`, `key`, or `url` keyword arguments must be specified'))
        try:
            if title:
                self.sheet = self.client.open(title)
            elif key:
                self.sheet = self.client.open_by_key(key)
            elif url:
                self.sheet = self.client.open_by_url(url)
        except (pygsheets.SpreadsheetNotFound, IndexError) as e:
            self.log.info(f'Spreadsheet not found by {title if title else key if key else url}: {e}')
            raise(KeyError('Spreadsheet not found'))
        self.set_ws(index=0)
        return self

    def set_or_create_sheet(self, title: str) -> 'GoogleSheets':
        """Activates the Google Spreadsheet by title if it exists, otherwise create it

        Args:
            title (str): The title of the Google Spreadsheet to set or create

        Returns:
            A copy of the current object; this allows call chaining
        """
        if title in self.list_sheets():
            self.set_sheet(title=title)
        else:
            self.create_sheet(title)
        return self

    def list_ws(self) -> List[pygsheets.Worksheet]:
        """Lists all worksheets in the active Google Spreadsheet
                https://pygsheets.readthedocs.io/en/stable/spreadsheet.html#pygsheets.Spreadsheet.worksheets

        Returns:
            List of all pygsheets Worksheet objects in active Spreadsheet

        Raises:
            TypeError: If a Spreadsheet has not been activated
        """
        if not self.sheet:
            raise(TypeError('You must activate a Spreadsheet before listing worksheets'))
        return self.sheet.worksheets()

    def set_ws(self, *, title: str=None, index: int=None, ws_id: str=None) -> 'GoogleSheets':
        """Activates the worksheet that subsequent actions will be performed on
                At least one parameter must be specified. If more that one is specified, then only one will be used
                https://pygsheets.readthedocs.io/en/stable/spreadsheet.html#pygsheets.Spreadsheet.worksheet

        Args:
            title (str, optional): The title of the worksheet to activate
            index (int, optional): The index of the worksheet to activate
            ws_id (str, optional): The ID of the worksheet to activate

        Returns:
            A copy of the current object; this allows call chaining

        Raises:
            ValueError: If a Spreadsheet has not been activated
            TypeError: If no keyword arguments are specified
            KeyError: If the specified worksheet does not exist
        """
        if not self.sheet:
            raise(ValueError('You must activate a Spreadsheet before activating a worksheet'))
        try:
            if title is not None:
                self.ws = self.sheet.worksheets('title', title, force_fetch=True)[0]
            elif index is not None:
                self.ws = self.sheet.worksheets('index', index, force_fetch=True)[0]
            elif ws_id is not None:
                self.ws = self.sheet.worksheets('id', ws_id, force_fetch=True)[0]
        except pygsheets.WorksheetNotFound as e:
            self.log.info(f'Worksheet not found by {title if title is not None else index if index is not None else ws_id}: {e}')
            raise(KeyError('Worksheet not found'))
        return self

    def create_ws(self, title: str) -> 'GoogleSheets':
        """Creates and activates a worksheet within the active Google Spreadsheet,
                https://pygsheets.readthedocs.io/en/stable/spreadsheet.html#pygsheets.Spreadsheet.add_worksheet

        Args:
            title (str): Title of the worksheet to create

        Returns:
            A copy of the current object; this allows call chaining

        Raises:
            TypeError: If a Spreadsheet has not been activated
        """
        if not self.sheet:
            raise(TypeError('You must activate a Spreadsheet before activating a worksheet'))
        self.ws = self.sheet.add_worksheet(title)
        return self

    def set_or_create_ws(self, title: str) -> 'GoogleSheets':
        """Activates the worksheet by title if it exists, otherwise create it

        Args:
            title (str): The title of the worksheet to set or create

        Returns:
            A copy of the current object; this allows call chaining

        Raises:
            TypeError: If a Spreadsheet has not been activated
        """
        if not self.sheet:
            raise(TypeError('You must activate a Spreadsheet before activating/creating a worksheet'))
        if title in [ws.title for ws in self.list_ws()]:
            self.set_ws(title=title)
        else:
            self.create_ws(title)
        return self

    def delete_ws(self, ws_id: str) -> 'GoogleSheets':
        """Deletes the specified worksheet by ID only

        Args:
            ws_id (str): The worksheet ID to delete

        Returns:
            A copy of the current object; this allows call chaining

        Raises:
            TypeError: If a Spreadsheet has not been activated
            KeyError: If the specified worksheet does not exist
        """
        if not self.ws:
            raise(TypeError('You must activate a Spreadsheet before deleting a worksheet'))
        try:
            if self.ws.id == ws_id:
                self.sheet.del_worksheet(self.ws)
                self.ws = None
            else:
                current_ws = self.ws
                self.sheet.del_worksheet(self.set_ws(ws_id=ws_id).ws)
                self.ws = current_ws
        except pygsheets.WorksheetNotFound as e:
            self.log.info(f'Worksheet not found by {ws_id}: {e}')
            raise(KeyError('Worksheet not found'))
        return self

    def get_row(self, index: int) -> List:
        """Get all values in row `index` from the active worksheet
                https://pygsheets.readthedocs.io/en/stable/worksheet.html#pygsheets.Worksheet.get_row

        Args:
            index (int): The integer index of the row to access

        Returns:
            List of values from row

        Raises:
            TypeError: If a worksheet has not been activated
        """
        if not self.ws:
            raise(TypeError('You must activate a worksheet before getting a row'))
        return self.ws.get_row(index, include_tailing_empty=False)

    def get_column(self, index: int) -> List:
        """Get all values in column `index` from the active worksheet
                https://pygsheets.readthedocs.io/en/stable/worksheet.html#pygsheets.Worksheet.get_col

        Args:
            index (int): The integer index of the column to access

        Returns:
            List of values from column

        Raises:
            TypeError: If a worksheet has not been activated
        """
        if not self.ws:
            raise(TypeError('You must activate a worksheet before getting a column'))
        return self.ws.get_col(index, include_tailing_empty=False)

    def _last_dimension(self, dimension: str) -> int:
        """Provides the index of last non-empty row or column

        Args:
            dimension (str): String representation of whether to find the last row or column

        Returns:
            Integer row or column index of last non-empty row or column
        """
        # Get all non-empty values from row or column 1, assume that the length == # of rows or columns
        # (It's backwards, list the values in the opposite index to find length of desired index)
        if dimension == 'ROWS':
            return len(self.get_column(1))
        elif dimension == 'COLUMNS':
            return len(self.get_row(1))
        else:
            raise(ValueError('You must specify the dimension, either `ROWS` or `COLUMNS`'))

    def update_row_by_index(self, values: List[List], row_offset: int, column_offset: int=1) -> bool:
        """Update values in a row based on the numerical index. (row_offset, column_offset) specifies the starting
            point in the worksheet for where to start updating values. Will add additional rows to fit provided data

        Args:
            values (List of lists): Adds values in order at the specified `row_offset`
                                    Additional lists in list will update the next row(s)
            row_offset (int): Adds row at specified position in the worksheet
            column_offset (int, optional): Starts the row at specified column, defaults to first column (far left)

        Returns:
            True on successful update; false otherwise

        Raises:
            TypeError: If a worksheet has not been activated
        """
        if not self.ws:
            raise(TypeError(f'You must activate a worksheet before updating row values'))

        try:
            if not isinstance(values[0], list):
                values = [values]
            self.ws.update_values(crange=(row_offset, column_offset), values=values, extend=True, majordim='ROWS')
            return True
        except Exception:
            return False

    def update_column_by_index(self, values: List[List], column_offset: int, row_offset: int=1) -> bool:
        """Update values in a column based on the numerical index. (row_offset, column_offset) specifies the starting
            point in the worksheet for where to start updating values. Will add additional columns to fit provided data

        Args:
            values (List of lists): Adds values in order at the specified `column_offset`
                                    Additional lists in list will update the next column(s)
            column_offset (int): Adds column at specified position in the worksheet
            row_offset (int, optional): Starts the column at specified row, defaults to first row (top)

        Returns:
            True on successful update; false otherwise

        Raises:
            TypeError: If a worksheet has not been activated
        """
        if not self.ws:
            raise(TypeError(f'You must activate a worksheet before updating column values'))

        try:
            if not isinstance(values[0], list):
                values = [values]
            self.ws.update_values(crange=(row_offset, column_offset), values=values, extend=True, majordim='COLUMNS')
            return True
        except Exception:
            return False

    def update_row_by_header(self, values: List[Dict], row_offset: int, header_row: int=1,
                             case_sensitive: bool=True) -> bool:
        """Update values in a row based on the specified header keys. Will add additional rows to fit provided data

        Args:
            values (List of dicts): Adds values starting at the `row_offset` based on position of matching key in `header_row`
                                    Additional dicts in list will update the next row(s)
            row_offset (int): Adds row at specified position in the worksheet
            header_row (int, optional): Specifies which row to use for matching headers, defaults to the first row
            case_sensitive (bool, optional): Specifies whether the header-key matching is case sensitive, it is by default

        Returns:
            True on successful update; false otherwise

        Raises:
            TypeError: If a worksheet has not been activated
        """
        if not self.ws:
            raise(TypeError(f'You must activate a worksheet before updating row values'))
        return self._update_dimension_by_header('ROWS', values, row_offset, header_row, case_sensitive)

    def update_column_by_header(self, values: List[Dict], column_offset: int, header_column: int=1,
                                case_sensitive: bool=True) -> bool:
        """Update values in a column based on the specified header keys. Will add additional columns to fit provided data

        .. seealso:: blabla

        VersionAdded:
            1.2

        Args:
            values (List of dicts): Adds values starting at the `column_offset` based on position
                of matching key in `header_column`. Additional dicts in list will update the next column(s)
            column_offset (int): Adds column at specified position in the worksheet
            header_column (int, optional): Specifies which column to use for matching headers, defaults to the first column
            case_sensitive (bool, optional): Specifies whether the header-key matching is case sensitive, it is by default

        Returns:
            bool: True on successful update; false otherwise

        Raises:
            TypeError: If a worksheet has not been activated
        """
        if not self.ws:
            raise(TypeError(f'You must activate a worksheet before updating column values'))
        return self._update_dimension_by_header('COLUMNS', values, column_offset, header_column, case_sensitive)

    def add_data_to_ws_rows(self, worksheet: str, data: List[Dict], preserve_blanks: bool=False) -> str:
        """Add data to an existing or new worksheet by header value

        Args:
            worksheet (str): The name of the worksheet to add to or create
            data (List of dicts): Adds values to specified `worksheet`
            preserve_blanks (bool, optional): Whether or not to preserve empty strings ('') when exporting data

        Returns:
           str: `<worksheet>!<starting range>:<ending range>`
        """
        self.set_or_create_ws(worksheet)
        height = len(self.get_column(1))
        height = 2 if height == 0 else height + 1
        if preserve_blanks:
            for i, d in enumerate(data):
                for k, v in d.items():
                    if v == '':
                        data[i][k] = '<blank>'
        success = self.update_row_by_header(data, height)
        if success:
            end_range = self.format_addr((height + len(data) - 1, len(self.get_row(1))))
            return f'{worksheet}!A{height}:{end_range}'
        else:
            return ''

    def get_data_from_ws_range(self, ws_range: Union[str, Pattern]) -> List[Dict]:
        """Recursively retrieves data from other ranges in other worksheets of the same Spreadsheet

        Args:
            ws_range (str or Pattern): A `<worksheet>!<starting range>:<ending range>` formatted range string or a `re` regex
                                parsed object

        Returns:
            List of dicts with header values as the key
        """
        if isinstance(ws_range, str):
            ws_range = self.ws_range_format.match(ws_range)
        if ws_range is None:
            return []
        previous_ws = self.ws.title
        self.set_ws(title=ws_range.group('worksheet'))
        starting_point = self.format_addr(ws_range.group('start_range'))
        ending_point = self.format_addr(ws_range.group('end_range'))
        headers = self.get_row(1)
        range_list = []
        for row in range(starting_point[0], ending_point[0] + 1):
            row_list = self.get_row(row)
            row_dict = {}
            for column, cell in enumerate(row_list):
                match = self.ws_range_format.match(cell)
                if match is not None:
                    row_list[column] = self.get_data_from_ws_range(match)
                else:
                    converted_cell = self._convert_cell_str_to_python(cell)
                    row_list[column] = converted_cell
                if cell == '<blank>' or row_list[column] not in ['', [], {}, ()]:
                    row_dict[headers[column]] = row_list[column]
            if row_dict:
                range_list.append(row_dict)
        self.set_ws(title=previous_ws)
        return range_list

    def _convert_cell_str_to_python(self, cell_str: str) -> object:
        """Takes a string retrieved from a worksheet cell and converts it to the matching python object

        Args:
            cell_str (str): The string value retrieved from a worksheet cell

        Returns:
            The converted python object, returns an empty string ('') if the value is blank
        """
        python_obj = ''
        if cell_str == 'TRUE':
            python_obj = True
        elif cell_str == 'FALSE':
            python_obj = False
        elif cell_str == 'None':
            python_obj = None
        elif cell_str == '<blank>':
            python_obj = ''
        elif cell_str and cell_str[0] == '[' and cell_str[-1] == ']':
            cell_list = cell_str[1:][:-1].split(',')
            if cell_list:
                for index, item in enumerate(cell_list):
                    cell_list[index] = self._convert_cell_str_to_python(item)
                # prune empty objects
                cell_list = [item for item in cell_list if item not in ['', [], {}, ()]]
                if cell_list:
                    python_obj = cell_list
        else:
            try:
                python_obj = int(cell_str)
            except (TypeError, ValueError):
                python_obj = cell_str
        return python_obj

    def _update_dimension_by_header(self, dimension: str, values: List[Dict], dimension_offset: int, header_index: int,
                                    case_sensitive: bool=True) -> bool:
        """Update values in either a row or a column

        Args:
            dimension (str): String representation of whether to update values in `ROWS` or `COLUMNS`
            values (List of dicts): Adds values starting at the `dimension_offset` based on position of matching key
                                        in `header_index`
                                    Additional dicts in list will update the next dimension(s)
                                    Missing headers will be added at the end of `dimension`
                                    A value of `None` within the list will cause that dimension to be skipped
            dimension_offset (int, optional): Adds dimension at specified position in the worksheet
            header_index (int, optional): Specifies what dimension index to use as headers when updating values
            case_sensitive (bool, optional): Specifies whether the header-key matching is case sensitive, it is by default

        Returns:
            True on successful update; false otherwise
        """
        if dimension == 'ROWS':
            headers = self.get_row(header_index)
            update_range = (dimension_offset, 1)
        elif dimension == 'COLUMNS':
            headers = self.get_column(header_index)
            update_range = (1, dimension_offset)

        if not case_sensitive:
            headers = [header.lower() for header in headers]

        value_matrix = []
        for item in values:
            ordered_items = {}
            if item is None:
                value_matrix.append([])
                continue
            # Match keys to header index
            for key, value in item.items():
                try:
                    if not case_sensitive:
                        cell_index = headers.index(key.lower()) + 1
                    else:
                        cell_index = headers.index(key) + 1
                except ValueError:
                    # Add missing header to `header_index` dimension
                    headers.append(key)
                    if dimension == 'ROWS':
                        if len(headers) > self.ws.cols:
                            self.add_column()
                        self.ws.update_value((header_index, len(headers)), key)
                    elif dimension == 'COLUMNS':
                        if len(headers) > self.ws.rows:
                            self.add_row()
                        self.ws.update_value((len(headers), header_index), key)
                    cell_index = len(headers)
                # Save key for updating the headers later
                ordered_items[cell_index] = str(value)

            # Add blank values for any missing indices
            for index in range(1, max(ordered_items)):
                if index not in ordered_items:
                    ordered_items[index] = None
            value_matrix.append([ordered_items[key] for key in sorted(ordered_items.keys())])

        try:
            self.ws.update_values(crange=update_range, values=value_matrix, extend=True, majordim=dimension)
            return True
        except Exception:
            return False

    def add_row(self, at_row: int=-1) -> 'GoogleSheets':
        """Adds a row to the active worksheet at position `at_row`
                https://pygsheets.readthedocs.io/en/stable/worksheet.html#pygsheets.Worksheet.insert_rows

        Args:
            at_row (int, optional): Adds row at specified position in the worksheet, -1 (default) specifies end
                of active rows

        Returns:
            A copy of the current object; this allows call chaining

        Raises:
            TypeError: If a worksheet has not been activated
        """
        if not self.ws:
            raise(TypeError('You must activate a worksheet before adding a row'))
        if at_row < 0:
            at_row = self.ws.rows
        else:
            at_row -= 1
        self.ws.insert_rows(at_row, inherit=True)
        return self

    def add_column(self, at_column: int=-1) -> 'GoogleSheets':
        """Adds a columns to the active worksheet at position `at_column`
                https://pygsheets.readthedocs.io/en/stable/worksheet.html#pygsheets.Worksheet.insert_cols

        Args:
            at_column (int, optional): Adds column at specified position in the worksheet, -1 (default) specifies end
                of active columns

        Returns:
            A copy of the current object; this allows call chaining

        Raises:
            TypeError: If a worksheet has not been activated
        """
        if not self.ws:
            raise(TypeError('You must activate a worksheet before adding a column'))
        if at_column < 0:
            at_column = self.ws.cols
        else:
            at_column -= 1
        self.ws.insert_cols(at_column, inherit=True)
        return self

    def find_cells(self, value, match_case: bool=True, match_entire_cell: bool=True) -> List[pygsheets.Cell]:
        """Finds all cells that contains `value` across all worksheets
                https://pygsheets.readthedocs.io/en/stable/worksheet.html#pygsheets.Worksheet.find

        Args:
            value: Value to find in active worksheet (supports a compiled regular expression)
            match_case (bool, optional): Whether or not match cells based on case. Default is case sensitive
            match_entire_cell (bool, optional): Whether or not match the entire value of the cell. Default is full cell matching

        Returns:
            A list of all pygsheets Cell objects that contains `value`
                https://pygsheets.readthedocs.io/en/stable/cell.html
        """
        cells = []
        for worksheet in self.list_ws():
            cells += worksheet.find(value, matchCase=match_case, matchEntireCell=match_entire_cell)
        return cells

    def replace_value(self, find_value: str, replacement: str) -> 'GoogleSheets':
        """Replace the `find_value` with `replacement` across all worksheets. This performs only a replacement on
                the specified `find_value`, leaving the rest of the cell contents intact.
                https://pygsheets.readthedocs.io/en/stable/worksheet.html#pygsheets.Worksheet.find

        Args:
            find_value: The string or regex value to find within the worksheet
            replacement (str): The string used to replace all found values

        Returns:
            A copy of the current object; this allows call chaining
        """
        for worksheet in self.list_ws():
            worksheet.replace(str(find_value), str(replacement))
        return self
