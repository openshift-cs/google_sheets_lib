from pytest import mark, raises
from google_sheets_lib import GoogleSheets  # noqa F401


def test_failed_googlesheets_worksheet_functions(gs_client: 'GoogleSheets'):
    with raises(TypeError):
        gs_client.list_ws()
        gs_client.set_ws()
        gs_client.create_ws('')


def test_failed_googlesheets_set_ws(gs_client: 'GoogleSheets'):
    with raises(ValueError):
        gs_client.set_ws()


def test_failed_googlesheets_set_sheet(gs_client: 'GoogleSheets'):
    with raises(ValueError):
        gs_client.set_sheet()
    with raises(KeyError):
        gs_client.set_sheet(title='1')


def test_failed_googlesheets_delete_sheet(gs_client: 'GoogleSheets'):
    with raises(KeyError):
        gs_client.delete_sheet(title='1')


@mark.dependency()
def test_successful_googlesheets_initiation(gs_client: 'GoogleSheets'):
    assert gs_client is not None
    [gs_client.delete_sheet(title=sheet) for sheet in gs_client.list_sheets()]


@mark.dependency(depends=['test_successful_googlesheets_initiation'])
def test_successful_googlesheets_list_sheets(gs_client: 'GoogleSheets'):
    assert len(gs_client.list_sheets()) == 0


@mark.dependency(depends=['test_successful_googlesheets_initiation'])
def test_successful_googlesheets_set_or_create_sheet(gs_client: 'GoogleSheets'):
    assert gs_client.sheet is None
    assert len(gs_client.list_sheets()) == 0
    assert len(gs_client.set_or_create_sheet('My Test Sheet').list_sheets()) == 1
    assert gs_client.sheet is not None
    assert gs_client.sheet.title == 'My Test Sheet'
    assert len(gs_client.set_or_create_sheet('My Test Sheet').list_sheets()) == 1
    assert gs_client.sheet.title == 'My Test Sheet'


@mark.dependency(depends=['test_successful_googlesheets_set_or_create_sheet'])
def test_successful_googlesheets_create_sheet(gs_client: 'GoogleSheets'):
    assert gs_client.sheet is None
    assert len(gs_client.create_sheet('New Sheet').list_sheets()) == 2
    assert gs_client.sheet is not None


@mark.dependency(depends=['test_successful_googlesheets_create_sheet'])
def test_successful_googlesheets_set_sheet(gs_client: 'GoogleSheets'):
    assert gs_client.sheet is None
    gs_client.set_sheet(title='New Sheet')
    assert gs_client.sheet is not None


@mark.dependency(depends=['test_successful_googlesheets_create_sheet'])
def test_successful_googlesheets_list_ws(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    # 1 worksheet by default on new spreadsheets
    ws_list = gs_client.list_ws()
    assert len(ws_list) == 1
    assert ws_list[0].title == 'Sheet1'
    assert 'Sheet1' in [ws.title for ws in ws_list]


@mark.dependency(depends=['test_successful_googlesheets_create_sheet'])
def test_successful_googlesheets_set_ws(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    # The default worksheet is Sheet1
    gs_client.set_ws(title='Sheet1')
    assert gs_client.ws is not None
    gs_client.set_ws(index=0)
    assert gs_client.ws.title == 'Sheet1'


@mark.dependency(depends=['test_successful_googlesheets_create_sheet'])
def test_successful_googlesheets_create_ws(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    assert len(gs_client.create_ws('My New WS').list_ws()) == 2
    assert gs_client.ws.title == 'My New WS'


@mark.dependency(depends=['test_successful_googlesheets_create_ws'])
def test_successful_googlesheets_set_or_create_ws(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    assert len(gs_client.list_ws()) == 2
    assert len(gs_client.set_or_create_ws('Created ws').list_ws()) == 3
    assert gs_client.ws.title == 'Created ws'
    gs_client.set_ws(index=0)
    assert gs_client.ws.title == 'Sheet1'
    assert len(gs_client.set_or_create_ws('Created ws').list_ws()) == 3
    assert gs_client.ws.title == 'Created ws'


@mark.dependency(depends=['test_successful_googlesheets_set_or_create_ws'])
def test_successful_googlesheets_delete_ws(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    ws_len = len(gs_client.list_ws())
    gs_client.delete_ws(gs_client.ws.id)
    assert len(gs_client.list_ws()) == ws_len - 1
    assert gs_client.ws is None


@mark.dependency(depends=['test_successful_googlesheets_delete_ws'])
def test_successful_googlesheets_add_row(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    num_rows = gs_client.ws.rows
    gs_client.add_row()
    assert gs_client.ws.rows == num_rows + 1


@mark.dependency(depends=['test_successful_googlesheets_add_row'])
def test_successful_googlesheets_add_column(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    num_cols = gs_client.ws.cols
    gs_client.add_column()
    assert gs_client.ws.cols == num_cols + 1


@mark.dependency(depends=['test_successful_googlesheets_add_row'])
def test_successful_googlesheets__last_dimension(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    # _last_dimension returns the last row/column that has data, on an empty sheet, this should be 0 for both
    assert gs_client._last_dimension('ROWS') == 0
    assert gs_client._last_dimension('COLUMNS') == 0


@mark.dependency(depends=['test_successful_googlesheets__last_dimension'])
def test_successful_googlesheets_update_row_by_index(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    row_1_values = ['colA', 'colB', None, 'colD']
    assert gs_client.update_row_by_index([row_1_values], 1) is True
    row_1 = gs_client.get_row(1)
    assert row_1 == ['' if val is None else str(val) for val in row_1_values]
    row_2_values = ['Hello', 'you', 'Crazy', 'person']
    assert gs_client.update_row_by_index([row_2_values], 2) is True
    row_2 = gs_client.get_row(2)
    assert row_2 == ['' if val is None else str(val) for val in row_2_values]


@mark.dependency(depends=['test_successful_googlesheets__last_dimension'])
def test_successful_googlesheets_update_column_by_index(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    col_1_values = ['row3', 'row4', None, 'row6']
    assert gs_client.update_column_by_index([col_1_values], 1, row_offset=3) is True
    col_1 = gs_client.get_column(1)[2:]
    assert col_1 == ['' if val is None else str(val) for val in col_1_values]
    col_2_values = ['Howdy', 'my', 'Sane', 'padre']
    assert gs_client.update_column_by_index([col_2_values], 2, row_offset=3) is True
    col_2 = gs_client.get_column(2)[2:]
    assert col_2 == ['' if val is None else str(val) for val in col_2_values]


@mark.dependency(depends=['test_successful_googlesheets_update_column_by_index'])
def test_successful_googlesheets_update_row_by_header(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    row_3_values = {'colA': "I'm below 'Hello'", 'colB': "I'm below 'you'", 'colE': "I'm a new row header"}
    assert gs_client.update_row_by_header([row_3_values], 3) is True
    row_3 = gs_client.get_row(3)
    assert row_3 == ["I'm below 'Hello'", "I'm below 'you'", '', '', "I'm a new row header"]
    # Purposefully changed casing
    row_4_values = {'hello': 'row4 colA', 'You': 'row4 colB', 'crazy': 'row4 colC', 'Person': 'row4 colD'}
    assert gs_client.update_row_by_header([row_4_values], 4, header_row=2, case_sensitive=False) is True
    row_4 = gs_client.get_row(4)
    assert row_4 == ['row4 colA', 'row4 colB', 'row4 colC', 'row4 colD']


@mark.dependency(depends=['test_successful_googlesheets_update_row_by_header'])
def test_successful_googlesheets_update_column_by_header(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    col_3_values = {"I'm below 'Hello'": "I'm next to \"I'm below 'you'\"",
                    'row4 colA': "I'm next to 'row4 colB'", 'row7': "I'm a new column header"}
    assert gs_client.update_column_by_header([col_3_values], 3) is True
    col_3 = gs_client.get_column(3)
    assert col_3 == ['', 'Crazy', "I'm next to \"I'm below 'you'\"", "I'm next to 'row4 colB'", '', '', "I'm a new column header"]
    col_4_values = {'sane': 'sane', 'Padre': 'Padre', 'ROW7': "I'm next to \"I'm a new column header\""}
    assert gs_client.update_column_by_header([col_4_values], 4, header_column=2, case_sensitive=False) is True
    col_4 = gs_client.get_column(4)[4:]
    assert col_4 == ['sane', 'Padre', "I'm next to \"I'm a new column header\""]
    col_5_values = {'row6': 'row6'}
    assert gs_client.update_column_by_header([col_5_values], 5) is True
    col5 = gs_client.get_column(5)
    assert col5 == ['colE', '', "I'm a new row header", '', '', 'row6']
    # Testing updates when grouped together, and a skip
    assert gs_client.update_column_by_header([col_3_values, None, col_5_values], 3) is True
    assert gs_client.get_column(3) == ['', 'Crazy', "I'm next to \"I'm below 'you'\"",
                                       "I'm next to 'row4 colB'", '', '', "I'm a new column header"]
    assert gs_client.get_column(5) == ['colE', '', "I'm a new row header", '', '', 'row6']


@mark.dependency(depends=['test_successful_googlesheets_update_column_by_header'])
def test_successful_googlesheets_add_data_to_ws_row(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')
    data = [
        {'colA': 'a', 'colB': 'b', 'colC': 'c'},
        {'colA': 1, 'colB': 2, 'colD': 4}
    ]
    updated_range = gs_client.add_data_to_ws_rows('Test WS', data)
    assert updated_range == 'Test WS!A2:D3'
    data = [
        {'colA': 'a', 'colB': 'b', 'colC': 'c', 'colE': 'e'},
        {'colA': 1, 'colB': 2, 'colD': 4}
    ]
    updated_range = gs_client.add_data_to_ws_rows('Test WS', data)
    assert updated_range == 'Test WS!A4:E5'


@mark.dependency(depends=['test_successful_googlesheets_update_column_by_header'])
def test_successful_googlesheets_find_cell(gs_client: 'GoogleSheets'):
    gs_client.set_sheet(title='New Sheet')

    cell1 = gs_client.find_cells('ROW7')
    cell2 = gs_client.find_cells("I'm next to \"I'm below 'you'\"")
    cell3 = gs_client.find_cells('sane')
    cell4 = gs_client.find_cells('you')
    cell5 = gs_client.find_cells('ROW7', match_case=False)
    cell6 = gs_client.find_cells('row6')
    cell7 = gs_client.find_cells('Not found')

    assert len(cell1) == 1
    assert cell1[0].label == 'B7'
    assert len(cell2) == 1
    assert cell2[0].label == 'C3'
    assert len(cell3) == 1
    assert cell3[0].label == 'D5'
    assert len(cell4) == 1
    assert cell4[0].label == 'B2'
    assert len(cell5) == 2
    assert cell5[0].label == 'A7'
    assert cell5[1].label == 'B7'
    # This should be found in 2 cells
    assert len(cell6) == 2
    assert cell6[0].label == 'A6'
    assert cell6[1].label == 'E6'
    # This shouldn't be found
    assert len(cell7) == 0
