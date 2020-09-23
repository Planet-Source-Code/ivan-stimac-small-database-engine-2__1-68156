Attribute VB_Name = "modGlobals"
Public Const dbIdent As String * 4 = "SDBE"
'contains table name, col count, col names...
Public Type tbl_params
    tbl_name As String      'table name
    col_count As Integer    'column count
    cols_arr() As String    'column names
    last_position As Long   'position of last row in file
    last_id As Long         'id of last row
    id_col As Long          'index of primary key column
    rows_cnt As Long        'count of rows in table
    indCols_arr() As Integer 'indexed rows array
    lastIndPos_arr() As Long     'position of last index (for search engine)
    step_small As Long      'last record in range of 1 000 records
    step_med As Long        '                       10 000
    step_big As Long        '                       50 000
End Type
'contains username and password for access to db
Public Type db_params
    db_userName As String * 10
    db_password As String * 10
End Type
'contains data about row (fields, where start previus row, is
'   row deleted)
Public Type row_data
    row_contain As String       'fields that row contains
    row_prev As Long            'position in file of previous row
    row_id As Long              'row id
    step_big As Long            'step for 50 000 records
    step_med As Long            'step for 10 000 records
    step_small As Long          'step for  1 000 records
    row_using As String * 1     'if + then row is not deleted, else it's deleted
End Type
'
'contains data of one column and all rows that contain this data
Public Type srch_indexes
    fld_data As String          'data in field
    fld_row As Long             'pointer to row
    prev_index As Long          'position in file of previous index
End Type
'
'contains pointers to indexes that contains data which same
'   first char
'Public Type ind_indexes
'    ind_char As String * 1      'char to look
'    ind_position As Long        'position of last index that contain data with first char ind_char
'    ind_prePos As Long          'position of previous char index
'End Type

