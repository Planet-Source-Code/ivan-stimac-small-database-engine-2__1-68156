Attribute VB_Name = "modSqlDelete"
Option Explicit
'
Public Function ExecuteSqlDelete(ByRef strSql As String, ByRef ffDB As Integer, ByRef mtbl_params() As tbl_params, ByVal db_file As String, ByVal db_access As String) As Long
    'On Error Resume Next
    Dim i As Long
    Dim tbl_ind As Integer
    Dim mRow As row_data
    Dim mRecSet As New clsRecordset
    Dim mDB As New clsDatabaseEngine
    
    tbl_ind = getTblIndex(strSql, mtbl_params)
    If tbl_ind < 0 Then
        Err.Raise -20, , "Invalid table name!"
    End If
    Err.Clear
    
    i = InStr(1, strSql, "FROM")
    'rewrite DELETE sql to SELECT sql
    strSql = "SELECT * " & Mid(strSql, i)
    'open database with new DatabaseEngine object
    If mDB.OpenDB(db_file, db_access) = 0 Then
        'get records with recordset
        Set mRecSet = mDB.OpenRecordSet(strSql)
        'set recodset property isForDelete as true, if it's true
        '   then recordset wil not parse columns, it will only read
        '   records position in file
        mRecSet.isForDelete = True
        '
        'if there is no records then exit
        If mRecSet.Rows = 0 Then
            mDB.CloseDB
            Set mDB = Nothing
            Set mRecSet = Nothing
            Exit Function
        End If
        'if there is records to delete
        mRecSet.MoveFirst
        For i = 1 To mRecSet.Rows
            'firs read each record
            Get #ffDB, mRecSet.RecordPosition, mRow
            'change row_using to '*', if it's '+' then row is
            '   not deleted, else it is
            mRow.row_using = "*"
            'then rewrite
            Put #ffDB, mRecSet.RecordPosition, mRow
            'move to next record
            mRecSet.MoveNext
        Next i
        
        'If Err.Number = 0 Then
        'refresh tbl def
        mtbl_params(tbl_ind).rows_cnt = mtbl_params(tbl_ind).rows_cnt - mRecSet.Rows
        mDB.setRowsCnt tbl_ind, mtbl_params(tbl_ind).rows_cnt
        mDB.RefreshDB
       ' End If
    End If
    'close database
    mDB.CloseDB
    'free memory
    Set mDB = Nothing
    Set mRecSet = Nothing
    'if there si some errors
'    If Err.Number <> 0 Then
'        ExecuteSqlDelete = -2
'        Err.Raise -2, , "Fail to open/write in file!"
'    End If
End Function

'''' function messages:
''''   -1 : there is no selected DB
''''   -2 : can't execute sql (invalid column name detected)
''''   -3 : fail to execute
''''    0 : execute success
'''Private Sub PrivateSqlSelect(ByRef strSql As String, ByRef ffDB As Integer, ByRef mtbl_params() As tbl_params, ByRef tbl_index As Integer)
'''        Dim colIndex_arr() As Integer
'''        Dim val_arr() As String, operator_arr() As String, condOper_arr() As String
'''        Dim msel_cols() As Integer
'''        Dim rec_pass() As Boolean, tmp_pass As Boolean, is_pass As Boolean
'''
'''        '
'''        Dim str_order As String
'''        Dim usingWhere As Boolean
'''        Dim f_ret As Integer
'''        Dim tmp_recSet As New clsRecSet ', tmp_rec2 As New clsRecSet
'''        Dim prev As Long, row_count As Long, curr_col As Integer
'''        Dim mRow As row_data
'''        Dim str_data() As String
'''        Dim i As Long, j As Long, c As Long, k As Integer, lastS As Integer, z As Integer
'''        Dim real_rowCnt As Long
'''
'''        Dim tmp_rows() As Long, tmp_rows2() As Long
'''
'''
'''
'''        'find table index
'''        tbl_index = getTblIndex(strSql, mtbl_params)
'''        'first check is WHERE statment correct or exists
'''        f_ret = parseSqlWhere(strSql, mtbl_params, tbl_index, colIndex_arr, val_arr, operator_arr, condOper_arr, str_order)
'''        usingWhere = False
'''        If f_ret = 0 Then
'''            usingWhere = True
'''        ElseIf f_ret = -2 Then
'''            Exit Sub
'''        End If
'''
'''        If tbl_index < 0 Then
'''            Err.Raise -20, , "Invalid table name detected!"
'''        End If
'''
'''        ReDim msel_cols(mtbl_params(tbl_index).col_count - 1)
'''        For i = 0 To mtbl_params(tbl_index).col_count - 1
'''            msel_cols(i) = i
'''        Next i
'''        'getSelectedCols strSql, mtbl_params, tbl_index, msel_cols
'''        '
'''        '
'''        'count rows
'''        prev = mtbl_params(tbl_index).last_position
'''        row_count = 0
'''        Do While prev > 0
'''            Get #ffDB, prev, mRow
'''            'if not deleted then
'''            If mRow.row_using = "+" Then
'''                row_count = row_count + 1
'''            End If
'''            prev = mRow.row_prev
'''        Loop
'''        'there is no data
'''        If row_count < 1 Then
'''            Exit Sub
'''        End If
'''        '
'''        'set arrays size
'''        ReDim str_data(row_count - 1)
'''        ReDim tmp_rows(row_count - 1)
'''        ReDim tmp_rows2(row_count - 1)
'''        'redim
'''        '
'''        'set recordset column count
'''        tmp_recSet.Columns = mtbl_params(tbl_index).col_count ' UBound(msel_cols) + 1
'''        tmp_recSet.Rows = row_count
'''
'''
'''        'tmp_rec2.Columns = mtbl_params(tbl_index).col_count 'UBound(msel_cols) + 1
'''        'tmp_rec2.Rows = row_count
'''        'now read to array
'''        row_count = 0
'''        prev = mtbl_params(tbl_index).last_position
'''        tmp_rows(0) = prev
'''        Do While prev > 0
'''            Get #ffDB, prev, mRow
'''            If mRow.row_using = "+" Then
'''                'If forDelete = True Then lngColl2(rowNum) = prev 'mColl2.Add prev
'''                str_data(row_count) = mRow.row_contain
'''                row_count = row_count + 1
'''                If row_count <= UBound(tmp_rows) Then tmp_rows(row_count) = mRow.row_prev
'''            End If
'''            prev = mRow.row_prev
'''        Loop
'''        'parse values and add them to tmp recordset
'''        For i = 0 To UBound(str_data)
'''            j = 0
'''            k = 1
'''            c = 0
'''            Do While k > 0
'''                'each value is in ' ', so we searching for '
'''                k = InStr(k, str_data(i), "'")
'''                If k > 0 Then
'''                    c = c + 1
'''                    If c = 2 Then
'''                        'when we find start and end of data, read data and save
'''                        tmp_recSet.Data(i, j) = Mid(str_data(i), lastS, k - lastS)
'''                        c = 0
'''                        j = j + 1
'''                    Else
'''                        lastS = k + 1
'''                    End If
'''                    k = k + 1
'''                End If
'''            Loop
'''        Next i
'''        'if we need to search for resuld that match query
'''        If usingWhere <> False Then
'''            'create array that contains for all columns in WHAT statment
'''            '   info about it, does current row match query
'''            ReDim rec_pass(UBound(colIndex_arr))
'''            'set column count of tmp recordset 2
'''            'we will use this variable for counting rows that
'''            '   match query
'''            row_count = 0
'''            'check first count of data that match query
'''            For i = 0 To tmp_recSet.Rows - 1
'''                curr_col = 0
'''                For j = 0 To UBound(colIndex_arr)
'''                    rec_pass(j) = False
'''                    'check values for current row with values in query
'''                    Select Case Trim(operator_arr(j))
'''                        Case "="
'''                            If tmp_recSet.Data(i, colIndex_arr(j)) = val_arr(j) Then rec_pass(j) = True
'''                        Case "<>"
'''                            If tmp_recSet.Data(i, colIndex_arr(j)) <> val_arr(j) Then rec_pass(j) = True
'''                        Case ">"
'''                            If Val(tmp_recSet.Data(i, colIndex_arr(j))) > Val(val_arr(j)) Then rec_pass(j) = True
'''                        Case "<"
'''                            If Val(tmp_recSet.Data(i, colIndex_arr(j))) < Val(val_arr(j)) Then rec_pass(j) = True
'''                    End Select
'''                Next j
'''                '
'''                tmp_pass = False
'''                is_pass = False
'''                For j = 0 To UBound(condOper_arr)
'''                    If condOper_arr(j) = "OR" Then
'''                        tmp_pass = True
'''                        Exit For
'''                    End If
'''                Next j
'''                is_pass = Not tmp_pass
'''                'check in array what match query
'''                For j = 0 To UBound(rec_pass)
'''                    If tmp_pass = True And rec_pass(j) = True Then
'''                        is_pass = True
'''                        Exit For
'''                    ElseIf tmp_pass = False And rec_pass(j) = False Then
'''                        is_pass = False
'''                        Exit For
'''                    End If
'''                Next j
'''                'save data if match
'''                If is_pass = True Then
'''                    tmp_rows2(row_count) = tmp_rows(i)
'''                    row_count = row_count + 1
'''                End If
'''            Next i
'''            'at the end fill recordset with data that match query
'''            ReDim sel_rows(row_count - 1)
'''            For i = 0 To row_count - 1
'''                sel_rows(i) = tmp_rows2(i)
'''            Next i
'''        End If
'''        'set function return
'''        'Set ExecuteSqlSelect = mRecSet
'''        'free memory
'''        Erase tmp_rows2
'''        Erase tmp_rows
'''        Erase colIndex_arr
'''        Erase val_arr
'''        Erase operator_arr
'''        Erase condOper_arr
'''        Erase msel_cols
'''        Erase str_data
'''        Erase rec_pass
'''End Sub
'''
'function find table index from sql query (find table name in query and search for
'   index)
Private Function getTblIndex(ByRef strSql As String, ByRef mtbl_params() As tbl_params) As Integer
    getTblIndex = -1
    On Error GoTo errH:
    Dim i As Integer, tmp_int As Integer, tmp_int2 As Integer
    Dim tmp_str As String

    tmp_int = InStr(1, StrConv(strSql, vbUpperCase), " FROM ")
    tmp_int2 = InStr(1, StrConv(strSql, vbUpperCase), " WHERE ")
    If tmp_int = 0 Then Exit Function

    tmp_int = tmp_int + Len(" FROM ")
    If tmp_int2 > tmp_int Then
        tmp_str = Trim(Mid(strSql, tmp_int, tmp_int2 - tmp_int))
    Else
        tmp_str = Trim(Mid(strSql, tmp_int))
    End If
    

    For i = 0 To UBound(mtbl_params)
        If tmp_str = mtbl_params(i).tbl_name Then
            getTblIndex = i
            Exit Function
        End If
    Next i
errH:
End Function

