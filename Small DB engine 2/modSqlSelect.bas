Attribute VB_Name = "modSqlSelect"
Option Explicit
' function messages:
'   -1 : there is no selected DB
'   -2 : can't execute sql (invalid column name detected)
'   -3 : fail to execute
'    0 : execute success
Public Function ExecuteSqlSelect(ByRef strSql As String, ByRef ffDB As Integer, ByRef mtbl_params() As tbl_params) As clsRecordset  ' clsRecSet
        Dim tbl_index As Integer, colIndex_arr() As Integer
        Dim val_arr() As String, operator_arr() As String, condOper_arr() As String
        '
        Dim msel_cols() As Integer
        'Dim rec_pass() As Boolean, tmp_pass As Boolean, is_pass As Boolean
        'Dim for_save As Boolean
        '
        Dim str_order As String
        Dim usingWhere As Boolean
        Dim f_ret As Integer
       ' Dim mRecSet As New clsRecSet, tmp_recSet As New clsRecSet, tmp_rec2 As New clsRecSet
        'Dim prev As Long, row_count As Long, curr_col As Integer
       ' Dim mRow As row_data
       ' Dim str_data() As String
        'Dim i As Long, j As Long, c As Long, k As Integer, lastS As Integer, z As Integer
        
        'find table index
        tbl_index = getTblIndex(strSql, mtbl_params)
        'first check is WHERE statment correct or exists
        f_ret = parseSqlWhere(strSql, mtbl_params, tbl_index, colIndex_arr, val_arr, operator_arr, condOper_arr, str_order)
        usingWhere = False
        If f_ret = 0 Then
            usingWhere = True
        'if there is some invalid column names
        ElseIf f_ret = -30 Then
            Exit Function
        End If
        
        If tbl_index < 0 Then
            Err.Raise -20, , "Invalid table name detected!"
        End If
        
        getSelectedCols strSql, mtbl_params, tbl_index, msel_cols
        
        Dim tmp_records As New clsRecordset
        tmp_records.db_FileNumber = ffDB
        tmp_records.db_LastPosition = mtbl_params(tbl_index).last_position
        tmp_records.have_where = usingWhere

        tmp_records.set_arrs msel_cols, colIndex_arr, val_arr, operator_arr, condOper_arr, _
                    mtbl_params(tbl_index).indCols_arr, mtbl_params(tbl_index).lastIndPos_arr
        
        tmp_records.tbl_records = mtbl_params(tbl_index).rows_cnt
        'get result count
        tmp_records.result_count
        'set steps
'''''        tmp_records.db_StepBig = mtbl_params(tbl_index).step_big
'''''        tmp_records.db_StepMed = mtbl_params(tbl_index).step_med
'''''        tmp_records.db_StepSmall = mtbl_params(tbl_index).step_small
        '
        tmp_records.db_LastID = mtbl_params(tbl_index).last_id
        tmp_records.db_idCol = mtbl_params(tbl_index).id_col
        
        
        Set ExecuteSqlSelect = tmp_records
''''''''
''''''''        '
''''''''        tmp_recSet.Columns = mtbl_params(tbl_index).col_count
''''''''        tmp_recSet.Rows = 1
''''''''        mRecSet.Rows = 0
''''''''        'find last item position
''''''''        prev = mtbl_params(tbl_index).last_position
''''''''        row_count = 0
''''''''        'set str data array size to 500
''''''''        ReDim str_data(500)
''''''''        'loop until we find first item
''''''''        Do While prev > 0
''''''''            Get ffDB, prev, mRow
''''''''            'if not deleted then save it
''''''''            If mRow.row_using = "+" Then
''''''''                row_count = row_count + 1
''''''''                'increase array size if need
''''''''                If row_count > UBound(str_data) Then ReDim Preserve str_data(UBound(str_data) + 500)
''''''''                'ReDim Preserve str_data(row_count - 1)
''''''''                'save row fields
''''''''                str_data(row_count - 1) = mRow.row_contain
''''''''            End If
''''''''            prev = mRow.row_prev
''''''''        Loop
''''''''        'there is no data
''''''''        If row_count < 1 Then
''''''''            Exit Function
''''''''        End If
''''''''        '
''''''''        'set recordset column count
''''''''        tmp_recSet.Columns = mtbl_params(tbl_index).col_count ' UBound(msel_cols) + 1
''''''''        tmp_recSet.Rows = row_count
''''''''        tmp_rec2.Rows = row_count
''''''''        'parse values and add them to tmp recordset
''''''''        For i = 0 To row_count - 1 'UBound(str_data)
''''''''            j = 0
''''''''            k = 1
''''''''            c = 0
''''''''            Do While k > 0
''''''''                'each value is in ' ', so we searching for '
''''''''                k = InStr(k, str_data(i), "'")
''''''''                If k > 0 Then
''''''''                    c = c + 1
''''''''                    If c = 2 Then
''''''''                        'when we find start and end of data, read data and save
''''''''                        tmp_recSet.Data(i, j) = Mid(str_data(i), lastS, k - lastS)
''''''''                        c = 0
''''''''                        j = j + 1
''''''''                    Else
''''''''                        lastS = k + 1
''''''''                    End If
''''''''                    k = k + 1
''''''''                End If
''''''''            Loop
''''''''        Next i
''''''''        'if there is no WHERE statment sel all data to recordset
''''''''        If usingWhere = False Then
''''''''            'if we read all columns
''''''''            If UBound(msel_cols) = mtbl_params(tbl_index).col_count - 1 Then
''''''''                Set mRecSet = tmp_recSet
''''''''            'if not, then read only selected columns
''''''''            Else
''''''''                mRecSet.Columns = UBound(msel_cols) + 1
''''''''                mRecSet.Rows = tmp_recSet.Rows
''''''''                'read only column user list in query
''''''''                For i = 0 To tmp_recSet.Rows - 1
''''''''                    For j = 0 To UBound(msel_cols)
''''''''                        mRecSet.Data(i, j) = tmp_recSet.Data(i, msel_cols(j))
''''''''                    Next j
''''''''                Next i
''''''''            End If
''''''''        'if we need to search for resuld that match query
''''''''        Else
''''''''            'create array that contains for all columns in WHAT statment
''''''''            '   info about it, does current row match query
''''''''            ReDim rec_pass(UBound(colIndex_arr))
''''''''            'set column count of tmp recordset 2
''''''''            tmp_rec2.Columns = UBound(msel_cols) + 1
''''''''            'we will use this variable for counting rows that
''''''''            '   match query
''''''''            row_count = 0
''''''''            'check first count of data that match query
''''''''            For i = 0 To tmp_recSet.Rows - 1
''''''''                curr_col = 0
''''''''                For j = 0 To UBound(colIndex_arr)
''''''''                    rec_pass(j) = False
''''''''                    'check values for current row with values in query
''''''''                    Select Case Trim(operator_arr(j))
''''''''                        Case "="
''''''''                            If tmp_recSet.Data(i, colIndex_arr(j)) = val_arr(j) Then rec_pass(j) = True
''''''''                        Case "<>"
''''''''                            If tmp_recSet.Data(i, colIndex_arr(j)) <> val_arr(j) Then rec_pass(j) = True
''''''''                        Case ">"
''''''''                            If Val(tmp_recSet.Data(i, colIndex_arr(j))) > Val(val_arr(j)) Then rec_pass(j) = True
''''''''                        Case "<"
''''''''                            If Val(tmp_recSet.Data(i, colIndex_arr(j))) < Val(val_arr(j)) Then rec_pass(j) = True
''''''''                    End Select
''''''''                Next j
''''''''                '
''''''''                tmp_pass = False
''''''''                is_pass = False
''''''''                For j = 0 To UBound(condOper_arr)
''''''''                    If condOper_arr(j) = "OR" Then
''''''''                        tmp_pass = True
''''''''                        Exit For
''''''''                    End If
''''''''                Next j
''''''''                is_pass = Not tmp_pass
''''''''                'check in array what match query
''''''''                For j = 0 To UBound(rec_pass)
''''''''                    If tmp_pass = True And rec_pass(j) = True Then
''''''''                        is_pass = True
''''''''                        Exit For
''''''''                    ElseIf tmp_pass = False And rec_pass(j) = False Then
''''''''                        is_pass = False
''''''''                        Exit For
''''''''                    End If
''''''''                Next j
''''''''                'save data if match
''''''''                If is_pass = True Then
''''''''                    For j = 0 To tmp_recSet.Columns - 1
''''''''                        'check is if we need this column
''''''''                        For z = 0 To UBound(msel_cols)
''''''''                            If msel_cols(z) = j Then
''''''''                                tmp_rec2.Data(row_count, curr_col) = tmp_recSet.Data(i, j)
''''''''                                curr_col = curr_col + 1
''''''''                            End If
''''''''                        Next z
''''''''                    Next j
''''''''                    row_count = row_count + 1
''''''''                End If
''''''''            Next i
''''''''            'at the end fill recordset with data that match query
''''''''            mRecSet.Rows = row_count
''''''''            mRecSet.Columns = tmp_rec2.Columns
''''''''            For i = 0 To row_count - 1
''''''''                For j = 0 To tmp_rec2.Columns - 1
''''''''                    mRecSet.Data(i, j) = tmp_rec2.Data(i, j)
''''''''                Next j
''''''''            Next i
''''''''        End If
        'set function return
'        Set ExecuteSqlSelect = mRecSet
        'free memory
        Erase colIndex_arr
        Erase val_arr
        Erase operator_arr
        Erase condOper_arr
        Erase msel_cols
       ' Erase rec_pass
End Function
'read columns that we need to read data from them
Private Sub getSelectedCols(ByRef strSql As String, ByRef mtbl_params() As tbl_params, ByRef tbl_index As Integer, ByRef colIndex_arr() As Integer)
    Dim i As Integer, j As Integer, start_read As Integer, end_read As Integer
    Dim tmp_str As String, str_all As String
    Dim tmp_arr() As String
    
    If tbl_index < 0 Then Exit Sub
    
    start_read = InStr(1, strSql, "SELECT", vbTextCompare) + Len("SELECT")
    end_read = InStr(1, strSql, "FROM", vbTextCompare)
    'use only list of columns
    str_all = Mid(strSql, start_read, end_read - start_read)
    'then find column count
    i = InStrCharCount(str_all, ",")
    ReDim colIndex_arr(i)
    ReDim tmp_arr(i)
    'if there is only one selected column or * (all columns)
    If i = 0 Then
        'if it's all columns
        tmp_arr(0) = Trim(str_all)
        If Trim(str_all) = "*" Then
            ReDim colIndex_arr(mtbl_params(tbl_index).col_count - 1)
            'save all column indexes in array
            For j = 0 To mtbl_params(tbl_index).col_count - 1
                colIndex_arr(j) = j
            Next j
        Else
            'if only one column is selected by column name then
            '   find index of it
            For j = 0 To mtbl_params(tbl_index).col_count - 1
                If tmp_arr(0) = mtbl_params(tbl_index).cols_arr(j) Then
                    colIndex_arr(0) = j
                    Exit For
                End If
            Next j
        End If
        Erase tmp_arr
        Exit Sub
    'if there is more columns selected
    Else
        i = 0
        start_read = 1
        'read them all
        Do While end_read > 0
            end_read = InStr(start_read, str_all, ",")
            If end_read = 0 Then
                end_read = Len(str_all)
                tmp_arr(i) = Mid(str_all, start_read, end_read - start_read)
                tmp_arr(i) = Trim(tmp_arr(i))
                Exit Do
            Else
                tmp_arr(i) = Mid(str_all, start_read, end_read - start_read)
                tmp_arr(i) = Trim(tmp_arr(i))
            End If
            
            start_read = end_read + 1
            i = i + 1
        Loop
    End If
    'at the end copy all from tmp array to array in main function "ExecuteSqlSelect"
    For i = 0 To UBound(tmp_arr)
        For j = 0 To mtbl_params(tbl_index).col_count - 1
            If tmp_arr(i) = mtbl_params(tbl_index).cols_arr(j) Then
                colIndex_arr(i) = j
            End If
        Next j
    Next i
    
    Erase tmp_arr
End Sub
'function find table index from sql query (find table name in query and search for
'   index)
Private Function getTblIndex(ByRef strSql As String, ByRef mtbl_params() As tbl_params) As Integer
    getTblIndex = -1
    On Error GoTo errH:
    Dim i As Integer, tmp_int As Integer, tmp_int2 As Integer
    Dim tmp_str As String
    
    tmp_int = InStr(1, strSql, " FROM ")
    tmp_int2 = InStr(1, strSql, " WHERE ")
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
