Attribute VB_Name = "modSqlWhere"
'
Option Explicit
Private i As Integer
'messages:
'    0 : success
'   -1 : no where
'   -2 : can't execute sql (invalid column name detected)
Public Function parseSqlWhere(ByRef strSql As String, ByRef mtbl_params() As tbl_params, ByRef tbl_index As Integer, ByRef colIndex_arr() As Integer, ByRef val_arr() As String, _
                         ByRef operator_arr() As String, ByRef condOper_arr() As String, ByRef str_order As String) As Integer
    
    Dim start_read As Integer, end_read As Integer, tmp_instr As Integer
    Dim arr_ind As Integer, col_num As Integer
    Dim tmp_int As Integer, tmp_arr() As String
    Dim tmp_str As String
    
    parseSqlWhere = -1
    
    'first use only string part after where so we can count columns
    tmp_int = InStr(1, strSql, " WHERE ", vbTextCompare)
    If tmp_int = 0 Then Exit Function
    tmp_str = Mid(strSql, tmp_int)
    'find column count
    col_num = InStrCharCount(tmp_str, "'") / 2
    If col_num = 0 Then Exit Function
    'free memory first
    Erase colIndex_arr
    Erase val_arr
    Erase operator_arr
    Erase condOper_arr
    
    'redim
    ReDim tmp_arr(col_num - 1)
    ReDim colIndex_arr(col_num - 1)
    ReDim val_arr(col_num - 1)
    ReDim operator_arr(col_num - 1)
    If col_num > 1 Then
        ReDim condOper_arr(col_num - 2)
    Else
        ReDim condOper_arr(0)
        condOper_arr(0) = "OR"
    End If
    
    '************** parse column_names and save their index **************
    start_read = InStr(1, strSql, " WHERE ", vbTextCompare)
    'if there is no WHERE then close function
    If start_read = 0 Then
        Erase tmp_arr
        Exit Function
    End If
    'set start at place after WHERE and if this is bigger than len of strSql then close function
    start_read = start_read + Len(" WHERE ")
    If start_read > Len(strSql) Then
        Erase tmp_arr
        Exit Function
    End If
    '
    arr_ind = 0
    'now find all column names
    Do While start_read > 0
        end_read = InStr(start_read + 1, strSql, "'")
        If end_read > start_read Then
            tmp_str = Mid(strSql, start_read, end_read - start_read - 1)
            'if there is operators then delete them from this string
            clearString tmp_str, " AND "
            clearString tmp_str, " OR "
            clearString tmp_str, "="
            clearString tmp_str, "<"
            clearString tmp_str, ">"
            tmp_int = getColumnIndex(mtbl_params, tbl_index, tmp_str)
            'if column exitst then save it's index
            If tmp_int >= 0 Then
                tmp_arr(arr_ind) = tmp_str
                colIndex_arr(arr_ind) = tmp_int
                'MsgBox "IND:" & tmp_int
                arr_ind = arr_ind + 1
            Else
                Err.Raise -30, , "There is some invalid column names in WHERE statment"
                parseSqlWhere = -30
                Erase tmp_arr
                Exit Function
            End If
        End If
        start_read = InStr(end_read + 1, strSql, " ")
        'start_read = InStr(start_read + 1, strSql, " ")
    Loop
    '
    '************** parse values and save them **************
    start_read = InStr(1, strSql, " WHERE ", vbTextCompare)
    start_read = InStr(start_read, strSql, "'")
    '
    If start_read = 0 Then Exit Function
    arr_ind = 0
    '
    Do While start_read > 0
        end_read = InStr(start_read + 1, strSql, "'")
        If end_read > 0 Then
            val_arr(arr_ind) = Mid(strSql, start_read + 1, end_read - start_read - 1)
            'MsgBox "DATA:" & val_arr(arr_ind)
            arr_ind = arr_ind + 1
        End If
        start_read = InStr(end_read + 1, strSql, "'")
    Loop
    '
    '************** parse operators (<, >, <>, =) **************
    arr_ind = 0
    '
    tmp_instr = InStr(1, strSql, " WHERE ")
    start_read = tmp_instr
    For i = 0 To UBound(tmp_arr)
        start_read = InStr(start_read + 1, strSql, tmp_arr(i), vbTextCompare) + Len(tmp_arr(i))
        end_read = InStr(start_read + 1, strSql, "'", vbTextCompare)
        operator_arr(i) = Mid(strSql, start_read, end_read - start_read)
    Next i
    '
    '************** parse operators (AND, OR) **************
    start_read = InStr(1, strSql, " WHERE ", vbTextCompare)
    start_read = InStr(start_read, strSql, "'")
    start_read = InStr(start_read + 1, strSql, "'")
    '
    arr_ind = 0
    '
    Do While start_read > 0
        'find space after value
        start_read = InStr(start_read + 1, strSql, " ")
        'and find space after operator
        end_read = InStr(start_read + 1, strSql, " ")
        If end_read = 0 Or start_read = 0 Then Exit Do
        'then if all ok save it
        If end_read > start_read Then
            condOper_arr(arr_ind) = Trim(Mid(strSql, start_read, end_read - start_read))
            'MsgBox condOper_arr(arr_ind)
            arr_ind = arr_ind + 1
        End If
        start_read = InStr(end_read + 1, strSql, "'")
        start_read = InStr(start_read + 1, strSql, "'")
    Loop
    '
    parseSqlWhere = 0
    
    Erase tmp_arr
End Function
'
'function return column index from column name
Private Function getColumnIndex(ByRef mtbl_params() As tbl_params, ByRef tbl_index As Integer, ByRef mcol_name As String) As Integer
    mcol_name = Trim(mcol_name)
    If tbl_index < 0 Then Exit Function
    For i = 0 To mtbl_params(tbl_index).col_count - 1
        If mcol_name = mtbl_params(tbl_index).cols_arr(i) Then
            getColumnIndex = i
            Exit Function
        End If
    Next i
    getColumnIndex = -1
End Function

                         
