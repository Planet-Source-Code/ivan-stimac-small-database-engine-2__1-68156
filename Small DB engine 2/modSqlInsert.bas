Attribute VB_Name = "modSqlInsert"
' function messages:
Option Explicit
'   -1 : there is no selected DB
'   -2 : fail to execute
'    0 : execute success
Public Function ExecuteSqlInsert(ByRef strSql As String, ByRef ffDB As Integer, ByRef mtbl_params() As tbl_params, Optional force_write As Boolean = False) As Long
    'On Error Resume Next
    Dim fRet As String
    Dim m_tblIndex As Integer, i As Integer, j As Integer
    Dim tmp_lastID As Long, tmp_lastPos As Long, row_pointer As Long
    Dim mrow_data As row_data
    Dim data_arr() As String
    'save to tmp variables lastID and lastPos to
    '   restore if happens error
    tmp_lastPos = mtbl_params(m_tblIndex).last_position
    tmp_lastID = mtbl_params(m_tblIndex).last_id
    'parse insert into query to string that will be writen to file : 'val1'|'val2'...
    fRet = parseInsetIntoSql(strSql, mtbl_params, m_tblIndex, data_arr, force_write)
    'set row id
    mrow_data.row_id = mtbl_params(m_tblIndex).last_id
    'increase row_count
    mtbl_params(m_tblIndex).rows_cnt = mtbl_params(m_tblIndex).rows_cnt + 1
    
    mrow_data.row_contain = fRet
    mrow_data.row_prev = mtbl_params(m_tblIndex).last_position
    mrow_data.row_using = "+"
    
    'check if we need to set new step values
    If mtbl_params(m_tblIndex).last_id Mod 50000 = 0 Then
        mtbl_params(m_tblIndex).step_big = LOF(ffDB) + 1
        mtbl_params(m_tblIndex).step_med = LOF(ffDB) + 1
        mtbl_params(m_tblIndex).step_small = LOF(ffDB) + 1
'        mrow_data.step_big = LOF(ffDB) + 1
'        mrow_data.step_med = LOF(ffDB) + 1
'        mrow_data.step_small = LOF(ffDB) + 1
    ElseIf mtbl_params(m_tblIndex).last_id Mod 10000 = 0 Then
        mtbl_params(m_tblIndex).step_med = LOF(ffDB) + 1
        mtbl_params(m_tblIndex).step_small = LOF(ffDB) + 1
'        mrow_data.step_med = LOF(ffDB) + 1
'        mrow_data.step_small = LOF(ffDB) + 1
    ElseIf mtbl_params(m_tblIndex).last_id Mod 1000 = 0 Then
        mtbl_params(m_tblIndex).step_small = LOF(ffDB) + 1
'        mrow_data.step_small = LOF(ffDB) + 1
    End If
    
'    If mrow_data.step_small > 0 Then
'        MsgBox mrow_data.step_big & vbCrLf & _
'            mrow_data.step_med & vbCrLf & _
'            mrow_data.step_small
'    End If
'
    'set row step values
    mrow_data.step_big = mtbl_params(m_tblIndex).step_big
    mrow_data.step_med = mtbl_params(m_tblIndex).step_med
    mrow_data.step_small = mtbl_params(m_tblIndex).step_small
    
    mtbl_params(m_tblIndex).last_position = LOF(ffDB) + 1
    
    row_pointer = LOF(ffDB) + 1
    Put ffDB, LOF(ffDB) + 1, mrow_data
    'write indexes for all indexed rows
    For i = 0 To UBound(mtbl_params(m_tblIndex).indCols_arr)
        'if there is no yet char indexes for this column then create them
        If mtbl_params(m_tblIndex).lastIndPos_arr(i) = 0 Then
           mtbl_params(m_tblIndex).lastIndPos_arr(i) = createCharIndexes(ffDB)
        End If
        'write new index
        writeIndex ffDB, m_tblIndex, mtbl_params(m_tblIndex).indCols_arr(i), mtbl_params(m_tblIndex).lastIndPos_arr(i), _
                            row_pointer, data_arr(mtbl_params(m_tblIndex).indCols_arr(i))
    Next i
    
    If Err.Number <> 0 Then
        'restore values
        mtbl_params(m_tblIndex).last_id = tmp_lastID
        mtbl_params(m_tblIndex).last_position = tmp_lastPos
        mtbl_params(m_tblIndex).rows_cnt = mtbl_params(m_tblIndex).rows_cnt - 1
        ExecuteSqlInsert = -2
        Err.Raise -2, , "Fail to open/write in file!"
    End If
End Function


'parse INSERT INTO SQL query
'   INSERT INTO tblName (colName1, colName2...) VALUES ('value1','value2'...) => TO :
'   'value1'|'value2'|...
Public Function parseInsetIntoSql(ByVal strSql As String, ByRef tbls() As tbl_params, ByRef ret_tblInd As Integer, ByRef str_dataSorted() As String, Optional force_write As Boolean = False) As String
    'On Error Resume Next
    Dim i As Long, z As Long, k As Long, data_cnt As Integer, len1 As Integer, len2 As Integer
    Dim strChr As String, tmpStr As String
    Dim str_fields() As String, str_data() As String ', str_dataSorted() As String
    Dim mTbl As String, mTbl_index As Integer
    
    'get table name
    mTbl = Mid(strSql, Len("INSERT INTO "), InStr(1, strSql, "(") - Len("INSERT INTO "))
    mTbl = Trim(mTbl)
    '
    'find table index in tblParams
    For i = 0 To UBound(tbls)
        If Trim(tbls(i).tbl_name) = mTbl Then
            mTbl_index = i
            Exit For
        ElseIf i = UBound(tbls) Then
            parseInsetIntoSql = "INVALID TABLE NAME!"
            Err.Raise -20, , "Invalid table name detected!"
            mTbl_index = -1
            Exit Function
        End If
    Next i
    ret_tblInd = mTbl_index
    
    'count fields
    data_cnt = InStrCharCount(strSql, ",")
    data_cnt = data_cnt + 1
    '
    ReDim str_fields(data_cnt)
    ReDim str_data(data_cnt)

    'read fields
    z = 0
    i = InStr(1, strSql, "(")
    k = InStr(1, strSql, ")")
    len1 = 0
    Do While len1 < k
        len2 = InStr(len1 + 1, strSql, ",")
        If len2 = 0 Then len2 = Len(strSql)
        If len2 > k Then len2 = k
        If len1 > len2 Then Exit Do
        
        If len1 = 0 Then
            str_fields(z) = Trim(Mid(strSql, i + 1, len2 - i - 1))
        Else
            str_fields(z) = Trim(Mid(strSql, len1 + 1, len2 - len1 - 1))
        End If
        len1 = len2
        z = z + 1
    Loop

    'find data values
    len1 = InStr(10, strSql, "VALUES", vbTextCompare)
    z = 0
    i = 0
    'k = len1
    Do While len1 > 0
        len1 = InStr(len1, strSql, "'") 'mColl.Item(i), "'")
        If len1 > 0 Then
            i = i + 1
            If i = 2 Then
                str_data(z) = Mid(strSql, len2, len1 - len2)
                z = z + 1
                i = 0
            Else
                len2 = len1 + 1
            End If
            len1 = len1 + 1
        End If
    Loop
    'increase last id
    tbls(mTbl_index).last_id = tbls(mTbl_index).last_id + 1
    'now we must sort data at this order as tbl definition columns order
    ReDim str_dataSorted(tbls(mTbl_index).col_count)
    For i = 0 To tbls(mTbl_index).col_count - 1
        For z = 0 To UBound(str_fields)
            'MsgBox str_fields(z)
            'if current field is primary key (auto increment) then auto set
            '   value
            If tbls(mTbl_index).id_col = i And force_write <> True Then
                '
                str_dataSorted(i) = tbls(mTbl_index).last_id
                Exit For
            'else use value from array
            Else
                If tbls(mTbl_index).cols_arr(i) = str_fields(z) Then
                    str_dataSorted(i) = str_data(z)
                    Exit For
                End If
            End If
        Next z
    Next i
    
    'set return value
    parseInsetIntoSql = ""
    For i = 0 To UBound(str_dataSorted) - 1
        If i > 0 Then parseInsetIntoSql = parseInsetIntoSql & "|"
        parseInsetIntoSql = parseInsetIntoSql & "'" & str_dataSorted(i) & "'"
    Next i
    'str_data = str_dataSorted(ind_data)
    'at the end free memory
    Erase str_fields
    Erase str_data
End Function

