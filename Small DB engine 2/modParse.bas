Attribute VB_Name = "modParse"
Option Explicit
'get username and password from strDbAccess string (username=sdgasg;password=sdgsd)
Public Sub parseDbAccess(ByVal strDbAccess As String, ByRef strUserName As String, ByRef strPass As String)
    Dim i As Integer, j As Integer
    Dim tmpStr As String, strChr As String ', strPass As String, strUserName As String
    
    'check is there username or password
    i = InStr(1, strDbAccess, "username", vbTextCompare)
    j = InStr(1, strDbAccess, "password", vbTextCompare)
    If i = 0 Or j = 0 Then
        strUserName = ""
        strPass = ""
        Exit Sub
    End If
    'username
    i = InStr(i + 1, strDbAccess, "=", vbTextCompare)
    j = InStr(i + 1, strDbAccess, ";", vbTextCompare)
    If j > 0 Then
        strUserName = Mid(strDbAccess, i + 1, j - i - 1)
    Else
        strUserName = Mid(strDbAccess, i + 1)
    End If
    'password
    i = InStr(1, strDbAccess, "password", vbTextCompare)
    i = InStr(i + 1, strDbAccess, "=", vbTextCompare)
    j = InStr(i + 1, strDbAccess, ";", vbTextCompare)
    If j > 0 Then
        strPass = Mid(strDbAccess, i + 1, j - i - 1)
    Else
        strPass = Mid(strDbAccess, i + 1)
    End If
End Sub

'parse string with tables params and store data in tblParams array
' tbl_name[(col_count,id_row)(col_name1,...):(last_id,row_count,las_position)(indexed_row1,..)(index_pos1,...)]
Public Sub parseTables(ByVal strToParse As String, ByRef storeParams() As tbl_params)
    Dim i As Integer, colCnt As Integer, readLevel As Integer, readSubLevel As Integer
    Dim tblsNum As Integer, currTbl As Integer, z As Integer
    Dim strChr As String, tmpStr As String
    Dim tmpName As String, tmpLastID As Long, tmpLast As Long, tmpCols() As String
    Dim id_col As Integer
    Dim tmp_rows As Long
    'indexed columns
    Dim tmp_indexes() As Integer, indexes_cnt As Integer
    'where index starts
    Dim ind_start() As Long
    
    Dim tbl_start As Integer, tbl_end As Integer, lvl_start As Integer, lvl_end As Integer
    Dim tmp_arr() As String
    
    readLevel = 0
    readSubLevel = 0
    tblsNum = 0
    indexes_cnt = 0
    '
    i = InStr(1, strToParse, "]")
    tblsNum = 1
    If i = 0 Then Exit Sub
    Do While i > 0
        i = InStr(i + 1, strToParse, "]")
        If i > 0 Then tblsNum = tblsNum + 1
    Loop
    tblsNum = tblsNum '- 1
    
    'MsgBox tblsNum
    
    If tblsNum = 0 Then Exit Sub
    ReDim storeParams(tblsNum)
    
    
    tbl_start = 0
    tbl_end = 0
    
    For i = 0 To UBound(storeParams) - 1
        tbl_start = InStr(tbl_start + 1, strToParse, "[")
        'find table name
        If tbl_end = 0 Then
            storeParams(i).tbl_name = Mid(strToParse, 1, tbl_start - 1)
        Else
            storeParams(i).tbl_name = Mid(strToParse, tbl_end + 1, tbl_start - tbl_end - 1)
        End If
        'MsgBox storeParams(i).tbl_name
        tbl_end = InStr(tbl_end + 1, strToParse, "]")
        '
        lvl_start = tbl_start
        readLevel = 1
        Do While readLevel <= 6
            lvl_start = InStr(lvl_start + 1, strToParse, "(")
            lvl_end = InStr(lvl_start + 1, strToParse, ")")
            tmpStr = Mid(strToParse, lvl_start + 1, lvl_end - lvl_start - 1)
            'MsgBox tmpStr & vbCrLf & readLevel
            Select Case readLevel
                'columns count and index of primary key row
                Case 1
                    tmp_arr = Split(tmpStr, ",")
                    storeParams(i).col_count = tmp_arr(0)
                    storeParams(i).id_col = tmp_arr(1)
                'column names
                Case 2
                    If InStr(1, tmpStr, ",") > 0 Then
                        storeParams(i).cols_arr = Split(tmpStr, ",")
                    'if there is only one column
                    Else
                        ReDim storeParams(i).cols_arr(0)
                        storeParams(i).cols_arr(0) = tmpStr
                    End If
                'last id, rows count, last position in file
                Case 3
                    tmp_arr = Split(tmpStr, ",")
                    storeParams(i).last_id = tmp_arr(0)
                    storeParams(i).rows_cnt = tmp_arr(1)
                    storeParams(i).last_position = tmp_arr(2)
                'indexed columns
                Case 4
                    If InStr(1, tmpStr, ",") > 0 Then
                        tmp_arr = Split(tmpStr, ",")
                        ReDim storeParams(i).indCols_arr(UBound(tmp_arr))
                        For z = 0 To UBound(tmp_arr)
                            storeParams(i).indCols_arr(z) = tmp_arr(z)
                        Next z
                    Else
                        ReDim storeParams(i).indCols_arr(0)
                        If Trim(tmpStr) = "" Then
                            storeParams(i).indCols_arr(0) = -1
                        Else
                            storeParams(i).indCols_arr(0) = tmpStr
                        End If
                    End If
                'position of last index
                Case 5
                    If InStr(1, tmpStr, ",") > 0 Then
                        tmp_arr = Split(tmpStr, ",")
                        
                        ReDim storeParams(i).lastIndPos_arr(UBound(tmp_arr))
                        For z = 0 To UBound(tmp_arr)
                           storeParams(i).lastIndPos_arr(z) = tmp_arr(z)
                        Next z
                    Else
                        ReDim storeParams(i).lastIndPos_arr(0)
                        If Trim(tmpStr) = "" Then
                            storeParams(i).lastIndPos_arr(0) = 0
                        Else
                            storeParams(i).lastIndPos_arr(0) = tmpStr
                        End If
                    End If
                'where starts last small, medium and big step :
                '   small  : last record in range of 1 000 records
                '   medium :            -||-        10 000 records
                '   big    :            -||-        50 000 records
'                Case 6
'                    tmp_arr = Split(tmpStr, ",")
'                    storeParams(i).step_small = tmp_arr(0)
'                    storeParams(i).step_med = tmp_arr(1)
'                    storeParams(i).step_big = tmp_arr(2)
            End Select
            readLevel = readLevel + 1
        Loop

    Next i
    
'    currTbl = 0
'    For i = 1 To Len(strToParse)
'        'MsgBox "POC:" & tmpStr
'        strChr = Mid(strToParse, i, 1)
'        'after [ start tblDef
'        If strChr = "[" Then
'            tmpName = tmpStr
'            tmpStr = ""
'        'then comes ( and inside column count and index of id column
'        ElseIf strChr = "(" Then
'            readLevel = readLevel + 1
'            readSubLevel = 0
'            tmpStr = ""
'        'column count
'        ElseIf strChr = "," And readLevel = 1 Then
'            colCnt = tmpStr
'            ReDim tmpCols(colCnt)
'            tmpStr = ""
'        'index of id column
'        ElseIf strChr = ")" And readLevel = 1 Then
'            id_col = tmpStr
'            tmpStr = ""
'        'save last column name to tmp array
'        ElseIf strChr = ")" And readLevel = 2 Then
'            tmpCols(readSubLevel) = tmpStr
'            tmpStr = ""
'        'last position in file
'        ElseIf strChr = ")" And readLevel = 3 Then
'            tmpLast = tmpStr
'            tmpStr = ""
'        'last indexed col id
'        ElseIf strChr = ")" And readLevel = 4 Then
'            tmpLast = tmpStr
'            tmpStr = ""
'        'if there is , then
'        ElseIf strChr = "," Then
'            'if reading column names
'            If readLevel = 2 Then
'                tmpCols(readSubLevel) = tmpStr
'                tmpStr = ""
'            'or reading lastID, row count and last position in file
'            ElseIf readLevel = 3 Then
'                'first data is lastID
'                If readSubLevel = 0 Then
'                    tmpLastID = tmpStr
'                'second is row count
'                Else
'                    tmp_rows = tmpStr
'                End If
'                tmpStr = ""
'            'read indexed column indexes
'            ElseIf readLevel = 4 Then
'                ReDim Preserve tmp_indexes(indexes_cnt)
'                indexes_cnt = indexes_cnt + 1
'            'read where indexes of each row starts in file
'            ElseIf readLevel = 5 Then
'
'            End If
'            readSubLevel = readSubLevel + 1
'        'at then end save to tblParams variable
'        ElseIf strChr = "]" Then
'            readLevel = 0
'            readSubLevel = 0
'            tmpStr = ""
'
'            storeParams(currTbl).tbl_name = tmpName
'            storeParams(currTbl).col_count = colCnt
'            storeParams(currTbl).last_id = tmpLastID
'            storeParams(currTbl).last_position = tmpLast
'            storeParams(currTbl).id_col = id_col
'            storeParams(currTbl).rows_cnt = tmp_rows
'            ReDim storeParams(currTbl).cols_arr(colCnt)
'            'save column names
'            For z = 1 To colCnt
'                storeParams(currTbl).cols_arr(z - 1) = tmpCols(z - 1)
'            Next z
'            currTbl = currTbl + 1
'        Else
'            tmpStr = tmpStr & strChr
'        End If
'    Next i
    
    Erase tmpCols
    Erase tmp_arr
    Erase ind_start
    Erase tmp_indexes
End Sub

'
Public Function createStrTblDef(ByRef mtbl_params() As tbl_params) As String
    On Error Resume Next
    Dim i As Integer, z As Integer
    'check for error (if there is no tables)
    i = UBound(mtbl_params)
    If Err.Number <> 0 Then Exit Function
    '
    createStrTblDef = ""
    For i = 0 To UBound(mtbl_params) - 1
        createStrTblDef = createStrTblDef & Trim(mtbl_params(i).tbl_name) & "[(" & mtbl_params(i).col_count & "," & mtbl_params(i).id_col & ")("
        For z = 0 To mtbl_params(i).col_count - 1
            createStrTblDef = createStrTblDef & mtbl_params(i).cols_arr(z)
            If z < mtbl_params(i).col_count - 1 Then
                createStrTblDef = createStrTblDef & ","
            End If
        Next z
        createStrTblDef = createStrTblDef & "):(" & mtbl_params(i).last_id & "," & mtbl_params(i).rows_cnt & "," & mtbl_params(i).last_position & ")("
        For z = 0 To UBound(mtbl_params(i).indCols_arr)
            If z > 0 Then createStrTblDef = createStrTblDef & ","
            createStrTblDef = createStrTblDef & mtbl_params(i).indCols_arr(z)
        Next z
        createStrTblDef = createStrTblDef & ")("
        For z = 0 To UBound(mtbl_params(i).lastIndPos_arr)
            If z > 0 Then createStrTblDef = createStrTblDef & ","
            createStrTblDef = createStrTblDef & mtbl_params(i).lastIndPos_arr(z)
        Next z
        createStrTblDef = createStrTblDef & ")]"
        'createStrTblDef = createStrTblDef & mtbl_params(i).step_small & "," & _
                        mtbl_params(i).step_med & "," & mtbl_params(i).step_big & ")]"
    Next i
End Function
