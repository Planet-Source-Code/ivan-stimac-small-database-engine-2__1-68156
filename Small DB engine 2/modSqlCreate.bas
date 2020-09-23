Attribute VB_Name = "modSqlCreate"
Option Explicit
' function messages:
'   -1 : there is no selected DB
'  -22 : fail to create
'  -23 : table exists
'    0 : execute success
Public Function ExecuteSqlCreate(ByRef strSql As String, ByRef dbTables As String, ByRef ffDB As Integer, ByRef tblStart As Long, mtbl_params() As tbl_params) As Long
    'On Error Resume Next
    Dim tmpStr As String, fReturn As String
    Dim i As Integer, strChr As String, tmpTblName  As String
    
    fReturn = getTblParamsFromSQL(strSql)
    
    If fReturn <> "INVALID SQL QUERY" Then
        'find table name
        tmpTblName = ""
        i = InStr(1, fReturn, "[")
        tmpTblName = Trim(Mid(fReturn, 1, i - 1))
        'save table
        If Asc(Mid(dbTables, 1, 1)) = 0 Then
            tmpStr = fReturn
        Else
            tmpStr = Trim(dbTables) & fReturn
        End If
        'cript and save table
        dbTables = Crypt(tmpStr)
        Put ffDB, tblStart, dbTables
        'restore string
        dbTables = tmpStr
    Else
        ExecuteSqlCreate = -22
        Err.Raise -22, , "Fail to create table!"
    End If

End Function

'

'parse CREATE TABLE SQL query
'   CREATE TABLE tblName (colName1, colName2...) => TO :
'   tblName[(colNums)(colName1, colName2...):(firstItem_pos,lastItem_pos)]
Private Function getTblParamsFromSQL(ByVal strSql As String) As String
    Dim i As Integer, z As Integer, indexed As Integer
    Dim strChr As String, tmpStr As String
    Dim tmpTblDef As tbl_params
    tmpTblDef.col_count = 0
    indexed = 0
    If Format(Mid(strSql, 1, Len("CREATE TABLE ")), ">") = "CREATE TABLE " Then
        'first we need to find coll count
        For i = Len("CREATE TABLE ") To Len(strSql)
            If Mid(strSql, i, 1) = "," Then tmpTblDef.col_count = tmpTblDef.col_count + 1
        Next i
        tmpTblDef.col_count = tmpTblDef.col_count + 1
        '
        z = 0
        'and then reserve tblCols(col_count) for column names
        ReDim tmpTblDef.cols_arr(tmpTblDef.col_count - 1)
        'fisrs and last data is 0 because there is no data yet
        tmpTblDef.last_position = 0
        tmpTblDef.last_id = 0
        'read table name and column names
        tmpTblDef.id_col = -1
        
        For i = Len("CREATE TABLE ") To Len(strSql)
            strChr = Mid(strSql, i, 1)
            If strChr = "(" Then
                tmpTblDef.tbl_name = Trim(tmpStr)
                tmpStr = ""
            ElseIf strChr = "," Or strChr = ")" Then
                tmpStr = Trim(tmpStr)
                'checking is current row primary key
                If Mid(tmpStr, 1, 1) = "*" Then
                    tmpTblDef.id_col = z
                    tmpStr = Mid(tmpStr, 2)
                'creating indexes for other indexed columns
                ElseIf Mid(tmpStr, 1, 1) = "#" Then
                    ReDim Preserve tmpTblDef.indCols_arr(indexed)
                    tmpTblDef.indCols_arr(indexed) = z
                    ReDim Preserve tmpTblDef.lastIndPos_arr(indexed)
                    tmpTblDef.lastIndPos_arr(indexed) = 0
                    tmpStr = Mid(tmpStr, 2)
                    indexed = indexed + 1
                End If
                tmpTblDef.cols_arr(z) = tmpStr
                tmpStr = ""
                z = z + 1
            Else
                tmpStr = tmpStr & strChr
            End If
        Next i
        'set return value for that will be encriped and writen to file
        getTblParamsFromSQL = tmpTblDef.tbl_name & "[(" & tmpTblDef.col_count & "," & tmpTblDef.id_col & ")("
        For i = 0 To tmpTblDef.col_count - 1
            getTblParamsFromSQL = getTblParamsFromSQL & tmpTblDef.cols_arr(i)
            If i < tmpTblDef.col_count - 1 Then getTblParamsFromSQL = getTblParamsFromSQL & ","
        Next i
        getTblParamsFromSQL = getTblParamsFromSQL & "):(0,0,0)("
        For i = 0 To UBound(tmpTblDef.indCols_arr)
            If i > 0 Then getTblParamsFromSQL = getTblParamsFromSQL & ","
            getTblParamsFromSQL = getTblParamsFromSQL & tmpTblDef.indCols_arr(i)
        Next i
        getTblParamsFromSQL = getTblParamsFromSQL & ")("
        For i = 0 To UBound(tmpTblDef.lastIndPos_arr)
            If i > 0 Then getTblParamsFromSQL = getTblParamsFromSQL & ","
            getTblParamsFromSQL = getTblParamsFromSQL & tmpTblDef.lastIndPos_arr(i)
        Next i
        getTblParamsFromSQL = getTblParamsFromSQL & ")]"
        MsgBox getTblParamsFromSQL
    'if we have invalid query
    Else
        getTblParamsFromSQL = "INVALID SQL QUERY"
    End If
End Function

