Attribute VB_Name = "modSqlDropTable"
Option Explicit

Public Function ExecuteSqlDrop(ByRef strSql As String, ByRef ffDB As Integer, ByRef mtbl_params() As tbl_params, ByVal tblParams_start As Long) As Long
    Dim tmp_tbl As String
    Dim mStr As String
    Dim i As Integer, i1 As Integer
    Dim tmp_params() As tbl_params
    Dim dbTables As String * 10000
    '
    strSql = Trim(strSql)
    tmp_tbl = Mid$(strSql, InStr(1, strSql, "DROP TABLE ") + Len("DROP TABLE "))
    tmp_tbl = Trim(tmp_tbl)
    '
    ReDim tmp_params(UBound(mtbl_params))
    'copy tables that stays in tmp variable
    i1 = 0
    For i = 0 To UBound(mtbl_params) - 1
        If mtbl_params(i).tbl_name <> tmp_tbl Then
            tmp_params(i1) = mtbl_params(i)
            i1 = i1 + 1
        End If
    Next i
    'then copy all from tmp variable to tbl_params variable
    ReDim mtbl_params(i1)
    For i = 0 To i1 - 1
        mtbl_params(i) = tmp_params(i)
    Next i
    '
    'create string to write in file
    mStr = createStrTblDef(mtbl_params)
    'cript this string
    dbTables = Crypt(mStr)
    'write to file
    Put ffDB, tblParams_start, dbTables
End Function
