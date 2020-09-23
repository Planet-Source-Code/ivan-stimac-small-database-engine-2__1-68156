VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabaseEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================
'       class_name           : clsDatabaseEngine
'       class_version        : v. 1.0.0
'       app_name and version : Small Database Engine v. 2
'       class_description    : this is main DB engine class
'                              it links all other modules
'==============================================================
'
Option Explicit
'##############################################################
'      variables to store info about opened file/database
'##############################################################
'file number
Private ffDB As Integer
'store table names, colum names for each table, where is writen
'   last data in file...
Private mtbl_params() As tbl_params
'store userName and password for access to datase
Private mDb_params As db_params
'string that contains all tables and their columns
Private dbTables As String * 10000
'
'store position of tblParams in file
Private tbl_paramsStart As Long
'
'sore db file and dbAccess
Private mDB_file As String, mDB_access As String
'if true then user can alone write id values
Private force_write As Boolean


'------------------------------------------------------------------------
'                       functions
'------------------------------------------------------------------------
'open database
'   -1 : invalid file
'   -2 : fail to open
'   -3 : invalid username or password
'    0 : open success
Public Function OpenDB(ByVal dbFile As String, ByVal dbAccess As String) As Long
    'On Error Resume Next
    Dim mdb_ident As String * 4
    Dim str_userName As String, str_pass As String
    Dim tmp_str As String
    'find user name and passwor from string
    parseDbAccess dbAccess, str_userName, str_pass
    '
    ffDB = FreeFile
    'opne file
    Open dbFile For Binary As #ffDB
    'if there is some errors
    If Err.Number <> 0 Then
        OpenDB = -2
        Err.Raise -2, , "Fail to open file!"
        Exit Function
    End If
    'check dbIdent (is this db file or some invalid file)
    Get ffDB, , mdb_ident
    If mdb_ident <> dbIdent Then
        OpenDB = -3
        Err.Raise -3, , "Invalid file type!"
        Close #ffDB
        ffDB = 0
        Exit Function
    End If
    '
    mDB_file = dbFile
    mDB_access = dbAccess
    'load username and passwor from file
    Get ffDB, , mDb_params
    'decrypt user name and password
    mDb_params.db_password = Decrypt(mDb_params.db_password)
    mDb_params.db_userName = Decrypt(mDb_params.db_userName)
    'then check them
    If str_userName <> Trim(mDb_params.db_userName) Or str_pass <> Trim(mDb_params.db_password) Then
        OpenDB = -4
        Err.Raise -4, , "Invalid UserName or Password!"
        Close #ffDB
        ffDB = 0
        Exit Function
    End If
    'find where start data about tables and columns
    tbl_paramsStart = Loc(ffDB) + 1
    'then read this
    Get ffDB, , dbTables
    
    'use only string
    If InStr(1, dbTables, Chr(0)) <> 0 Then
        tmp_str = Mid(dbTables, 1, InStr(1, dbTables, Chr(0)) - 1)
    Else
        tmp_str = Trim(dbTables)
    End If
    'decrypt it
    tmp_str = Decrypt(Trim(tmp_str))
    'save to dbTables decrypted
    'MsgBox tmp_str
    dbTables = tmp_str
    'and parse
    parseTables Trim(tmp_str), mtbl_params
    
    'MsgBox mtbl_params(1).tbl_name
End Function
'close database
Public Function CloseDB() As Long
    'refresh tables if user didn't do this
    RefreshDB
    'close file
    Close #ffDB
    'reset variables
    ffDB = 0
    dbTables = ""
    mDb_params.db_password = ""
    mDb_params.db_userName = ""
    'free memory
    Erase mtbl_params
End Function
'create a new database
Public Function CreateDB(ByVal dbFile As String, ByVal dbAccess As String) As Long
    On Error Resume Next
    ffDB = FreeFile
    
    If dbAccess = "" Then
        mDb_params.db_password = Space(10)
        mDb_params.db_userName = Space(10)
    Else
        parseDbAccess dbAccess, mDb_params.db_userName, mDb_params.db_password
        mDb_params.db_userName = Crypt(mDb_params.db_userName)
        mDb_params.db_password = Crypt(mDb_params.db_password)
    End If
    
    'write appIdent, dbParams (user name and pass) and tables into file
    Open dbFile For Binary As ffDB
        Put ffDB, , dbIdent
        Put ffDB, , mDb_params
        Put ffDB, , dbTables
    Close ffDB
    
    'messages
    If Err.Number <> 0 Then
        CreateDB = -2
        Err.Raise -2, , "Fail to open/write in file!"
    Else
        CreateDB = 0
    End If
    
    ffDB = 0
End Function

'need to call after writing fields to update tables
Public Sub RefreshDB() 'As Long
    Dim mStr As String
    mStr = createStrTblDef(mtbl_params)

    dbTables = Crypt(mStr)
'   write to file
    Put ffDB, tbl_paramsStart, dbTables
'    'restore (encriped) tables def to string
    dbTables = mStr
    'MsgBox mStr
End Sub

'messages:
'   0 : success
'  -1 :
'  -2 :
'  -3 :
'  -4 :
'  -5 :
'  -6 :
'  -7 :
'  -8 :
'  -9 : invalid column names in where statment
Public Function ExecuteSql(ByVal strSql As String) As Long
    If ffDB = 0 Then
        ExecuteSql = -1
        Err.Raise -1, , "There is no selected database!"
        Exit Function
    End If
    '
    If StrComp(Mid(Trim(strSql), 1, Len("CREATE TABLE")), "CREATE TABLE", vbTextCompare) = 0 Then
        If validateCreateTable(strSql) = True Then
            ExecuteSql = ExecuteSqlCreate(strSql, dbTables, ffDB, tbl_paramsStart, mtbl_params)
            parseTables Trim(dbTables), mtbl_params
        Else
            ExecuteSql = -21
            Err.Raise -21, , "Invalid CREATE TABLE query or table with this name exist!"
        End If
    ElseIf StrComp(Mid(Trim(strSql), 1, Len("INSERT INTO")), "INSERT INTO", vbTextCompare) = 0 Then
        ExecuteSql = ExecuteSqlInsert(strSql, ffDB, mtbl_params, force_write)
    ElseIf StrComp(Mid(Trim(strSql), 1, Len("DELETE FROM")), "DELETE FROM", vbTextCompare) = 0 Then
        ExecuteSql = ExecuteSqlDelete(strSql, ffDB, mtbl_params, mDB_file, mDB_access)
    ElseIf StrComp(Mid(Trim(strSql), 1, Len("DROP TABLE")), "DROP TABLE", vbTextCompare) = 0 Then
        ExecuteSql = ExecuteSqlDrop(strSql, ffDB, mtbl_params, tbl_paramsStart)
    End If
    
End Function

'
Public Function OpenRecordSet(ByVal strSql As String) As clsRecordset ' clsRecSet
    If ffDB = 0 Then
        Err.Raise -1, , "There is no selected database!"
        Exit Function
    End If
    If StrComp(Mid(Trim(strSql), 1, Len("SELECT")), "SELECT", vbTextCompare) = 0 Then
        Set OpenRecordSet = ExecuteSqlSelect(strSql, ffDB, mtbl_params)
    End If
End Function

Friend Sub setRowsCnt(ByVal tbl_ind As Integer, ByVal rws_cnt As Long)
    mtbl_params(tbl_ind).rows_cnt = rws_cnt
End Sub

'-----------------------------------------------------------------------
'check if we have this table
Private Function validateCreateTable(ByRef strSql As String) As Boolean
    On Error Resume Next
    Dim tmpTblName As String
    Dim i As Integer, j As Integer
    
    validateCreateTable = True
    
    i = InStr(1, strSql, "CREATE TABLE")
    i = i + Len("CREATE TABLE")
    j = InStr(i + 1, strSql, "(")
    tmpTblName = Trim(Mid(strSql, i, j - i))
    
    ' check if table with this name exists
    For i = 0 To UBound(mtbl_params)
        If Err.Number <> 0 Then Exit For
        If mtbl_params(i).tbl_name = Trim(tmpTblName) Then
            validateCreateTable = False
            Exit Function
        End If
    Next i
    
End Function

'------------------------------------------------------------------------
'                       properties
'------------------------------------------------------------------------
Public Property Get PrimaryKey(ByVal tblIndex As Integer) As Integer
    On Error Resume Next
    PrimaryKey = mtbl_params(tblIndex).id_col
End Property

Public Property Get TablesCount() As Integer
    On Error Resume Next
    TablesCount = UBound(mtbl_params) '- 1
End Property

Public Property Get TableName(ByVal tblIndex As Integer) As String
    On Error Resume Next
    TableName = mtbl_params(tblIndex).tbl_name
End Property

Public Property Get ColCount(ByVal tblIndex As Integer) As Integer
    On Error Resume Next
    ColCount = mtbl_params(tblIndex).col_count
End Property

Public Property Get ColName(ByVal tblIndex As Integer, ByVal colIndex As Integer) As String
    On Error Resume Next
    ColName = mtbl_params(tblIndex).cols_arr(colIndex)
End Property

Public Property Get IndexedColumnsCount(ByVal tblIndex As Integer) As Integer
    IndexedColumnsCount = UBound(mtbl_params(tblIndex).indCols_arr) + 1
End Property

Public Property Get IndexedColumn(ByVal tblIndex As Integer, ByVal colIndex As Integer) As Integer
    IndexedColumn = mtbl_params(tblIndex).indCols_arr(colIndex)
End Property

'---------------------------------------------------------------
Friend Property Let ForceWrite(ByVal nV As Boolean)
    force_write = nV
End Property
Friend Property Get ForceWrite() As Boolean
    ForceWrite = force_write
End Property


