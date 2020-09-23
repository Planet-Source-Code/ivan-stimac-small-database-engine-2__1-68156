VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function exportToSql(ByVal dbFile As String, ByVal dbAccess As String, ByVal outputFile As String)
    Dim i As Long, j As Long, k As Integer, ffOP As Integer
    Dim str_tmp As String, str_tmp2 As String
    Dim mDB As New clsDatabaseEngine
    Dim mRecSet As New clsRecordset

    If mDB.OpenDB(dbFile, dbAccess) = 0 Then
        ffOP = FreeFile
        Open outputFile For Output As #ffOP
            'create table queryes
            For i = 0 To mDB.TablesCount - 1
                str_tmp = "CREATE TABLE " & mDB.TableName(i) & " ("
                For j = 0 To mDB.ColCount(i) - 1
                    If j > 0 Then
                        str_tmp = str_tmp & ","
                    End If
                    '
                    If mDB.PrimaryKey(i) = j Then
                        str_tmp = str_tmp & "*"
                    End If
                    '
                    For k = 0 To mDB.IndexedColumnsCount(i) - 1
                        If mDB.IndexedColumn(i, k) = j Then
                            str_tmp = str_tmp & "#"
                        End If
                    Next k
                    
                    str_tmp = str_tmp & mDB.ColName(i, j)
                Next j
                str_tmp = str_tmp & ")"
                Print #ffOP, str_tmp
            Next i
            'create records queries
            For i = 0 To mDB.TablesCount - 1
                str_tmp = "INSERT INTO " & mDB.TableName(i) & " ("
                For j = 0 To mDB.ColCount(i) - 1
                    If j > 0 Then
                        str_tmp = str_tmp & ","
                    End If
                    str_tmp = str_tmp & mDB.ColName(i, j)
                Next j
                str_tmp = str_tmp & ") VALUES ("
                '
                Set mRecSet = mDB.OpenRecordSet("SELECT * FROM " & mDB.TableName(i))
                
                If mRecSet.Rows > 0 Then
                    mRecSet.MoveFirst
                    
                    For j = 0 To mRecSet.Rows - 1
                        str_tmp2 = str_tmp
                        For k = 0 To mRecSet.Columns - 1
                            If k > 0 Then
                                str_tmp2 = str_tmp2 & ","
                            End If
                            str_tmp2 = str_tmp2 & "'" & mRecSet.Fields(k) & "'"
                        Next k
                        str_tmp2 = str_tmp2 & ")"
                        Print #ffOP, str_tmp2
                        mRecSet.MoveNext
                    Next j
                    
                End If
            Next i
        Close #ffOP
    End If
    
    Set mDB = Nothing
    Set mRecSet = Nothing
End Function
'
Public Function ImportSql(ByVal sourceFile As String, ByVal dbFile As String, ByVal dbAccess As String, Optional keep_IDs As Boolean = False) As Long
    Dim mDB As New clsDatabaseEngine
    Dim i As Long, ffOP As Integer
    Dim str_tmp As String, str_tmp2() As String
    
    i = 0
    If mDB.OpenDB(dbFile, dbAccess) = 0 Then
        ffOP = FreeFile
        Open sourceFile For Input As #ffOP
            mDB.ForceWrite = keep_IDs
            Do Until EOF(ffOP)
                Input #ff, str_tmp
                'naci redak s najvecim id-om i postaviti last id na taj
'                If Left$(str_tmp, Len("CREATE TABLE")) = "CREATE TABLE" Then
'
'                    i = i + 1
'                End If
                mDB.ExecuteSql str_tmp
            Loop
        Close #ffOP
    End If
    mDB.RefreshDB
    mDB.CloseDB
    Set mDB = Nothing
End Function

