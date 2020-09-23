VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================
'       class_name        : clsRecSet
'       class_version     : v. 1.0.0
'       class_            : for Small Database Engine v. 2
'       class_description : can store table data
'==============================================================

Private strData() As String
Private mRows As Long, mCols As Long
Private currRow As Long

'rows
Public Property Get Rows() As Long
    Rows = mRows
End Property
Public Property Let Rows(ByVal nV As Long)
    mRows = nV
    setFields
End Property
'
'columns
Public Property Get Columns() As Long
    Columns = mCols
End Property
Public Property Let Columns(ByVal nV As Long)
    mCols = nV
    setFields
End Property
'
'data
Public Property Get Data(ByVal cRow As Long, ByVal cCol As Long) As String
    If cRow > mRows - 1 Then cRow = mRows - 1
    If cCol > mCols - 1 Then cCol = mCols - 1
    Data = strData(cRow, cCol)
End Property
Public Property Let Data(ByVal cRow As Long, ByVal cCol As Long, ByVal nData As String)
    If cRow > mRows - 1 Then cRow = mRows - 1
    If cCol > mCols - 1 Then cCol = mCols - 1
    strData(cRow, cCol) = nData
End Property
'
'data
Public Property Get DataRow(ByVal Index As Long) As String
    If currRow > mRows - 1 Then currRow = mRows - 1
    If Index > mCols - 1 Then Index = mCols - 1
    DataRow = strData(currRow, Index)
End Property
Public Property Let DataRow(ByVal Index As Long, ByVal nData As String)
    strData(currRow, Index) = nData
End Property
'
'columns
Public Property Get CurrentRow() As Long
    CurrentRow = currRow
End Property
Public Property Let CurrentRow(ByVal nV As Long)
    currRow = nV
    If currRow > mRows - 1 Then currRow = mRows - 1
End Property
'
'eof
Public Property Get EndOfData() As Boolean
    EndOfData = False
    If currRow > mRows - 1 Then EndOfData = True
End Property

'
Private Sub setFields()
    'Delete strData
    Erase strData
    ReDim strData(mRows, mCols)
    currRow = 0
End Sub
'
Public Function NextRow() As Long
    NextRow = 1
    If currRow < mRows Then
        currRow = currRow + 1
    Else
        'NextRow = EOF
        NextRow = 0
    End If
End Function

Public Function PrevRow() As Long
    PrevRow = 1
    If currRow > 0 Then
        currRow = currRow - 1
    Else
        'NextRow = EOF
        PrevRow = 0
    End If
End Function

Public Sub FirstRow()
    currRow = 0
End Sub

Public Function LastRow()
    currRow = mRows - 1
End Function
'
'
Public Function InsertData(ByRef cRow As Long, ByRef cCol As Long, ByRef mData As String)
    If cRow > mRows - 1 Or cCol > mCols - 1 Then
        ReDim Preserve strData(cRow, cCol)
        mRows = UBound(strData, 1) + 1
        mCols = UBound(strData, 2) + 1
    End If
    strData(cRow, cCol) = nData
End Function


'--------------------------------------------------------------------
Private Sub Class_Initialize()
    mRows = 0
    mCols = 0
    currRow = 0
End Sub

Private Sub Class_Terminate()
    Erase strData
End Sub

