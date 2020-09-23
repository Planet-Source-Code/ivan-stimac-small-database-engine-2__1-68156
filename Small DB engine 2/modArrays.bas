Attribute VB_Name = "modArrays"
Option Explicit
'check does arr (string data type) contain value 'whatVal'
'Public Function isInArrStr(ByRef mArr() As String, ByVal whatVal As String) As Boolean
'    On Error GoTo errH
'    Dim i As Long
'    isInArr = False
'    For i = LBound(mArr) To UBound(mArr)
'        If mArr(i) = whatVal Then
'            isInArr = True
'            Exit Function
'        End If
'    Next i
'errH:
'End Function

'check does arr (long data type)  contain value 'whatVal'
Public Function isInArrLng(ByRef mArr() As Long, ByVal whatVal As Long) As Boolean
    On Error GoTo errH
    Dim i As Long
    isInArrLng = False
    For i = LBound(mArr) To UBound(mArr)
        If mArr(i) = whatVal Then
            isInArrLng = True
            Exit Function
        End If
    Next i
errH:
End Function

