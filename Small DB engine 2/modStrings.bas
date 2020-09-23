Attribute VB_Name = "modStrings"
Option Explicit
Public Function InStrCharCount(ByRef mString As String, ByVal strChar As String) As Integer
    Dim mPos As Long
    mPos = 1
    InStrCharCount = 0
    Do While mPos > 0
        mPos = InStr(mPos, mString, strChar, vbBinaryCompare)
        If mPos > 0 Then
            InStrCharCount = InStrCharCount + 1
            mPos = mPos + 1
        End If
    Loop
End Function

Public Sub clearString(ByRef mString As String, ByVal from_what As String)
    Dim i As Integer
    Dim tmp_str As String
    mString = " " & Trim(mString) & " "
    
    i = InStr(1, mString, from_what, vbTextCompare)
    If i > 0 Then
        tmp_str = Mid(mString, 1, i - 1) & Mid(mString, i + Len(from_what))
        mString = Trim(tmp_str)
    End If
    
End Sub


