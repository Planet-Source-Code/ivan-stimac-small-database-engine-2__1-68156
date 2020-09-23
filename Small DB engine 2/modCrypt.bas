Attribute VB_Name = "modCrypt"
Option Explicit

Public Function Crypt(ByVal strToCrypt As String) As String
    Dim i As Integer, tmpASC As Integer, divLen As Long
    Dim tmpStr As String, tmpChr As String
    tmpStr = ""
    divLen = Len(strToCrypt)
    '
    Do While divLen > 255
        divLen = divLen - 200
    Loop
    '
    'MsgBox divLen
    For i = 1 To Len(strToCrypt)
        tmpChr = Mid(strToCrypt, i, 1)
        If Asc(tmpChr) <> 0 Then
            tmpASC = Asc(tmpChr) - divLen
            If tmpASC < 1 Then tmpASC = 255 + tmpASC
            tmpStr = tmpStr & Chr(tmpASC)
        Else
            'tmpStr = tmpStr & tmpChr
            Exit Function
        End If
    Next i
    Crypt = tmpStr
End Function

Public Function Decrypt(ByVal strToDecrypt As String) As String
    Dim i As Integer, tmpASC As Integer
    Dim tmpStr As String, tmpChr As String
    Dim divLen As Long
    tmpStr = ""
    '
    divLen = Len(strToDecrypt)
    Do While divLen > 255
        divLen = divLen - 200
    Loop
    '

    For i = 1 To Len(strToDecrypt)
        tmpChr = Mid(strToDecrypt, i, 1)
        If Asc(tmpChr) <> 0 Then
            tmpASC = Asc(tmpChr) + divLen
            If tmpASC > 255 Then tmpASC = tmpASC - 255
            tmpStr = tmpStr & Chr(tmpASC)
        Else
            'tmpStr = tmpStr & tmpChr
            Exit Function
        End If
    Next i
    Decrypt = tmpStr
End Function
