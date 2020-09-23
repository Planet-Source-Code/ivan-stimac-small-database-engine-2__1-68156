Attribute VB_Name = "modIndexes"
'==============================================================
'       module_name           : modIndexes
'       module_version        : v. 1.0.0
'       app_name and version  : Small Database Engine v. 2
'       module_description    : work with indexes and indexed
'                               columns
'==============================================================
'
'
'   ABOUT INDEXING
'
'       For info about indexing look for:
'       MSDN: Indexes
'       MSDN: Indexes Collection
'
'
'   This database uses simple indexing to improve search performance.
'   There is 2 types of indexes, char_indexing and field indexing.
'   Char indexing links field indexes with same first char.
'   Each table column (selected as indexed column) have char indexes
'   and filed indexes.
'
'   Char index is one string that contains last field index
'   for each char (ascii 32 - 255 supported)
'   Shape in file:
'   last_index_pos(ascii32);last_index_pos(ascii33),...,last_index_pos(ascii255)
'   exampe: 11254;54788;45548;5454424;....
'           11254 is position in file of last index that have first char chr(ascii 32)
'
'   Field index contains field data, pointer to row and pointer to previous
'   index
'
'   ---------------------------------------------------------------------------------------
'   Ilustarition:
'
'           a                              b                                c                      ....
'   [anie, ana, anie,..]             [bob, boris...]                   [cat, car...]            ...........
'
'
'   Pointers:
'
'      a -> pointer to [anie]
'   anie -> pointer to [ana]
'         > pointer to /row/
'    ana -> pointer to [anie] : but this anie is new index (pointer to row that also have anie)
'         > pointer to /row/
'   anie -> pointer to [next index]
'         > pointer to /row/
Option Explicit


'this sub check all indexes and searching for current, and if
'   this index exists then add to him pointer to current row, if
'   not exist create him
'                         opened bd file        last position in file  position of char indexes    index to save (rof indexed field)   position in file of this index
Public Sub writeIndex(ByVal ffDB As Integer, ByVal tblInd As Integer, ByVal colInd As Integer, ByVal chrInd_pos As Long, ByVal row_pointer As Long, ByVal str_index As String) ', ByRef ret_pos As Long)
    Dim new_ind As srch_indexes
    Dim i As Integer, j As Integer, haveCol As Boolean
    Dim tmp_arr() As String
    
    'Dim sel_tblInd As String
    'Dim sel_cols(0 To 3) As String
    'Dim sel_clsNum As Integer
    Dim chrInd_last(32 To 255) As Long
    '
    If ffDB = 0 Then Exit Sub
    
    'store char indexes (contains position of last indexes for each char)
    '   support asc 32 - 255
    'check is selected column in cache
'''    haveCol = False
'''    For i = 0 To 3
'''        If sel_cols(i) = Str(colInd) Then
'''            haveCol = True
'''            sel_clsNum = i
'''        End If
'''    Next i
    
    'check is selected table in chache
   ' If sel_tblInd <> Str(tblInd) Or colInd <> sel_clsNum Then 'haveCol <> True Then
'''        If sel_tblInd <> Str(tblInd) Then
'''            sel_clsNum = 0
'''            sel_cols(0) = ""
'''            sel_cols(1) = ""
'''            sel_cols(2) = ""
'''            sel_cols(3) = ""
'''        End If

        'sel_clsNum = colInd
        'sel_tblInd = Str(tblInd)
        '
        Dim readCharIndexes As String

        readCharIndexes = Space(2000)
        '
        Get ffDB, chrInd_pos, readCharIndexes
        'remove spaces
        If InStr(1, readCharIndexes, Chr(0)) <> 0 Then
            readCharIndexes = Mid(readCharIndexes, 1, InStr(1, readCharIndexes, Chr(0)) - 1)
        Else
            readCharIndexes = Trim(readCharIndexes)
        End If
        'parse
        If InStr(1, readCharIndexes, ";") > 0 Then
            tmp_arr = Split(readCharIndexes, ";")
            For i = 0 To UBound(tmp_arr)
                chrInd_last(i + 32) = CLng(tmp_arr(i))
            Next i
            'sel_clsNum = 1
        End If
    'End If
    '
    i = Asc(Left$(str_index, 1))
    If i >= 32 Then
        Dim mStr As String * 2000
        Err.Clear
        
        new_ind.fld_data = str_index
        new_ind.prev_index = chrInd_last(i)
        new_ind.fld_row = row_pointer
        
        'save new last char index position
        chrInd_last(i) = LOF(ffDB) + 1
        'write new index
        Put #ffDB, LOF(ffDB) + 1, new_ind
        'return position of new index
        'ret_pos = LOF(ffDB) + 1
        
        'load with selected index data
        ReDim tmp_arr(32 To 255)
        For j = 32 To 255
            tmp_arr(j) = chrInd_last(j)
        Next j
        'rewrite char indexes
        mStr = Join(tmp_arr, ";")
        Put #ffDB, chrInd_pos, mStr
    End If
    Erase tmp_arr
End Sub
'                       file number             search for             position of char indexes    position of first index that match query, last index in file
Public Sub loadIndex(ByVal ffDB As Integer, ByVal strSearch As String, ByVal chrInd_pos As Long, ByRef firstInd_pos As Long)
    Dim mSrch_ind As srch_indexes
    'Dim mRow As row_data
    Dim tmp_arr() As String, readCharIndexes As String
    Dim i As Integer, chr_pos As Long, prev_row As Long
    '
    Dim chrInd_posTmp(32 To 255) As Long
    '
    firstInd_pos = 0
    'read char indexes from file
    readCharIndexes = Space(2000)
    
    If chrInd_pos = 0 Then Exit Sub
    
    Get #ffDB, chrInd_pos, readCharIndexes
    '
    'remove spaces
    If InStr(1, readCharIndexes, Chr(0)) <> 0 Then
        readCharIndexes = Mid(readCharIndexes, 1, InStr(1, readCharIndexes, Chr(0)) - 1)
    Else
        readCharIndexes = Trim(readCharIndexes)
    End If
    
    'if no empty
    If InStr(1, readCharIndexes, ";") > 0 Then
        'parse
        tmp_arr = Split(readCharIndexes, ";")
        For i = 0 To UBound(tmp_arr)
            chrInd_posTmp(i + 32) = CLng(tmp_arr(i))
        Next i
        'find where starts first index that containd data with
        '   same first cahr as string strSearch
        i = Asc(Left$(strSearch, 1))
        chr_pos = chrInd_posTmp(i)
        'read last index
        prev_row = chr_pos
        Do While prev_row > 0
            Get #ffDB, prev_row, mSrch_ind
            'if we find data that match then save it's
            '   position and exit sub
            'MsgBox mSrch_ind.fld_data
            If mSrch_ind.fld_data = strSearch Then
                
                firstInd_pos = prev_row 'mSrch_ind.fld_row
                Erase tmp_arr
                Exit Sub
            End If
            prev_row = mSrch_ind.prev_index
        Loop
    End If
    Erase tmp_arr
End Sub

'create char indexes
Public Function createCharIndexes(ByVal ffDB As Integer) As Long
    Dim tmp_arr1(32 To 255) As String
    Dim j As Integer
    Dim mStr As String * 2000
    '
    For j = 32 To 255
        tmp_arr1(j) = Str(0)
    Next j
    'rewrite char indexes
    mStr = Join(tmp_arr1, ";")
    createCharIndexes = LOF(ffDB) + 1
    '
    Put #ffDB, LOF(ffDB) + 1, mStr
End Function
