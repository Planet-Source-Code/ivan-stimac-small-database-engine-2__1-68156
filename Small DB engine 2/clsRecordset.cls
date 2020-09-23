VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================
'       class_name           : clsRecordset
'       class_version        : v. 1.0.0
'       app_name and version : Small Database Engine v. 2
'       class_description    : this class search for number
'                              of rows that match query, read
'                              selected row from file
'==============================================================

Option Explicit
'opened file
Private ffDB As Integer
'last position, position of first data, pos of last data
Private last_pos As Long, first_start As Long, last_start As Long, last_id As Long, id_col As Integer
'step big, med, small
'Private stp_big As Long, stp_med As Long, stp_small As Long
'records count, columns count, current_row
Private record_cnt As Long, column_cnt As Long, curr_row As Long
'if there is <>
Private skip_row As Long
'all records in table
Private tbl_recs As Long
'other data from clsDatabase
Private strSql As String, haveWhere As Boolean
'selected columns
Private msel_cols() As Integer
'columns, values and operators in where statment
Private colIndex_arr() As Integer, val_arr() As String, operator_arr() As String, condOper_arr() As String
'to store row fields
Private fileds_arr() As String
'indexed columns and char index positions
Private ind_cols() As Integer, chrInd_pos() As Long
'positions of columns that should not be in recordest
Private dontUse_arr() As Long
'position of columns that match query if we using muliple condition seach
Private use_arr() As Long, readFrom_arr As Boolean
'position in file of current record (need if we use recordset for
'       deleting)
Private currRec_pos As Long
Private is_delete As Boolean
'
Private r_data As row_data, prev_row As Long, tmp_row As Long
Private ind_data As srch_indexes, prev_ind As Long, first_ind As Long
'
Private tmp_rowData As String
'
Dim i As Integer, j As Integer, k As Integer, c As Integer, lastS As Integer
Dim i1 As Integer


'##############################################################
'                  friend properties
'##############################################################
'opened db file number
Friend Property Let db_FileNumber(ByVal nV As Integer)
    ffDB = nV
End Property
'last position in file
Friend Property Let db_LastPosition(ByVal nV As Long)
    last_pos = nV
    prev_row = nV
End Property
'last id
Friend Property Let db_LastID(ByVal nV As Long)
    last_id = nV
End Property
'id column index
Friend Property Let db_idCol(ByVal nV As Integer)
    id_col = nV
End Property
'~~~~~~~~~~~~~ steps ~~~~~~~~~~~~~~~~~~~~~~~
'Friend Property Let db_StepBig(ByVal nV As Long)
'    stp_big = nV
'End Property
'Friend Property Let db_StepMed(ByVal nV As Long)
'    stp_med = nV
'End Property
'Friend Property Let db_StepSmall(ByVal nV As Long)
'    stp_small = nV
'End Property
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'allRecordsCount
Friend Property Let tbl_records(ByVal nV As Long)
    tbl_recs = nV
End Property
'sql query
Friend Property Let str_sql(ByVal nV As String)
    strSql = nV
End Property
'is there what statment in query
Friend Property Let have_where(ByVal nV As Boolean)
    haveWhere = nV
End Property
'
'
'is there what statment in query
Friend Property Let isForDelete(ByVal nV As Boolean)
    is_delete = nV
End Property
'if we need record position to delete it
Friend Property Get RecordPosition() As Long
    RecordPosition = currRec_pos
End Property
'##############################################################
'                  friend subs/functions
'##############################################################
Friend Sub set_arrs(ByRef sel_cols() As Integer, ByRef col_ind() As Integer, _
        ByRef cnd_vals() As String, ByRef cnd_oper() As String, ByRef oper2() As String, _
        indexedCols_arr() As Integer, charIndexes_pos() As Long)
    '
    msel_cols = sel_cols
    colIndex_arr() = col_ind
    val_arr() = cnd_vals
    operator_arr() = cnd_oper
    condOper_arr() = oper2
    'redim array that store row fields
    ReDim fileds_arr(UBound(msel_cols))
    column_cnt = UBound(msel_cols) + 1
    ind_cols = indexedCols_arr
    chrInd_pos = charIndexes_pos
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'               search for results count
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Friend Sub result_count()
    readFrom_arr = False
    prev_ind = -1
    skip_row = 0
'-->
    'if there is no WHERE wuery
    If haveWhere = False Then
        record_cnt = tbl_recs
        first_start = last_pos
'-->
    'if searching by primary key only
    ElseIf UBound(colIndex_arr) = 0 And colIndex_arr(0) = id_col Then
        search_primKey
'-->
    'if search in one column but this isn't primary key column
    ElseIf UBound(colIndex_arr) = 0 And colIndex_arr(0) <> id_col Then
        search_IndexedRow
'-->
    'if searching by primary key only, but 2 conditions
    ElseIf UBound(colIndex_arr) = 1 And colIndex_arr(0) = id_col And colIndex_arr(1) = id_col Then
        Dim rec1 As Long, rec2 As Long
        Dim pos1 As Long, pos2 As Long
        Dim tmp_last1 As Long, tmp_last2 As Long
        'search first condition
        search_primKey 0
        pos1 = first_start
        rec1 = record_cnt
        tmp_last1 = last_start
        'search second condition
        search_primKey 1
        pos2 = first_start
        rec2 = record_cnt
        tmp_last2 = last_start
        '
        'get where starts first and last row in file
        If pos1 < pos2 Then
            first_start = pos1
        Else
            first_start = pos2
        End If

        If tmp_last1 > tmp_last2 Then
            last_start = tmp_last1
        Else
            last_start = tmp_last2
        End If
        '
        'get unused records if looking for <
        If Trim(operator_arr(0)) = "<" Then
            rec1 = tbl_recs - rec1
        End If
        If Trim(operator_arr(1)) = "<" Then
            rec2 = tbl_recs - rec2
        End If
        'get record sount
        If Trim(operator_arr(0)) <> Trim(operator_arr(1)) Then
            record_cnt = Abs(rec1 - rec2)
        Else
            If rec1 < rec2 Then
                record_cnt = rec1
                first_start = pos1
            Else
                record_cnt = rec2
                first_start = pos2
            End If
        End If
'-->
    'if there is any other contition idCol='val' AND/OR col2='val2'...
    Else
        search_muliple 0
        readFrom_arr = True
    End If
End Sub

'?????????????????????????????????????????????????????????????????
'               search by muliple condition
'?????????????????????????????????????????????????????????????????
Private Sub search_muliple(Optional arrInd1 As Integer = 0)
    'results positions
    Dim tmp_res() As Long ' (number of columns to check, number of rows that match)
    Dim col_res() As Long ' number of rows for each column
    Dim tmp_res2() As Long, tmp_arr1() As Long
    Dim tmp_cnt As Long
    '
    Dim i As Integer, z As Integer, j As Long, last_val As Long
    Dim tmp_isForDel As Boolean
    '
    ReDim tmp_res(UBound(colIndex_arr), 1000)
    ReDim col_res(UBound(colIndex_arr))
    last_val = 1000
    '
    'set is for delete to True to avoid parsing row fields and
    '   get faster execute
    tmp_isForDel = is_delete
    isForDelete = True
    tmp_cnt = 0
    For i = 0 To UBound(colIndex_arr)
        If colIndex_arr(i) = id_col Then
            search_primKey (i)
            'move to first record that match query
            If record_cnt > 0 Then
                MoveFirst
                For j = 0 To record_cnt - 1
                    'if we need to increse array size
                    If j > last_val Then
                        last_val = last_val + 1000
                        ReDim Preserve tmp_res(UBound(colIndex_arr), last_val)
                    End If
                    '

                    'save position of record
                    tmp_res(i, j) = RecordPosition
                    'increse number of records that match query for current row
                    col_res(i) = col_res(i) + 1
                    'move to next record
                    MoveNext
                Next j
            End If
        Else
            search_IndexedRow (i)
            If record_cnt > 0 Then
                'move to first record that match query
                MoveFirst
                For j = 0 To record_cnt - 1
                    'if we need to increse array size
                    If j > last_val Then
                        last_val = last_val + 1000
                        ReDim Preserve tmp_res(UBound(colIndex_arr), last_val)
                    End If
                    'save position of record

                    tmp_res(i, j) = RecordPosition
                    'increse number of records that match query for current row
                    col_res(i) = col_res(i) + 1
                    'move to next record
                    MoveNext
                Next j
            End If
        End If
        tmp_cnt = tmp_cnt + last_val
    Next i
    'restore is for delete
    isForDelete = tmp_isForDel
    '
    ReDim tmp_res2(tmp_cnt)
    tmp_cnt = 0
    'number of AND/OR operators
    For z = 0 To UBound(condOper_arr)
        ReDim tmp_arr1(col_res(z + 1))
        For j = 0 To col_res(z + 1) - 1
            tmp_arr1(j) = tmp_res(z + 1, j)
        Next j
        'whatch column name before operator and after operator ''''number of columns to check
        For i = 0 To 1 'UBound(colIndex_arr)
            'number of results for current column
            For j = 0 To col_res(z + i) - 1
                'select operator ( AND / OR )
                Select Case Trim(condOper_arr(z))
                    Case "AND", "and"
                        If isInArrLng(tmp_res2, tmp_res(z, j)) <> True Then
                            If isInArrLng(tmp_arr1, tmp_res(z, j)) = True Then
                                tmp_res2(tmp_cnt) = tmp_res(z, j)
                                tmp_cnt = tmp_cnt + 1
                            End If
                        End If
                        If j = col_res(z + i) - 1 Then GoTo exitCh
                    Case "OR", "or"
                        'MsgBox tmp_res(i, j) & vbCrLf & i
                        If isInArrLng(tmp_res2, tmp_res(z + i, j)) <> True Then
                            tmp_res2(tmp_cnt) = tmp_res(z + i, j)
                            tmp_cnt = tmp_cnt + 1
                        End If
                End Select
            Next j
        Next i
exitCh:
    Next z
    
    record_cnt = tmp_cnt
    

    If tmp_cnt > 0 Then
        ReDim use_arr(tmp_cnt - 1)
        For i = 0 To tmp_cnt - 1
            use_arr(i) = tmp_res2(i)
        Next i
    End If
    'free memory
    Erase col_res
    Erase tmp_res
    Erase tmp_res2
    Erase tmp_arr1
    DoEvents
End Sub


'?????????????????????????????????????????????????????????????????
'               search by indexed row
'?????????????????????????????????????????????????????????????????
Private Sub search_IndexedRow(Optional arrInd1 As Integer = 0)
    Dim i As Integer
    Dim tmp_int As Long
    Dim ret_pos As Long
    '
    Erase dontUse_arr
        '
    tmp_int = -1
    For i = 0 To UBound(ind_cols)
        If ind_cols(i) = colIndex_arr(arrInd1) Then
            tmp_int = colIndex_arr(arrInd1)
        End If
    Next i
    'get position of char indexes
    For i = 0 To UBound(ind_cols)
        If tmp_int = ind_cols(i) Then
            tmp_int = chrInd_pos(i)
            Exit For
        End If
    Next i
    
    If tmp_int < 0 Then Exit Sub
    'get position of first index
    loadIndex ffDB, val_arr(arrInd1), tmp_int, ret_pos
    '
    record_cnt = 0
'-->
    If Trim(operator_arr(arrInd1)) = "=" Then
        'count records
        Do While ret_pos > 0
            Get #ffDB, ret_pos, ind_data
            'if index contain pointer to row that match query
            If ind_data.fld_data = val_arr(arrInd1) Then
                'read row
                Get #ffDB, ind_data.fld_row, r_data
                'check is deleted
                If r_data.row_using = "+" Then
                    'save position of first record
                    If record_cnt = 0 Then
                        first_start = ind_data.fld_row
                        prev_ind = ret_pos ' ind_data.prev_index
                    End If
                    'save current position as last position and do this
                    '   for each record that match query (because we dont know
                    '   realy what record is last
                    last_start = ret_pos
                    'insrease record count
                    record_cnt = record_cnt + 1
                End If
            End If
            ret_pos = ind_data.prev_index
        Loop
        
'-->
    ElseIf Trim(operator_arr(arrInd1)) = "<>" Then
         'count records
        Do While ret_pos > 0
            Get #ffDB, ret_pos, ind_data
            'if index contain pointer to row that match query
            If ind_data.fld_data = val_arr(arrInd1) Then
                'read row
                Get #ffDB, ind_data.fld_row, r_data
                'check is deleted
                If r_data.row_using = "+" Then
                    ReDim Preserve dontUse_arr(record_cnt)
                    dontUse_arr(record_cnt) = ind_data.fld_row
                    record_cnt = record_cnt + 1
                End If
            End If
            ret_pos = ind_data.prev_index
        Loop
        first_start = last_pos
        prev_ind = -1
        record_cnt = tbl_recs - record_cnt
    End If
End Sub


'?????????????????????????????????????????????????????????????????
'               search only by primary key
'?????????????????????????????????????????????????????????????????
Private Sub search_primKey(Optional arrInd1 As Integer = 0)
    record_cnt = 0
'-->
    If Trim(operator_arr(arrInd1)) = "=" Then
        prev_row = last_pos
        Do While prev_row > 0
            Get ffDB, prev_row, r_data
            'if we are on step row then move to prevous because it's pointer is to himself
            If (r_data.row_id Mod 50000 = 0 Or r_data.row_id Mod 10000 = 0 Or r_data.row_id Mod 1000 = 0) And r_data.row_id <> val_arr(arrInd1) Then
                prev_row = r_data.row_prev
            Else
                'MsgBox prev_row & vbCrLf & r_data.step_med
                'checking what jump we can do
                If r_data.row_id - 50000 > val_arr(arrInd1) Then
                    prev_row = r_data.step_big
                ElseIf r_data.row_id - 10000 > val_arr(arrInd1) Then
                    prev_row = r_data.step_med
                ElseIf r_data.row_id - 1000 > val_arr(arrInd1) Then
                    prev_row = r_data.step_small
                'if we find record
                ElseIf r_data.row_id = val_arr(arrInd1) Then
                    'and if record is not deleted
                    If r_data.row_using = "+" Then
                        record_cnt = 1
                        first_start = prev_row
                        last_start = prev_row
                        Exit Sub
                    'but if this record is deleted then record count
                    '   is 0
                    Else
                        record_cnt = 0
                        Exit Sub
                    End If
                Else
                    prev_row = r_data.row_prev
                End If
            End If
        Loop
'-->
    ElseIf Trim(operator_arr(arrInd1)) = "<" Then
        prev_row = last_pos
        last_start = 0
        record_cnt = 0
        Do While prev_row > 0
            Get ffDB, prev_row, r_data
            'if we are on step row then move to prevous because it's pointer is to himself
            If (r_data.row_id Mod 50000 = 0 Or r_data.row_id Mod 10000 = 0 Or r_data.row_id Mod 1000 = 0) And r_data.row_id <> val_arr(arrInd1) Then
                prev_row = r_data.row_prev
            Else
                'checking what jump we can do
                If r_data.row_id - 50000 > val_arr(arrInd1) Then
                    prev_row = r_data.step_big
                ElseIf r_data.row_id - 10000 > val_arr(arrInd1) Then
                    prev_row = r_data.step_med
                ElseIf r_data.row_id - 1000 > val_arr(arrInd1) Then
                    prev_row = r_data.step_small
                'if we find record
                ElseIf r_data.row_id < val_arr(arrInd1) Then
                    If r_data.row_using = "+" Then
                        record_cnt = record_cnt + 1
                        If record_cnt = 1 Then first_start = prev_row
                        'Exit Sub
                    End If
                    prev_row = r_data.row_prev
                Else
                    prev_row = r_data.row_prev
                End If
            End If
        Loop
'-->
    ElseIf Trim(operator_arr(arrInd1)) = ">" Then
        prev_row = last_pos
        first_start = prev_row
        record_cnt = 0
        Do While prev_row > 0
            Get ffDB, prev_row, r_data
            If r_data.row_id > val_arr(arrInd1) Then
                If r_data.row_using = "+" Then
                    record_cnt = record_cnt + 1
                    last_start = prev_row
                End If
                prev_row = r_data.row_prev
            Else
                Exit Sub
            End If
        Loop
'-->
    ElseIf Trim(operator_arr(arrInd1)) = "<>" Then
        prev_row = last_pos
        last_start = 0
        Do While prev_row > 0
            Get ffDB, prev_row, r_data
            'if we are on step row then move to prevous because it's pointer is to himself
            If (r_data.row_id Mod 50000 = 0 Or r_data.row_id Mod 10000 = 0 Or r_data.row_id Mod 1000 = 0) And r_data.row_id <> val_arr(arrInd1) Then
                prev_row = r_data.row_prev
            Else
                'checking what jump we can do
                If r_data.row_id - 50000 > val_arr(arrInd1) Then
                    prev_row = r_data.step_big
                ElseIf r_data.row_id - 10000 > val_arr(arrInd1) Then
                    prev_row = r_data.step_med
                ElseIf r_data.row_id - 1000 > val_arr(arrInd1) Then
                    prev_row = r_data.step_small
                'if we find record
                ElseIf r_data.row_id = val_arr(arrInd1) Then
                    If r_data.row_using = "+" Then
                        record_cnt = tbl_recs - 1
                        skip_row = prev_row
                        first_start = last_pos
                        Exit Sub
                    End If
                Else
                    prev_row = r_data.row_prev
                End If
            End If
        Loop
    End If
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'               load data from file
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Friend Sub set_arr(ByVal data_index As Long, Optional arrInd1 As Integer = 0)
    'find and read selected row
     'if search is not by indexed row
     
'-->
    'if we have positions of selected columns in array
    If readFrom_arr = True Then
        If record_cnt > 0 Then
            If data_index <= UBound(use_arr) Then
                Get ffDB, use_arr(data_index), r_data
                'curr_row = data_index + 1
                tmp_rowData = r_data.row_contain
            End If
        End If
'-->
    'if searching by primary key
    ElseIf prev_ind = -1 Then
        Do While curr_row <= data_index
            'tmp_row = prev_row
            If prev_row = 0 Then Exit Do
            Get ffDB, prev_row, r_data
            'if not deleted then save it
            If r_data.row_using = "+" And skip_row <> prev_row Then
                If isInArrLng(dontUse_arr, prev_row) <> True Then
                    'save row fields
                    If curr_row = data_index Then
                        tmp_rowData = r_data.row_contain
                        currRec_pos = prev_row
                        prev_row = r_data.row_prev
                        '
                        Exit Do
                    End If
                    curr_row = curr_row + 1
                    If curr_row > record_cnt Or curr_row <= last_start Then Exit Sub
                End If
            End If
            prev_row = r_data.row_prev
        Loop
'-->
    Else
        If prev_ind = 0 Then Exit Sub
        Do While prev_ind > 0
            'read index
            Get ffDB, prev_ind, ind_data
            'If ind_data.fld_data = val_arr(arrInd1) Then
                'find row pointer
                prev_row = ind_data.fld_row
                'check does current index data match query
                If ind_data.fld_data = val_arr(arrInd1) Then
                    'if this row is not in dont use arr
                    If isInArrLng(dontUse_arr, prev_row) <> True Then
                        Get ffDB, prev_row, r_data
                        'check is row deleted
                        If r_data.row_using = "+" Then
                            tmp_rowData = r_data.row_contain
                            prev_ind = ind_data.prev_index
                            currRec_pos = prev_row
                            'curr_row = curr_row + 1
                            GoTo exitCheck
                        End If
                    End If
                End If
            'End If
            prev_ind = ind_data.prev_index
        Loop
    End If
    'if searching by primary key
    
    'End If
exitCheck:
    'if we using recordet for deleting then we only need record position
    '   we don't need to parse columns values
    If is_delete <> True Then
        j = 0
        k = 1
        c = 0
        i1 = 0
        Do While k > 0
            'each value is in ' ', so we searching for '
            k = InStr(k, tmp_rowData, "'")
            If k > 0 Then
                c = c + 1
                If c = 2 Then
                    For i = 0 To UBound(msel_cols)
                        'when we find start and end of data, read data and save
                        If msel_cols(i) = j Then
                            fileds_arr(i1) = Mid(tmp_rowData, lastS, k - lastS)
                            i1 = i1 + 1
                        End If
                    Next i
                    'MsgBox fileds_arr(j)
                    c = 0
                    j = j + 1
                Else
                    lastS = k + 1
                End If
                k = k + 1
            End If
        Loop
    End If

End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                
                



'##############################################################
'                  public properties
'##############################################################
'get row count
Public Property Get Rows() As Long
    Rows = record_cnt
End Property
'get column count
Public Property Get Columns() As Long
    Columns = column_cnt
End Property
'get current row
Public Property Get CurrentRow() As Long
    CurrentRow = curr_row
End Property
'get data of selected field
Public Property Get Fields(ByVal fld_index As Long) As String
    '/////////////////////////////////////////////
    Fields = fileds_arr(fld_index)
End Property
'get current row
Public Property Get IsEOF() As Boolean
    If curr_row >= record_cnt - 1 Then
        IsEOF = True
    Else
        IsEOF = False
    End If
End Property

'##############################################################
'                  public functions
'##############################################################
'go to first row
Public Function MoveFirst()
    prev_row = first_start
    curr_row = 0
    set_arr 0
End Function
'go to last row
Public Function MoveLast()
    set_arr record_cnt - 1
End Function
'go to next row
Public Function MoveNext()
    curr_row = curr_row + 1
    If curr_row > record_cnt Then
        MoveLast
        Err.Raise 380, "clsRecordset"
    End If
    set_arr curr_row
End Function
'go to prev row
Public Function MovePrevious()
    Dim tmp_curr As Long
    If curr_row = 0 Then
        Err.Raise 380, "clsRecordset"
    End If
    
    prev_ind = first_ind
    prev_row = first_start
    tmp_curr = curr_row - 1
    curr_row = 0
    
    set_arr tmp_curr
End Function
'move to row
Public Function MoveToRow(ByVal row_index As Long)
    
    If row_index < 0 Or row_index > record_cnt Then
        Err.Raise 380, "clsRecordset"
    Else
        If row_index > curr_row Then
            'curr_row = row_index
            set_arr row_index
        ElseIf row_index < curr_row Then
            prev_ind = first_ind
            prev_row = first_start
            curr_row = 0
            
            set_arr row_index
        End If
    End If
End Function

Private Sub Class_Terminate()
    'free memory
    Erase colIndex_arr
    Erase val_arr
    Erase operator_arr
    Erase condOper_arr
    
    Erase fileds_arr
    Erase ind_cols
    Erase chrInd_pos
    Erase dontUse_arr
    Erase use_arr
End Sub

