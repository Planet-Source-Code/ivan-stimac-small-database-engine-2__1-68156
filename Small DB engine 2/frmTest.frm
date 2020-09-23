VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12405
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   46
      Top             =   7020
      Width           =   12075
   End
   Begin VB.Frame frmWriteNew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Write new data"
      Height          =   6315
      Left            =   -11100
      TabIndex        =   39
      Top             =   480
      Visible         =   0   'False
      Width           =   12075
      Begin VB.CommandButton buttAddRow 
         Caption         =   "New Row"
         Height          =   495
         Left            =   5580
         TabIndex        =   44
         Top             =   5580
         Width           =   1995
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   43
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton buttWrite 
         Caption         =   "Write"
         Height          =   495
         Left            =   9780
         TabIndex        =   41
         Top             =   5580
         Width           =   1995
      End
      Begin VB.CommandButton buttCancelWrite 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   7680
         TabIndex        =   40
         Top             =   5580
         Width           =   1995
      End
      Begin VB.Label lblColName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "##"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   42
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame frmNewTable 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create new table"
      Height          =   6315
      Left            =   11760
      TabIndex        =   24
      Top             =   420
      Visible         =   0   'False
      Width           =   12075
      Begin VB.CommandButton Command7 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   7680
         TabIndex        =   36
         Top             =   5580
         Width           =   1995
      End
      Begin VB.CommandButton buttAddTable 
         Caption         =   "Create table"
         Height          =   495
         Left            =   9780
         TabIndex        =   35
         Top             =   5580
         Width           =   1995
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Delete column"
         Height          =   375
         Left            =   5700
         TabIndex        =   34
         Top             =   3180
         Width           =   1815
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add new column"
         Height          =   2055
         Left            =   5700
         TabIndex        =   29
         Top             =   1080
         Width           =   5175
         Begin VB.CheckBox chIndexed 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Select as indexed row"
            Height          =   255
            Left            =   1560
            TabIndex        =   54
            Top             =   1140
            Width           =   3075
         End
         Begin VB.CommandButton buttAddColumn 
            Caption         =   "Add"
            Height          =   375
            Left            =   3540
            TabIndex        =   33
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CheckBox chPK 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Primary key (auto increment)"
            Height          =   255
            Left            =   1560
            TabIndex        =   32
            Top             =   840
            Width           =   3075
         End
         Begin VB.TextBox txtColName 
            Height          =   285
            Left            =   1560
            TabIndex        =   31
            Top             =   420
            Width           =   3255
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Column name:"
            Height          =   315
            Left            =   180
            TabIndex        =   30
            Top             =   420
            Width           =   3015
         End
      End
      Begin VB.ListBox lstCols 
         Height          =   4545
         ItemData        =   "frmTest.frx":0000
         Left            =   1500
         List            =   "frmTest.frx":000D
         TabIndex        =   28
         Top             =   1140
         Width           =   4095
      End
      Begin VB.TextBox txtTblName 
         Height          =   285
         Left            =   1500
         TabIndex        =   26
         Top             =   540
         Width           =   4155
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Columns:"
         Height          =   255
         Left            =   300
         TabIndex        =   27
         Top             =   960
         Width           =   3315
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Table name:"
         Height          =   255
         Left            =   300
         TabIndex        =   25
         Top             =   540
         Width           =   3315
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data"
      Height          =   3195
      Left            =   180
      TabIndex        =   21
      Top             =   3780
      Width           =   12075
      Begin VB.CommandButton Command12 
         Caption         =   "Drop table 'test2'"
         Height          =   315
         Left            =   9660
         TabIndex        =   56
         Top             =   300
         Width           =   1875
      End
      Begin VB.TextBox txtWriteCnt 
         Height          =   315
         Left            =   5760
         TabIndex        =   52
         Text            =   "20000"
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Write rows"
         Height          =   375
         Left            =   4200
         TabIndex        =   51
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Left            =   9240
         TabIndex        =   50
         Text            =   "*"
         Top             =   1740
         Width           =   2175
      End
      Begin VB.TextBox txtWhere 
         Height          =   285
         Left            =   9240
         TabIndex        =   48
         Text            =   "WHERE username = 'cyber'"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Read  data"
         Height          =   435
         Left            =   9060
         TabIndex        =   45
         Top             =   2580
         Width           =   2415
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Delete data"
         Height          =   435
         Left            =   9780
         TabIndex        =   38
         Top             =   1020
         Width           =   1575
      End
      Begin VB.CommandButton buttWriteNew 
         Caption         =   "Write new data"
         Height          =   435
         Left            =   8040
         TabIndex        =   37
         Top             =   1020
         Width           =   1635
      End
      Begin VB.ComboBox cmbTbls2 
         Height          =   315
         Left            =   300
         TabIndex        =   23
         Text            =   "Combo1"
         Top             =   300
         Width           =   3675
      End
      Begin VB.ListBox lstResult 
         Height          =   2205
         Left            =   300
         TabIndex        =   22
         Top             =   720
         Width           =   7575
      End
      Begin VB.Label lblDataCnt 
         Caption         =   "Label12"
         Height          =   195
         Left            =   8040
         TabIndex        =   53
         Top             =   720
         Width           =   2355
      End
      Begin VB.Label Label11 
         Caption         =   "Read fields"
         Height          =   195
         Left            =   8100
         TabIndex        =   49
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label Label10 
         Caption         =   "Read where"
         Height          =   195
         Left            =   8100
         TabIndex        =   47
         Top             =   2220
         Width           =   1755
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Opened database"
      Height          =   3495
      Left            =   5220
      TabIndex        =   16
      Top             =   120
      Width           =   6975
      Begin VB.Frame Frame4 
         Caption         =   "Tables and columns"
         Height          =   2595
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Width           =   6615
         Begin VB.CommandButton Command2 
            Caption         =   "Create new table"
            Height          =   435
            Left            =   4020
            TabIndex        =   20
            Top             =   300
            Width           =   2415
         End
         Begin VB.ListBox lstTblCols 
            Height          =   1620
            Left            =   180
            TabIndex        =   19
            Top             =   780
            Width           =   3675
         End
         Begin VB.ComboBox cmbTables 
            Height          =   315
            Left            =   180
            TabIndex        =   18
            Text            =   "Combo1"
            Top             =   300
            Width           =   3675
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Open database"
      Height          =   1635
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   4935
      Begin VB.CommandButton Command11 
         Caption         =   "Export to SQL"
         Height          =   375
         Left            =   3060
         TabIndex        =   55
         Top             =   1140
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Text            =   "\testDb.txt"
         Top             =   300
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OpenDB"
         Height          =   375
         Left            =   3060
         TabIndex        =   12
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtUserName2 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   780
         Width           =   1755
      End
      Begin VB.TextBox txtPassword2 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         Top             =   1140
         Width           =   1755
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "File name:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "User name:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "New database"
      Height          =   1635
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtPassword 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1140
         Width           =   1755
      End
      Begin VB.TextBox txtUserName 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   780
         Width           =   1755
      End
      Begin VB.CommandButton buttCreateDb 
         Caption         =   "CreateDB"
         Height          =   675
         Left            =   3060
         TabIndex        =   3
         Top             =   780
         Width           =   1575
      End
      Begin VB.TextBox txtFileName 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Text            =   "\testDb.txt"
         Top             =   300
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "User name:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File name:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1035
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim mDB As New clsDatabaseEngine
Dim mRecSet As clsRecordset
Dim Rows As Integer

Private Sub buttAddColumn_Click()
    If Me.txtColName <> "" Then
        If Me.chPK.Value = 1 Then
            Me.lstCols.AddItem "*" & Me.txtColName.Text
            Me.chPK.Enabled = False
            Me.chPK.Value = 0
        ElseIf Me.chIndexed.Value = 1 Then
            Me.lstCols.AddItem "#" & Me.txtColName.Text
            Me.chPK.Enabled = False
            Me.chPK.Value = 0
        Else
            Me.lstCols.AddItem Me.txtColName.Text
        End If
        
        Me.txtColName.Text = ""
        Me.txtColName.SetFocus
    End If
End Sub

Private Sub buttAddRow_Click()
    Rows = Rows + 1
    createGrid Rows
End Sub

Private Sub buttAddTable_Click()
    Dim strQuery As String
    Dim i As Integer, ret As Long
    If Me.txtTblName.Text <> "" Then
        strQuery = "CREATE TABLE " & Me.txtTblName.Text & " ("
        For i = 0 To Me.lstCols.ListCount - 1
            If i > 0 Then strQuery = strQuery & ","
            strQuery = strQuery & Me.lstCols.List(i)
        Next i
        strQuery = strQuery & ")"
        ret = mDB.ExecuteSql(strQuery)
        Me.List1.AddItem strQuery
        If ret = 0 Then
            frmNewTable.Visible = False
            Me.chPK.Enabled = True
            mDB.CloseDB
            Command1_Click
        Else
            MsgBox "Can't create table!"
        End If
    Else
        MsgBox "Please enter table name!", vbExclamation
    End If
End Sub

Private Sub buttCancelWrite_Click()
    frmWriteNew.Visible = False
End Sub

Private Sub buttCreateDb_Click()
    Dim ret As Long
    ret = mDB.CreateDB(App.Path & Me.txtFileName.Text, "username=" & Me.txtUserName.Text & ";password=" & Me.txtPassword.Text)
    
    If ret = 0 Then
        MsgBox "Create success!", vbInformation
    Else
        MsgBox "Can't write to file! Please check file name!", vbCritical
    End If
End Sub

Private Sub buttWrite_Click()
    Dim i As Integer, j As Integer
    Dim strTmp As String, strQuery As String
    Dim ret As Long
    strTmp = "INSERT INTO " & Me.cmbTbls2.List(Me.cmbTbls2.ListIndex) & " ("
    For i = 0 To Me.lblColName.Count - 1
        If i > 0 Then strTmp = strTmp & ","
        If Left$(Me.lblColName(i).Caption, 1) <> "#" Then
            strTmp = strTmp & Me.lblColName(i).Caption
        Else
            strTmp = strTmp & Mid(Me.lblColName(i).Caption, 2)
        End If
    Next i
    strTmp = strTmp & ") VALUES ("
    strQuery = strTmp

    For i = 0 To Me.txtData.Count - 1
        
        If j = Me.lblColName.Count Then
            j = 0
            strQuery = strQuery & ")"
            ret = mDB.ExecuteSql(strQuery)
            Me.List1.AddItem strQuery
            strQuery = strTmp
            strQuery = strQuery & "'" & Me.txtData(i).Text & "'"
        Else
            If j > 0 Then strQuery = strQuery & ","
            strQuery = strQuery & "'" & Me.txtData(i).Text & "'"
        End If
        j = j + 1
        If Me.txtData(i).Text = "" Then Exit For
    Next i
    If ret = 0 Then
        frmWriteNew.Visible = False
    Else
        MsgBox "Maybe some data is not writen!"
    End If
    mDB.RefreshDB
End Sub

Private Sub buttWriteNew_Click()
   ' mDB.ExecuteSql "INSERT INTO proba (prezime, ime) VALUES ('ivan','stimac')"
    createGrid 1, True
    Me.frmWriteNew.Visible = True
End Sub

Private Sub cmbTables_Click()
    Dim i As Integer, j As Integer
    Me.lstTblCols.Clear
    For i = 0 To mDB.ColCount(Me.cmbTables.ListIndex) - 1
        If i <> mDB.PrimaryKey(Me.cmbTables.ListIndex) Then
            Me.lstTblCols.AddItem mDB.ColName(Me.cmbTables.ListIndex, i)
        Else
            Me.lstTblCols.AddItem "*" & mDB.ColName(Me.cmbTables.ListIndex, i)
        End If
        'select indexed columns
        For j = 0 To mDB.IndexedColumnsCount(Me.cmbTables.ListIndex) - 1
            If mDB.IndexedColumn(Me.cmbTables.ListIndex, j) = i Then
                Me.lstTblCols.List(Me.lstTblCols.ListCount - 1) = "#" & Me.lstTblCols.List(Me.lstTblCols.ListCount - 1)
            End If
        Next j
    Next i
End Sub


Private Sub cmbTbls2_Click()
    Me.cmbTables.ListIndex = Me.cmbTbls2.ListIndex
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Dim ret As Long, i As Integer
    ret = mDB.OpenDB(App.Path & Me.txtFileName.Text, "username=" & Me.txtUserName2.Text & ";password=" & Me.txtPassword2.Text)
    
    If ret = 0 Then
        Me.cmbTables.Clear
        Me.cmbTbls2.Clear
        For i = 0 To mDB.TablesCount - 1
            Me.cmbTables.AddItem mDB.TableName(i)
            Me.cmbTbls2.AddItem mDB.TableName(i)
        Next i
        Me.cmbTables.ListIndex = 0
        Me.cmbTbls2.ListIndex = 0
    ElseIf ret = -1 Then
        MsgBox "Invalid database file!", vbExclamation
    ElseIf ret = -2 Then
        MsgBox "Fail to open file!", vbCritical
    Else
        MsgBox "Invalid username or password!", vbExclamation
    End If
End Sub

Private Sub Command10_Click()
    Dim i As Long
    
    Dim tick As Long
    tick = GetTickCount
    For i = 1 To Val(txtWriteCnt.Text)
       mDB.ExecuteSql ("INSERT INTO " & Me.cmbTbls2.List(Me.cmbTbls2.ListIndex) & " (username,password) VALUES ('test','test')")
    Next i
    mDB.RefreshDB
    tick = GetTickCount - tick
    MsgBox "Elapsed " & Format$(tick / 1000, "0.000") & " second", 32
End Sub



Private Sub Command11_Click()
    Dim mExport As New clsExport
    mExport.exportToSql App.Path & Me.txtFileName.Text, "username=" & Me.txtUserName2.Text & ";password=" & Me.txtPassword2.Text, App.Path & "\export.txt"
End Sub

Private Sub Command12_Click()
    mDB.ExecuteSql "DROP TABLE test2"
End Sub

Private Sub Command2_Click()
'    Dim mTblName As String
'    Dim ret As Long
'    mTblName = InputBox("Please enter table name")
'    If mTblName <> "" Then
'        ret=mdb.ExecuteSql("CREATE TABLE
'    End If
    frmNewTable.Visible = True
End Sub

Private Sub Command6_Click()
    If Me.lstCols.ListIndex >= 0 Then
        If Mid(Me.lstCols.List(Me.lstCols.ListIndex), 1, 1) = "*" Then chPK.Enabled = True
        Me.lstCols.RemoveItem Me.lstCols.ListIndex
    End If
End Sub

Private Sub Command7_Click()
    frmNewTable.Visible = False
End Sub


Private Sub Command8_Click()
    'On Error Resume Next
    Dim i As Long, j As Long
    Dim tmp_str As String
    Dim tick As Long
    tick = GetTickCount
    
    
    Set mRecSet = mDB.OpenRecordSet("SELECT " & Me.txtFields.Text & " FROM " & Me.cmbTbls2.List(Me.cmbTbls2.ListIndex) & " " & Me.txtWhere.Text)     ' WHERE ime = 'ivan' AND prezime = 'stimac' OR ID = '10'"
    
    tick = GetTickCount - tick
    MsgBox "Elapsed " & Format$(tick / 1000, "0.000") & " second", 32
    
    Me.lblDataCnt.Caption = mRecSet.Rows
    
    Me.List1.AddItem "SELECT " & Me.txtFields.Text & " FROM " & Me.cmbTbls2.List(Me.cmbTbls2.ListIndex) & " " & Me.txtWhere.Text    ' WHERE ime = 'ivan' AND prezime = 'stimac' OR ID = '10'"
    Me.lstResult.Clear
    
   ' MsgBox "IDE1"
    mRecSet.MoveFirst
    For i = 0 To mRecSet.Rows - 1
        tmp_str = ""
       ' MsgBox "IDE_i:" & i
        For j = 0 To mRecSet.Columns - 1
            tmp_str = tmp_str & mRecSet.Fields(j) & "          "  '(i, j) & "       "
        Next j
        mRecSet.MoveNext
        Me.lstResult.AddItem tmp_str
    Next i
End Sub

Private Sub Command9_Click()
    mDB.ExecuteSql "DELETE FROM " & Me.cmbTbls2.List(Me.cmbTbls2.ListIndex) & " " & txtWhere.Text
End Sub

Private Sub Form_Load()
    frmNewTable.Left = 120
    frmWriteNew.Left = 120
End Sub


Private Sub createGrid(ByVal rowNum As Integer, Optional resetValues As Boolean = False)
    Dim i As Integer, z As Integer, colNum As Integer
    Dim mY As Long, mX As Long
    Rows = rowNum
    colNum = 0
    For i = 0 To Me.lstTblCols.ListCount - 1
        If Mid(Me.lstTblCols.List(i), 1, 1) <> "*" Then colNum = colNum + 1
    Next i
    
    If resetValues = True Then
        For i = 0 To Me.txtData.Count - 1
            If i > 0 Then Unload Me.txtData(i)
        Next i
        Me.txtData(0).Text = ""
    Else
        If rowNum * colNum < Me.txtData.Count Then
            For i = rowNum * colNum + 1 To Me.txtData.Count - 1
                Unload Me.txtData(i)
            Next i
        End If
    End If
    
    For i = 1 To Me.lblColName.Count - 1
        If i > 0 Then Unload Me.lblColName(i)
    Next i
    
    For i = 1 To colNum - 1
        Load Me.lblColName(i)
        Me.lblColName(i).Visible = True
    Next i
    DoEvents
    
    
    For i = 1 To (colNum) * (rowNum) - 1
        If txtData.Count - 1 < i Then
            Load Me.txtData(i)
            Me.txtData(i).Text = ""
        End If
        Me.txtData(i).Visible = True
    Next i
    
    For i = 1 To colNum - 1
        Me.lblColName(i).Left = Me.lblColName(i - 1).Left + Me.lblColName(i).Width
    Next i
    
    z = 0
    For i = 0 To Me.lstTblCols.ListCount - 1
        If Mid(Me.lstTblCols.List(i), 1, 1) <> "*" Then
            Me.lblColName(z).Caption = Me.lstTblCols.List(i)
            z = z + 1
        End If
    Next i
    
    mY = Me.txtData(0).Top
    mX = Me.txtData(0).Left
    z = 0
    For i = 1 To Me.txtData.Count - 1
        z = z + 1
        mX = mX + Me.txtData(i - 1).Width
        
        If z = colNum Then
            mY = mY + Me.txtData(i).Height
            mX = Me.txtData(0).Left
            z = 0
        End If
        
        Me.txtData(i).Top = mY
        Me.txtData(i).Left = mX
    Next i
End Sub

