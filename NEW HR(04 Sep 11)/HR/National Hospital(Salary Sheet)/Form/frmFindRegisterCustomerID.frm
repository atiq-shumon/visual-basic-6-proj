VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFindRegisterCustomerID 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Finde Customer ID"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Searched By Your Choice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   6015
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "By Employee Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   300
         Width           =   2235
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "By Employee ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   650
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   330
         TabIndex        =   4
         Top             =   675
         Width           =   495
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   1335
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFindRegisterCustomerID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intputsel1, colval1 As Integer
Public MyRecoredSet As New Recordset
Public oWnerForm1 As Form
Dim frmRecordset As New Recordset
Dim mbCtrlKey As Integer
Dim AccountInitial$
Dim mbSortCol As String
Private Sub grdDataGrid_DblClick()
On Error GoTo ErrDes
DataSelect
Exit Sub
ErrDes:
    MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub grdDataGrid_HeadClick(ByVal ColIndex As Integer)
On Error GoTo ErrDes
    If mbCtrlKey Then
        msSortCol = "[" & MyRecoredSet(ColIndex).Name & "] desc"
        mbCtrlKey = 0 'reset it
    Else
        msSortCol = "[" & MyRecoredSet(ColIndex).Name & "]"
    End If
    MyRecoredSet.Sort = msSortCol
    msSortCol = vbNullString 'reset it
Exit Sub
ErrDes:
    MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub grdDataGrid_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrDes
If KeyCode = vbKeyReturn Then DataSelect
If KeyCode = vbKeyEscape Then Unload Me
Exit Sub
ErrDes:
    MsgBox Err.Description, vbInformation, App.Title
End Sub
Public Sub DataSelect()
On Error GoTo DataSelectError
    Dim varData As Variant
    Dim showtexData(5) As Variant
    
    If (MyRecoredSet.RecordCount <> 0) Or (Not (MyRecoredSet.EOF Or MyRecoredSet.BOF)) Then
        varData = MyRecoredSet(0)
        oWnerForm1.txtEmp_Name = varData
        Unload Me
        oWnerForm1.txtEmp_Name.SetFocus
    End If
    colval1 = 0
    Exit Sub
DataSelectError:
    MsgBox Err.Description
End Sub

Private Sub Option1_Click()
On Error GoTo ErrDes
Label1.Caption = "Employee Name"
Text2.Text = ""
Text2.SetFocus
Exit Sub
ErrDes:
    MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub Option2_Click()
On Error GoTo ErrDes
Label1.Caption = "Employee Name"
Text2.Text = ""
Text2.SetFocus
Exit Sub
ErrDes:
    MsgBox Err.Description, vbInformation, App.Title
End Sub

Public Sub DataQuery()
On Error GoTo ErrDes
Set MyRecoredSet = New Recordset
Dim Hints$, SQLStatement$, AccNoLike$

If Option1.Value = True Then
    Hints = "%" + Text2.Text + "%"
    SQLStatement = " SELECT  Emp_id from Emp_info where  (Emp_ID" + _
                          " Like '" & Hints & "')"
ElseIf Option2.Value = True Then
      Hints = "%" + Text2.Text + "%"
      SQLStatement = "SELECT EMP_NM from Emp_Info where  (Emp_Nm" + _
                          " Like '" & Hints & "')"
End If

    MyRecoredSet.CursorLocation = adUseClient
   'MyRecoredSet.Open SQLStatement, gsConnect
    MyRecoredSet.Open SQLStatement, strCN.Connection_String
    Set frmRecordset = MyRecoredSet
    Set grdDataGrid.DataSource = frmRecordset
    grdDataGrid.Columns(0).Width = 4500
                'grdDataGrid.Columns(1).Width = 2000
Exit Sub
ErrDes:
    MsgBox Err.Description, vbInformation, App.Title
End Sub



Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrDes
If KeyCode = vbKeyReturn Then
    DataQuery
    grdDataGrid.SetFocus
ElseIf KeyCode = vbKeyEscape Then
    Unload Me
End If
Exit Sub
ErrDes:
    MsgBox Err.Description, vbInformation, App.Title
End Sub
