VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form90 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create User"
   ClientHeight    =   4425
   ClientLeft      =   2250
   ClientTop       =   2610
   ClientWidth     =   7470
   Icon            =   "frmCreate_User.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7470
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   60
      Left            =   495
      TabIndex        =   13
      Top             =   450
      Visible         =   0   'False
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   106
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   "ID"
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
         Caption         =   "Name"
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
      BeginProperty Column02 
         DataField       =   ""
         Caption         =   "Designation"
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
            DividerStyle    =   6
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   6
            ColumnWidth     =   3135.118
         EndProperty
         BeginProperty Column02 
            DividerStyle    =   6
            ColumnWidth     =   2069.858
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2115
      ScaleHeight     =   240
      ScaleWidth      =   3435
      TabIndex        =   15
      Top             =   3825
      Width           =   3435
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00F3F3F3&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2565
         Picture         =   "frmCreate_User.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Close "
         Top             =   -45
         Width           =   915
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00F3F3F3&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1710
         Picture         =   "frmCreate_User.frx":1700
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   " Clear "
         Top             =   -45
         Width           =   915
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00F3F3F3&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   855
         Picture         =   "frmCreate_User.frx":2582
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "  To insert a new record press clear button prior to write anything  "
         Top             =   -45
         Width           =   915
      End
      Begin VB.CommandButton cmdDel 
         Appearance      =   0  'Flat
         BackColor       =   &H00F3F3F3&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -45
         Picture         =   "frmCreate_User.frx":33D4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " Close "
         Top             =   -45
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00F3F3F3&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   855
         Picture         =   "frmCreate_User.frx":4126
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   " Save "
         Top             =   -45
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1995
      Left            =   435
      TabIndex        =   10
      Top             =   1665
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   3519
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BorderStyle     =   0
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   "Id"
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
         Caption         =   "Name"
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
      BeginProperty Column02 
         DataField       =   ""
         Caption         =   "Permission"
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
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3929.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1110.047
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   450
      TabIndex        =   0
      Top             =   1305
      Width           =   1440
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1980
      TabIndex        =   1
      Top             =   1305
      Width           =   3870
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5940
      TabIndex        =   4
      Top             =   1305
      Width           =   1080
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   180
      Top             =   3465
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   225
      Top             =   3195
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblUser_No 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3690
      TabIndex        =   17
      Top             =   225
      Width           =   3300
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F3F3&
      Caption         =   "Create User"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   2970
      TabIndex        =   16
      Top             =   900
      Width           =   1185
   End
   Begin VB.Label lblNE_User 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F3F3&
      Caption         =   " N.E User"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   495
      TabIndex        =   14
      Top             =   3780
      Width           =   780
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FAEDF1&
      BorderColor     =   &H00FDD9E8&
      Height          =   60
      Index           =   3
      Left            =   1890
      Top             =   630
      Visible         =   0   'False
      Width           =   5145
   End
   Begin VB.Label lblGetEmp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F3F3F3&
      Caption         =   " Get Employee "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4275
      TabIndex        =   12
      Top             =   900
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FAEDF1&
      BorderColor     =   &H00FDD9E8&
      Height          =   2085
      Index           =   2
      Left            =   405
      Top             =   1620
      Width           =   6595
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FAEDF1&
      BorderColor     =   &H00FDD9E8&
      Height          =   285
      Index           =   1
      Left            =   1935
      Top             =   1260
      Width           =   3930
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FAEDF1&
      BorderColor     =   &H00FDD9E8&
      Height          =   285
      Index           =   0
      Left            =   405
      Top             =   1260
      Width           =   1500
   End
   Begin VB.Label lblPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00F3F3F3&
      Caption         =   "  Set Password  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   5655
      TabIndex        =   11
      Top             =   900
      Width           =   1350
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1980
      TabIndex        =   9
      Top             =   945
      Width           =   525
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   405
      TabIndex        =   7
      Top             =   945
      Width           =   405
   End
End
Attribute VB_Name = "Form90"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''Dim Rs As New Recordset
''Dim myrst  As New ADODB.Recordset
''Dim permit   As String
''Dim psl    As String
''
''Private Sub cmdClear_Click()
''    Text1 = ""
''    Text2 = ""
''    Check1.Value = 0
''    Grid_Click False, Me
''End Sub
''
''Private Sub cmdClose_Click()
''    Close_Msg Me
''End Sub
''
''Private Sub cmdDel_Click()
''    opr = "D"
''    cmdSave_Click
''    Me.Text1 = ""
''    Me.Text2 = ""
''    Text1.Locked = False
''End Sub
''
''Private Sub cmdEdit_Click()
''
''    cmdSave_Click
''
''
''End Sub
''
''Private Sub cmdSave_Click()
''
''If Text1 = Empty Or Text2 = Empty Then
''    MsgBox "Incomplete Information", vbInformation + vbOKOnly, "Attention"
''    Exit Sub
''End If
''    con.ConnectionString = strCN.Connection
''    con.Open
''    Set cmd.ActiveConnection = con
''
''    cmd.CommandText = "exec pro_soft_pass'" _
''    + Trim(Text1) + "','" _
''    + Trim(Text2) + "',' ','" _
''    + U_Id + "'," + CStr(Check1.Value) + ",'" + opr + "'"
''
''
''    Set Rs = cmd.Execute
''
''    MsgBox Rs!Message, vbExclamation + vbOKOnly
''
''    con.Close
''
''    If opr = "I" Then
''        Get_Emp = 1 ''User has been created and now password can be set
''    End If
''
''        Check1.Value = 0
''
''    Populate_User
''
''    Grid_Click False, Me
''
''    End Sub
''
''
''
''
''Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''
''If Button = 1 Then
''
''    If Adodc1.Recordset.RecordCount > 0 Then
''
''        Text1 = Adodc1.Recordset!U_Id
''        Text2 = Adodc1.Recordset!U_Name
''
''            If Adodc1.Recordset!Cancel = True Then
''                Check1.Value = 1
''            Else
''                Check1.Value = 0
''            End If
''
''        Text1.Locked = True
''
''        Grid_Click True, Me
''
''    End If
''End If
''
''If Button = 2 Then
''    Text1 = ""
''    Text2 = ""
''    Text1.Locked = False
''    Grid_Click False, Me
''End If
''
''End Sub
''
''Private Sub DataGrid2_Click()
''On Error Resume Next
''
''If Adodc2.Recordset.RecordCount > 0 Then
''    Text1 = ""
''    Text2 = ""
''    Text1 = Adodc2.Recordset!Emp_id
''    Text2 = Adodc2.Recordset!nm
''End If
''
''DataGrid2.Visible = False
''Shape1(3).Visible = False
''
''End Sub
''
''Private Sub Form_Load()
''    Grid_Click False, Me
''    Populate_User
''End Sub
''
''
''Private Sub Label4_Click()
''Dim resp As String
''
''resp = MsgBox("Are you sure you want to make all employee user of the software ?", vbQuestion + vbYesNo)
''
''    If resp = vbYes Then
''        MousePointer = 13
''        Create_User
''        Populate_User
''        MousePointer = 1
''    End If
''End Sub
''
''Private Sub lblGetEmp_Click()
''
''DataGrid2.Visible = True
''With DataGrid2
''    .Height = 3340
''    .Width = 6540
''    .Top = 775
''    .Left = 435
''End With
''
''    Shape1(3).Visible = True
''With Shape1(3)
''    .Height = 3405
''    .Width = 6595
''    .Top = 730
''    .Left = 405
''End With
''
''''-----------------------------------------------------------
''    Adodc2.ConnectionString = strCN.Connection
''
''        Adodc2.RecordSource = "exec POP_Emp_NonUser"
''        Adodc2.Refresh
''                If Adodc2.Recordset.RecordCount > 0 Then
''                Adodc2.Recordset.MoveLast
''                End If
''                Set DataGrid2.DataSource = Adodc2
''
''
''                DataGrid2.Columns(0).DataField = "Emp_ID"
''                DataGrid2.Columns(1).DataField = "NM"
''                DataGrid2.Columns(2).DataField = "Emp_desig"
''                DataGrid2.ReBind
''                DataGrid2.Refresh
''''-----------------------------------------------------------
''
''
''End Sub
''
''Private Sub lblNE_User_Click()
''Text1 = ""
''Text2 = ""
''Text1.SetFocus
''End Sub
''
''Private Sub lblPassword_Click()
''
''If Get_Emp <> 1 Then
''    MsgBox "Create user first", vbOKOnly + vbExclamation
''    Exit Sub
''Else
''    Carry_Emp_ID = Me.Text1     'emp_ID
''    pick = Carry_Emp_ID
''    Tflag = True
''    Unload Me
''   ' Form91.Show
''End If
''End Sub
''
''Public Sub Populate_User()
''
''Adodc1.ConnectionString = strCN.Connection
''Adodc1.RecordSource = "Exec POP_User"
''Adodc1.Refresh
''Set DataGrid1.DataSource = Adodc1
''
''    DataGrid1.Columns(0).DataField = "u_id"
''    DataGrid1.Columns(1).DataField = "u_name"
''    DataGrid1.Columns(2).DataField = "permit"
''    DataGrid1.ReBind
''    DataGrid1.Refresh
''
''    lblUser_No = "Number of current user = " & CStr(Adodc1.Recordset.RecordCount)
''
''DataGrid1.Refresh
''End Sub
''
''Private Sub Text2_KeyPress(KeyAscii As Integer)
''If KeyAscii = 13 And Text2 <> "" Then SendKeys Chr(9)
''End Sub
''
''
''Private Sub Text1_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 And Text1 <> "" Then
''
''        SendKeys Chr(9)
''
''        If Len(Trim(Text1)) > 3 Then
''            Find_User
''        End If
''
''    End If
''End Sub
''
''Public Sub Find_User()
''        Adodc1.Recordset.MoveFirst
''        Dim i As Integer
''        For i = 0 To Adodc1.Recordset.RecordCount - 1
''
''            If Trim(Text1) = Trim(DataGrid1.Columns(0).CellValue(DataGrid1.Bookmark)) Then
''                DataGrid1.SetFocus
''                Exit Sub
''            Else
''            Adodc1.Recordset.MoveNext
''
''            End If
''        Next i
''
''End Sub
''Public Sub Create_User()
''
''    con.ConnectionString = strCN.Connection
''    con.Open
''
''        Set cmd.ActiveConnection = con
''
''        cmd.CommandText = "exec Create_User_AtOnce"
''
''        Set Rs = cmd.Execute
''
''        MsgBox Rs!Message, vbExclamation + vbOKOnly
''
''    con.Close
''
''End Sub
