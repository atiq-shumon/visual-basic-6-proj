VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form93 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Software privilege"
   ClientHeight    =   5640
   ClientLeft      =   1710
   ClientTop       =   1815
   ClientWidth     =   8730
   Icon            =   "frmSoftware privilege.frx":0000
   LinkTopic       =   "Form22"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8730
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   4350
      Picture         =   "frmSoftware privilege.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4905
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   3150
      Picture         =   "frmSoftware privilege.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4905
      Width           =   1185
   End
   Begin VB.ComboBox cmbEmp_id 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      ItemData        =   "frmSoftware privilege.frx":3CDE
      Left            =   5535
      List            =   "frmSoftware privilege.frx":3CE0
      TabIndex        =   0
      Top             =   270
      Width           =   2265
   End
   Begin VB.ComboBox cboAccess_Area 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      ItemData        =   "frmSoftware privilege.frx":3CE2
      Left            =   1485
      List            =   "frmSoftware privilege.frx":3CE4
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2445
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmSoftware privilege.frx":3CE6
      Height          =   3480
      Left            =   405
      TabIndex        =   11
      Top             =   1260
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6138
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      ColumnHeaders   =   0   'False
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   0
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
      Caption         =   "Available Previleges"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "scr_no"
         Caption         =   "Screen"
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
         DataField       =   "descript"
         Caption         =   "Privileges"
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
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5609.764
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmSoftware privilege.frx":3CFB
      Height          =   3480
      Left            =   4770
      TabIndex        =   8
      Top             =   1260
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6138
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ColumnHeaders   =   0   'False
      ForeColor       =   192
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   0
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
      Caption         =   "Given Previleges"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "scr_no"
         Caption         =   "Screen no."
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
         DataField       =   "descript"
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            DividerStyle    =   0
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   0
            ColumnWidth     =   4034.835
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   6570
      Top             =   2250
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
      Caption         =   "Adodc3"
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
      Left            =   5670
      Top             =   2250
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4815
      Top             =   2250
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
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3390
      Left            =   4005
      ScaleHeight     =   3390
      ScaleWidth      =   690
      TabIndex        =   12
      Top             =   1305
      Width           =   690
      Begin VB.CommandButton cmdAll_Out 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   615
         Left            =   45
         Picture         =   "frmSoftware privilege.frx":3D10
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "  Delete all permission  "
         Top             =   2565
         Width           =   600
      End
      Begin VB.CommandButton cmdSingle_Out 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   615
         Left            =   45
         Picture         =   "frmSoftware privilege.frx":401A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "  Delete permission  "
         Top             =   1755
         Width           =   600
      End
      Begin VB.CommandButton cmdAll_In 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   615
         Left            =   45
         Picture         =   "frmSoftware privilege.frx":4324
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "  Give all permission  "
         Top             =   945
         Width           =   600
      End
      Begin VB.CommandButton cmdSingle_In 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   615
         Left            =   45
         Picture         =   "frmSoftware privilege.frx":462E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "  Give permission  "
         Top             =   135
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7830
      ScaleHeight     =   375
      ScaleWidth      =   510
      TabIndex        =   9
      Top             =   225
      Width           =   510
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Height          =   435
         Index           =   0
         Left            =   -45
         Picture         =   "frmSoftware privilege.frx":4938
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   6165
      Top             =   2295
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
      Caption         =   "Adodc3"
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Access Area"
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
      Index           =   2
      Left            =   405
      TabIndex        =   14
      Top             =   765
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Index           =   0
      Left            =   4860
      TabIndex        =   13
      Top             =   765
      Width           =   405
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FECCC7&
      BorderColor     =   &H00FECCC7&
      Height          =   3570
      Index           =   4
      Left            =   360
      Top             =   1215
      Width           =   7980
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FECCC7&
      BorderColor     =   &H00FECCC7&
      Height          =   330
      Index           =   2
      Left            =   5535
      Top             =   720
      Width           =   2805
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
      Index           =   1
      Left            =   4860
      TabIndex        =   7
      Top             =   315
      Width           =   525
   End
   Begin VB.Label txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   255
      Left            =   5670
      TabIndex        =   6
      Top             =   765
      Width           =   2655
   End
End
Attribute VB_Name = "Form93"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New Connection
Dim cmd As New Command
Dim strPortion As String

Private Sub cboAccess_Area_Change()
    POP_Previleges
End Sub

Private Sub cboAccess_Area_Click()
    POP_Previleges
End Sub

Private Sub cmdAll_In_Click()
On Error Resume Next
    
    If txtName = "" Then
        Invalid_User_Msg
        cmbEmp_id.SetFocus
        Exit Sub
    End If

    con.ConnectionString = strCN.Connection
    con.Open
    cmd.CommandText = "exec give_pmt 'All','" + strPortion + "'," + str(Adodc1.Recordset!code) + ",'" + Trim(cmbEmp_id) + "'"
    cmd.ActiveConnection = con
    cmd.Execute
    con.Close
    
    POP_Previleges
    
End Sub

Private Sub cmdAll_Out_Click()
On Error Resume Next
    
    If txtName = "" Then
        Invalid_User_Msg
        cmbEmp_id.SetFocus
        Exit Sub
    End If

    con.ConnectionString = strCN.Connection
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandText = "exec size_pmt 'All','" + strPortion + "'," + CStr(Adodc2.Recordset!code) + ",'" + cmbEmp_id + "'"
    cmd.Execute
    con.Close
    
    POP_Previleges

End Sub

Private Sub cmdClear_Click()
    cmbEmp_id = ""
    txtName = ""
    cmbEmp_id.Refresh
    cmbEmp_id.SetFocus
    
    cboAccess_Area = "Regardless"
    POP_Previleges
End Sub

Private Sub cmdClose_Click()
Close_Msg Me
End Sub

Private Sub cmdSingle_In_Click()
On Error Resume Next

    If txtName = "" Then
        Invalid_User_Msg
        cmbEmp_id.SetFocus
        Exit Sub
    End If
    
    con.ConnectionString = strCN.Connection
    con.Open
    cmd.CommandText = "exec give_pmt 'Single','" + strPortion + "'," + str(Adodc1.Recordset!code) + ",'" + Trim(cmbEmp_id) + "'"
    cmd.ActiveConnection = con
    cmd.Execute
    con.Close
    
    POP_Previleges
End Sub

Private Sub cmdSingle_Out_Click()
On Error Resume Next
    
    If txtName = "" Then
        Invalid_User_Msg
        cmbEmp_id.SetFocus
        Exit Sub
    End If
    
    con.ConnectionString = strCN.Connection
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandText = "exec size_pmt 'Single','" + strPortion + "'," + CStr(Adodc2.Recordset!code) + ",'" + Trim(cmbEmp_id) + "'"
    cmd.Execute
    con.Close
    
    POP_Previleges
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
    POP_Previleges
    POP_User
    
    
    With cboAccess_Area
        .AddItem "Regardless"
        .AddItem "Employee Information"
        .AddItem "Payroll"
        .AddItem "Provident Fund"
        .AddItem "Loan"
        .AddItem "General Setup"
        .AddItem "Reports"
        .AddItem "Software Security"
        .ListIndex = 0
    End With
    
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub cmbEmp_id_KeyPress(KeyAscii As Integer)
On Error Resume Next

If cmbEmp_id <> "" And KeyAscii = 13 Then

    Adodc3.ConnectionString = strCN.Connection
    Adodc3.RecordSource = "exec POP_UserName_SPrivilege '" + Trim(cmbEmp_id) + " '"
    Adodc3.Refresh
    
    If Adodc3.Recordset.RecordCount > 0 Then
        txtName = Adodc3.Recordset!U_Name
    Else
        txtName = ""
    End If
    
    POP_Previleges
    cboAccess_Area.SetFocus

End If
End Sub


Public Sub POP_Previleges()
On Error Resume Next

Get_Access_Area

Adodc1.ConnectionString = strCN.Connection
Adodc1.RecordSource = "exec Software_Previleges 'Not_In','" + Trim(cmbEmp_id) + "','" + strPortion + "'"
Adodc1.Refresh

Adodc2.ConnectionString = strCN.Connection
Adodc2.RecordSource = "exec Software_Previleges 'In','" + Trim(cmbEmp_id) + "','" + strPortion + "'"
Adodc2.Refresh

    
    
End Sub

Public Sub Get_Access_Area()
    Select Case Trim(cboAccess_Area)
        Case "Regardless"
            strPortion = "ALL"
        Case "Employee Information"
            strPortion = "E"
        Case "Attendance"
            strPortion = "A"
        Case "Loan"
            strPortion = "L"
        Case "Movement"
            strPortion = "M"
        Case "Provident Fund"
            strPortion = "PF"
        Case "Payroll"
            strPortion = "P"
        Case "Software Security"
            strPortion = "S"
        Case "Reports"
            strPortion = "R"
        Case "General Setup"
            strPortion = "T"
    End Select

End Sub

Public Sub POP_User()

    Adodc4.ConnectionString = strCN.Connection
    Adodc4.RecordSource = "exec POP_User"
    Adodc4.Refresh
    
    cmbEmp_id.Clear
    Adodc4.Recordset.MoveFirst
    
    Do Until Adodc4.Recordset.EOF = True
        
        cmbEmp_id.AddItem Adodc4.Recordset!U_Id
        Adodc4.Recordset.MoveNext
    Loop
    
    
    

End Sub

Public Sub Invalid_User_Msg()
    MsgBox "Invalid user or user not selected!", vbOKOnly + vbInformation, "Message"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub
