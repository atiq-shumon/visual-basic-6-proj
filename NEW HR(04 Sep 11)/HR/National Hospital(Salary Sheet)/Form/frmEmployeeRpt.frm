VERSION 5.00
Begin VB.Form frmEmployeeRpt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employee Information Report"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7680
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Salary Disbursee Staff"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   12
      Top             =   2400
      Width           =   7515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6255
      Picture         =   "frmEmployeeRpt.frx":0000
      TabIndex        =   9
      Top             =   5880
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&View Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5010
      Picture         =   "frmEmployeeRpt.frx":08CA
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Report Type "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5610
      Left            =   15
      TabIndex        =   1
      Top             =   -15
      Width           =   7515
      Begin VB.ComboBox cboYear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5700
         TabIndex        =   23
         Text            =   "Combo2"
         Top             =   5010
         Width           =   1605
      End
      Begin VB.ComboBox cboMonth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4020
         TabIndex        =   22
         Text            =   "Combo2"
         Top             =   5010
         Width           =   1635
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         ItemData        =   "frmEmployeeRpt.frx":1194
         Left            =   3990
         List            =   "frmEmployeeRpt.frx":119E
         TabIndex        =   21
         Top             =   3870
         Width           =   3345
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         ItemData        =   "frmEmployeeRpt.frx":11B0
         Left            =   3990
         List            =   "frmEmployeeRpt.frx":11C0
         TabIndex        =   20
         Top             =   3435
         Width           =   3345
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         ItemData        =   "frmEmployeeRpt.frx":11F4
         Left            =   3990
         List            =   "frmEmployeeRpt.frx":11F6
         TabIndex        =   19
         Top             =   4275
         Width           =   3345
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   7
         Left            =   3990
         TabIndex        =   18
         Top             =   2655
         Width           =   3345
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Designation wise Employee's Report"
         Enabled         =   0   'False
         Height          =   465
         Index           =   5
         Left            =   270
         TabIndex        =   17
         Top             =   4200
         Width           =   3045
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sex wise Employee's Report"
         Enabled         =   0   'False
         Height          =   465
         Index           =   6
         Left            =   270
         TabIndex        =   16
         Top             =   3795
         Width           =   2445
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Class Wise Employee's Report"
         Enabled         =   0   'False
         Height          =   465
         Index           =   7
         Left            =   270
         TabIndex        =   15
         Top             =   3360
         Width           =   2445
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Employees"
         Enabled         =   0   'False
         Height          =   465
         Index           =   8
         Left            =   270
         TabIndex        =   14
         Top             =   2970
         Width           =   2445
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Deaprtmentwise All Employee"
         Height          =   465
         Index           =   9
         Left            =   270
         TabIndex        =   13
         Top             =   2580
         Width           =   2445
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         ItemData        =   "frmEmployeeRpt.frx":11F8
         Left            =   4020
         List            =   "frmEmployeeRpt.frx":1202
         TabIndex        =   11
         Top             =   1665
         Width           =   3345
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         ItemData        =   "frmEmployeeRpt.frx":1214
         Left            =   4020
         List            =   "frmEmployeeRpt.frx":1224
         TabIndex        =   10
         Top             =   1230
         Width           =   3345
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "frmEmployeeRpt.frx":1258
         Left            =   4020
         List            =   "frmEmployeeRpt.frx":125A
         TabIndex        =   7
         Top             =   2070
         Width           =   3345
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   4020
         TabIndex        =   0
         Top             =   450
         Width           =   3345
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Designation wise Employee's Report"
         Height          =   465
         Index           =   4
         Left            =   300
         TabIndex        =   6
         Top             =   1995
         Width           =   3045
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sex wise Employee's Report"
         Height          =   465
         Index           =   3
         Left            =   300
         TabIndex        =   5
         Top             =   1590
         Width           =   2445
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Class Wise Employee's Report"
         Height          =   465
         Index           =   2
         Left            =   300
         TabIndex        =   4
         Top             =   1155
         Width           =   2445
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Employees"
         Height          =   465
         Index           =   1
         Left            =   300
         TabIndex        =   3
         Top             =   765
         Width           =   2445
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Deaprtmentwise All Employee"
         Height          =   465
         Index           =   0
         Left            =   300
         TabIndex        =   2
         Top             =   375
         Value           =   -1  'True
         Width           =   2445
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5940
         TabIndex        =   25
         Top             =   4740
         Width           =   690
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Month : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4170
         TabIndex        =   24
         Top             =   4740
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmEmployeeRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Errdes
On Error GoTo Errdes

If Option1(0).Value = True Then
    ReportStatusofEmployee = 0
    DEPARMENTNAMEFORTPT = Trim(Combo1(0))
    SEXFORREPORT = 5
    DESIGNATIONFORRPT = ""
    CheckStatusofEmployee = 5
    rptmode = 23
ElseIf Option1(1).Value = True Then
    ReportStatusofEmployee = 1
    DEPARMENTNAMEFORTPT = ""
    SEXFORREPORT = 5
    DESIGNATIONFORRPT = ""
    CheckStatusofEmployee = 5
    rptmode = 23
ElseIf Option1(2).Value = True Then
    ReportStatusofEmployee = 2
    SEXFORREPORT = 5
    DEPARMENTNAMEFORTPT = ""
    DESIGNATIONFORRPT = ""
    rptmode = 23
    If Combo1(2).Text = "First Class" Then
        CheckStatusofEmployee = 1
    ElseIf Combo1(2).Text = "2nd Class" Then
        CheckStatusofEmployee = 2
    ElseIf Combo1(2).Text = "3rd Class" Then
        CheckStatusofEmployee = 3
    Else
        CheckStatusofEmployee = 4
    End If

ElseIf Option1(3).Value = True Then
    ReportStatusofEmployee = 3
    DEPARMENTNAMEFORTPT = ""
    DESIGNATIONFORRPT = ""
    CheckStatusofEmployee = 5
    If Combo1(3).Text = "Male" Then
        SEXFORREPORT = 1
    Else
        SEXFORREPORT = 0
    End If
    rptmode = 23
ElseIf Option1(4).Value = True Then
    ReportStatusofEmployee = 4
    DEPARMENTNAMEFORTPT = ""
    DESIGNATIONFORRPT = Trim(Combo1(1))
    CheckStatusofEmployee = 5
    SEXFORREPORT = 5
ElseIf Option1(9).Value = True Then
      paramMonth = cboMonth.Text
      paramYear = cboYear.Text
      paramDepartment = Combo1(7).Text
      rptmode = 49
End If




Form20.Show vbModal
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
get_Department
get_Designation
Load_Yr Me
Load_MonthNm Me
cboMonth.Text = MonthName(Month(Now))
cboYear.Text = YEAR(Now)

End Sub
Private Sub get_Department()
On Error GoTo Errdes
Dim cmd As New Command
Dim conn4 As New Connection
Dim rs4 As New Recordset

conn4.ConnectionString = strCN.Connection_String
conn4.Open
cmd.ActiveConnection = conn4
cmd.CommandType = adCmdText

cmd.CommandText = "select DEPT_NM from ST_DEPT order by DEPT_NM "
rs4.CursorLocation = adUseClient
rs4.Open cmd.CommandText, conn4, adOpenDynamic, adLockOptimistic

    If rs4.RecordCount > 0 Then

        Do Until rs4.EOF
            Combo1(0).AddItem rs4.Fields(0)
            Combo1(7).AddItem rs4.Fields(0)
            rs4.MoveNext
        Loop
  End If

    rs4.Close
    conn4.Close
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub get_Designation()
On Error GoTo Errdes
Dim cmd5 As New Command
Dim conn5 As New Connection
Dim rs5 As New Recordset

conn5.ConnectionString = strCN.Connection_String
conn5.Open
cmd5.ActiveConnection = conn5
cmd5.CommandType = adCmdText

cmd5.CommandText = "select DESIGNATION from ST_DESIG  order by DESIGNATION "
rs5.CursorLocation = adUseClient
rs5.Open cmd5.CommandText, conn5, adOpenDynamic, adLockOptimistic

    If rs5.RecordCount > 0 Then

        Do Until rs5.EOF
            Combo1(1).AddItem rs5.Fields(0)
            rs5.MoveNext
        Loop
  End If

    rs5.Close
    conn5.Close
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Option1_Click(Index As Integer)
On Error GoTo Errdes
Select Case Index
Case 0
    For i = 1 To 3
        Combo1(i) = ""
    Next
    Combo1(0).SetFocus
    
Case 1
     For i = 0 To 3
        Combo1(i) = ""
    Next

Case 2
    
    Combo1(1) = ""
    Combo1(3) = ""
    Combo1(0) = ""
    Combo1(2).SetFocus

Case 3
    Combo1(1) = ""
    Combo1(2) = ""
    Combo1(0) = ""
    Combo1(3).SetFocus
Case 4
    
    Combo1(2) = ""
    Combo1(0) = ""
    Combo1(3) = ""
    Combo1(1).SetFocus
End Select
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
