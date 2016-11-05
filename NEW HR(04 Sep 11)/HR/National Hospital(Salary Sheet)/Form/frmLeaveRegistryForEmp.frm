VERSION 5.00
Begin VB.Form frmLeaveRegistryForEmp 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Earn Leave Entry Screen"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8415
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      Height          =   480
      Left            =   5805
      Picture         =   "frmLeaveRegistryForEmp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3330
      Width           =   1230
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   7125
      Picture         =   "frmLeaveRegistryForEmp.frx":1DCA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3330
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   3165
      Picture         =   "frmLeaveRegistryForEmp.frx":384C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3330
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   1845
      Picture         =   "frmLeaveRegistryForEmp.frx":51DE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   4485
      Picture         =   "frmLeaveRegistryForEmp.frx":6B70
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3330
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3120
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   8235
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
         Height          =   1275
         Left            =   180
         TabIndex        =   16
         Top             =   1620
         Width           =   7845
         Begin VB.TextBox Text3 
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   250
            Left            =   5730
            TabIndex        =   19
            Top             =   270
            Width           =   902
         End
         Begin VB.TextBox Text2 
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   250
            Left            =   1770
            TabIndex        =   18
            Text            =   "120"
            Top             =   315
            Width           =   902
         End
         Begin VB.ComboBox cboYear 
            Height          =   315
            Left            =   1770
            TabIndex        =   17
            Top             =   720
            Width           =   1095
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Left            =   5685
            Top             =   255
            Width           =   1095
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Left            =   1725
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Extra Leave Remain"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   3660
            TabIndex        =   7
            Top             =   255
            Width           =   1440
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total  Earn Leave"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   8
            Top             =   345
            Width           =   1275
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Year"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   105
            TabIndex        =   9
            Top             =   705
            Width           =   330
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00800000&
            Height          =   360
            Left            =   1725
            Top             =   690
            Width           =   1155
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1275
         TabIndex        =   2
         Top             =   360
         Width           =   1410
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1275
         TabIndex        =   1
         Top             =   1270
         Width           =   1935
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   3840
         TabIndex        =   10
         Top             =   420
         Width           =   420
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   6
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   3840
         TabIndex        =   5
         Top             =   855
         Width           =   825
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   4
         Top             =   870
         Width           =   840
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   375
         Left            =   1245
         Top             =   330
         Width           =   1455
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Type"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   3
         Top             =   1320
         Width           =   660
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFC0C0&
         Height          =   285
         Left            =   1245
         Top             =   1245
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmLeaveRegistryForEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Leave_Application As New Cls_Leave_Info
'Private Sub cmdClear_Click()
'Text2.Text = ""
'Text3.Text = ""
'lblName(1) = ""
'lblDesig(1) = ""
'lblDept(1) = ""
'Combo1(1).SetFocus
'End Sub
'
'Private Sub cmdClose_Click()
'    Close_Msg Me
'End Sub
'
'Private Sub cmdSave_Click()
'With Leave_Application
'    .Connstring = strCN.Connection_String
'    .Emp_ID = Combo1(1)
'    .TOTAL_EARN_LEAVE = Text2
'    .EXTRA_LEAVE_REMAIN = Text3
'    .YEAR = cboYear
'    .UPDATE_DATE = Date$
'    .Earn_Leave_Save
'    .Show_Message
'End With
'End Sub
'
'Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'Select Case Index
'Case 1
'    Get_Global_Info_All_Employee
'    Get_EXTRA_LEAVE_REMAIN_Of_EMPLOYEE
'End Select
'End If
'End Sub
'
'Private Sub Form_Load()
'Load_Yr Me
'Dim cmd As New Command
'Dim conn10 As New Connection
'Dim rs10 As New Recordset
'Dim conn11 As New Connection
'Dim rs11 As New Recordset
'
'conn10.ConnectionString = strCN.Connection_String
'conn10.Open
'cmd.ActiveConnection = conn10
'cmd.CommandType = adCmdText
'
'cmd.CommandText = "select emp_id from emp_info order by emp_id"
'
'
'rs10.CursorLocation = adUseClient
'rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic
'
'If rs10.RecordCount > 0 Then
'    Do Until rs10.EOF
'        Combo1(1).AddItem rs10.Fields(0)
'        rs10.MoveNext
'    Loop
'    Combo1(1) = Combo1(1).List(0)
'End If
'rs10.Close
'conn10.Close
'
'
'End Sub
'Private Sub Get_Global_Info_All_Employee()
'On Error GoTo Errdes
'Dim conn5 As New Connection
'Dim cmd As New Command
'Dim RS2 As New ADODB.Recordset
'conn5.ConnectionString = strCN.Connection_String
'conn5.Open
'cmd.ActiveConnection = conn5
'cmd.CommandType = adCmdText
'cmd.CommandText = "SELECT EMP_INFO.EMP_NM,ST_DESIG.DESIGNATION, " + _
'                    " ST_DEPT.DEPT_NM , St_JbType.JType_Nm " + _
'                    " FROM EMP_INFO, EMP_JOB_INFO," + _
'                    " ST_DEPT , St_Desig, St_JbType " + _
'                    " WHERE ((EMP_JOB_INFO.EMP_ID=EMP_INFO.EMP_ID)" + _
'                    " AND (EMP_JOB_INFO.DESIG=ST_DESIG.DESIG_CODE) " + _
'                    " AND (EMP_JOB_INFO.DEPT=ST_DEPT.DEPT_CODE) " + _
'                    " AND (EMP_JOB_INFO.JTYPE=ST_JBTYPE.JTYPE_CODE) AND EMP_INFO.Emp_Id='" & Combo1(1) & "')"
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'RS2.CursorLocation = adUseClient
'
'RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic
'
'If Not RS2.EOF Then
'    lblName(1) = RS2.Fields(0)
'    lblDesig(1) = RS2.Fields(1)
'    lblDept(1) = RS2.Fields(2)
'    Text1 = RS2.Fields(3)
'End If
'    Text2.SetFocus
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Get_EXTRA_LEAVE_REMAIN_Of_EMPLOYEE()
'On Error GoTo Errdes
'Dim conn5 As New Connection
'Dim cmd As New Command
'Dim RS2 As New ADODB.Recordset
'conn5.ConnectionString = strCN.Connection_String
'conn5.Open
'cmd.ActiveConnection = conn5
'cmd.CommandType = adCmdText
'cmd.CommandText = "select EXTRA_LEAVE_REMAIN from Earn_Leave where TRACK_ID=(select max(TRACK_ID) from Earn_Leave where EMP_ID='" & Trim(Combo1(1)) & "')"
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'RS2.CursorLocation = adUseClient
'
'RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic
'
'If Not RS2.EOF Then
'        Text3 = RS2.Fields(0)
'End If
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'
'
