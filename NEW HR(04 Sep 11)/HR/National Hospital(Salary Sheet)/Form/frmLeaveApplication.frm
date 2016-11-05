VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLeaveApplication 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Leave Application"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8340
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   4350
      Picture         =   "frmLeaveApplication.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5580
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   1710
      Picture         =   "frmLeaveApplication.frx":1A0A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5610
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   3030
      Picture         =   "frmLeaveApplication.frx":339C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5580
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   6990
      Picture         =   "frmLeaveApplication.frx":4D2E
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5580
      Width           =   1185
   End
   Begin VB.CommandButton cmdPreview 
      Height          =   480
      Left            =   5670
      Picture         =   "frmLeaveApplication.frx":67B0
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5580
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Leave Application"
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
      Height          =   5325
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
         Height          =   3105
         Left            =   150
         TabIndex        =   13
         Top             =   2040
         Width           =   7725
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   315
            Index           =   0
            Left            =   1830
            TabIndex        =   37
            Top             =   2130
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Text6 
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   6450
            TabIndex        =   36
            Top             =   2640
            Width           =   975
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Height          =   30
            Left            =   180
            TabIndex        =   35
            Top             =   1980
            Width           =   7305
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enjoy the Recreation"
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   3990
            TabIndex        =   34
            Top             =   2130
            Width           =   3465
         End
         Begin VB.TextBox Text5 
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   250
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1215
            Width           =   902
         End
         Begin VB.TextBox Text4 
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   250
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   390
            Width           =   902
         End
         Begin VB.TextBox Text3 
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   250
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   800
            Width           =   902
         End
         Begin VB.TextBox Text2 
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   250
            Left            =   1845
            TabIndex        =   18
            Top             =   390
            Width           =   902
         End
         Begin MSComCtl2.DTPicker dtpEmp_join_date 
            Height          =   315
            Left            =   5745
            TabIndex        =   19
            Top             =   810
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   8388608
            CalendarTitleForeColor=   8388608
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   63963139
            CurrentDate     =   36998
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   5745
            TabIndex        =   30
            Top             =   1280
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   8388608
            CalendarTitleForeColor=   8388608
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   63963139
            CurrentDate     =   36998
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   315
            Index           =   1
            Left            =   1830
            TabIndex        =   3
            Top             =   2550
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label7 
            Caption         =   "<Total>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1800
            TabIndex        =   4
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Total Earned Leave"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   1560
            Width           =   1410
         End
         Begin VB.Shape Shape12 
            BorderColor     =   &H00FFC0C0&
            Height          =   360
            Left            =   1800
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Shape Shape11 
            BorderColor     =   &H00FFC0C0&
            Height          =   360
            Left            =   1800
            Top             =   2100
            Width           =   1695
         End
         Begin VB.Shape Shape10 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Left            =   6360
            Top             =   2610
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "No.of Days Enjoyed(Recreation)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3990
            TabIndex        =   33
            Top             =   2670
            Width           =   2280
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Next Recreation Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   180
            TabIndex        =   32
            Top             =   2580
            Width           =   1545
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Last Recreation Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   180
            TabIndex        =   31
            Top             =   2160
            Width           =   1515
         End
         Begin VB.Shape Shape9 
            BorderColor     =   &H00FFC0C0&
            Height          =   360
            Left            =   5715
            Top             =   1245
            Width           =   1785
         End
         Begin VB.Shape Shape8 
            BorderColor     =   &H00FFC0C0&
            Height          =   375
            Left            =   5715
            Top             =   780
            Width           =   1785
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Join"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   4590
            TabIndex        =   29
            Top             =   1245
            Width           =   855
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Of Applied"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   4320
            TabIndex        =   28
            Top             =   825
            Width           =   1125
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Left            =   1800
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Left            =   5745
            Top             =   360
            Width           =   1095
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Left            =   1800
            Top             =   780
            Width           =   1095
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            Height          =   300
            Left            =   1800
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " Extra Leave Remain"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   17
            Top             =   1230
            Width           =   1485
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " Earn Leave Remain"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   16
            Top             =   828
            Width           =   1455
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Casul Leave Remain(Current Year)"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   3060
            TabIndex        =   15
            Top             =   405
            Width           =   2445
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Days Leave Applied"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   14
            Top             =   405
            Width           =   1425
         End
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1275
         TabIndex        =   12
         Top             =   1270
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1680
         Width           =   6525
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1275
         TabIndex        =   1
         Top             =   360
         Width           =   1410
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFC0C0&
         Height          =   285
         Left            =   1245
         Top             =   1245
         Width           =   2085
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
         TabIndex        =   11
         Top             =   1320
         Width           =   660
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         Height          =   375
         Left            =   1230
         Top             =   1635
         Width           =   6600
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Type"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   9
         Top             =   1680
         Width           =   855
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
         Caption         =   "Designation"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   8
         Top             =   870
         Width           =   840
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
         TabIndex        =   7
         Top             =   855
         Width           =   825
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
         TabIndex        =   2
         Top             =   420
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmLeaveApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Job_Info As New clsEmp_Job_Detail
'Private Leave_Application As New Cls_Leave_Info
'Dim GET_REMAIN_CASUALLEAVE
'Dim Emp_Status As String
'Dim Get_No_of_Leave, No_Of_Days
'Dim CurentYearTotalLeave As Integer
'Private Sub cmdClear_Click()
'On Error GoTo Errdes
'Text2.Text = ""
'Text4.Text = ""
'Text3.Text = ""
'Text5.Text = ""
'DTPicker1 = Format(Date, "dd-mmm-yyyy")
'dtpEmp_join_date = Format(Date, "dd-mmm-yyyy")
'Combo1(1) = Combo1(1).List(0)
'lblName(1) = ""
'lblDesig(1) = ""
'lblDept(1) = ""
'Text1 = ""
'For i = 0 To 1
'    Me.MaskEdBox1(i).Text = "__/__/__"
'Next i
'Text6.Text = ""
'Label7.Caption = "N/A"
'Combo1(1).SetFocus
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub cmdClose_Click()
'    Close_Msg Me
'End Sub
'Private Sub cmdDelete_Click()
'On Error GoTo Errdes
'With Leave_Application
'    .Connstring = strCN.Connection_String
'    .Emp_ID = Combo1(1)
'    .DATE_OF_APPLIED = dtpEmp_join_date
'    .Delete
'    MsgBox "Data Deleted Successfully", vbInformation, "IT Division, DNMIH"
'    cmdClear_Click
'End With
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub cmdPreview_Click()
'On Error GoTo Errdes
'Dim f As New frmLeaveApplicationReport
'f.Show 1
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub cmdSave_Click()
'On Error GoTo Errdes
'If Mid(Combo2, 1, 3) <> "006" Then
'    With Leave_Application
'        .Connstring = strCN.Connection_String
'        .Emp_ID = Combo1(1)
'        .LEAVE_CODE = Mid(Combo2, 1, Get_For_Concatenate(Combo2) - 1)
'        .NO_OF_DAYS_LEAVE = Text2
'        .DATE_OF_APPLIED = dtpEmp_join_date
'        .DATE_OF_JOINFROMLF = DTPicker1
'        .REMAIN_CASUALLEAVE = Text4
'        .Save
'        .Show_Message
'    End With
'Else
'    Recreation_Info_Save
'End If
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo Errdes
'If KeyCode = 13 Then
'Select Case Index
'Case 1
'    Get_Casual_LEAVE_Of_EMPLOYEE
'    Get_EXTRA_LEAVE_REMAIN_Of_EMPLOYEE
'    Get_Global_Info_All_Employee
'    Label7.Caption = "<Total>"
'End Select
'End If
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'
'End Sub
'Private Sub Combo2_Click()
'Get_Casual_LEAVE_Of_EMPLOYEE
'If Mid(Combo2, 1, 3) = "006" Then
'    Get_Last_and_Next_RecreationDate
'    dtpEmp_join_date.SetFocus
'    Label7.Caption = "N/A"
'End If
'Except_recreation_Leave
'
'If Mid(Combo2, 1, 3) = "002" Then
'    Get_Current_Year_Leave_Calculation
'End If
'Casual_Leave_Validation
'End Sub
'Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    dtpEmp_join_date.SetFocus
'End If
'End Sub
'Private Sub dtpEmp_join_date_Change()
'Dateof_Join_after_Leave
'End Sub
'Private Sub dtpEmp_join_date_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    Text6.SetFocus
'End If
'End Sub
'Private Sub Form_Load()
'On Error GoTo Errdes
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
'conn11.ConnectionString = strCN.Connection_String
'conn11.Open
'cmd.ActiveConnection = conn11
'cmd.CommandType = adCmdText
'cmd.CommandText = "Select LEAVE_CODE,LEAVE_NAME from ST_LEAVE"
'rs11.CursorLocation = adUseClient
'rs11.Open cmd.CommandText, conn11, adOpenDynamic, adLockOptimistic
'
'If rs11.RecordCount > 0 Then
'    Do Until rs11.EOF
'        Combo2.AddItem rs11.Fields(0) & " - " & rs11.Fields(1)
'        rs11.MoveNext
'    Loop
'    Combo2 = Combo2.List(0)
'End If
'
'dtpEmp_join_date.Value = Format(Date, "dd-mmm-yyyy")
'DTPicker1.Value = Format(Date, "dd-mmm-yyyy")
''dtpEmp_join_date.Enabled = False
'DTPicker1.Enabled = False
'
'
'
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
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
'    Combo2.SetFocus
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Function Get_For_Concatenate(str As String)
'Dim txt$, FTxt$
'For i = 1 To Len(str)
'    txt = Mid(str, i, 1)
'    If txt = "-" Then
'       Get_For_Concatenate = Len(FTxt)
'    Else
'        FTxt = FTxt + txt
'    End If
'Next i
'
'End Function
'Private Sub Text2_Change()
'On Error GoTo Errdes
'If Mid(Combo2, 1, 3) = "001" Then
'
'    If Val(Text2.Text) >= 7 Then
'        Get_Casual_LEAVE_Of_EMPLOYEE
'        Text4 = GET_REMAIN_CASUALLEAVE
'        MsgBox "More than 6-days is not Allowed at a time for this type of Leave ! ", vbCritical, "IT Division, DNMIH"
'        Text2.SetFocus
'        Exit Sub
'    ElseIf Val(Trim(Text4.Text)) < 0 Then
'        MsgBox "Cusual Leave is not Avaiable any more", vbCritical, "Daffodil Sotware Ltd"
'        Get_Casual_LEAVE_Of_EMPLOYEE
'        Text4 = GET_REMAIN_CASUALLEAVE
'        Text2.SetFocus
'        Exit Sub
'    Else
'        Get_Casual_LEAVE_Of_EMPLOYEE
'        Text4 = GET_REMAIN_CASUALLEAVE
'        Text4 = Val(Trim(Text4.Text)) - Text2.Text
'        'Label7.Caption = Label7.Caption - Val(Text2.Text)
'        Dateof_Join_after_Leave
'    End If
'ElseIf Mid(Combo2, 1, 3) = "002" Then
'        Text4 = Val(Trim(Text4.Text)) - Text2.Text
'        Dateof_Join_after_Leave
'        Casual_Leave_Validation
'ElseIf Mid(Combo2, 1, 3) = "006" Then
'        Dateof_Join_after_Leave
'        Label7.Caption = "N/A"
'Else
'        Get_Other_Leave_Calculation
'        'Text4 = Val(Trim(Text4.Text)) - Text2.Text
'        Dateof_Join_after_Leave
'        Label7.Caption = "N/A"
'End If
'Exit Sub
'Errdes:
'    If Err.Number = 13 Then
'        Text2 = 0
'    Else
'        MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'    End If
'End Sub
'Private Sub Dateof_Join_after_Leave()
'On Error GoTo Errdes
'Dim conn5 As New Connection
'Dim cmd As New Command
'Dim RS2 As New ADODB.Recordset
'Dim Rec As String
'conn5.ConnectionString = strCN.Connection_String
'conn5.Open
'cmd.ActiveConnection = conn5
'cmd.CommandType = adCmdText
'Rec = " + " & Text2
'
'
'If Rec = " + " Then Rec = " + " & 0
'
'
''------------------- HAS TO INSERT VALUE OF TIME
'If Mid(Combo2, 1, 3) <> "006" Then
'    cmd.CommandText = " Select to_date('" & dtpEmp_join_date & "','DD-mm-yy" & "')" & Rec & " from dual"
'Else
'   If Trim(Text6) = "" Then
'        Rec = " + " & Trim(Text4)
'   Else
'       Rec = " + " & Trim(Text6)
'   End If
'
'   cmd.CommandText = " Select to_date('" & dtpEmp_join_date & "','DD-mm-yy" & "')" & Rec & " from dual"
'
'End If
'
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'RS2.CursorLocation = adUseClient
'
'RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic
'
'If Not RS2.EOF Then
'    DTPicker1 = Format(RS2.Fields(0), "dd-mmm-yyyy")
'End If
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
'cmd.CommandText = "select EXTRA_LEAVE_REMAIN,TOTAL_EARN_LEAVE from Earn_Leave where TRACK_ID=(select max(TRACK_ID) from Earn_Leave where EMP_ID='" & Trim(Combo1(1)) & "')"
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'RS2.CursorLocation = adUseClient
'
'RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic
'
'If Not RS2.EOF Then
'        Text5 = RS2.Fields(0)
'        Text3 = RS2.Fields(1)
'End If
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Get_Casual_LEAVE_Of_EMPLOYEE()
'On Error GoTo Errdes
'Dim conn5 As New Connection
'Dim cmd As New Command
'Dim RS2 As New ADODB.Recordset
'conn5.ConnectionString = strCN.Connection_String
'conn5.Open
'cmd.ActiveConnection = conn5
'cmd.CommandType = adCmdText
'
'cmd.CommandText = "select REMAIN_CASUALLEAVE from leave_application" + _
'                        " where   Track_ID=(Select max(Track_id)from leave_application where emp_iD='" & Combo1(1) & "' and LEAVE_CODE='" & Mid(Combo2, 1, Get_For_Concatenate(Combo2) - 1) & "')"
'
'
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'RS2.CursorLocation = adUseClient
'
'RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic
'
'If Not RS2.EOF Then
'        GET_REMAIN_CASUALLEAVE = RS2.Fields(0)
'        Text4 = GET_REMAIN_CASUALLEAVE
'Else
'    Dim conn007 As New Connection
'    Dim cmd1 As New Command
'    Dim RS007 As New ADODB.Recordset
'    conn007.ConnectionString = strCN.Connection_String
'    conn007.Open
'    cmd1.ActiveConnection = conn007
'    cmd1.CommandType = adCmdText
'    cmd1.CommandText = "SELECT ST_JBTYPE.JTYPE_NM From EMP_JOB_INFO,St_JbType" + _
'        " Where (EMP_JOB_INFO.JType = St_JbType.JType_Code) and EMP_JOB_INFO.EMP_ID='" & Combo1(1) & "'"
'
'    cmd1.Properties("iRowsetChange") = True
'    cmd1.Properties("updatability") = 7
'    RS007.CursorLocation = adUseClient
'
'    RS007.Open cmd1.CommandText, conn007, adOpenDynamic, adLockOptimistic
'    If Not RS007.EOF Then
''''        If RS007.Fields(0) = "Permanent" Then
''''            GET_REMAIN_CASUALLEAVE = 33
''''            Text4.Text = GET_REMAIN_CASUALLEAVE
''''        Else
''''            GET_REMAIN_CASUALLEAVE = 16
''''            Text4.Text = GET_REMAIN_CASUALLEAVE
''''        End If
'
'    '===============CHANGES FROM HERE
'
'        If RS007.Fields(0) = "Permanent" Then
'            Emp_Status = "Permanent"
'              If Mid(Combo2, 1, 3) = "001" Then
'                   Label6(3).Caption = "Casual Leave Remain"
'                   GET_REMAIN_CASUALLEAVE = 20
'                   Text4.Text = GET_REMAIN_CASUALLEAVE
'               End If
'
'
'              If Mid(Combo2, 1, 3) = "002" Then
'
'                   Label6(3).Caption = "Earn Leave Remain(Current Year)"
'                   'GET_REMAIN_CASUALLEAVE = 33
'                   Text4.Text = GET_REMAIN_CASUALLEAVE
'               End If
'
'
'
'               If Mid(Combo2, 1, 3) = "005" Then
'
'                   Label6(3).Caption = "Maternal Leave Remain"
'                   'GET_REMAIN_CASUALLEAVE = 120
'                   Text4.Text = GET_REMAIN_CASUALLEAVE
'               End If
'
'
'               '''------------------------------FOR THE RECREATION LEAVE
'
'               If Mid(Combo2, 1, 3) = "006" Then
'                   Label6(3).Caption = "Recreation Leave Remain"
'                   Except_recreation_Leave
'                   Text4.Text = 15
'                   Text6 = 15
'               End If
'
'
'
'
'        Else
'
'               If Mid(Combo2, 1, 3) = "001" Then
'                    Label6(3).Caption = "Casual Leave Remain"
'                    GET_REMAIN_CASUALLEAVE = 20
'                    Text4.Text = GET_REMAIN_CASUALLEAVE
'               End If
'
'
'
'                If Mid(Combo2, 1, 3) = "002" Then
'                    Label6(3).Caption = "Earn Leave Remain"
'                    'GET_REMAIN_CASUALLEAVE = 16
'                    Text4.Text = GET_REMAIN_CASUALLEAVE
'               End If
'
'
'
'               If Mid(Combo2, 1, 3) = "005" Then
'                   Label6(3).Caption = "Maternal Leave Remain"
'                   'GET_REMAIN_CASUALLEAVE = 120
'                   Text4.Text = GET_REMAIN_CASUALLEAVE
'               End If
'
'        End If
'
'        If Mid(Combo2, 1, 3) = "003" Then
'            Label6(3).Caption = "Leave without Pay"
'        End If
'
'
'        If Mid(Combo2, 1, 3) = "004" Then
'            Label6(3).Caption = "Half Avearge Pay"
'        End If
'
'    '==================================END OF CHANGES
'
'
'    End If
'
'End If
'
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Text2_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Len(Trim(Text2.Text)) = 0 Then Text2.Text = 0
'End If
'End Sub
'Private Sub Except_recreation_Leave()
'If Mid(Combo2, 1, 3) = "006" Then
'    Text2.Locked = True
'    Text3.Locked = True
'    Text4.Locked = True
'    Text5.Locked = True
'    Text6.Text = Val(Text4.Text)
'    MaskEdBox1(0).Enabled = True
'    MaskEdBox1(1).Enabled = True
'Else
'    Text2.Locked = False
'    Text3.Locked = False
'    Text4.Locked = False
'    Text5.Locked = False
'    MaskEdBox1(0).Text = "__/__/__"
'    MaskEdBox1(1).Text = "__/__/__"
'
'    MaskEdBox1(0).Enabled = False
'    MaskEdBox1(1).Enabled = False
'    Text6.Text = ""
'    Text6.Locked = True
'
'End If
'End Sub
'Private Sub Text6_Change()
'Dateof_Join_after_Leave
'End Sub
'Private Sub Recreation_Info_Save()
'Dim CheckStatusValue
'With Leave_Application
'        .Connstring = strCN.Connection_String
'        .Emp_ID = Combo1(1)
'        .LAST_RCREAT_DATE = MaskEdBox1(0)
'        .NEXT_RCREAT_DATE = MaskEdBox1(1)
'        .ENJOED_RECRT = Text6
'
'    If Me.Check1.Value Then
'        CheckStatusValue = 1
'    Else
'        CheckStatusValue = 0
'    End If
'
'        .ENJOED_STATUS = CheckStatusValue
'        .RECRE_ST_DT = dtpEmp_join_date
'        .RECRE_END_DT = DTPicker1
'        .CreatedDate = Date
'        .CreatedBy = "DSL"
'        .UpdateBy = "DSL"
'        .UPDATE_DATE = Date
'        .Creation_Leave_Save
'        .Show_Message
'    End With
'End Sub
'Private Sub Get_Last_and_Next_RecreationDate()
'On Error GoTo Errdes
'Dim conn5 As New Connection
'Dim cmd As New Command
'Dim RS2 As New ADODB.Recordset
'conn5.ConnectionString = strCN.Connection_String
'conn5.Open
'cmd.ActiveConnection = conn5
'cmd.CommandType = adCmdText
'cmd.CommandText = "select LAST_RCREAT_DATE,NEXT_RCREAT_DATE from Emp_Recreation" + _
'        " where sequence_id=(select max(sequence_id) from Emp_Recreation where emp_id='" & Trim(Combo1(1)) & "')"
'
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'RS2.CursorLocation = adUseClient
'
'RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic
'
'If Not RS2.EOF Then
'    MaskEdBox1(0) = Format(RS2.Fields(1), "dd/mm/yy")
'    MaskEdBox1(1) = Format(DateAdd("yyyy", 3, Format(Me.MaskEdBox1(0), "dd/mm/yy")), "dd/mm/yy") '-----------------next increment after 3-years
'Else
'    MaskEdBox1(0) = Format(Date, "dd/mm/yy")
'    MaskEdBox1(1) = Format(DateAdd("yyyy", 3, Format(Me.MaskEdBox1(0), "dd/mm/yy")), "dd/mm/yy") '-----------------next increment after 3-years
'End If
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Get_Earned_Leave_Curernt_year()
'On Error GoTo Errdes
'Dim conn7 As New Connection
'Dim cmd As New Command
'Dim RS7 As New ADODB.Recordset
'conn7.ConnectionString = strCN.Connection_String
'conn7.Open
'cmd.ActiveConnection = conn7
'cmd.CommandType = adCmdText
'cmd.CommandText = "select LAST_RCREAT_DATE,NEXT_RCREAT_DATE from Emp_Recreation" + _
'        " where sequence_id=(select max(sequence_id) from Emp_Recreation where emp_id='" & Trim(Combo1(1)) & "')"
'
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'RS7.CursorLocation = adUseClient
'
'RS7.Open cmd.CommandText, conn7, adOpenDynamic, adLockOptimistic
'
'If Not RS7.EOF Then
'    MaskEdBox1(0) = Format(RS2.Fields(1), "dd/mm/yy")
'    MaskEdBox1(1) = Format(DateAdd("yyyy", 3, Format(Me.MaskEdBox1(0), "dd/mm/yy")), "dd/mm/yy") '-----------------next increment after 3-years
'Else
'    MaskEdBox1(0) = Format(Date, "dd/mm/yy")
'    MaskEdBox1(1) = Format(DateAdd("yyyy", 3, Format(Me.MaskEdBox1(0), "dd/mm/yy")), "dd/mm/yy") '-----------------next increment after 3-years
'End If
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Get_Current_Year_Leave_Calculation()
'On Error GoTo Errdes
'Dim YeartoGet, DatePart
'DatePart = "01/01/"
'YeartoGet = YEAR(Now)
'Dim StartDate
'StartDate = DatePart & YEAR(Now)
'Dim conn8 As New Connection
'Dim cmd As New Command
'Dim rs8 As New ADODB.Recordset
'conn8.ConnectionString = strCN.Connection_String
'conn8.Open
'cmd.ActiveConnection = conn8
'cmd.CommandType = adCmdText
'
'If Mid(Combo2, 1, 3) = "002" Then
'     cmd.CommandText = "select to_number(to_date('" & Date & "','dd-mm-yyyy" & "')" & "- " & "to_date('" & StartDate & "','dd-mm-yyyy" & "'))  from dual"
'End If
'
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'rs8.CursorLocation = adUseClient
'
'rs8.Open cmd.CommandText, conn8, adOpenDynamic, adLockOptimistic
'
'Get_Curr_Yr_Prev_Calc
'
'If Not rs8.EOF Then
'     No_Of_Days = rs8.Fields(0)
'     If Emp_Status = "Permanent" Then
'        Text4 = Round(No_Of_Days / 11) - CurentYearTotalLeave  ''''''''''''''1day Leave/14days for Permanent Employee
'        GET_REMAIN_CASUALLEAVE = Round(No_Of_Days / 22) - CurentYearTotalLeave
'     Else
'        Text4 = Round(No_Of_Days / 22) - CurentYearTotalLeave  ''''''''''''''1day Leave/14days for Permanent Employee
'        GET_REMAIN_CASUALLEAVE = Round(No_Of_Days / 22) - CurentYearTotalLeave
'     End If
'Else
'   No_Of_Days = 0
'   If Emp_Status = "Permanent" Then
'      Text4 = Round(No_Of_Days / 11) - CurentYearTotalLeave  ''''''''''''''1day Leave/14days for Permanent Employee
'      GET_REMAIN_CASUALLEAVE = Round(No_Of_Days / 22) - CurentYearTotalLeave
'   Else
'       Text4 = Round(No_Of_Days / 22) - CurentYearTotalLeave  ''''''''''''''1day Leave/14days for Permanent Employee
'       GET_REMAIN_CASUALLEAVE = Round(No_Of_Days / 22) - CurentYearTotalLeave
'   End If
'End If
'
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Casual_Leave_Validation()
'If Val(Text3.Text) > 120 Then
'    Text5.Text = 120 - Val(Text3.Text)
'End If
'Label7.Caption = Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)
'End Sub
'Private Sub Get_Curr_Yr_Prev_Calc()
'On Error GoTo Errdesc
'DatePart1 = "01/01/"
'Dim TodaysDay$
'YeartoGet1 = YEAR(Now)
'Dim StartDate1
'StartDate1 = Format(DatePart1 & YEAR(Now), "dd-mmm-yyyy")
'TodaysDay = Format(Date, "dd-mmm-yyyy")
'Dim cmd As New Command
'Dim Myrecord As New ADODB.Recordset
'Dim dayNo As Double
'Dim MyConn As New Connection
'MyConn.ConnectionString = strCN.Connection_String
'MyConn.Open
'cmd.ActiveConnection = MyConn
'cmd.CommandType = adCmdText
'
'cmd.CommandType = adCmdText
'
'cmd.CommandText = "select sum(NO_OF_DAYS_LEAVE) as days from LEAVE_APPLICATION where emp_id='" & Combo1(1) & "' and " + _
'               " (to_date(to_char(DATE_OF_APPLIED,'dd-mon-yyyy'),'dd-mon-yyyy') >= to_date('" & StartDate1 & "','dd-mon-yyyy') and to_date(to_char(DATE_OF_APPLIED,'dd-mon-yyyy'),'dd-mon-yyyy') <= to_date('" & TodaysDay & "','dd-mon-yyyy')) and  LEAVE_CODE='002'"
'
'    cmd.Properties("iRowsetChange") = True
'    cmd.Properties("updatability") = 7
'    Myrecord.CursorLocation = adUseClient
'
' Myrecord.Open cmd.CommandText, MyConn, adOpenDynamic, adLockOptimistic
'
' If Myrecord.BOF = False Then
'    CurentYearTotalLeave = "" & Myrecord(0)
'
' Else
'   CurentYearTotalLeave = 0
' End If
'
'    Myrecord.Close
'    MyConn.Close
'
'Exit Sub
'Errdesc:
'If Err.Number = 13 Then
'    CurentYearTotalLeave = 0
'Else
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End If
'End Sub
'Private Sub Get_Other_Leave_Calculation()
'On Error GoTo Errdes
'If Mid(Combo2, 1, 3) = "004" Then
'      Text3.Text = Val(Text3.Text) - Round(Val(Text2.Text)) / 2
'ElseIf Mid(Combo2, 1, 3) = "005" Or Mid(Combo2, 1, 3) = "006" Then
'    If Val(Text3.Text) > Val(Text2.Text) Then
'        Text3 = Val(Text3.Text) - Val(Text2.Text)
'    Else
'        MsgBox "Invalid Input", vbCritical, "IT Division, DNMIH."
'        Text2.SetFocus
'        Exit Sub
'    End If
'End If
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'
