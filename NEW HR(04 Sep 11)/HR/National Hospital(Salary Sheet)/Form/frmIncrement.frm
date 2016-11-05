VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIncrement 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employee Increment Record"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8175
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   975
      Left            =   8160
      TabIndex        =   28
      Top             =   2400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1720
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   4320
      Picture         =   "frmIncrement.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4170
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   5565
      Picture         =   "frmIncrement.frx":1A0A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4170
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   1800
      Picture         =   "frmIncrement.frx":35F4
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4170
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   3030
      Picture         =   "frmIncrement.frx":4F86
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4170
      Width           =   1140
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   6930
      Picture         =   "frmIncrement.frx":6918
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4170
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Increment Information"
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
      Height          =   3900
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7965
      Begin VB.TextBox txtfields 
         BorderStyle     =   0  'None
         Height          =   585
         Index           =   5
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   3090
         Width           =   6015
      End
      Begin VB.TextBox txtfields 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   6630
         TabIndex        =   9
         Top             =   2595
         Width           =   1125
      End
      Begin VB.TextBox txtfields 
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   3
         Left            =   6630
         TabIndex        =   27
         Text            =   "0"
         Top             =   2190
         Width           =   1035
      End
      Begin VB.TextBox txtfields 
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   6630
         TabIndex        =   26
         Text            =   "0"
         Top             =   1740
         Width           =   1125
      End
      Begin VB.TextBox txtfields 
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   1800
         TabIndex        =   25
         Top             =   1710
         Width           =   1935
      End
      Begin VB.TextBox txtfields 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   210
         Index           =   0
         Left            =   5130
         TabIndex        =   24
         Top             =   1350
         Width           =   2520
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1770
         TabIndex        =   2
         Top             =   360
         Width           =   1410
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1785
         TabIndex        =   1
         Top             =   1320
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker LDtofIncrem 
         Height          =   315
         Left            =   1770
         TabIndex        =   12
         Top             =   2160
         Width           =   2040
         _ExtentX        =   3598
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
      Begin MSComCtl2.DTPicker Next_Incre_Dt 
         Height          =   315
         Left            =   1770
         TabIndex        =   13
         Top             =   2625
         Width           =   2040
         _ExtentX        =   3598
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
      Begin VB.Shape Shape10 
         BorderColor     =   &H00FFC0C0&
         Height          =   330
         Left            =   6600
         Top             =   2580
         Width           =   1200
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FFC0C0&
         Height          =   660
         Left            =   1740
         Top             =   3060
         Width           =   6120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Comments"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   3150
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Register Page No"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4050
         TabIndex        =   29
         Top             =   2670
         Width           =   1260
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No.of Days Leave (Half- Pay)"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   4050
         TabIndex        =   18
         Top             =   2213
         Width           =   2070
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No.of Days Leave(Without Pay)"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   4050
         TabIndex        =   17
         Top             =   1770
         Width           =   2250
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Next Dt. of Increment"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   16
         Top             =   2685
         Width           =   1515
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FFC0C0&
         Height          =   390
         Left            =   1755
         Top             =   2115
         Width           =   2100
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Date of Increment"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   15
         Top             =   2220
         Width           =   1320
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00FFC0C0&
         Height          =   390
         Left            =   1740
         Top             =   2595
         Width           =   2100
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00FFC0C0&
         Height          =   285
         Left            =   6615
         Top             =   1710
         Width           =   1185
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FFC0C0&
         Height          =   285
         Left            =   6615
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFC0C0&
         Height          =   300
         Left            =   1755
         Top             =   1695
         Width           =   2085
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Increment Amount"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   1755
         Width           =   1290
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
         Left            =   4065
         TabIndex        =   11
         Top             =   375
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
         Left            =   180
         TabIndex        =   7
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
         Left            =   4065
         TabIndex        =   6
         Top             =   900
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
         Left            =   180
         TabIndex        =   5
         Top             =   900
         Width           =   840
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   375
         Left            =   1740
         Top             =   330
         Width           =   1455
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Basic"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   4050
         TabIndex        =   4
         Top             =   1358
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         Height          =   285
         Left            =   5085
         Top             =   1290
         Width           =   2715
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
         Left            =   180
         TabIndex        =   3
         Top             =   1335
         Width           =   660
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFC0C0&
         Height          =   285
         Left            =   1755
         Top             =   1290
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmIncrement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Increment_Rc As New Cls_IncrementPro
'Dim NextYrIncrementDate As Date
''Dim Next_Increment_Date As String
'Dim Next_Increment_Date As Date
'Dim GetNoOfDaysLvWithoutPay, GetNoOfDaysHalfLv, TotalDaysOfLeave
'Private Sub cmdClear_Click()
''On Error GoTo Errdes
'For i = 0 To 5
'    txtfields(i).Text = ""
'Next
'
'lblDept(1) = ""
'lblName(1) = ""
'lblDesig(1) = ""
'lblDept(1) = ""
'Text1 = ""
'LDtofIncrem = Format(Date$, "dd-mmm-yyyy")
'Next_Incre_Dt = Date$
'Combo1(1).SetFocus
'Exit Sub
'Errdes:
'MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub cmdClose_Click()
'Unload Me
'End Sub
'Private Sub cmdDelete_Click()
''On Error GoTo Errdes
'With Increment_Rc
'    .Connstring = strCN.Connection_String
'    .Emp_ID = Trim(Combo1(1).Text)
'    .LAST_DT_INCRE = LDtofIncrem
'    .NEXT_DT_INCRE = Next_Incre_Dt
'    .Increment_Delete
'End With
'MsgBox "Data Deleted Successfully", vbInformation, "IT Division, DNMIH"
'cmdClear_Click
'
'Exit Sub
'Errdes:
'MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'
'End Sub
'Private Sub cmdPrint_Click()
'Dim f As New frmIncrementReport
'f.Show 1
'End Sub
'Private Sub cmdSave_Click()
''On Error GoTo Errdes
'With Increment_Rc
'    .Connstring = strCN.Connection_String
'    .Emp_ID = Trim(Combo1(1).Text)
'    .Amount = txtfields(1).Text
'    .LAST_DT_INCRE = LDtofIncrem
'    .NEXT_DT_INCRE = Next_Incre_Dt
'    .LEAVE_WITHOUTPAY = txtfields(2).Text
'    .LEAVE_HALFPAY = txtfields(3).Text
'    .UPDATE_DATE = Date
'    .RegistrationPgNo = txtfields(4).Text
'    .CommentsonIncre = txtfields(5).Text
'    .Increment_Save
'End With
'MsgBox "Data Save Successfully", vbInformation, "IT Division, DNMIH"
'
'Exit Sub
'Errdes:
'MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'
'End Sub
'Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo Errdes
'If KeyCode = 13 Then
'    Get_Global_Info_All_Employee
'    Get_Last_IncrementDt
'    Dateof_Next_Increment
'    Get_WithoutPay_Leave
'    Get_With_HalfPay_Leave
'    Get_Increment_Amount
'End If
'Exit Sub
'Errdes:
'MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Form_Load()
'On Error GoTo Errdes
'
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
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
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
'                    " ST_DEPT.DEPT_NM , St_JbType.JType_Nm,EMP_JOB_INFO.BASIC_SAL " + _
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
'    txtfields(0) = RS2.Fields(4)
'End If
'    txtfields(1).SetFocus
'
'RS2.Close
'conn5.Close
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Get_Increment_Amount()
'On Error GoTo Errdes
'Dim conn6 As New Connection
'Dim cmd As New Command
'Dim rs3 As New ADODB.Recordset
'conn6.ConnectionString = strCN.Connection_String
'conn6.Open
'cmd.ActiveConnection = conn6
'cmd.CommandType = adCmdText
'cmd.CommandText = "SELECT ST_PAYSCALE.INCR FROM EMP_JOB_INFO,ST_PAYSCALE " + _
'                " Where (EMP_JOB_INFO.Scale_code = St_Payscale.Scale_code) " + _
'                " AND EMP_JOB_INFO.EMP_ID='" & Combo1(1) & "'"
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'rs3.CursorLocation = adUseClient
'
'rs3.Open cmd.CommandText, conn6, adOpenDynamic, adLockOptimistic
'
'If Not rs3.EOF Then
'    txtfields(1) = rs3.Fields(0)
'End If
'rs3.Close
'conn6.Close
'
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Dateof_Next_Increment()
'On Error GoTo Errdes
'Dim Myconn1 As New Connection
'Dim cmd1 As New Command
'Dim myrs1 As New ADODB.Recordset
'Myconn1.ConnectionString = strCN.Connection_String
'Myconn1.Open
'cmd1.ActiveConnection = Myconn1
'cmd1.CommandType = adCmdText
'
'cmd1.CommandText = "select add_months(to_date('" & LDtofIncrem & "','dd-mm-yy'), 12) from dual"
'
'
'cmd1.Properties("iRowsetChange") = True
'cmd1.Properties("updatability") = 7
'myrs1.CursorLocation = adUseClient
'myrs1.Open cmd1.CommandText, Myconn1, adOpenDynamic, adLockOptimistic
'
'If Not myrs1.EOF Then
'     Next_Increment_Date = Format(myrs1.Fields(0), "dd-mmm-yyyy")
'     Next_Incre_Dt = Next_Increment_Date
'End If
'myrs1.Close
'Myconn1.Close
'
'Exit Sub
'Errdes:
'    If Err.Number = 13 Then
'        MsgBox "Date is not in Proper Format"
'    Else
'        MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'    End If
'End Sub
'Private Sub LDtofIncrem_Change()
'On Error GoTo Errdes
'Dateof_Next_Increment
'Get_Leave_Affect_On_IncrementDate
'Real_IncrementDt_Considering_Leave
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Get_Leave_Affect_On_IncrementDate()
'If Len(Trim(txtfields(2))) = 0 Then txtfields(2) = 0
'GetNoOfDaysLvWithoutPay = Val(txtfields(2))
'If Val(txtfields(3)) <> 0 Then
'    GetNoOfDaysHalfLv = Round(Val(txtfields(3) \ 2))
'Else
'    GetNoOfDaysHalfLv = 0
'End If
'TotalDaysOfLeave = GetNoOfDaysLvWithoutPay + GetNoOfDaysHalfLv
'End Sub
'Private Sub Real_IncrementDt_Considering_Leave()
'On Error GoTo Errdes
'Dim conn5 As New Connection
'Dim cmd As New Command
'Dim RS2 As New ADODB.Recordset
'Dim Rec As String
'conn5.ConnectionString = strCN.Connection_String
'conn5.Open
'cmd.ActiveConnection = conn5
'cmd.CommandType = adCmdText
'
'If TotalDaysOfLeave = "" Then
'    TotalDaysOfLeave = 0
'End If
'
'Rec = "+" & TotalDaysOfLeave
'cmd.CommandText = " Select to_date('" & Next_Increment_Date & "','dd-mm-yy" & "')" & Rec & " from dual"
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'RS2.CursorLocation = adUseClient
'
'RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic
'
'If Not RS2.EOF Then
'      Next_Incre_Dt = Format(RS2.Fields(0), "dd-mmm-yyyy")
'End If
'RS2.Close
'conn5.Close
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'
'Private Sub LDtofIncrem_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'  txtfields(3).SetFocus
'End If
'End Sub
'
'Private Sub Next_Incre_Dt_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'  txtfields(4).SetFocus
'End If
'End Sub
'
'Private Sub txtfields_Change(Index As Integer)
'Select Case Index
'
'Case 2
'    Get_Leave_Affect_On_IncrementDate
'    Real_IncrementDt_Considering_Leave
'
'Case 3
'    Get_Leave_Affect_On_IncrementDate
'    Real_IncrementDt_Considering_Leave
'
'End Select
'
'End Sub
'Private Sub Get_Last_IncrementDt()
'On Error GoTo Errdes
'Dim conn5 As New Connection
'Dim cmd As New Command
'Dim RS2 As New ADODB.Recordset
'conn5.ConnectionString = strCN.Connection_String
'conn5.Open
'cmd.ActiveConnection = conn5
'cmd.CommandType = adCmdText
'cmd.CommandText = " select LAST_DT_INCRE from INCREMENT_RECORD where TRACK_ID=(select max(TRACK_ID) from INCREMENT_RECORD where EMP_ID='" & Combo1(1) & "')"
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'RS2.CursorLocation = adUseClient
'RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic
'If Not RS2.EOF Then
'    LDtofIncrem = Format(RS2.Fields(0), "dd/mm/yyyy")
'Else
'    LDtofIncrem = Format(Date$, "dd-mm-yyyy")
'End If
'RS2.Close
'conn5.Close
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Get_WithoutPay_Leave()
'On Error GoTo Errdes
'Dim conn5 As New Connection
'Dim cmd As New Command
'Dim RS2 As New ADODB.Recordset
'conn5.ConnectionString = strCN.Connection_String
'conn5.Open
'cmd.ActiveConnection = conn5
'cmd.CommandType = adCmdText
''-----------FOR GETTING THE LEAVE OF WITHOUT PAY
'cmd.CommandText = "SELECT sum(LEAVE_APPLICATION.NO_OF_DAYS_LEAVE) From Leave_Application " + _
'                " where LEAVE_APPLICATION.LEAVE_CODE='003' And LEAVE_APPLICATION.DATE_OF_APPLIED " + _
'                " Between '" & Format(LDtofIncrem, "dd-mm-yyyy") & "' and '" & Format(Next_Incre_Dt, "dd-mm-yyyy") & "'"
'
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'RS2.CursorLocation = adUseClient
'RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic
'If Not RS2.EOF Then
'    GetNoOfDaysLvWithoutPay = RS2.Fields(0)
'
'    If IsNull(GetNoOfDaysLvWithoutPay) Then
'        GetNoOfDaysLvWithoutPay = 0
'    End If
'    txtfields(2).Text = GetNoOfDaysLvWithoutPay
'Else
'     GetNoOfDaysLvWithoutPay = 0
'     txtfields(2).Text = GetNoOfDaysLvWithoutPay
'End If
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub Get_With_HalfPay_Leave()
'On Error GoTo Errdes
'Dim conn5 As New Connection
'Dim cmd As New Command
'Dim RS2 As New ADODB.Recordset
'conn5.ConnectionString = strCN.Connection_String
'conn5.Open
'cmd.ActiveConnection = conn5
'cmd.CommandType = adCmdText
''-----------FOR GETTING THE LEAVE OF half  PAY
'cmd.CommandText = "SELECT sum(LEAVE_APPLICATION.NO_OF_DAYS_LEAVE) From Leave_Application " + _
'                " where LEAVE_APPLICATION.LEAVE_CODE='004' And LEAVE_APPLICATION.DATE_OF_APPLIED " + _
'                " Between '" & Format(LDtofIncrem, "dd-mm-yyyy") & "' and '" & Format(Next_Incre_Dt, "dd-mm-yyyy") & "'"
'
'cmd.Properties("iRowsetChange") = True
'cmd.Properties("updatability") = 7
'RS2.CursorLocation = adUseClient
'RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic
'If Not RS2.EOF Then
'    GetNoOfDaysHalfLv = RS2.Fields(0)
'    If IsNull(GetNoOfDaysHalfLv) Then
'        GetNoOfDaysHalfLv = 0
'        txtfields(3).Text = GetNoOfDaysHalfLv
'    Else
'        txtfields(3).Text = Round(GetNoOfDaysHalfLv \ 2)
'    End If
'Else
'    GetNoOfDaysHalfLv = 0
'    txtfields(2).Text = GetNoOfDaysHalfLv
'End If
'
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'End Sub
'Private Sub txtfields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'Select Case Index
'Case 1
'    txtfields(2).SetFocus
'Case 2
'    LDtofIncrem.SetFocus
'Case 3
'    Next_Incre_Dt.SetFocus
'
'Case 4
'  txtfields(5).SetFocus
'Case 5
'  cmdSave.SetFocus
'End Select
'End If
'End Sub
