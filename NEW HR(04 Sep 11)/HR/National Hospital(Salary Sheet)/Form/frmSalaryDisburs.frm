VERSION 5.00
Begin VB.Form frmSalaryDisburs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pay Slip"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8220
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "&Exit"
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
      Left            =   6795
      Picture         =   "frmSalaryDisburs.frx":0000
      TabIndex        =   3
      Top             =   1575
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
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
      Left            =   5535
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSalaryDisburs.frx":08CA
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pay Slip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1365
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   7980
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1125
         TabIndex        =   0
         Top             =   375
         Width           =   1590
      End
      Begin VB.CommandButton cmdView 
         Height          =   315
         Left            =   2745
         Picture         =   "frmSalaryDisburs.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   4275
         TabIndex        =   6
         Top             =   810
         Width           =   1095
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   1125
         TabIndex        =   5
         Top             =   810
         Width           =   1725
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3150
         TabIndex        =   4
         Top             =   375
         Width           =   4605
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FECCC7&
         Height          =   380
         Left            =   1080
         Top             =   765
         Width           =   4335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FECCC7&
         Height          =   385
         Left            =   1080
         Top             =   315
         Width           =   6720
      End
      Begin VB.Label Label1 
         Caption         =   "Emp ID"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   270
         TabIndex        =   10
         Top             =   390
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Month"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   870
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3465
         TabIndex        =   8
         Top             =   855
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmSalaryDisburs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Salaray_Pre As New Cls_salary_Preparation

Private Sub cboMonth_Click()
Get_Employee_Whom_Slary_has_Been_Prepared
End Sub

Private Sub cboYear_Click()
 '' Get_Employee_Whom_Slary_has_Been_Prepared
End Sub

Private Sub cmdView_Click()
On Error GoTo Errdes
Dim f2 As New frmDataSelect
Dim getconnected As New Connection
Dim cmd As New Command
Dim myrs As New ADODB.Recordset

    getconnected.ConnectionString = strCN.Connection_String
    getconnected.Open
    cmd.ActiveConnection = getconnected
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT EMP_INFO.EMP_ID, EMP_INFO.EMP_NM FROM EMP_INFO  "
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs.CursorLocation = adUseClient
    
    myrs.Open cmd.CommandText, getconnected, adOpenDynamic, adLockOptimistic
    
    
     Set f2.adoRecordset = myrs
     Set f2.OwnerForm = Me
     f2.Width = 6500
     f2.grdDataGrid.Columns(0).Caption = "Emp ID"
     f2.grdDataGrid.Columns(1).Caption = "Name"
     f2.grdDataGrid.Columns(0).Width = 1800
     f2.grdDataGrid.Columns(1).Width = 5500
     f2.intPutSel = 0
     f2.Show 1
     Combo1(0).Text = myrs.Fields(0)
     Text1 = myrs.Fields(1)

Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "Daffodil Sotware Ltd"

End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Select Case Index
Case 0
    If KeyCode = 13 Then
        Get_All_Information
    End If
End Select
End If
End Sub

Private Sub Command1_Click()
'Update_Salry_Preparation_Table
If ReportTracker = 3 Then
    rptmode = 12
Else
    rptmode = 28

End If
GetMonthOftheYear = cboMonth
Rpt_Year = cboYear
Emp_ID_Value = Trim(Combo1(0).Text)
Form20.Show vbModal
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo Errdes
Load_Yr Me
Load_MonthNm Me

If ReportTracker = 3 Then
    Me.Caption = "Payslip Preparation(For Salary)"
Else
    Me.Caption = "Payslip Preparation(For Bonus)"
End If

Dim cmd As New Command
Dim conn10 As New Connection
Dim rs10 As New Recordset

conn10.ConnectionString = strCN.Connection_String
conn10.Open
cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText

cmd.CommandText = "select emp_id from emp_info order by emp_id"
rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic

If rs10.RecordCount > 0 Then

    Do Until rs10.EOF
        Combo1(0).AddItem rs10.Fields(0)
        rs10.MoveNext
    Loop


End If

rs10.Close
conn10.Close

Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "Daffodil Sotware Ltd"
End Sub
Private Sub Update_Salry_Preparation_Table()
On Error GoTo Errdes
With Salaray_Pre
    .Connstring = strCN.Connection_String
    .Emp_ID = Combo1(0)
    .PAY_MONTH = cboMonth
    .PAY_YEAR = cboYear
    .Salary_Update_Save
End With
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "Daffodil Sotware Ltd"
End Sub
Private Sub Get_All_Information()
On Error GoTo Errdesc
Dim conn1 As New Connection
Dim rs1 As New Recordset
Dim cmd As New Command
conn1.ConnectionString = strCN.Connection_String
conn1.Open
cmd.ActiveConnection = conn1
cmd.CommandType = adCmdText
cmd.CommandText = "SELECT Emp_Nm  From emp_info  WHERE EMP_ID='" & Combo1(0) & "'"
rs1.CursorLocation = adUseClient
rs1.Open cmd.CommandText, conn1, adOpenDynamic, adLockOptimistic

    If rs1.RecordCount > 0 Then
        Text1 = rs1.Fields(0)
        Command1.SetFocus
    Else
        MsgBox "Invalid Employee No.", vbInformation, "Warning:IT Division, DNMIH"
        rs1.Close
        conn1.Close
        Exit Sub
    End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Update_Accounts_Table()
On Error GoTo Errdes
With Salaray_Pre
    .Connstring = strCN.Connection_String
    .Emp_ID = Combo1(0)
    .PAY_MONTH = cboMonth
    .PAY_YEAR = cboYear
    .Salary_Update_Save
End With
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Get_Employee_Whom_Slary_has_Been_Prepared()
Dim cmd As New Command
Dim conn As New Connection
Dim Rs As New Recordset

conn.ConnectionString = strCN.Connection_String
conn.Open
cmd.ActiveConnection = conn
cmd.CommandType = adCmdText


cmd.CommandText = "Select EMP_ID from salary_preparation WHERE SALARY_DISBURSE='0' " + _
                " and PAY_MONTH ='" & cboMonth & "' and PAY_YEAR='" & cboYear & "'"
Rs.CursorLocation = adUseClient
Rs.Open cmd.CommandText, conn, adOpenDynamic, adLockOptimistic

If Rs.RecordCount > 0 Then
    Combo1(0).Clear
    Do Until Rs.EOF
        Combo1(0).AddItem Rs.Fields(0)
        Rs.MoveNext
    Loop
    
  If Not Combo1(0).ListCount <> 0 Then
    Combo1(0) = Combo1(0).List(0)
  End If
  
End If

Rs.Close
conn.Close


End Sub

