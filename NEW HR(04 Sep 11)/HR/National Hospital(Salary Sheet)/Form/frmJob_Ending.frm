VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form23 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Job Ending & Others"
   ClientHeight    =   3660
   ClientLeft      =   660
   ClientTop       =   1560
   ClientWidth     =   8070
   Icon            =   "frmJob_Ending.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   2760
      Index           =   2
      Left            =   180
      TabIndex        =   6
      Top             =   135
      Width           =   7755
      Begin VB.TextBox TextBox1 
         Height          =   615
         Left            =   1260
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1920
         Width           =   5685
      End
      Begin VB.ComboBox cboEndType 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1440
         Width           =   1845
      End
      Begin VB.TextBox txtEmpID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         TabIndex        =   13
         Top             =   390
         Width           =   1755
      End
      Begin MSComCtl2.DTPicker dtpIssue_Dt 
         Height          =   330
         Index           =   0
         Left            =   4455
         TabIndex        =   0
         Top             =   1350
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   582
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   12582912
         CalendarTrailingForeColor=   16576
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   64946177
         CurrentDate     =   37722
      End
      Begin VB.Label lblCost 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Left            =   4470
         TabIndex        =   18
         Top             =   960
         Width           =   2625
      End
      Begin VB.Label lblDesig 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Left            =   1380
         TabIndex        =   15
         Top             =   960
         Width           =   2625
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Left            =   4350
         TabIndex        =   14
         Top             =   450
         Width           =   2625
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   1
         Top             =   1875
         Width           =   795
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ending type"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   315
         TabIndex        =   11
         Top             =   1425
         Width           =   840
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   3555
         TabIndex        =   12
         Top             =   1425
         Width           =   675
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3555
         TabIndex        =   10
         Top             =   945
         Width           =   825
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Name"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   3555
         TabIndex        =   9
         Top             =   435
         Width           =   465
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   8
         Top             =   435
         Width           =   900
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   7
         Top             =   945
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdPreview 
      Height          =   480
      Left            =   2610
      Picture         =   "frmJob_Ending.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3060
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   3915
      Picture         =   "frmJob_Ending.frx":2694
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3060
      Width           =   1140
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   1350
      Picture         =   "frmJob_Ending.frx":4116
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3060
      Width           =   1140
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   135
      Picture         =   "frmJob_Ending.frx":5AA8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3060
      Width           =   1095
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objEmp_Info As New emp_info
Private Job_Info As New clsEmp_Job_Detail
Dim job_ending As New JobEndingInfo

Private Sub cmdClear_Click()
    Clear_Screen
    txtEmpID.SetFocus
End Sub

Private Sub cmdClose_Click()
    Close_Msg Me
End Sub
Private Sub cmdPreview_Click()
    rptmode = 8
    Form20.Show vbModal
End Sub
Private Sub cmdSave_Click()
With job_ending

    .Connstring = strCN.Connection_String
    .Emp_ID = Trim(txtEmpID)
    .Emp_Name = Trim(lblName)
    .designation = Trim(lblDesig)
    .JobEnding_Type = Trim(cboEndType)
    .JobEndingDate = Format(dtpIssue_Dt(0), "dd-mmm-yyyy")
    .Desciption = Trim(TextBox1.Text)
    
    If cboEndType.Text = cboEndType.List(0) Then
        .JobEndingStatus = 0    '-----None
    ElseIf cboEndType.Text = cboEndType.List(1) Then
        .JobEndingStatus = 1    '-----Retirement
    ElseIf cboEndType.Text = cboEndType.List(2) Then
        .JobEndingStatus = 2    '----Golden HandShake
    ElseIf cboEndType.Text = cboEndType.List(3) Then
        .JobEndingStatus = 3    '----Transfer
    ElseIf cboEndType.Text = cboEndType.List(4) Then
        .JobEndingStatus = 4    '----Suspend
    ElseIf cboEndType.Text = cboEndType.Text = cboEndType.List(5) Then
        .JobEndingStatus = 5    '----Death
    Else
        .JobEndingStatus = 6    '----Others
    End If
        
    .Department = Trim(lblCost)
    .Save
    MsgBox "Data Saved Successfully", vbInformation, "IT Division, DNMIH"
 End With
  
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

On Error Resume Next
    Screen_Position Me
    txtEmpID.MaxLength = Id_Len

    With cboEndType
        .AddItem "None"
        .AddItem "Retirement"
        .AddItem "Golden Handshake"
        .AddItem "Transfer"
        .AddItem "Suspend"
        .AddItem "Death"
        .AddItem "Others"
        .ListIndex = 0
    End With
    
  ' dtpEnd_Date = Now
  'dtpIssue_Dt(0) = Date
    
End Sub
'Private Sub txtEmpID_KeyDown(index As Integer, keycode As MSForms.ReturnInteger, Shift As Integer)
'
'End Sub
Private Sub txtEmp_ID_KeyDown(KeyCode As Integer)
If KeyCode = 13 Then
    cboEndType.SetFocus
End If
End Sub

Public Sub Flash_Data()
  'On Error Resume Next

    With objEmp_Info
        .Connstring = strCN.Connection_String
        .Emp_ID = txtEmpID
        '.get_Job_End_X
        '.GetX
'            lblName = .Emp_Nm
'            lblUnit = .Unit_Nm
'            lblCost = .Cost_Nm
'            lblDesig = .Desig_Nm
'            cboEndType = .End_Type
'            dtpEnd_Date = .End_date
    End With

    txtEmpID.SetFocus
    
End Sub

Private Sub txtEmpId_LostFocus()
    Get_Global_Info_All_Employee
End Sub
Private Sub Get_Global_Info_All_Employee()
On Error GoTo Errdes
Dim conn5 As New Connection
Dim cmd As New Command
Dim RS2 As New ADODB.Recordset
conn5.ConnectionString = strCN.Connection_String
conn5.Open
cmd.ActiveConnection = conn5
cmd.CommandType = adCmdText
cmd.CommandText = "SELECT  a.Emp_Id, a.Emp_Nm,c.Designation,d.DEPT_NM " + _
                " from Emp_Info a, Emp_Job_Info b ,St_Desig c, St_Dept d " + _
                " Where A.Emp_Id = b.Emp_Id and a.Emp_ID=b.Emp_ID " + _
                " and b.Desig= c.Desig_code and b.Dept= d.Dept_code and a.Emp_Id='" & txtEmpID & "'"

cmd.Properties("iRowsetChange") = True
cmd.Properties("updatability") = 7
RS2.CursorLocation = adUseClient

RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic

If Not RS2.EOF Then
    lblName = RS2.Fields(1)
    lblDesig = RS2.Fields(2)
    lblCost = RS2.Fields(3)
End If
    
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"


End Sub
