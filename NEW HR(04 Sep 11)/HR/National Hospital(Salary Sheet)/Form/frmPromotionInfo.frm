VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPromotionInfo 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employee's Promotion Record"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8970
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      Height          =   480
      Left            =   6360
      Picture         =   "frmPromotionInfo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4365
      Width           =   1230
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   7680
      Picture         =   "frmPromotionInfo.frx":1DCA
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4350
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   3720
      Picture         =   "frmPromotionInfo.frx":384C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4365
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   2400
      Picture         =   "frmPromotionInfo.frx":51DE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4365
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   5040
      Picture         =   "frmPromotionInfo.frx":6B70
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4365
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Promotion Information"
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
      Height          =   4080
      Left            =   135
      TabIndex        =   7
      Top             =   180
      Width           =   8730
      Begin VB.Frame Frame3 
         BackColor       =   &H80000009&
         Height          =   2175
         Left            =   135
         TabIndex        =   12
         Top             =   180
         Width           =   8460
         Begin VB.TextBox txtfields 
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   4
            Left            =   6210
            TabIndex        =   33
            Top             =   360
            Width           =   1965
         End
         Begin VB.TextBox txtfields 
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   6210
            TabIndex        =   32
            Top             =   750
            Width           =   1950
         End
         Begin VB.TextBox txtfields 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   2
            Left            =   6210
            TabIndex        =   14
            Top             =   1650
            Width           =   1125
         End
         Begin VB.TextBox txtfields 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   0
            Left            =   6210
            TabIndex        =   13
            Top             =   1155
            Width           =   1350
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   0
            Top             =   300
            Width           =   1410
         End
         Begin MSComCtl2.DTPicker LDtofIncrem 
            Height          =   315
            Left            =   1710
            TabIndex        =   1
            Top             =   1650
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
            Left            =   1680
            TabIndex        =   15
            Top             =   1170
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin VB.Label lblName 
            BackStyle       =   0  'Transparent
            Height          =   255
            Index           =   1
            Left            =   1650
            TabIndex        =   36
            Top             =   750
            Width           =   2535
         End
         Begin VB.Shape Shape10 
            BorderColor     =   &H00FFC0C0&
            Height          =   285
            Left            =   6120
            Top             =   330
            Width           =   2085
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00FFC0C0&
            Height          =   285
            Left            =   6120
            Top             =   720
            Width           =   2085
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Designation"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   4980
            TabIndex        =   27
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date of Join"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   23
            Top             =   1230
            Width           =   855
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00FFC0C0&
            Height          =   390
            Left            =   1665
            Top             =   1605
            Width           =   2145
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Promotion Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   22
            Top             =   1710
            Width           =   1440
         End
         Begin VB.Shape Shape9 
            BorderColor     =   &H00FFC0C0&
            Height          =   390
            Left            =   1635
            Top             =   1140
            Width           =   2145
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
            Left            =   90
            TabIndex        =   21
            Top             =   810
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
            Left            =   90
            TabIndex        =   19
            Top             =   360
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
            Left            =   4995
            TabIndex        =   18
            Top             =   765
            Width           =   825
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFC0C0&
            Height          =   375
            Index           =   0
            Left            =   1650
            Top             =   270
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
            Left            =   4995
            TabIndex        =   17
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Scale"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   4995
            TabIndex        =   16
            Top             =   1650
            Width           =   750
         End
         Begin VB.Shape Shape8 
            BorderColor     =   &H00FFC0C0&
            Height          =   285
            Left            =   6120
            Top             =   1605
            Width           =   2085
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FFC0C0&
            Height          =   285
            Left            =   6120
            Top             =   1140
            Width           =   2085
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
         Height          =   1635
         Left            =   135
         TabIndex        =   8
         Top             =   2340
         Width           =   8445
         Begin VB.ComboBox Combo2 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmPromotionInfo.frx":857A
            Left            =   6255
            List            =   "frmPromotionInfo.frx":858A
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1080
            Width           =   2040
         End
         Begin VB.ComboBox cboDept 
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2610
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1080
            Width           =   2040
         End
         Begin VB.ComboBox cboDesig 
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   6255
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   675
            Width           =   2040
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   6255
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   225
            Width           =   2040
         End
         Begin VB.TextBox txtfields 
            BorderStyle     =   0  'None
            Height          =   210
            Index           =   3
            Left            =   2610
            TabIndex        =   4
            Text            =   "0"
            Top             =   720
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   2610
            TabIndex        =   2
            Top             =   270
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
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Emp Class"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   5175
            TabIndex        =   35
            Top             =   1170
            Width           =   735
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   375
            Index           =   1
            Left            =   6210
            Top             =   1035
            Width           =   2130
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "New Department"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   31
            Top             =   1140
            Width           =   1200
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "New Designation"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   4695
            TabIndex        =   30
            Top             =   765
            Width           =   1215
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   375
            Index           =   33
            Left            =   2565
            Top             =   1035
            Width           =   2130
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   375
            Index           =   37
            Left            =   6210
            Top             =   630
            Width           =   2130
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Scale"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   4950
            TabIndex        =   11
            Top             =   315
            Width           =   960
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Present Promotion Effective Date"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   10
            Top             =   330
            Width           =   2355
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00FFC0C0&
            Height          =   285
            Left            =   2565
            Top             =   690
            Width           =   1185
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00FFC0C0&
            Height          =   375
            Left            =   6210
            Top             =   180
            Width           =   2130
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Current Basic"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   9
            Top             =   735
            Width           =   945
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FFC0C0&
            Height          =   375
            Left            =   2565
            Top             =   235
            Width           =   2130
         End
      End
   End
End
Attribute VB_Name = "frmPromotionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Increment_Rc As New Cls_IncrementPro
Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Combo2.SetFocus
End If
End Sub

Private Sub cboDesig_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cboDept.SetFocus
End If
End Sub

Private Sub cmdClear_Click()
    'lblName(1) = ""
    'lblDept(1) = ""
    txtfields(1) = ""
    txtfields(4) = ""
    txtfields(2) = ""
    txtfields(0) = ""
    txtfields(3) = ""
    Next_Incre_Dt = Date$
    LDtofIncrem = Date$
    DTPicker1 = Date$
    Combo1(0) = Combo1(0).List(0)
    Combo1(1) = Combo1(1).List(0)
    Combo1(1).SetFocus
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim f As New frmPromotionRecordReport
f.Show 1
End Sub

Private Sub cmdSave_Click()
With Increment_Rc
    .Connstring = strCN.Connection_String
    .Emp_ID = Trim(Combo1(1).Text)
    .LAST_PROM_DATE = LDtofIncrem
    .P_PROMOTION_EFF_DT = DTPicker1
    .CURRENT_BASIC = txtfields(3)
    .CURRENT_SCALE = Trim(Combo1(0))
    .ENTRY_DATE = Date$
    .ENTRY_BY = U_Id
    .FROMDEGINATION = Trim(txtfields(4).Text)
    .TODESIGNATION = Trim(cboDesig)
    .FROMDEPARTMENT = Trim(txtfields(1))
    .TODEPARTMENT = Trim(cboDept)
    .LastBasic = txtfields(0)
    .LastScale = txtfields(2)
    
    If Combo2.Text = Combo2.List(0) Then
        .EmpClass = 1
    ElseIf Combo2.Text = Combo2.List(1) Then
         .EmpClass = 2
    ElseIf Combo2.Text = Combo2.List(2) Then
         .EmpClass = 3
    Else
         .EmpClass = 4
    End If


    .Promotion_Save
End With
MsgBox "Data Save Successfully", vbInformation, "IT Division, DNMIH"
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Select Case Index
Case 1
    Get_Global_Info_All_Employee
    Get_Last_Promotion_Date
Case 0
    txtfields(3).SetFocus
End Select
End If
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdSave.SetFocus
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Combo1(0).SetFocus
End If
End Sub

Private Sub Form_Load()
On Error GoTo Errdes

Dim cmd As New Command
Dim conn10 As New Connection
Dim rs10 As New Recordset
Dim conn11 As New Connection
Dim rs11 As New Recordset

conn10.ConnectionString = strCN.Connection_String
conn10.Open
cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText

cmd.CommandText = "select emp_id from emp_info order by emp_id"


rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic

If rs10.RecordCount > 0 Then
    Do Until rs10.EOF
        Combo1(1).AddItem rs10.Fields(0)
        rs10.MoveNext
    Loop
    Combo1(1) = Combo1(1).List(0)
End If
rs10.Close
conn10.Close
Get_Pay_Scale

Load_Desig Me
Load_Department Me

Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
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
'cmd.CommandText = "SELECT EMP_INFO.EMP_NM,EMP_JOB_INFO.JDATE,EMP_JOB_INFO.SCALE_CODE, " + _
'                " EMP_JOB_INFO.Basic_Sal , ST_DEPT.DEPT_NM " + _
'                " From Emp_info, EMP_JOB_INFO, ST_DEPT " + _
'                " WHERE ((EMP_JOB_INFO.EMP_ID=EMP_INFO.EMP_ID) " + _
'                " AND (EMP_JOB_INFO.DEPT=ST_DEPT.DEPT_CODE) " + _
'                " AND EMP_INFO.EMP_ID='" & Combo1(1) & "')"
cmd.CommandText = "SELECT EMP_INFO.EMP_NM,EMP_JOB_INFO.JDATE,EMP_JOB_INFO.SCALE_CODE," + _
                    " EMP_JOB_INFO.Basic_Sal,ST_DEPT.DEPT_NM,st_desig.DESIGNATION From Emp_info, EMP_JOB_INFO, " + _
                    " ST_DEPT,st_desig WHERE ((EMP_JOB_INFO.EMP_ID=EMP_INFO.EMP_ID) " + _
                    " AND (EMP_JOB_INFO.DEPT=ST_DEPT.DEPT_CODE) " + _
                    " and (emp_job_info.DESIG=st_desig.DESIG_CODE) AND EMP_INFO.EMP_ID='" & Combo1(1) & "') "


cmd.Properties("iRowsetChange") = True
cmd.Properties("updatability") = 7
RS2.CursorLocation = adUseClient

RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic

If Not RS2.EOF Then
    lblName(1) = RS2.Fields(0)
    txtfields(1) = RS2.Fields(4)
    Next_Incre_Dt = RS2.Fields(1)
    txtfields(2) = RS2.Fields(2)
    txtfields(0) = RS2.Fields(3)
    txtfields(4) = RS2.Fields(5)
Else
    MsgBox "Invalid Emp ID !", vbCritical, "IT Division, DNMIH"
    Combo1(1).SetFocus
    Exit Sub
End If
    LDtofIncrem.SetFocus
    
RS2.Close
conn5.Close
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Get_Pay_Scale()
On Error GoTo Errdes
Dim conn005 As New Connection
Dim cmd As New Command
Dim RS002 As New ADODB.Recordset
conn005.ConnectionString = strCN.Connection_String
conn005.Open
cmd.ActiveConnection = conn005
cmd.CommandType = adCmdText
cmd.CommandText = "select SCALE_CODE from ST_PAYSCALE"
cmd.Properties("iRowsetChange") = True
cmd.Properties("updatability") = 7
RS002.CursorLocation = adUseClient

RS002.Open cmd.CommandText, conn005, adOpenDynamic, adLockOptimistic

If RS002.RecordCount > 0 Then
    Do Until RS002.EOF
        Combo1(0).AddItem RS002.Fields(0)
        RS002.MoveNext
    Loop
    Combo1(0) = Combo1(0).List(0)
End If
    
RS002.Close
conn005.Close
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Get_Last_Promotion_Date()
On Error GoTo Errdes
Dim conn5 As New Connection
Dim cmd As New Command
Dim RS2 As New ADODB.Recordset
conn5.ConnectionString = strCN.Connection_String
conn5.Open
cmd.ActiveConnection = conn5
cmd.CommandType = adCmdText
cmd.CommandText = "Select LAST_PROM_DATE from promotion_info where emp_id='" & Combo1(1) & "' and track_id=(select max(TRACK_ID) from promotion_info where emp_id='" & Combo1(1) & "')"
cmd.Properties("iRowsetChange") = True
cmd.Properties("updatability") = 7
RS2.CursorLocation = adUseClient

RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic

If RS2.RecordCount > 0 Then
    LDtofIncrem = RS2.Fields(0)
Else
    LDtofIncrem = Date$
End If
    
RS2.Close
conn5.Close
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub LDtofIncrem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    DTPicker1.SetFocus
End If
End Sub

Private Sub txtfields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Select Case Index
Case 3
  cboDesig.SetFocus

End Select
End If
End Sub
