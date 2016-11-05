VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProvidentFund 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "P&rocess"
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
      Left            =   2790
      TabIndex        =   1
      Top             =   6465
      Width           =   1275
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
      Left            =   5355
      Picture         =   "frmProvidentFund.frx":0000
      TabIndex        =   3
      Top             =   6465
      Width           =   1215
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
      Left            =   4095
      Picture         =   "frmProvidentFund.frx":08CA
      TabIndex        =   2
      Top             =   6465
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Provident Fund Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   6270
      Left            =   30
      TabIndex        =   4
      Top             =   60
      Width           =   6885
      Begin VB.Frame Frame4 
         Height          =   1005
         Left            =   3090
         TabIndex        =   30
         Top             =   4140
         Visible         =   0   'False
         Width           =   3675
         Begin VB.ComboBox cmbPaymentPurpose 
            Height          =   315
            Left            =   1380
            TabIndex        =   32
            Top             =   570
            Width           =   2175
         End
         Begin VB.ComboBox cmbPFSourceofFund 
            Height          =   315
            Left            =   1380
            TabIndex        =   31
            Top             =   180
            Width           =   2175
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Payment Purpose"
            Height          =   195
            Left            =   90
            TabIndex        =   34
            Top             =   630
            Width           =   1245
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Source of Fund"
            Height          =   195
            Left            =   90
            TabIndex        =   33
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "PF Payment (Purpose Wise)"
         Height          =   315
         Index           =   7
         Left            =   210
         TabIndex        =   29
         Top             =   4590
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "PF Receive (Source of Fund Wise)"
         Height          =   315
         Index           =   6
         Left            =   210
         TabIndex        =   28
         Top             =   4200
         Width           =   2955
      End
      Begin VB.Frame Frame3 
         Caption         =   "Date"
         Height          =   675
         Left            =   150
         TabIndex        =   17
         Top             =   5430
         Width           =   6165
         Begin MSComCtl2.DTPicker Begin_Fiscal_date 
            Height          =   315
            Left            =   930
            TabIndex        =   18
            Top             =   210
            Width           =   1800
            _ExtentX        =   3175
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
            Format          =   58523651
            CurrentDate     =   37073
         End
         Begin MSComCtl2.DTPicker End_Fiscal_Year 
            Height          =   315
            Left            =   3420
            TabIndex        =   19
            Top             =   210
            Width           =   1800
            _ExtentX        =   3175
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
            Format          =   58523651
            CurrentDate     =   37072
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "To"
            Height          =   195
            Left            =   2880
            TabIndex        =   27
            Top             =   270
            Width           =   195
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "From"
            Height          =   195
            Left            =   270
            TabIndex        =   26
            Top             =   270
            Width           =   345
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "PF Balance Sheet"
         Height          =   315
         Index           =   5
         Left            =   180
         TabIndex        =   16
         Top             =   3540
         Width           =   1635
      End
      Begin VB.Frame Frame2 
         Height          =   1485
         Left            =   3120
         TabIndex        =   15
         Top             =   2640
         Visible         =   0   'False
         Width           =   3645
         Begin VB.ComboBox cmbAccountType 
            Height          =   315
            Left            =   1350
            TabIndex        =   22
            Top             =   630
            Width           =   2175
         End
         Begin VB.ComboBox cmbBankCode 
            Height          =   315
            Left            =   1350
            TabIndex        =   21
            Top             =   270
            Width           =   2175
         End
         Begin VB.ComboBox cmbAccountNo 
            Height          =   315
            Left            =   1350
            TabIndex        =   20
            Top             =   990
            Width           =   2175
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Bank Name"
            Height          =   195
            Left            =   300
            TabIndex        =   25
            Top             =   330
            Width           =   840
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Account Type"
            Height          =   195
            Left            =   300
            TabIndex        =   24
            Top             =   660
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Account No"
            Height          =   195
            Left            =   300
            TabIndex        =   23
            Top             =   1020
            Width           =   855
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bank Statement "
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   14
         Top             =   3240
         Width           =   2385
      End
      Begin VB.OptionButton Option1 
         Caption         =   "PF Cash Book"
         Height          =   315
         Index           =   3
         Left            =   180
         TabIndex        =   13
         Top             =   2910
         Width           =   2385
      End
      Begin VB.OptionButton Option1 
         Caption         =   "PF Income Expendeture"
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   2550
         Width           =   2715
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   810
         Width           =   1290
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1950
         Width           =   4410
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Top             =   1950
         Width           =   1230
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Provident Fund Details at the End of Year"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   1350
         Width           =   4740
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Latest Provident Fund Details"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   4740
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Index           =   1
         Left            =   1980
         Top             =   810
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Emp ID"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   690
         TabIndex        =   9
         Top             =   870
         Width           =   945
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C0C0&
         Height          =   555
         Left            =   195
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   6270
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BorderColor     =   &H00FECCC7&
         Height          =   405
         Index           =   0
         Left            =   585
         Top             =   1920
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Employee Name"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3105
         TabIndex        =   8
         Top             =   1710
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Emp ID"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   900
         TabIndex        =   7
         Top             =   1680
         Width           =   525
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         Height          =   825
         Left            =   225
         Shape           =   4  'Rounded Rectangle
         Top             =   1650
         Width           =   6225
      End
   End
End
Attribute VB_Name = "frmProvidentFund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PV_Preparation As New Cls_salary_Preparation



Private Sub Combo1_Click(Index As Integer)
Get_Emp_Name
End Sub
Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Get_Emp_Name
End If
End Sub
Private Sub Command1_Click()
On Error GoTo Errdes
If Option1(0).Value = True Then
   rptmode = 10
   Emp_ID_Value = Trim(Combo1(1))
   Form20.Show vbModal
ElseIf Option1(1).Value = True Then
   rptmode = 16
   Emp_ID_Value = Trim(Combo1(0).Text)
   Form20.Show vbModal
ElseIf Option1(2).Value = True Then
   rptmode = 40
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
ElseIf Option1(3).Value = True Then
   rptmode = 41
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
ElseIf Option1(4).Value = True Then
'MsgBox "Under Development"
    If cmbBankCode.Text = "" Then
    MsgBox "Bank Name Required"
    ElseIf cmbAccountType.Text = "" Then
    MsgBox "Account Type Required"
    ElseIf cmbAccountNo.Text = "" Then
    MsgBox "Account No Required"
    Else
    rptmode = 42
    sBankCode = Get_Code(cmbBankCode.Text)
    sBankName = Get_Description(cmbBankCode.Text)
    sAccountType = Get_Code(cmbAccountType.Text)
    sAccountTypeName = Get_Description(cmbAccountType.Text)
    sAccountNo = Get_Code(cmbAccountNo.Text)
    BeginDateForReport = Begin_Fiscal_date
    EnddateforReport = End_Fiscal_Year
    Form20.Show vbModal
   End If
ElseIf Option1(5).Value = True Then
   rptmode = 43
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
ElseIf Option1(6).Value = True Then
   If cmbPFSourceofFund.Text = "" Then
   MsgBox "Source of Fund Required"
   Else
   rptmode = 44
   sSourceId = Get_Code(cmbPFSourceofFund.Text)
   sSourceName = Get_Description(cmbPFSourceofFund.Text)
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
   End If
ElseIf Option1(7).Value = True Then
   If cmbPaymentPurpose.Text = "" Then
   MsgBox "Purpose of Payment Required"
   Else
   rptmode = 45
   sPurposeId = Get_Code(cmbPaymentPurpose.Text)
   sPurposeName = Get_Description(cmbPaymentPurpose.Text)
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
   End If
End If
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Command3_Click()
On Error GoTo Errdes
'If Trim(Len(Combo1(0))) = 0 Then
'    MsgBox "Process Terminated as Emp ID is not avialable", vbInformation, "IT Division, DNMIH"
'    Combo1(0).SetFocus
'    Exit Sub
' Else
    Emp_ID_Value = Trim(Combo1(0).Text)
    End_Of_Year_PV_Calculation
    MsgBox "Process has been Completed for the Employee", vbInformation, "IT Division, DNMIH"
' End If
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = Asc(Chr(96))
End If
End Sub

Private Sub Form_Load()
On Error GoTo Errdes
get_Value_Into_Bank_Name
get_Value_Into_Account_type
get_Value_Into_Account_No
get_Value_Into_PF_Payment_Purpose
get_Value_Into_PF_Source_of_Fund
Dim cmd As New Command
Dim conn10 As New Connection
Dim rs10 As New Recordset
Dim conn11 As New Connection
Dim rs11 As New Recordset

conn10.ConnectionString = strCN.Connection_String
conn10.Open
cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText

cmd.CommandText = "select emp_id from emp_info order by emp_id "


rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic

If rs10.RecordCount > 0 Then
    Do Until rs10.EOF
        Combo1(0).AddItem rs10.Fields(0)
        Combo1(1).AddItem rs10.Fields(0)
        rs10.MoveNext
    Loop
    Combo1(0) = Combo1(0).List(0)
End If
Combo1(0).Text = Combo1(0).List(0)
Combo1(1).Text = Combo1(1).List(0)
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub End_Of_Year_PV_Calculation()
On Error GoTo Errdes
With PV_Preparation
    .Connstring = strCN.Connection_String
    .Emp_ID = Combo1(0)
    .PF_EndofYear_Save
End With
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Get_Emp_Name()
On Error GoTo Errdesc
Dim cmd As New Command
Dim conn2 As New ADODB.Connection
Dim RS2 As New ADODB.Recordset
conn2.ConnectionString = strCN.Connection_String
conn2.Open
cmd.ActiveConnection = conn2
cmd.CommandType = adCmdText

cmd.CommandText = "select EMP_NM from emp_info where Emp_Id='" & Combo1(0) & "'"

RS2.CursorLocation = adUseClient
RS2.Open cmd.CommandText, conn2, adOpenDynamic, adLockOptimistic

    If RS2.RecordCount > 0 Then
        txtfields.Text = RS2.Fields(0)
        RS2.Close
        conn2.Close
    Else
        txtfields.Text = ""
        RS2.Close
        conn2.Close
    End If
Exit Sub
Errdesc:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub


Private Sub get_Value_Into_Account_type()
On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim Rs As New ADODB.Recordset
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select TYPE_ID,TYPE_NAME from L_ACCOUNT_TYPE order by TYPE_ID"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    Rs.CursorLocation = adUseClient
    Rs.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
       If Rs.RecordCount > 0 Then
            Do Until Rs.EOF
            cmbAccountType.AddItem Rs.Fields(1) & "~" & Rs.Fields(0)
            Rs.MoveNext
            Loop
            'cmbAccountType.ListIndex = 0
        End If

    
    
    Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub get_Value_Into_Bank_Name()
On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim Rs As New ADODB.Recordset
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select BANK_ID,BANK_NAME from L_BANK order by BANK_ID"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    Rs.CursorLocation = adUseClient
    Rs.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
       If Rs.RecordCount > 0 Then
            Do Until Rs.EOF
            cmbBankCode.AddItem Rs.Fields(1) & "~" & Rs.Fields(0)
            Rs.MoveNext
            Loop
        End If

    
    
    Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub get_Value_Into_Account_No()
On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim Rs As New ADODB.Recordset
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText

    cmd.CommandText = "select ACCOUNT_NO from MEMBER_FUND"

    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    Rs.CursorLocation = adUseClient
    Rs.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
       If Rs.RecordCount > 0 Then
            Do Until Rs.EOF
            cmbAccountNo.AddItem Rs.Fields(0)
            Rs.MoveNext
            Loop
        End If
    Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
Frame2.Visible = False
Frame4.Visible = False
Case 1
Frame2.Visible = False
Frame4.Visible = False
Case 2
Frame2.Visible = False
Frame4.Visible = False
Case 3
Frame2.Visible = False
Frame4.Visible = False
Case 4
Frame2.Visible = True
Frame4.Visible = False
Case 5
Frame2.Visible = False
Frame4.Visible = False
Case 6
Frame2.Visible = False
Frame4.Visible = True
Label9.Visible = True
cmbPFSourceofFund.Visible = True
cmbPaymentPurpose.Visible = False
Label10.Visible = False
Frame4.Height = 600
Case 7
Frame2.Visible = False
Frame4.Visible = True
cmbPFSourceofFund.Visible = False
cmbPaymentPurpose.Visible = True
Label9.Visible = False
Label10.Visible = True
Frame4.Height = 600
cmbPaymentPurpose.Top = 180
Label10.Top = 240
End Select
End Sub
Private Sub get_Value_Into_PF_Source_of_Fund()
On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim Rs As New ADODB.Recordset
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
        
    cmd.CommandText = "select SOURCE_ID,SOURCE_NAME from L_PF_SOURCEOF_FUND order by SOURCE_ID"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    Rs.CursorLocation = adUseClient
    Rs.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
       If Rs.RecordCount > 0 Then
            Do Until Rs.EOF
            cmbPFSourceofFund.AddItem Rs.Fields(1) & "~" & Rs.Fields(0)
            Rs.MoveNext
            Loop
        End If

    
    
    Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub get_Value_Into_PF_Payment_Purpose()
On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim Rs As New ADODB.Recordset
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select PURPOSE_ID,PURPOSE_NAME from L_PF_PAYMENT_PURPOSE order by PURPOSE_ID"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    Rs.CursorLocation = adUseClient
    Rs.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
       If Rs.RecordCount > 0 Then
            Do Until Rs.EOF
            cmbPaymentPurpose.AddItem Rs.Fields(1) & "~" & Rs.Fields(0)
            Rs.MoveNext
            Loop
        End If

    
    
    Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
