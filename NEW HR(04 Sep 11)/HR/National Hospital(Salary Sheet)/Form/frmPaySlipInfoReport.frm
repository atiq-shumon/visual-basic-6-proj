VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPaySlipInfoReport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Salary Disbursment"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7530
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6165
      Picture         =   "frmPaySlipInfoReport.frx":0000
      TabIndex        =   3
      Top             =   2475
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
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
      Left            =   4860
      Picture         =   "frmPaySlipInfoReport.frx":08CA
      TabIndex        =   12
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Salary Disbursement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2265
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   7350
      Begin VB.OptionButton optPayment_Mode 
         Caption         =   "Cash"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1710
         TabIndex        =   13
         Top             =   1350
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton optPayment_Mode 
         Caption         =   "Bank"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2715
         TabIndex        =   11
         Top             =   1350
         Width           =   780
      End
      Begin VB.TextBox Text3 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   5220
         TabIndex        =   10
         Top             =   1320
         Width           =   1995
      End
      Begin VB.ComboBox cboAcc_Name 
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
         Height          =   315
         Left            =   1650
         TabIndex        =   8
         Top             =   390
         Width           =   5565
      End
      Begin VB.ComboBox cboMonth 
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
         Height          =   315
         Left            =   1665
         TabIndex        =   5
         Top             =   1700
         Width           =   2220
      End
      Begin VB.ComboBox cboYear 
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
         Height          =   315
         Left            =   5220
         TabIndex        =   4
         Top             =   1700
         Width           =   2040
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         ItemData        =   "frmPaySlipInfoReport.frx":1194
         Left            =   1650
         List            =   "frmPaySlipInfoReport.frx":11A4
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   765
         Width           =   5565
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Check No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4140
         TabIndex        =   15
         Top             =   1350
         Width           =   765
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   360
         Index           =   2
         Left            =   5175
         Top             =   1665
         Width           =   2100
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   360
         Index           =   1
         Left            =   1620
         Top             =   1665
         Width           =   2280
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mode of Payment"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   1305
         Width           =   1245
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   285
         Index           =   24
         Left            =   5175
         Top             =   1305
         Width           =   2085
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   405
         Index           =   0
         Left            =   1620
         Top             =   765
         Width           =   5655
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Available Head"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   450
         Width           =   1080
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   360
         Index           =   23
         Left            =   1620
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4230
         TabIndex        =   7
         Top             =   1710
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Month"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   1710
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Employee Class"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   825
         Width           =   1110
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   135
      Top             =   2565
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
End
Attribute VB_Name = "frmPaySlipInfoReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Salaray_Pre As New Cls_salary_Preparation
Dim Get_Salary_Into_Text
Dim EmpClassType
Dim VoucherNumber As Double
Private Sub cboAcc_code_Click()
On Error GoTo Errdesc
Dim conn9 As New Connection
Dim rs9 As New Recordset
Dim cmd As New Command
conn9.ConnectionString = strCN.Connection_String
conn9.Open
cmd.ActiveConnection = conn9
cmd.CommandType = adCmdText
cmd.CommandText = "select acc_name from account.acct where acc_code not in(select acc_head from account.acct) and user_acc='" & cboAcc_code & "'"
rs9.CursorLocation = adUseClient
rs9.Open cmd.CommandText, conn9, adOpenDynamic, adLockOptimistic
    If rs9.RecordCount > 0 Then
        cboAcc_Name = rs9.Fields(0)
    End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub cboAcc_Name_Click()
On Error GoTo Errdesc
Dim conn10 As New Connection
Dim rs10 As New Recordset
Dim cmd As New Command
conn10.ConnectionString = strCN.Connection_String
conn10.Open
cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText
cmd.CommandText = "select user_acc from account.acct where acc_code not in(select acc_head from account.acct) and acc_name='" & cboAcc_Name & "'"
rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic
    If rs10.RecordCount > 0 Then
        cboAcc_code = rs10.Fields(0)
    End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case Index
Case 0
    If KeyCode = 13 Then
        'Get_All_Information
    End If
End Select
End Sub
Private Sub Command1_Click()
On Error GoTo Erres
Get_Employee_Status
Get_Salary_That_Employee
Get_Voucher_Number
'Update_Salry_Preparation_Table
Call Update_Account_TableSpace_Table
Call Update_Account_TableSpace_Table1
Call Post_to_Ledger
MsgBox "Process has Completted Successfully", vbInformation, "IT Division, DNMIH"
Exit Sub
Erres:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo Errdes
Load_Yr Me
Load_MonthNm Me
''''Call GetUserAcc ------------------ AFTER INTEGRATION
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "Daffodil Sotware Ltd"
End Sub
Private Sub GetUserAcc()

        cboAcc_code.Clear
        cboAcc_Name.Clear

        Adodc1.ConnectionString = strCN.Connection_String
        Adodc1.RecordSource = "select user_acc,acc_name from account.acct where acc_code not in(select acc_head from account.acct) order by user_acc"
        Adodc1.Refresh
        If Adodc1.Recordset.RecordCount > 0 Then
            Do Until Adodc1.Recordset.EOF
                cboAcc_code.AddItem Adodc1.Recordset!user_acc
                cboAcc_Name.AddItem Adodc1.Recordset!acc_name
                Adodc1.Recordset.MoveNext
            Loop
        End If
End Sub
Private Sub Update_Account_TableSpace_Table() '''''for debit entry
On Error GoTo Errdes
With Salaray_Pre
    .Connstring = strCN.Connection_String
    .Vou_No = VoucherNumber
    .AccountHitdate = Date
    .NarrationofAcc = "Salary For the Month of " & cboMonth
    .Acc_Code = "8101" ''''-----SALARY HEAD CODE

    .Acc_NetPayDr = Get_Salary_Into_Text
    .Acc_NetPayCr = 0

    If optPayment_Mode(0).Value = True Then
        .Acc_PaymenyCPBD = "CP"
    Else
        .Acc_PaymenyCPBD = "BP"
    End If
    .Acc_Check_No = Text3
    .Acc_GetNull = "Prj-Code"
    .Acc_UserID = U_Id
    .Accounts_Hit_Save
End With
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "Daffodil Sotware Ltd"
End Sub
Private Sub Get_Employee_Status()
On Error GoTo Errdesc
If Combo1(0).Text = "First Class" Then
    EmpClassType = 1
ElseIf Combo1(0).Text = "Second Class" Then
    EmpClassType = 2
ElseIf Combo1(0).Text = "Third Class" Then
    EmpClassType = 3
Else
    EmpClassType = 4
End If
Exit Sub
Errdesc:
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
Private Sub Update_Account_TableSpace_Table1() ''' for credit entry
On Error GoTo Errdes
With Salaray_Pre
    .Connstring = strCN.Connection_String
    .Vou_No = VoucherNumber
    .AccountHitdate = Date
    .NarrationofAcc = "Salary For the Month of " & cboMonth
    .Acc_Code = cboAcc_code

    .Acc_NetPayDr = 0
    .Acc_NetPayCr = Get_Salary_Into_Text

    If optPayment_Mode(0).Value = True Then
        .Acc_PaymenyCPBD = "CP"
    Else
        .Acc_PaymenyCPBD = "BP"
    End If
    .Acc_Check_No = Text3
    .Acc_GetNull = "Prj-Code"
    .Acc_UserID = U_Id
    .Accounts_Hit_Save
End With
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "Daffodil Sotware Ltd"
End Sub
Private Sub Get_Voucher_Number()
On Error GoTo Errdesc
Dim conn10 As New Connection
Dim rs10 As New Recordset
Dim cmd As New Command
conn10.ConnectionString = strCN.Connection_String
conn10.Open
cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText
cmd.CommandText = "select max(account.vou.vou_no)+1 from account.vou"
rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic
    If rs10.RecordCount > 0 Then
        If IsNull(rs10.Fields(0)) Then
            VoucherNumber = 1
        Else
            VoucherNumber = rs10.Fields(0)
        End If
    Else
        VoucherNumber = 1
    End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Post_to_Ledger()
On Error GoTo Errdes
With Salaray_Pre
    .Connstring = strCN.Connection_String
    .Vou_No = VoucherNumber
    
    If optPayment_Mode(0).Value = True Then
        .Acc_PaymenyCPBD = "CP"
    Else
        .Acc_PaymenyCPBD = "BP"
    End If
    .Post_to_VOUCHER_Save
End With
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "Daffodil Sotware Ltd"
End Sub
Private Sub Get_Salary_That_Employee()
On Error GoTo Errdesc
Dim conn10 As New Connection
Dim rs10 As New Recordset
Dim cmd As New Command
conn10.ConnectionString = strCN.Connection_String
conn10.Open
cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText
cmd.CommandText = "select sum(net_payable) from salary_preparation where PAY_MONTH='" & Trim(cboMonth) & "' and PAY_YEAR='" & Trim(cboYear) & "' and substr(Emp_Id,1,1)='" & EmpClassType & "'"
rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic
    If rs10.RecordCount > 0 Then
        Get_Salary_Into_Text = rs10.Fields(0)
    End If
 If conn10.State = 1 Then
    conn10.Close
 End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

