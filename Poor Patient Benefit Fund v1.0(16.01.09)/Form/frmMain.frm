VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00915411&
   Caption         =   "Poor Patient Benefit Fund 2009"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1710
      MaxLength       =   20
      TabIndex        =   14
      Top             =   6000
      Width           =   2400
   End
   Begin VB.TextBox txtUserID 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1695
      MaxLength       =   20
      TabIndex        =   0
      Top             =   5550
      Width           =   2400
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1695
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   6435
      Width           =   2400
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2190
      Top             =   7350
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   1530
      TabIndex        =   16
      Top             =   6030
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   510
      TabIndex        =   15
      Top             =   6000
      Width           =   945
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed && Maintenance by :IT Division,DNMIH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   5400
      TabIndex        =   13
      Top             =   3600
      Width           =   5085
   End
   Begin VB.Shape Shape2 
      Height          =   1755
      Left            =   0
      Top             =   2580
      Width           =   20865
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Poor Patient Benefit Fund'09"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   975
      Left            =   810
      TabIndex        =   12
      Top             =   2910
      Width           =   11025
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1755
      Left            =   240
      Top             =   5310
      Width           =   4065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   540
      TabIndex        =   11
      Top             =   6465
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   540
      TabIndex        =   10
      Top             =   5550
      Width           =   660
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   1530
      TabIndex        =   9
      Top             =   5580
      Width           =   75
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   1530
      TabIndex        =   8
      Top             =   6450
      Width           =   75
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   9630
      TabIndex        =   7
      Top             =   615
      Width           =   2265
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   9630
      TabIndex        =   6
      Top             =   1065
      Width           =   2265
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Log On Time :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   8280
      TabIndex        =   5
      Top             =   1050
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   9630
      TabIndex        =   4
      Top             =   210
      Width           =   1485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name    :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   8280
      TabIndex        =   3
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID         :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   8280
      TabIndex        =   1
      Top             =   180
      Width           =   1125
   End
   Begin VB.Menu mnuFile 
      Caption         =   " [File]"
      Enabled         =   0   'False
      Begin VB.Menu mnus1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProject_Info 
         Caption         =   "Project Information"
      End
      Begin VB.Menu mnus2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnitInfo 
         Caption         =   "Unit Information"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnus3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcct_Info 
         Caption         =   "Accounts Information"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuS4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpening_Bal 
         Caption         =   "Opening Balance"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuS5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditAccount 
         Caption         =   "Edit Account"
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnusepr123 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log off Current User"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuS6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "     [Transaction]"
      Enabled         =   0   'False
      Begin VB.Menu mnuVou_Entry 
         Caption         =   "Voucher Entry"
         Shortcut        =   ^T
      End
      Begin VB.Menu gsdfgsdfg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuchqreg 
         Caption         =   "Cheque Register"
      End
      Begin VB.Menu mnfdsfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBE 
         Caption         =   "Budget Entry"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "    [ Report]"
      Enabled         =   0   'False
      Begin VB.Menu mnuS10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptChartOfAccount 
         Caption         =   "Chart Of Accounts"
      End
      Begin VB.Menu mnus18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChartofAccount 
         Caption         =   "Chart of Account(Bengali)"
      End
      Begin VB.Menu gfdgfdg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOB 
         Caption         =   "Opening Balance"
      End
      Begin VB.Menu gfgsdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVouReg 
         Caption         =   "Voucher Register"
      End
      Begin VB.Menu mnuS11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheque 
         Caption         =   "Cheque Register"
      End
      Begin VB.Menu fsdafsdafsad 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCashbook 
         Caption         =   "Cash/Bank  Book(General)"
      End
      Begin VB.Menu trwetwertwer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCashBank 
         Caption         =   "&Cash/Bank Book(Account Specific)"
         Shortcut        =   ^B
      End
      Begin VB.Menu sepcash 
         Caption         =   "-"
      End
      Begin VB.Menu mnucbbaag 
         Caption         =   "Cash-bank book at a Glance"
      End
      Begin VB.Menu fdsfsdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptLedger 
         Caption         =   "Ledger"
      End
      Begin VB.Menu mnuS12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptSchedule 
         Caption         =   "Schedule"
      End
      Begin VB.Menu mnuS13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrialBalance 
         Caption         =   "Trial Balance"
      End
      Begin VB.Menu fdgdfgdg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRp 
         Caption         =   "Receipt and Payment"
      End
      Begin VB.Menu sep32423 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIncomeStatement 
         Caption         =   "Income/Expenditure Account Statement"
      End
      Begin VB.Menu gfdsgsdfgsdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBudgetVariance 
         Caption         =   "Budget Variance"
      End
      Begin VB.Menu fsdfsd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuF_schedule 
         Caption         =   "Fixed Asset Schedule"
      End
      Begin VB.Menu gfsdgfsd 
         Caption         =   "-"
      End
      Begin VB.Menu mnusandst 
         Caption         =   "Stock and Storage Schedule"
      End
      Begin VB.Menu sadfsdafdsa 
         Caption         =   "-"
      End
      Begin VB.Menu mnutV 
         Caption         =   "Tax/VAT schedule"
      End
      Begin VB.Menu tyet 
         Caption         =   "-"
      End
      Begin VB.Menu mnuparty 
         Caption         =   "Tax/VAT schedule(party wise)"
      End
      Begin VB.Menu mnuS16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRptFianlAccounts 
         Caption         =   "BalanceSheet"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "      [ Utility]"
      Enabled         =   0   'False
      Begin VB.Menu mnuS8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewUser 
         Caption         =   "Add New User"
      End
      Begin VB.Menu mnuS9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangPassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu GDDFGDFG 
         Caption         =   "-"
      End
      Begin VB.Menu mnuldv 
         Caption         =   "Unposted Voucher List"
         Enabled         =   0   'False
      End
      Begin VB.Menu fgdsgfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuchse 
         Caption         =   "Cheque Status Entry"
         Enabled         =   0   'False
      End
      Begin VB.Menu fdgfdsg 
         Caption         =   "-"
      End
      Begin VB.Menu mnudatabackup 
         Caption         =   "Data Backup"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "    [Help]"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuabout_Click()
 frmAbout.Show vbModal
End Sub

Private Sub mnuAcct_Info_Click()
    Form4.Show vbModal
End Sub

Private Sub mnuBE_Click()
 Form27.Show vbModal
End Sub

Private Sub mnuBudgetVariance_Click()
   Form28.Show vbModal
'    Screen.MousePointer = vbHourglass
'    rptMode = 15
'    CRViewer1.Show vbModal
End Sub

Private Sub mnuCashBank_Click()
  Form15.Show vbModal
End Sub

Private Sub mnuCashbook_Click()
 Form18.Show vbModal
End Sub

Private Sub mnucbbaag_Click()
 Form19.Show vbModal
End Sub

Private Sub mnuChangPassword_Click()
    Form8.Show vbModal
End Sub

Private Sub mnuChartofAccount_Click()
      rptMode = 11
    CRViewer1.Show vbModal
End Sub

Private Sub mnuCheque_Click()
 Form23.Show vbModal
End Sub

Private Sub mnuchqreg_Click()
  Form22.Show vbModal
End Sub

Private Sub mnuchse_Click()
 Form21.Show vbModal
End Sub

Private Sub mnudatabackup_Click()
  On Error GoTo err_desc
  Shell ("c:\WINNT\BACKUP_acc.BAT")
  Exit Sub
err_desc:
     MsgBox Err.Description, vbCritical, "IT Division,DNMIH"
End Sub

Private Sub mnuEditAccount_Click()
'    Form14.Show vbModal
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuF_schedule_Click()
   Form17.Show vbModal
End Sub

Private Sub mnuIncomeStatement_Click()
  Form13.Show vbModal
End Sub

Private Sub mnuLogOff_Click()
     Dim reply As String
    reply = MsgBox("Do you want to Log Off?", vbQuestion + vbYesNo, "Log Off...")
    If reply = vbYes Then
        mnuFile.Enabled = False
        mnuTransaction.Enabled = False
        mnuReport.Enabled = False
        mnuUtility.Enabled = False
        Frame2.Visible = True
        txtUserID.Text = ""
        txtPassword.Text = ""
        Form1.Label2(2).Caption = ""
        Form1.Label2(3).Caption = ""
        Form1.Label2(5).Caption = ""
        txtName.Text = ""
        
    End If
End Sub

Private Sub mnuNewUser_Click()
    Form7.Show vbModal
End Sub

Private Sub mnuOB_Click()
  rptMode = 100
  CRViewer1.Show vbModal
End Sub

Private Sub mnuOpening_Bal_Click()
    Form5.Show vbModal
End Sub

Private Sub mnuparty_Click()
 Form26.Show vbModal
End Sub

Private Sub mnuProject_Info_Click()
    Form2.Show vbModal
End Sub

Private Sub mnuRp_Click()
  Form16.Show vbModal
End Sub

Private Sub mnuRptChartOfAccount_Click()
    rptMode = 1
    CRViewer1.Show vbModal
End Sub

Private Sub mnuRptFianlAccounts_Click()
  Form29.Show 1
End Sub

Private Sub mnuRptLedger_Click()
    Form10.Show vbModal
End Sub

Private Sub mnuRptSchedule_Click()
  Form12.Show vbModal
End Sub

Private Sub mnusandst_Click()
  Form20.Show vbModal
End Sub

Private Sub mnuTrialBalance_Click()
    Form11.Show vbModal
End Sub

Private Sub mnutV_Click()
 Form25.Show vbModal
End Sub

Private Sub mnuUnitInfo_Click()
    Form3.Show vbModal
End Sub

Private Sub mnuVou_Entry_Click()
    Form6.Show vbModal
End Sub

Private Sub mnuVouReg_Click()
    Form9.Show vbModal
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Dim i As Integer
    If KeyAscii = 13 Then
       On Error GoTo err_desc
    If Adodc1.Recordset.RecordCount > 0 Then
        If Adodc1.Recordset!Pass_word = UCase(Trim(txtPassword.Text)) Then
            
            mnuFile.Enabled = True
            mnuTransaction.Enabled = True
            mnuReport.Enabled = True
            mnuUtility.Enabled = True
            Shape1.Visible = False
            For i = 0 To 2
             Label1(i).Visible = False
            Next i
            txtUserID.Visible = False
            txtName.Visible = False
            txtPassword.Visible = False
            For i = 0 To 2
             Label3(i).Visible = False
            Next i
            
            Form1.Label2(2) = txtUserID
            Form1.Label2(3).Caption = txtName
            Form1.Label2(5).Caption = Time
        Else
            MsgBox "Invalid Password", vbOKOnly + vbCritical, "Warning..."
            txtPassword = ""
            txtPassword.SetFocus
            Exit Sub
        End If
    
    End If
    Exit Sub
err_desc:
        MsgBox Err.Description, vbOKOnly + vbInformation, "IT Division,DNMIH."

    End If
    
End Sub
Private Sub txtUserID_Change()
    txtUserID.Text = UCase(txtUserID.Text)
End Sub



Private Sub txtUserID_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub txtUserID_LostFocus()
On Error GoTo err_desc
    If Len(Trim(txtUserID.Text)) = 0 Then Exit Sub
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select user_id,user_name,pass_word from security where user_id='" & UCase(Trim(txtUserID.Text)) & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        If Adodc1.Recordset!user_id = UCase(Trim(txtUserID.Text)) Then
              Form1.txtName = "" & Adodc1.Recordset!user_name
             
            Exit Sub
        Else
            MsgBox "Invalid User ID", vbOKOnly + vbCritical, "Warning..."
            txtUserID.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Invalid User ID", vbOKOnly + vbCritical, "Warning..."
        txtUserID.SetFocus
        Exit Sub
    End If
    
   Exit Sub
err_desc:
   MsgBox Err.Description, vbOKOnly + vbCritical, "IT Division,DNMIH"
End Sub
