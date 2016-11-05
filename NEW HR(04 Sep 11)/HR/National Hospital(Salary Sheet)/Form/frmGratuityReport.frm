VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGratuityReport 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7395
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Closing Process"
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
      Left            =   3180
      TabIndex        =   9
      Top             =   5340
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gratuity Fund Information"
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
      Height          =   5115
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   7275
      Begin VB.OptionButton Option1 
         Caption         =   "Fund Payment (Bank Wise)"
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
         Height          =   345
         Index           =   8
         Left            =   360
         TabIndex        =   25
         Top             =   3510
         Width           =   2925
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fund Receive (Bank Wise)"
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
         Height          =   345
         Index           =   7
         Left            =   360
         TabIndex        =   24
         Top             =   3090
         Width           =   2925
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fund Payment (Purpose Wise)"
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
         Height          =   315
         Index           =   6
         Left            =   360
         TabIndex        =   23
         Top             =   2670
         Width           =   2925
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fund Receive (Source Wise)"
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
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   22
         Top             =   2250
         Width           =   2925
      End
      Begin VB.Frame Frame3 
         Height          =   1065
         Left            =   3540
         TabIndex        =   19
         Top             =   2790
         Visible         =   0   'False
         Width           =   3615
         Begin VB.ComboBox cmbPurposeOfPayment 
            Height          =   315
            Left            =   1350
            TabIndex        =   21
            Top             =   600
            Width           =   2175
         End
         Begin VB.ComboBox cmbSourceofFund 
            Height          =   315
            Left            =   1350
            TabIndex        =   20
            Top             =   210
            Width           =   2175
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Payment Purpose"
            Height          =   195
            Left            =   60
            TabIndex        =   27
            Top             =   630
            Width           =   1245
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Source of Fund"
            Height          =   195
            Left            =   60
            TabIndex        =   26
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Gratuity Balance Sheet"
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
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   18
         Top             =   1920
         Width           =   2595
      End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   3510
         TabIndex        =   11
         Top             =   540
         Visible         =   0   'False
         Width           =   3615
         Begin VB.ComboBox cmbAccountNo 
            Height          =   315
            Left            =   1230
            TabIndex        =   14
            Top             =   1080
            Width           =   2175
         End
         Begin VB.ComboBox cmbBankCode 
            Height          =   315
            Left            =   1230
            TabIndex        =   13
            Top             =   360
            Width           =   2175
         End
         Begin VB.ComboBox cmbAccountType 
            Height          =   315
            Left            =   1230
            TabIndex        =   12
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Account No"
            Height          =   195
            Left            =   150
            TabIndex        =   17
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Account Type"
            Height          =   195
            Left            =   150
            TabIndex        =   16
            Top             =   750
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bank Name"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   420
            Width           =   840
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bank Account Statement"
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
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   1515
         Width           =   3075
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Gratuity Opening  Closing"
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
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   1050
         Width           =   2955
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cash Book"
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
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   735
         Width           =   2025
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Income and Expense Statement"
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
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker Begin_Fiscal_date 
         Height          =   315
         Left            =   1635
         TabIndex        =   5
         Top             =   4470
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   22675459
         CurrentDate     =   37073
      End
      Begin MSComCtl2.DTPicker End_Fiscal_Year 
         Height          =   315
         Left            =   3720
         TabIndex        =   6
         Top             =   4470
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
         Format          =   22675459
         CurrentDate     =   37072
      End
      Begin VB.Label Label3 
         Caption         =   " From                                                     To"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   4530
         Width           =   3615
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00C0C0FF&
         Height          =   675
         Left            =   150
         Shape           =   4  'Rounded Rectangle
         Top             =   4305
         Width           =   6975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   1590
         Top             =   2.00150e5
         Width           =   1680
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   1590
         Top             =   4455
         Width           =   3945
      End
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
      Left            =   4740
      Picture         =   "frmGratuityReport.frx":0000
      TabIndex        =   0
      Top             =   5325
      Width           =   1215
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
      Left            =   6000
      Picture         =   "frmGratuityReport.frx":08CA
      TabIndex        =   1
      Top             =   5325
      Width           =   1215
   End
End
Attribute VB_Name = "frmGratuityReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OGra As New clsGratuity
Private Sub Begin_Fiscal_date_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    End_Fiscal_Year.SetFocus
End If
End Sub
Private Sub Command1_Click()
If Option1(0).Value = True Then
   rptmode = 30
   'Emp_ID_Value_ForLoan = Combo1.Text
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
ElseIf Option1(1).Value = True Then
   'End_Of_Year_GratuityFund_Open_Closing
   rptmode = 31
   'Emp_ID_Value_ForLoan = Combo1.Text
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
ElseIf Option1(2).Value = True Then
   'End_Of_Year_GratuityFund_Open_Closing
   rptmode = 32
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
ElseIf Option1(3).Value = True Then
    If cmbBankCode.Text = "" Then
    MsgBox "Bank Name Required"
    ElseIf cmbAccountType.Text = "" Then
    MsgBox "Account Type Required"
    ElseIf cmbAccountNo.Text = "" Then
    MsgBox "Account No Required"
    Else
    rptmode = 33
    sBankCode = Get_Code(cmbBankCode.Text)
    sBankName = Get_Description(cmbBankCode.Text)
    sAccountType = Get_Code(cmbAccountType.Text)
    sAccountTypeName = Get_Description(cmbAccountType.Text)
    sAccountNo = Get_Code(cmbAccountNo.Text)
    BeginDateForReport = Begin_Fiscal_date
    EnddateforReport = End_Fiscal_Year
    Form20.Show vbModal
    End If
ElseIf Option1(4).Value = True Then
   rptmode = 34
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
  
ElseIf Option1(5).Value = True Then
   
   If cmbSourceofFund.Text = "" Then
   MsgBox "Source of Fund Required"
   Else
   rptmode = 35
   sSourceId = Get_Code(cmbSourceofFund.Text)
   sSourceName = Get_Description(cmbSourceofFund.Text)
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
  End If
ElseIf Option1(6).Value = True Then
   
   If cmbPurposeOfPayment.Text = "" Then
   MsgBox "Purpose of Payment Required"
   Else
   rptmode = 36
   sPurposeId = Get_Code(cmbPurposeOfPayment.Text)
   sPurposeName = Get_Description(cmbPurposeOfPayment.Text)
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
  End If
ElseIf Option1(7).Value = True Then
   If cmbBankCode.Text = "" Then
   MsgBox "Bank Name Required"
   Else
   rptmode = 37
   sBankCode = Get_Code(cmbBankCode.Text)
   sBankName = Get_Description(cmbBankCode.Text)
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
  End If
ElseIf Option1(8).Value = True Then
   If cmbBankCode.Text = "" Then
   MsgBox "Bank Name Required"
   Else
   rptmode = 38
   sBankCode = Get_Code(cmbBankCode.Text)
   sBankName = Get_Description(cmbBankCode.Text)
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
  End If


End If
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo Errdesc

End_Of_Year_GratuityFund_Open_Closing
MsgBox "Process Complete", vbInformation, "IT Division, DNMIH"

Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
    
End Sub
Private Sub End_Of_Year_GratuityFund_Open_Closing()
On Error GoTo Errdes
With OGra
    .Connstring = strCN.Connection_String
    .Save_Gratuity_Open_Close
End With
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub End_Fiscal_Year_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command1.SetFocus
End If
End Sub
Private Sub get_Value_RPT_Bank_Name()
'On Error GoTo Errdes
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
Private Sub get_Value_RPT_Account_Type()
'On Error GoTo Errdes
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
        End If

    
    
    Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub get_Value_RPT_Account_No()
'On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim Rs As New ADODB.Recordset
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select ACCOUNT_NO from GRA_CAPITAL_FUND order by TRACK_ID"
    
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



Private Sub Form_Load()
get_Value_RPT_Bank_Name
get_Value_RPT_Account_No
get_Value_RPT_Account_Type
get_Value_RPT_Source_of_Fund
get_Value_RPT_Purpose_of_Payment
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
Frame2.Visible = False
Frame3.Visible = False
Case 1
Frame2.Visible = False
Frame3.Visible = False
Case 2
Frame2.Visible = False
Frame3.Visible = False
Case 3
Frame2.Visible = True
Frame2.Height = 1575
cmbBankCode.Visible = True
cmbAccountType.Visible = True
cmbAccountNo.Visible = True
Label2.Visible = True
Label4.Visible = True
Frame3.Visible = False
Case 4
Frame2.Visible = False
Frame3.Visible = False

Case 5
Frame2.Visible = False
Frame3.Visible = True
Frame3.Height = 600
Frame3.Top = 2610
cmbSourceofFund.Visible = True
cmbPurposeOfPayment.Visible = False
cmbPurposeOfPayment.Top = 600
Label6.Top = 630
Label6.Visible = False
Label5.Visible = True
Case 6
Frame2.Visible = False
Frame3.Visible = True
Frame3.Height = 600
Frame3.Top = 2610
cmbPurposeOfPayment.Top = 200
Label6.Top = 250
cmbSourceofFund.Visible = False
cmbPurposeOfPayment.Visible = True
Label5.Visible = False
Label6.Visible = True
Case 7
cmbBankCode.Visible = True
cmbAccountType.Visible = False
cmbAccountNo.Visible = False
Label2.Visible = False
Label4.Visible = False
Frame2.Visible = True
Frame2.Height = 900
Frame3.Visible = False
Case 8
Frame2.Visible = True
cmbBankCode.Visible = True
cmbAccountType.Visible = False
cmbAccountNo.Visible = False
Label2.Visible = False
Label4.Visible = False
Frame2.Height = 900
Frame3.Visible = False
End Select

End Sub

Private Sub get_Value_RPT_Source_of_Fund()
'On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim Rs As New ADODB.Recordset
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select SOURCE_ID,SOURCE_NAME from L_SOURCE_OF_FUND order by SOURCE_ID"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    Rs.CursorLocation = adUseClient
    Rs.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
       If Rs.RecordCount > 0 Then
            Do Until Rs.EOF
            cmbSourceofFund.AddItem Rs.Fields(1) & "~" & Rs.Fields(0)
            Rs.MoveNext
            Loop
        End If
    Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub get_Value_RPT_Purpose_of_Payment()
'On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim Rs As New ADODB.Recordset
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select PURPOSE_ID,PURPOSE_NAME from L_GRA_PAYMENT_PURPOSE order by PURPOSE_ID"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    Rs.CursorLocation = adUseClient
    Rs.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
       If Rs.RecordCount > 0 Then
            Do Until Rs.EOF
            cmbPurposeOfPayment.AddItem Rs.Fields(1) & "~" & Rs.Fields(0)
            Rs.MoveNext
            Loop
        End If

    
    
    Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

