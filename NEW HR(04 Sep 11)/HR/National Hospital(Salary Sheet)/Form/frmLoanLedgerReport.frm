VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLoanLedgerReport 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7395
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Loan Ledger Information"
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
      Height          =   3495
      Left            =   45
      TabIndex        =   3
      Top             =   75
      Width           =   7275
      Begin VB.OptionButton Option1 
         Caption         =   "Loan (Against PF) Refund Statment (Yearly-All Employee)"
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
         Index           =   2
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   5295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Loan (Against PF) Taken Statment"
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
         TabIndex        =   8
         Top             =   720
         Width           =   3765
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Loan (Against PF) Refund Statment (Individual)"
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
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   4785
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1635
         TabIndex        =   0
         Top             =   2100
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H008080FF&
         Height          =   330
         Left            =   1630
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2460
         Width           =   5265
      End
      Begin MSComCtl2.DTPicker Begin_Fiscal_date 
         Height          =   315
         Left            =   1635
         TabIndex        =   9
         Top             =   2850
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
         Format          =   58720259
         CurrentDate     =   36998
      End
      Begin MSComCtl2.DTPicker End_Fiscal_Year 
         Height          =   315
         Left            =   3720
         TabIndex        =   10
         Top             =   2850
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
         Format          =   58720259
         CurrentDate     =   36998
      End
      Begin VB.Label Label3 
         Caption         =   " From                                                     To"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   2940
         Width           =   3615
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00C0C0FF&
         Height          =   1425
         Left            =   150
         Shape           =   4  'Rounded Rectangle
         Top             =   1965
         Width           =   6975
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   1590
         Top             =   2055
         Width           =   1785
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   1590
         Top             =   2.00150e5
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "Emp  Name"
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   405
         TabIndex        =   6
         Top             =   2535
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   1590
         Top             =   2445
         Width           =   3945
      End
      Begin VB.Label Label1 
         Caption         =   " Emp ID"
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   360
         TabIndex        =   5
         Top             =   2190
         Width           =   1185
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
      Picture         =   "frmLoanLedgerReport.frx":0000
      TabIndex        =   1
      Top             =   3705
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
      Left            =   6060
      Picture         =   "frmLoanLedgerReport.frx":08CA
      TabIndex        =   2
      Top             =   3705
      Width           =   1215
   End
End
Attribute VB_Name = "frmLoanLedgerReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Begin_Fiscal_date_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    End_Fiscal_Year.SetFocus
End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    
'=========== change by Zahid =====
    If Combo1.Text = "" Then
        Begin_Fiscal_date.Enabled = True
        End_Fiscal_Year.Enabled = True
        Else
        Begin_Fiscal_date.Enabled = False
        End_Fiscal_Year.Enabled = False
    End If
'=========== end of change=======
    Get_Emp_Name
        If Begin_Fiscal_date.Enabled = True Then
            Begin_Fiscal_date.SetFocus
        Else
            Command1.SetFocus
        End If
End If
End Sub

Private Sub Combo1_LostFocus()
Get_Emp_Name
'Command1.SetFocus
End Sub
Private Sub Command1_Click()
If Option1(0).Value = True Then
   rptmode = 17
   Emp_ID_Value_ForLoan = Combo1.Text
   Form20.Show vbModal
ElseIf Option1(1).Value = True Then
   rptmode = 20
   Emp_ID_Value_ForLoan = Combo1.Text
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal

Else
   rptmode = 21
   BeginDateForReport = Begin_Fiscal_date
   EnddateforReport = End_Fiscal_Year
   Form20.Show vbModal
End If
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub End_Fiscal_Year_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command1.SetFocus
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
        Combo1.AddItem rs10.Fields(0)
        rs10.MoveNext
    Loop
    Combo1 = Combo1.List(0)
End If
rs10.Close
conn10.Close
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

cmd.CommandText = "select EMP_NM from emp_info where Emp_Id='" & Combo1 & "'"

RS2.CursorLocation = adUseClient
RS2.Open cmd.CommandText, conn2, adOpenDynamic, adLockOptimistic

    If RS2.RecordCount > 0 Then
      Text1.Text = RS2.Fields(0)
        RS2.Close
        conn2.Close
    Else
        Text1.Text = ""
        RS2.Close
        conn2.Close
    End If
Exit Sub
Errdesc:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Option1_Click(Index As Integer)
Select Case Index

Case 0

    If Option1(0).Value = True Then
        Option1(0).ForeColor = &HFF8080
        Option1(1).ForeColor = &H404040
        Begin_Fiscal_date.Enabled = True
        End_Fiscal_Year.Enabled = True
        Combo1.Enabled = True
        Option1(2).ForeColor = &H404040
        Combo1.SetFocus
    Else
        Option1(0).ForeColor = &H404040
        Option1(1).ForeColor = &HFF8080
        Begin_Fiscal_date.Enabled = False
        End_Fiscal_Year.Enabled = False
    
    End If
    
Case 1
    If Option1(1).Value = True Then
       
        Option1(1).ForeColor = &HFF8080
        Option1(0).ForeColor = &H404040
        Combo1.Enabled = True
        Text1.Enabled = True
        Combo1.SetFocus
        Begin_Fiscal_date.Enabled = False
        End_Fiscal_Year.Enabled = False
        Option1(2).ForeColor = &H404040
    
    Else
        Option1(1).ForeColor = &H404040
        Option1(0).ForeColor = &HFF8080
        Begin_Fiscal_date.Enabled = True
        End_Fiscal_Year.Enabled = True
        Combo1.Enabled = False
        Text1.Enabled = False
          
    
    End If

Case 2
     If Option1(2).Value = True Then
        Option1(1).ForeColor = &H404040
        Option1(0).ForeColor = &H404040
        Option1(2).ForeColor = &HFF8080
        Begin_Fiscal_date.Enabled = True
        End_Fiscal_Year.Enabled = True
        Begin_Fiscal_date.SetFocus
        Combo1.Enabled = False
        Text1.Enabled = False
     Else
        Begin_Fiscal_date.Enabled = False
        End_Fiscal_Year.Enabled = False
        Option1(2).ForeColor = &H404040
        Combo1.Enabled = True
        Text1.Enabled = True
       
    End If
End Select
End Sub
