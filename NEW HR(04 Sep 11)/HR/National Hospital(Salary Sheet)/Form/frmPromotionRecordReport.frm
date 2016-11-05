VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPromotionRecordReport 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6510
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6510
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
      Left            =   5205
      Picture         =   "frmPromotionRecordReport.frx":0000
      TabIndex        =   9
      Top             =   2310
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
      Left            =   3975
      Picture         =   "frmPromotionRecordReport.frx":08CA
      TabIndex        =   8
      Top             =   2310
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Promotion Record of Employee"
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
      Height          =   2130
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   6375
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         ForeColor       =   &H008080FF&
         Height          =   330
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   4545
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1575
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker Begin_Fiscal_date 
         Height          =   315
         Left            =   1575
         TabIndex        =   3
         Top             =   1350
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
         Format          =   58523651
         CurrentDate     =   36998
      End
      Begin MSComCtl2.DTPicker End_Fiscal_Year 
         Height          =   315
         Left            =   3660
         TabIndex        =   4
         Top             =   1365
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
         CurrentDate     =   36998
      End
      Begin VB.Label Label1 
         Caption         =   "Employee ID"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   300
         TabIndex        =   7
         Top             =   645
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   300
         TabIndex        =   6
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   1590
         Top             =   2.00150e5
         Width           =   1680
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   1530
         Top             =   735
         Width           =   1785
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00C0C0FF&
         Height          =   1635
         Left            =   90
         Shape           =   4  'Rounded Rectangle
         Top             =   330
         Width           =   6165
      End
      Begin VB.Label Label3 
         Caption         =   " From                                                       To"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   255
         TabIndex        =   5
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Left            =   1530
         Top             =   1350
         Width           =   3945
      End
   End
End
Attribute VB_Name = "frmPromotionRecordReport"
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
    Get_Emp_Name
    Begin_Fiscal_date.SetFocus
End If
End Sub

Private Sub Command1_Click()
   rptmode = 22
   BeginDateOfIncremnt = Begin_Fiscal_date
   EndDateOfIncremnt = End_Fiscal_Year
   Emp_ID_Value = Trim(Combo1.Text)
   Form20.Show vbModal

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
    
End If
rs10.Close
conn10.Close

Begin_Fiscal_date = Format(Date, "dd/mm/yyyy")
End_Fiscal_Year = Format(Date, "dd/mm/yyyy")
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

