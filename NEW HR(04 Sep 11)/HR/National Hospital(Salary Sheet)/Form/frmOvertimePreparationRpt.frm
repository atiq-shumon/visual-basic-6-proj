VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOvertimePreparationRpt 
   BackColor       =   &H80000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6780
   ClipControls    =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Report On Overtime"
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
      Height          =   1635
      Left            =   75
      TabIndex        =   4
      Top             =   90
      Width           =   6585
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00FF8080&
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   390
         Width           =   1590
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         Height          =   335
         Left            =   1530
         TabIndex        =   5
         Top             =   750
         Width           =   4785
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   330
         Left            =   1530
         TabIndex        =   2
         Top             =   1110
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date of Payment                   "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   1170
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "Emp ID"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   225
         TabIndex        =   7
         Top             =   420
         Width           =   645
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   375
         Left            =   1500
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label3 
         Caption         =   "Emp Name"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   225
         TabIndex        =   6
         Top             =   795
         Width           =   885
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
      Left            =   4140
      Picture         =   "frmOvertimePreparationRpt.frx":0000
      TabIndex        =   3
      Top             =   1845
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
      Left            =   5400
      Picture         =   "frmOvertimePreparationRpt.frx":08CA
      TabIndex        =   1
      Top             =   1845
      Width           =   1215
   End
End
Attribute VB_Name = "frmOvertimePreparationRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BeginDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Get_Emp_Name
    MaskEdBox1.SetFocus
End If
End Sub

Private Sub Command1_Click()
   rptmode = 7
   Emp_ID_Value = Trim(Combo1.Text)
   EnddateforReport = MaskEdBox1
   Form20.Show vbModal




End Sub

Private Sub Command2_Click()
Unload Me
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

Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Command1.SetFocus
End If
End Sub
