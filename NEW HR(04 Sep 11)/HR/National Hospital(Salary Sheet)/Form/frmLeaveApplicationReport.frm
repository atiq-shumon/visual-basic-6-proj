VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLeaveApplicationReport 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8265
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   8265
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
      Left            =   6930
      Picture         =   "frmLeaveApplicationReport.frx":0000
      TabIndex        =   6
      Top             =   1515
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
      Left            =   5670
      Picture         =   "frmLeaveApplicationReport.frx":08CA
      TabIndex        =   5
      Top             =   1515
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report On Leave Application"
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
      Height          =   1320
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   8070
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         Height          =   335
         Left            =   2715
         TabIndex        =   8
         Top             =   385
         Width           =   5280
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00FF8080&
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   390
         Width           =   1590
      End
      Begin MSComCtl2.DTPicker BeginDate 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   750
         Width           =   1635
         _ExtentX        =   2884
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
         Format          =   59047939
         CurrentDate     =   36998
      End
      Begin MSComCtl2.DTPicker EndDate 
         Height          =   315
         Left            =   3150
         TabIndex        =   7
         Top             =   750
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
         Format          =   59047939
         CurrentDate     =   36998
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   375
         Left            =   1050
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "From Date                                            To"
         DragMode        =   1  'Automatic
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   825
         Width           =   6180
      End
      Begin VB.Label Label1 
         Caption         =   "Emp ID"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   225
         TabIndex        =   2
         Top             =   480
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmLeaveApplicationReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Change()
Get_Emp_Name
End Sub

Private Sub Combo1_LostFocus()
Get_Emp_Name
End Sub

Private Sub Command1_Click()
rptmode = 13
Emp_IDforLeave = Trim(Combo1.Text)
EnddateforReport = BeginDate
BeginDateForReport = EndDate
Form20.Show vbModal

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
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
BeginDate = Format(Date$, "dd-mm-yyyy")
EndDate = Format(Date$, "dd-mm-yyyy")
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


