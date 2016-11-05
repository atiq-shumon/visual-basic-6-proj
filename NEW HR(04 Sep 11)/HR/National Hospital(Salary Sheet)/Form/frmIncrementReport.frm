VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIncrementReport 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8220
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5565
      Picture         =   "frmIncrementReport.frx":0000
      TabIndex        =   5
      ToolTipText     =   "View Report"
      Top             =   2790
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
      Left            =   6840
      Picture         =   "frmIncrementReport.frx":08CA
      TabIndex        =   8
      ToolTipText     =   "Exit"
      Top             =   2790
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Increment Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   2610
      Left            =   90
      TabIndex        =   6
      Top             =   60
      Width           =   8010
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         ItemData        =   "frmIncrementReport.frx":1194
         Left            =   4440
         List            =   "frmIncrementReport.frx":1196
         TabIndex        =   11
         Top             =   825
         Width           =   3165
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "frmIncrementReport.frx":1198
         Left            =   4440
         List            =   "frmIncrementReport.frx":11A8
         TabIndex        =   10
         Top             =   435
         Width           =   3165
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         ItemData        =   "frmIncrementReport.frx":11DC
         Left            =   4440
         List            =   "frmIncrementReport.frx":11DE
         TabIndex        =   9
         Top             =   1200
         Width           =   3165
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Department wise Increment Report"
         Height          =   255
         Index           =   2
         Left            =   255
         TabIndex        =   2
         Top             =   1230
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Employee Specific Increment Report"
         Height          =   255
         Index           =   1
         Left            =   255
         TabIndex        =   1
         Top             =   855
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Employee Class Wise Increment Report"
         Height          =   255
         Index           =   0
         Left            =   255
         TabIndex        =   0
         Top             =   465
         Value           =   -1  'True
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker BeginDateofIncr 
         Height          =   315
         Left            =   2805
         TabIndex        =   3
         Top             =   1950
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
         Format          =   61538307
         CurrentDate     =   36998
      End
      Begin MSComCtl2.DTPicker EndDateOfIncr 
         Height          =   315
         Left            =   5550
         TabIndex        =   4
         Top             =   1905
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
         Format          =   61538307
         CurrentDate     =   36998
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   315
         TabIndex        =   12
         Top             =   1665
         Width           =   1035
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   615
         Left            =   195
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Width           =   7620
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   390
         Left            =   5490
         Top             =   1875
         Width           =   2100
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FFC0C0&
         Height          =   390
         Left            =   2760
         Top             =   1920
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "From                                                        To"
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   2205
         TabIndex        =   7
         Top             =   1980
         Width           =   4560
      End
   End
End
Attribute VB_Name = "frmIncrementReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Option1(0).Value = True Then
    BeginDateOfIncremnt = BeginDateofIncr
    EndDateOfIncremnt = EndDateOfIncr
       
    If Combo1(0).Text = "First Class" Then
        CheckStatusofEmployee = 1
    ElseIf Combo1(0).Text = "2nd Class" Then
        CheckStatusofEmployee = 2
    ElseIf Combo1(0).Text = "3rd Class" Then
        CheckStatusofEmployee = 3
    Else
        CheckStatusofEmployee = 4
    End If
    Emp_ID_Value = ""
    


ElseIf Option1(1).Value = True Then
    CheckStatusofEmployee = 0
    BeginDateOfIncremnt = BeginDateofIncr
    EndDateOfIncremnt = EndDateOfIncr
    DEPARMENTNAMEFORTPT = ""
    Emp_ID_Value = Trim(Combo1(2).Text)


ElseIf Option1(2).Value = True Then
    BeginDateOfIncremnt = BeginDateofIncr
    EndDateOfIncremnt = EndDateOfIncr
    CheckStatusofEmployee = 0
    DEPARMENTNAMEFORTPT = Trim(Combo1(1))
    Emp_ID_Value = ""
   
End If

rptmode = 14
Form20.Show vbModal











'On Error GoTo Errdes
'On Error GoTo Errdes
'rptmode = 23

'
'If Option1(0).Value = True Then
'    'ReportStatusofEmployee = ""
'    ReportStatusofEmployee = 0
'    DEPARMENTNAMEFORTPT = Trim(Combo1(0))
'    SEXFORREPORT = 5
'    DESIGNATIONFORRPT = ""
'    CheckStatusofEmployee = 5
'ElseIf Option1(1).Value = True Then
'    'ReportStatusofEmployee = ""
'    ReportStatusofEmployee = 1
'    DEPARMENTNAMEFORTPT = ""
'    SEXFORREPORT = 5
'    DESIGNATIONFORRPT = ""
'    CheckStatusofEmployee = 5
'
'ElseIf Option1(2).Value = True Then
'    'ReportStatusofEmployee = ""
'    ReportStatusofEmployee = 2
'    SEXFORREPORT = 5
'    DEPARMENTNAMEFORTPT = ""
'    DESIGNATIONFORRPT = ""
'
'    If Combo1(2).Text = "First Class" Then
'        CheckStatusofEmployee = 1
'    ElseIf Combo1(2).Text = "2nd Class" Then
'        CheckStatusofEmployee = 2
'    ElseIf Combo1(2).Text = "3rd Class" Then
'        CheckStatusofEmployee = 3
'    Else
'        CheckStatusofEmployee = 4
'    End If
'
'ElseIf Option1(3).Value = True Then
'    'ReportStatusofEmployee = ""
'    ReportStatusofEmployee = 3
'    DEPARMENTNAMEFORTPT = ""
'    DESIGNATIONFORRPT = ""
'    CheckStatusofEmployee = 5
'    If Combo1(3).Text = "Male" Then
'        SEXFORREPORT = 1
'    Else
'        SEXFORREPORT = 0
'    End If
'
'ElseIf Option1(4).Value = True Then
'    'ReportStatusofEmployee = ""
'    ReportStatusofEmployee = 4
'    DEPARMENTNAMEFORTPT = ""
'    DESIGNATIONFORRPT = Trim(Combo1(1))
'    CheckStatusofEmployee = 5
'    SEXFORREPORT = 5
'End If
'Form20.Show vbModal
'Exit Sub
'Errdes:
' MsgBox Err.Description, vbInformation, "IT Division, DNMIH"









End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
get_Department
Get_All_Employee
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
        Combo1(0).Enabled = True
        Combo1(0).SetFocus
        Combo1(1).Text = ""
        Combo1(1).Enabled = False
        Combo1(2).Text = ""
        Combo1(2).Enabled = False
Case 1
        Combo1(0).Enabled = False
        Combo1(0).Text = ""
        Combo1(2).Enabled = True
        Combo1(2).SetFocus
        Combo1(1).Text = ""
        Combo1(1).Enabled = False
Case 2
        Combo1(0).Enabled = False
        Combo1(0).Text = ""
        Combo1(2).Text = ""
        Combo1(2).Enabled = False
        Combo1(1).Enabled = True
        Combo1(1).SetFocus
        
End Select
End Sub
Private Sub get_Department()
On Error GoTo Errdes
Dim cmd As New Command
Dim conn4 As New Connection
Dim rs4 As New Recordset

conn4.ConnectionString = strCN.Connection_String
conn4.Open
cmd.ActiveConnection = conn4
cmd.CommandType = adCmdText

cmd.CommandText = "select DEPT_NM from ST_DEPT"
rs4.CursorLocation = adUseClient
rs4.Open cmd.CommandText, conn4, adOpenDynamic, adLockOptimistic

    If rs4.RecordCount > 0 Then

        Do Until rs4.EOF
            Combo1(1).AddItem rs4.Fields(0)
            rs4.MoveNext
        Loop
  End If

    rs4.Close
    conn4.Close
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Get_All_Employee()
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
        Combo1(2).AddItem rs10.Fields(0)
        rs10.MoveNext
    Loop
  
End If

rs10.Close
conn10.Close
End Sub
