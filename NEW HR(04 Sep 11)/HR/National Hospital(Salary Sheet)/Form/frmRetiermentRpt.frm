VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Retired Employee's Report"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7680
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7680
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
      Left            =   5040
      Picture         =   "frmRetiermentRpt.frx":0000
      TabIndex        =   11
      Top             =   3240
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
      Left            =   6360
      Picture         =   "frmRetiermentRpt.frx":08CA
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Report Type "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7425
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   4020
         TabIndex        =   9
         Top             =   450
         Width           =   3165
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   1
         ItemData        =   "frmRetiermentRpt.frx":1194
         Left            =   4020
         List            =   "frmRetiermentRpt.frx":1196
         TabIndex        =   8
         Top             =   2340
         Width           =   3165
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   2
         ItemData        =   "frmRetiermentRpt.frx":1198
         Left            =   4020
         List            =   "frmRetiermentRpt.frx":11A8
         TabIndex        =   7
         Top             =   1440
         Width           =   3165
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   3
         ItemData        =   "frmRetiermentRpt.frx":11DC
         Left            =   4020
         List            =   "frmRetiermentRpt.frx":11E6
         TabIndex        =   6
         Top             =   1920
         Width           =   3165
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Department wise Retired Employee's Report"
         ForeColor       =   &H000000FF&
         Height          =   465
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   375
         Value           =   -1  'True
         Width           =   3555
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Class wise Retired Employee's Report"
         ForeColor       =   &H000000FF&
         Height          =   465
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1365
         Width           =   3555
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Designation wise Retired Employee's Report"
         ForeColor       =   &H000000FF&
         Height          =   465
         Index           =   4
         Left            =   240
         TabIndex        =   3
         Top             =   2265
         Width           =   3555
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sex wise Retiretment Report"
         ForeColor       =   &H000000FF&
         Height          =   465
         Index           =   3
         Left            =   240
         TabIndex        =   2
         Top             =   1845
         Width           =   3555
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Retired Employee's Report"
         ForeColor       =   &H000000FF&
         Height          =   465
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   3555
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Errdes
'On Error GoTo Errdes
rptmode = 24

If Option1(0).Value = True Then
    
    ReportStatusofEmployee = 0
    DEPARMENTNAMEFORTPT = Trim(Combo1(0))
    SEXFORREPORT = 5
    DESIGNATIONFORRPT = ""
    CheckStatusofEmployee = 5
ElseIf Option1(1).Value = True Then
    
    ReportStatusofEmployee = 1
    DEPARMENTNAMEFORTPT = ""
    SEXFORREPORT = 5
    DESIGNATIONFORRPT = ""
    CheckStatusofEmployee = 5
    
ElseIf Option1(2).Value = True Then
    ReportStatusofEmployee = 2
    SEXFORREPORT = 5
    DEPARMENTNAMEFORTPT = ""
    DESIGNATIONFORRPT = ""
    
    If Combo1(2).Text = "First Class" Then
        CheckStatusofEmployee = 1
    ElseIf Combo1(2).Text = "2nd Class" Then
        CheckStatusofEmployee = 2
    ElseIf Combo1(2).Text = "3rd Class" Then
        CheckStatusofEmployee = 3
    Else
        CheckStatusofEmployee = 4
    End If

ElseIf Option1(3).Value = True Then
    ReportStatusofEmployee = 3
    DEPARMENTNAMEFORTPT = ""
    DESIGNATIONFORRPT = ""
    CheckStatusofEmployee = 5
    If Combo1(3).Text = "Male" Then
        SEXFORREPORT = 1
    Else
        SEXFORREPORT = 0
    End If

ElseIf Option1(4).Value = True Then
    ReportStatusofEmployee = 4
    DEPARMENTNAMEFORTPT = ""
    DESIGNATIONFORRPT = Trim(Combo1(1))
    CheckStatusofEmployee = 5
    SEXFORREPORT = 5
End If
Form20.Show vbModal
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
get_Designation
get_Department
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

cmd.CommandText = "select DEPT_NM from ST_DEPT order by DEPT_NM "
rs4.CursorLocation = adUseClient
rs4.Open cmd.CommandText, conn4, adOpenDynamic, adLockOptimistic

    If rs4.RecordCount > 0 Then

        Do Until rs4.EOF
            Combo1(0).AddItem rs4.Fields(0)
            rs4.MoveNext
        Loop
  End If

    rs4.Close
    conn4.Close
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub get_Designation()
On Error GoTo Errdes
Dim cmd5 As New Command
Dim conn5 As New Connection
Dim rs5 As New Recordset

conn5.ConnectionString = strCN.Connection_String
conn5.Open
cmd5.ActiveConnection = conn5
cmd5.CommandType = adCmdText

cmd5.CommandText = "select DESIGNATION from ST_DESIG order by DESIGNATION asc"
rs5.CursorLocation = adUseClient
rs5.Open cmd5.CommandText, conn5, adOpenDynamic, adLockOptimistic

    If rs5.RecordCount > 0 Then

        Do Until rs5.EOF
            Combo1(1).AddItem rs5.Fields(0)
            rs5.MoveNext
        Loop
  End If

    rs5.Close
    conn5.Close
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Option1_Click(Index As Integer)
On Error GoTo Errdes
Select Case Index
Case 0
    For i = 1 To 3
        Combo1(i) = ""
    Next
    Combo1(0).SetFocus
    
Case 1
     For i = 0 To 3
        Combo1(i) = ""
    Next

Case 2
    
    Combo1(1) = ""
    Combo1(3) = ""
    Combo1(0) = ""
    Combo1(2).SetFocus

Case 3
    Combo1(1) = ""
    Combo1(2) = ""
    Combo1(0) = ""
    Combo1(3).SetFocus
Case 4
    
    Combo1(2) = ""
    Combo1(0) = ""
    Combo1(3) = ""
    Combo1(1).SetFocus
End Select
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub
