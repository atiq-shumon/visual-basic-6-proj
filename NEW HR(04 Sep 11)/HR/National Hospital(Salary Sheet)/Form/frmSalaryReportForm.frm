VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalaryReportForm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6465
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Salary Type"
      ForeColor       =   &H008080FF&
      Height          =   615
      Left            =   150
      TabIndex        =   17
      Top             =   4170
      Width           =   6465
      Begin VB.OptionButton opnSalaryType 
         Appearance      =   0  'Flat
         Caption         =   "Dress Allowance Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3240
         TabIndex        =   9
         Top             =   270
         Width           =   2655
      End
      Begin VB.OptionButton opnSalaryType 
         Appearance      =   0  'Flat
         Caption         =   "Only Bonus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   10
         Top             =   270
         Width           =   1335
      End
      Begin VB.OptionButton opnSalaryType 
         Appearance      =   0  'Flat
         Caption         =   "Reg."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton opnSalaryType 
         Appearance      =   0  'Flat
         Caption         =   "Supp."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   18
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
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
      Height          =   425
      Left            =   3885
      Picture         =   "frmSalaryReportForm.frx":0000
      TabIndex        =   7
      Top             =   4860
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Left            =   5145
      Picture         =   "frmSalaryReportForm.frx":08CA
      TabIndex        =   6
      Top             =   4860
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Salary Preparation Report"
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
      Height          =   4050
      Left            =   135
      TabIndex        =   2
      Top             =   135
      Width           =   7020
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "Employee Specific Salary Preparation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   2
         Left            =   420
         TabIndex        =   21
         Top             =   1320
         Width           =   6615
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "Department wise Salary Preparation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   1
         Left            =   420
         TabIndex        =   20
         Top             =   2400
         Width           =   6615
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "Employee's Class wise Salary Preparation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   420
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   6615
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   330
         Index           =   0
         Left            =   1680
         TabIndex        =   14
         Top             =   1815
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmSalaryReportForm.frx":1194
         Left            =   1665
         List            =   "frmSalaryReportForm.frx":11A7
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   825
         Width           =   3975
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSalaryReportForm.frx":11EC
         Left            =   1665
         List            =   "frmSalaryReportForm.frx":11EE
         TabIndex        =   8
         Top             =   2880
         Width           =   3975
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
         Height          =   315
         Left            =   4545
         TabIndex        =   1
         Top             =   3450
         Width           =   1095
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
         Height          =   315
         Left            =   1665
         TabIndex        =   0
         Top             =   3450
         Width           =   2220
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   330
         Index           =   1
         Left            =   3720
         TabIndex        =   15
         Top             =   1815
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "From                                                         To"
         Height          =   375
         Left            =   450
         TabIndex        =   16
         Top             =   1845
         Width           =   3135
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FFC0C0&
         Height          =   390
         Left            =   1620
         Top             =   780
         Width           =   4065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Employee Class"
         Height          =   195
         Left            =   450
         TabIndex        =   11
         Top             =   870
         Width           =   1110
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFC0C0&
         Height          =   390
         Left            =   4500
         Top             =   3405
         Width           =   1185
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         Height          =   360
         Left            =   1620
         Top             =   2850
         Width           =   4065
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   390
         Left            =   1620
         Top             =   3405
         Width           =   2310
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Department"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   450
         TabIndex        =   5
         Top             =   2925
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Month"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   450
         TabIndex        =   4
         Top             =   3495
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   3
         Top             =   3495
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmSalaryReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1(1).Value = True Then '''dept wise
    rptmode = 9
    GetMonthOftheYear = cboMonth
    GetSalaryPreparationYaer = cboYear
    ComboValue_Dept = Trim(Combo1.Text)
    
    If Combo4.Text = "First Class" Then
        CheckStatusofEmployee = 1
    ElseIf Combo4.Text = "Second Class" Then
        CheckStatusofEmployee = 2
    ElseIf Combo4.Text = "Third Class" Then
        CheckStatusofEmployee = 3
    ElseIf Combo4.Text = "Fourth Class" Then
        CheckStatusofEmployee = 4
    Else
       CheckStatusofEmployee = 5 '''ALL CLASS
    End If
    
    Form20.Show vbModal
ElseIf Option1(0).Value = True Then '' class wise
    GetMonthOftheYear = cboMonth
    GetSalaryPreparationYaer = cboYear
    If Combo4.Text = "First Class" Then
        ComboValue_Dept = 1
    ElseIf Combo4.Text = "Second Class" Then
        ComboValue_Dept = 2
    ElseIf Combo4.Text = "Third Class" Then
        ComboValue_Dept = 3
   ElseIf Combo4.Text = "Fourth Class" Then
        ComboValue_Dept = 4
    Else
       ComboValue_Dept = 5 '''ALL CLASS
    End If
    
      rptmode = 18
    
Else
    rptmode = 25
    BeginDateForReport = MaskEdBox1(0)
    EnddateforReport = MaskEdBox1(1)
    GetMonthOftheYear = cboMonth
    GetSalaryPreparationYaer = cboYear
End If

Form20.Show vbModal
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()

Load_Yr Me
Load_MonthNm Me

localSalaryType = "R"

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
            Combo1.AddItem rs4.Fields(0)
            rs4.MoveNext
        Loop
  End If

    rs4.Close
    conn4.Close

End Sub

Private Sub MaskEdBox1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Select Case Index
Case 0
    MaskEdBox1(1).SetFocus
Case 1
    cboMonth.SetFocus
End Select
End If
End Sub

Private Sub opnSalaryType_Click(Index As Integer)
  Select Case Index
         Case 0
            localSalaryType = "R"
         Case 1
            localSalaryType = "S"
         Case 2
            localSalaryType = "B"
         Case 3
            localSalaryType = "D"
  End Select
  
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index

Case 0
        If Option1(0).Value = True Then
            Option1(0).ForeColor = &HFF0000
            Option1(1).ForeColor = &H80000012
            Combo4.Enabled = True
            Combo1.Enabled = False
            Label6.Enabled = True
            Label4.Enabled = False
            Me.MaskEdBox1(0).Enabled = False
            Me.MaskEdBox1(0).Enabled = False
            cboMonth.SetFocus
        Else
            Option1(1).ForeColor = &H80000012
            Combo4.Enabled = False
            Combo1.Enabled = True
            Label6.Enabled = False
            Label4.Enabled = True
        End If
    

Case 1
        If Option1(1).Value = True Then
            Option1(1).ForeColor = &HFF0000
            Option1(0).ForeColor = &H80000012
            Combo1.Enabled = True
            Combo4.Enabled = False
            Label6.Enabled = False
            Label4.Enabled = True
            Me.MaskEdBox1(0).Enabled = False
            Me.MaskEdBox1(0).Enabled = False
            cboMonth.SetFocus
        Else
            Option1(0).ForeColor = &H80000012
            Combo1.Enabled = False
            Combo4.Enabled = True
            Label6.Enabled = True
            Label4.Enabled = False
        End If



Case 2
        If Option1(2).Value = True Then
            Option1(2).ForeColor = &HFF0000
            Option1(2).ForeColor = &H80000012
            'Combo2.Enabled = True
            Combo4.Enabled = False
            Combo1.Enabled = False
            Label6.Enabled = False
            Label4.Enabled = True
            Me.MaskEdBox1(0).Enabled = True
            Me.MaskEdBox1(0).Enabled = True
            
            Me.MaskEdBox1(0).SetFocus
        Else
            Option1(0).ForeColor = &H80000012
            Combo1.Enabled = False
            Combo4.Enabled = True
            Label6.Enabled = True
            Label4.Enabled = False
            Me.MaskEdBox1(0).Enabled = False
            Me.MaskEdBox1(0).Enabled = False
            
        End If




End Select
End Sub
