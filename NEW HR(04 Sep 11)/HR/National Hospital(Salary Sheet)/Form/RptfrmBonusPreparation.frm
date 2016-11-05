VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RptfrmBonusPreparation 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report Bonus Preparation"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6150
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Bonus Preparation "
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
      Height          =   3390
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5910
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
         TabIndex        =   6
         Top             =   2535
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
         Height          =   315
         Left            =   4560
         TabIndex        =   5
         Top             =   2535
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "RptfrmBonusPreparation.frx":0000
         Left            =   1665
         List            =   "RptfrmBonusPreparation.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   825
         Width           =   3975
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   330
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Top             =   1770
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   330
         Index           =   1
         Left            =   3720
         TabIndex        =   7
         Top             =   1770
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
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         Height          =   855
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   2310
         Width           =   5655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   9
         Top             =   2580
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Month"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   405
         TabIndex        =   12
         Top             =   2580
         Width           =   450
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   390
         Left            =   1620
         Top             =   2490
         Width           =   2310
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFC0C0&
         Height          =   390
         Left            =   4500
         Top             =   2490
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Employee Class"
         Height          =   195
         Left            =   405
         TabIndex        =   10
         Top             =   870
         Width           =   1110
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FFC0C0&
         Height          =   390
         Left            =   1620
         Top             =   780
         Width           =   4065
      End
      Begin VB.Label Label5 
         Caption         =   "From                                                  To"
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   1845
         Width           =   3135
      End
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
      Left            =   4800
      Picture         =   "RptfrmBonusPreparation.frx":004A
      TabIndex        =   1
      Top             =   3615
      Width           =   1215
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
      Left            =   3480
      Picture         =   "RptfrmBonusPreparation.frx":0914
      TabIndex        =   0
      Top             =   3615
      Width           =   1215
   End
End
Attribute VB_Name = "RptfrmBonusPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1(0).Value = True Then

    If Len(Trim(Combo4.Text)) = 0 Or Len(Trim(cboMonth)) = 0 Or Len(Trim(cboYear)) = 0 Then
      MsgBox "Proper Value has not given!", vbInformation, "IT Division, DNMIH"
      Combo4.SetFocus
      Exit Sub
    Else
      rptmode = 26
      GetMonthOftheYear = cboMonth
      GetSalaryPreparationYaer = cboYear
      
      If Combo4.Text = "First Class" Then
          StatusofEmployee = 1
      ElseIf Combo4.Text = "Second Class" Then
          StatusofEmployee = 2
      ElseIf Combo4.Text = "Third Class" Then
          StatusofEmployee = 3
      Else
          StatusofEmployee = 4
      End If
    End If
  
Else
    If Me.MaskEdBox1(0).Text = "__/__/__" Or Me.MaskEdBox1(1).Text = "__/__/__" Or Len(Trim(cboMonth)) = 0 Or Len(Trim(cboYear)) = 0 Then
        MsgBox "Proper Value has not given!", vbInformation, "IT Division, DNMIH"
        MaskEdBox1(0).SetFocus
        Exit Sub
    Else
        rptmode = 26
        BeginDateForReport = MaskEdBox1(0)
        EnddateforReport = MaskEdBox1(1)
        GetMonthOftheYear = cboMonth
        GetSalaryPreparationYaer = cboYear
     End If
End If

Form20.Show vbModal
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
Load_Yr Me
Load_MonthNm Me
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
Private Sub Option1_Click(Index As Integer)
Select Case Index

Case 0
        If Option1(0).Value = True Then
            Option1(0).ForeColor = &HFF0000
            Combo4.Enabled = True
            Label6.Enabled = True
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

Case 2
        If Option1(2).Value = True Then
            Option1(2).ForeColor = &HFF0000
            Option1(2).ForeColor = &H80000012
            Combo4.Enabled = False
            Label6.Enabled = False
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

