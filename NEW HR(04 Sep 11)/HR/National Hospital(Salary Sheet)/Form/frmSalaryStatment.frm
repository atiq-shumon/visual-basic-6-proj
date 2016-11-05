VERSION 5.00
Begin VB.Form frmSalaryStatment 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6120
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Salary Type"
      ForeColor       =   &H008080FF&
      Height          =   555
      Left            =   240
      TabIndex        =   9
      Top             =   1530
      Width           =   5745
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
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   2205
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
         Left            =   2100
         TabIndex        =   12
         Top             =   240
         Width           =   1305
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
         Left            =   1140
         TabIndex        =   11
         Top             =   240
         Width           =   825
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
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee Salary Statment Send to Bank"
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
      Height          =   1455
      Left            =   225
      TabIndex        =   2
      Top             =   180
      Width           =   5730
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
         Left            =   1395
         TabIndex        =   5
         Top             =   405
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
         Left            =   4275
         TabIndex        =   4
         Top             =   405
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmSalaryStatment.frx":0000
         Left            =   1395
         List            =   "frmSalaryStatment.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   810
         Width           =   3975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3735
         TabIndex        =   8
         Top             =   450
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Month"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   450
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Type of Emp"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   855
         Width           =   900
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   395
         Left            =   1350
         Top             =   360
         Width           =   4065
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         Height          =   360
         Left            =   1350
         Top             =   785
         Width           =   4065
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
      Left            =   4725
      Picture         =   "frmSalaryStatment.frx":0050
      TabIndex        =   1
      Top             =   2220
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
      Left            =   3465
      Picture         =   "frmSalaryStatment.frx":091A
      TabIndex        =   0
      Top             =   2220
      Width           =   1215
   End
End
Attribute VB_Name = "frmSalaryStatment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If ReportTracker = 1 Then
    rptmode = 15
Else
    rptmode = 27
End If

GetMonthOftheYear = cboMonth
GetSalaryPreparationYaer = cboYear
If Combo1.Text = "First Class" Then
    StatusofEmployee = "1"
ElseIf Combo1.Text = "2nd Class" Then
    StatusofEmployee = "2"
ElseIf Combo1.Text = "3rd Class" Then
    StatusofEmployee = "3"
ElseIf Combo1.Text = "4th Class" Then
    StatusofEmployee = "4"
Else
    StatusofEmployee = "5"
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
If ReportTracker = 1 Then
    Me.Caption = "Salary Preparation Report Send to bank"
Else
    Me.Caption = "Bonus Preparation Report Send to bank"
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
