VERSION 5.00
Begin VB.Form frmSalarySummary 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboDept 
      Height          =   315
      ItemData        =   "frmSalarySummary.frx":0000
      Left            =   3240
      List            =   "frmSalarySummary.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1080
      Width           =   3255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Department Wise Break-down"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   16
      Top             =   1110
      Width           =   2865
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Department Wise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   2220
      TabIndex        =   15
      Top             =   810
      Width           =   1905
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Class Wise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   14
      Top             =   810
      Value           =   -1  'True
      Width           =   1905
   End
   Begin VB.ComboBox cboMonthTo 
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
      ItemData        =   "frmSalarySummary.frx":0004
      Left            =   3870
      List            =   "frmSalarySummary.frx":002C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1920
      Width           =   1560
   End
   Begin VB.ComboBox cboYearTo 
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
      ItemData        =   "frmSalarySummary.frx":0092
      Left            =   5430
      List            =   "frmSalarySummary.frx":00D8
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
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
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1890
      Width           =   1215
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
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1890
      Width           =   1560
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   -30
      TabIndex        =   6
      Top             =   2670
      Width           =   6825
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   375
         Left            =   5430
         TabIndex        =   8
         Top             =   210
         Width           =   1155
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "View Report"
         Height          =   375
         Left            =   4230
         TabIndex        =   7
         Top             =   210
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   -90
      TabIndex        =   4
      Top             =   -120
      Width           =   6885
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Disbursement Summary Statement"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   495
         Left            =   180
         TabIndex        =   5
         Top             =   210
         Width           =   6915
      End
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   210
      Top             =   690
      Width           =   6405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   5910
      TabIndex        =   13
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4080
      TabIndex        =   12
      Top             =   1680
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   390
      Left            =   3150
      TabIndex        =   11
      Top             =   1860
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   300
      TabIndex        =   10
      Top             =   1650
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   2010
      TabIndex        =   9
      Top             =   1650
      Width           =   330
   End
End
Attribute VB_Name = "frmSalarySummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
 Unload Me
End Sub
Private Sub cmdShow_Click()
         BEGINYEARFORWHOM = cboMonth + ", " + cboYear
         ENDDATEFORWHOM = cboMonthTo + ", " + cboYearTo
         GetFromMonthtoWhom = "01-" + Trim(Mid(cboMonth, 1, 3)) + "-" + Trim(cboYear)
         GetToMonthtoWhom = "28-" + Trim(Mid(cboMonthTo, 1, 3)) + "-" + Trim(cboYearTo)
        If CDate(GetFromMonthtoWhom) > CDate(GetToMonthtoWhom) Then
          MsgBox "Improper Date Range   " + Chr(13) + " Please put a valid date range.", vbInformation, "Date entry error..."
          cboMonth.SetFocus
          Exit Sub
        End If
   If Option1(1).Value = True Then
    currentOption = 2
   ElseIf Option1(2).Value = True Then
    currentOption = 3
    currentDept = Trim(cboDept.Text)
   End If
    
   
   
  If Option1(0).Value = True Then
        rptmode = 46
        Form20.Show 1
  Else
        rptmode = 47
        Form20.Show 1
  End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys (Chr(9))
  End If
End Sub

Private Sub Form_Load()
  Load_Yr Me
  Load_MonthNm Me
  cboMonthTo.Text = MonthName(Month(Now))
  cboYearTo.Text = YEAR(Now)
  Load_Departments Me
End Sub

