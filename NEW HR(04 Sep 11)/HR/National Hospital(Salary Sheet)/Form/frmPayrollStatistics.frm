VERSION 5.00
Begin VB.Form frmPayrollStatistics 
   Caption         =   "Payroll Statistics"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form6"
   ScaleHeight     =   2520
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   3900
      TabIndex        =   5
      Top             =   1830
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      Caption         =   "Statistics"
      Height          =   1545
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   5865
      Begin VB.ComboBox cboYear 
         Height          =   315
         ItemData        =   "frmPayrollStatistics.frx":0000
         Left            =   2850
         List            =   "frmPayrollStatistics.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   750
         Width           =   1755
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   450
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   750
         Width           =   1665
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Designation Wise Salary Disbursee"
         Height          =   405
         Left            =   270
         TabIndex        =   2
         Top             =   300
         Width           =   2835
      End
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   435
      Left            =   1920
      TabIndex        =   0
      Top             =   1830
      Width           =   1905
   End
End
Attribute VB_Name = "frmPayrollStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdShow_Click()
'   BEGINYEARFORWHOM = cboMonth + ", " + cboYear
'         ENDDATEFORWHOM = cboMonthTo + ", " + cboYearTo
'         GetFromMonthtoWhom = "01-" + Trim(Mid(cboMonth, 1, 3)) + "-" + Trim(cboYear)
'         GetToMonthtoWhom = "28-" + Trim(Mid(cboMonthTo, 1, 3)) + "-" + Trim(cboYearTo)
'        If CDate(GetFromMonthtoWhom) > CDate(GetToMonthtoWhom) Then
'          MsgBox "Improper Date Range   " + Chr(13) + " Please put a valid date range.", vbInformation, "Date entry error..."
'          cboMonth.SetFocus
'          Exit Sub
'        End If
'   If Option1(1).Value = True Then
'    currentOption = 2
'   ElseIf Option1(2).Value = True Then
'    currentOption = 3
'    currentDept = Trim(cboDept.Text)
'   End If
    
   
   
'  If Option1(0).Value = True Then
'        rptmode = 46
'        Form20.Show 1
'  Else
        paramMonth = cboMonth.Text
        paramYear = cboYear.Text
        rptmode = 48
        Form20.Show 1
'  End If



End Sub

Private Sub Form_Load()
  Load_Yr Me
  Load_MonthNm Me
  cboMonth.Text = MonthName(Month(Now))
  cboYear.Text = YEAR(Now)
  
End Sub
