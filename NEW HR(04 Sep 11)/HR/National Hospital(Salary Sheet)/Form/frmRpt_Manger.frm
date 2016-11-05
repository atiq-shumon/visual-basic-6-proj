VERSION 5.00
Begin VB.Form Form21 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Manager"
   ClientHeight    =   6795
   ClientLeft      =   2205
   ClientTop       =   2160
   ClientWidth     =   5100
   ForeColor       =   &H00800000&
   Icon            =   "frmRpt_Manger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   5100
   Begin VB.Frame Frame2 
      Height          =   5835
      Index           =   1
      Left            =   135
      TabIndex        =   2
      Top             =   90
      Width           =   4560
      Begin VB.OptionButton optReport 
         Caption         =   "Designation wise Employee Statistics"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   300
         TabIndex        =   13
         Top             =   5460
         Width           =   3210
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Salary Summary Statment"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   300
         TabIndex        =   17
         Top             =   5100
         Width           =   2760
      End
      Begin VB.OptionButton optReport 
         Caption         =   "Bonus Statment Send to Bank"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   300
         TabIndex        =   15
         Top             =   4200
         Width           =   2760
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   4905
         Left            =   225
         TabIndex        =   3
         Top             =   120
         Width           =   4245
         Begin VB.OptionButton optReport 
            Caption         =   "Pay Slip (Bonus)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   60
            TabIndex        =   16
            Top             =   2400
            Width           =   2760
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Yearly Salary Statment"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   14
            Top             =   4560
            Width           =   2760
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   2205
            TabIndex        =   12
            Top             =   300
            Width           =   1725
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Salary Statment Send to Bank"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   11
            Top             =   3645
            Width           =   2760
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Leave Application Information"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   10
            Top             =   3195
            Width           =   2760
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Pay Slip (Salary)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   60
            TabIndex        =   9
            Top             =   1990
            Width           =   2760
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Pro&vident Fund Statement Report"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   8
            Top             =   2775
            Width           =   2760
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Monthly Salary Preparation Report"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   7
            Top             =   1600
            Width           =   2760
         End
         Begin VB.OptionButton optReport 
            Caption         =   "&Loan Statement"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   6
            Top             =   1200
            Width           =   1815
         End
         Begin VB.OptionButton optReport 
            Caption         =   "Over time Preparation Report"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   5
            Top             =   800
            Width           =   2400
         End
         Begin VB.OptionButton optReport 
            Caption         =   "&Employee Information"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   4
            Top             =   345
            Value           =   -1  'True
            Width           =   1815
         End
      End
   End
   Begin VB.CommandButton cmdPreview 
      Height          =   480
      Left            =   2040
      Picture         =   "frmRpt_Manger.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1320
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   3465
      Picture         =   "frmRpt_Manger.frx":2694
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6225
      Width           =   1185
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Close_Msg Me
End Sub
Private Sub cmdPreview_Click()
On Error GoTo Errdesc
If optReport(6).Value = True Then
    Dim f As New frmProvidentFund
    f.Show 1
ElseIf optReport(5).Value = True Then
    Dim f3 As New frmSalaryReportForm
    f3.Show 1
ElseIf optReport(3).Value = True Then
    rptmode = 7
    Form20.Show vbModal
ElseIf optReport(4).Value = True Then
    Dim f2 As New frmLoanLedgerReport
        f2.Show 1
ElseIf optReport(0).Value = True Then
    rptmode = 4
    Form20.Show vbModal
ElseIf optReport(7).Value = True Then   ''''''''PAYSLIP FOR SALARY
    ReportTracker = 3
    Dim f4 As New frmSalaryDisburs
    f4.Show 1
ElseIf optReport(1).Value = True Then
   Dim f01 As New frmLeaveApplicationReport
    f01.Show 1
ElseIf optReport(2).Value = True Then
    ReportTracker = 1
    Dim f12 As New frmSalaryStatment
    f12.Show 1
ElseIf optReport(9).Value = True Then
    ReportTracker = 2
    Dim f13 As New frmSalaryStatment
    f12.Show 1
ElseIf optReport(10).Value = True Then
    ReportTracker = 4        ''''''''PAYSLIP FOR BONUS
    Dim f14 As New frmSalaryDisburs
    f14.Show 1
End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub

Private Sub optReport_Click(Index As Integer)
Select Case Index
Case 0
    
    If optReport(0).Value = True Then
        For i = 1 To 10
            optReport(i).Value = False
        Next i
        
        
        optReport(0).Value = True
        
        Text1.Text = ""
        lblDesig.Visible = True
        Text1.Visible = True
        Text1.SetFocus
    End If
    
Case 1

    For i = 2 To 10
        optReport(i).Value = False
    Next i
    
    optReport(0).Value = False
    optReport(1).Value = True
    lblDesig.Visible = False
    Text1.Visible = False

Case 2

    For i = 0 To 1
        optReport(i).Value = False
    Next i
    
    For i = 3 To 10
        optReport(i).Value = False
    Next i
    
    optReport(2).Value = True
'    lblDesig.Visible = False
    Text1.Visible = False
    
Case 3

    For i = 0 To 2
        optReport(i).Value = False
    Next i
    
    For i = 4 To 10
        optReport(i).Value = False
    Next i
    
    optReport(3).Value = True
    lblDesig.Visible = False
    Text1.Visible = False

Case 4

    For i = 0 To 3
        optReport(i).Value = False
    Next i
    
    For i = 5 To 10
        optReport(i).Value = False
    Next i
    
     
    optReport(4).Value = True
    ''lblDesig.Visible = False
    Text1.Visible = False
    
Case 5

    For i = 0 To 4
        optReport(i).Value = False
    Next i
    
    For i = 6 To 10
        optReport(i).Value = False
    Next i
    
    optReport(5).Value = True
   '' lblDesig.Visible = False
    Text1.Visible = False
    
Case 6

    For i = 0 To 5
        optReport(i).Value = False
    Next i
    
    For i = 7 To 10
        optReport(i).Value = False
    Next i
    
    optReport(6).Value = True
    lblDesig.Visible = False
    Text1.Visible = False

Case 7

    For i = 0 To 6
        optReport(i).Value = False
    Next i
    
    For i = 8 To 10
        optReport(i).Value = False
    Next i
    
    optReport(7).Value = True
    ''lblDesig.Visible = False
    Text1.Visible = False

Case 8

    For i = 0 To 7
        optReport(i).Value = False
    Next i
    optReport(10).Value = False
    optReport(9).Value = False
    optReport(8).Value = False
    
    Dim f As New frmTowhomitConcern
    f.Show 1

Case 9

    For i = 0 To 8
        optReport(i).Value = False
    Next i
    
    optReport(10).Value = False
    optReport(9).Value = True
   '''lblDesig.Visible = False
    Text1.Visible = False


Case 10

    For i = 0 To 9
        optReport(i).Value = False
    Next i
    
    optReport(10).Value = True
''    lblDesig.Visible = False
    Text1.Visible = False
    
Case 11

    For i = 0 To 10
        optReport(i).Value = False
    Next i
    
    optReport(11).Value = True
    '''lblDesig.Visible = False
    Text1.Visible = False
    Dim f15 As New frmSalarySummary
    f15.Show 1
    
Case 12
    For i = 0 To 11
        optReport(i).Value = False
    Next i
    
    optReport(12).Value = True
'    lblDesig.Visible = False
    Text1.Visible = False
    Dim f16 As New frmPayrollStatistics
    f16.Show 1

        
End Select
End Sub
