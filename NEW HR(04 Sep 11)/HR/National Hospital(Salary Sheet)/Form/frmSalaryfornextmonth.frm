VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSalaryfornextmonth 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salary Preparation for the Next Month...."
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3375
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   4935
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
         ItemData        =   "frmSalaryfornextmonth.frx":0000
         Left            =   2760
         List            =   "frmSalaryfornextmonth.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   1575
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
         ItemData        =   "frmSalaryfornextmonth.frx":008E
         Left            =   1560
         List            =   "frmSalaryfornextmonth.frx":00D4
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cmbyear 
         Height          =   315
         ItemData        =   "frmSalaryfornextmonth.frx":015C
         Left            =   1680
         List            =   "frmSalaryfornextmonth.frx":01A2
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox cmbmonth 
         Height          =   315
         ItemData        =   "frmSalaryfornextmonth.frx":022A
         Left            =   2760
         List            =   "frmSalaryfornextmonth.frx":0252
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   330
         Left            =   2640
         TabIndex        =   12
         Top             =   2445
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Year/Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   18
         Top             =   525
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Year/Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   1860
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   870
         Left            =   2280
         TabIndex        =   16
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Salary Preparation Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   2520
         Width           =   2040
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808000&
         Height          =   3015
         Left            =   120
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
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
      Index           =   1
      Left            =   4080
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "Prepare...."
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
         Index           =   0
         Left            =   2760
         TabIndex        =   1
         Top             =   3600
         Width           =   1095
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   435
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   435
      Index           =   1
      Left            =   720
      TabIndex        =   7
      Top             =   1200
      Width           =   225
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   435
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   1680
      Width           =   195
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   435
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   225
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   435
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   435
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   3120
      Width           =   240
   End
End
Attribute VB_Name = "frmSalaryfornextmonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PV_Preparation As New clsLoan
Dim VariabletoCheckSal As Boolean

Private Sub Command1_Click(Index As Integer)
On Error GoTo Errdesc
Select Case Index
Case 0

    If Trim(Len(cboYear)) = 0 Then
       MsgBox "Process Terminated due to Year", vbInformation, organizationInfo
       cboYear.SetFocus
       Exit Sub
    ElseIf Trim(Len(cboMonth)) = 0 Then
       MsgBox "Process Terminated due to Month", vbInformation, organizationInfo
       cboYear.SetFocus
       Exit Sub
       
    ElseIf Trim(Len(cmbyear)) = 0 Then
       MsgBox "Process Terminated due to Month", vbInformation, organizationInfo
       cmbyear.SetFocus
       Exit Sub
    ElseIf Trim(Len(cmbmonth)) = 0 Then
       MsgBox "Process Terminated due to Month", vbInformation, organizationInfo
       cmbmonth.SetFocus
       Exit Sub
    ElseIf MaskEdBox1 = "__/__/__" Then
       MsgBox "Date is not availabe !", vbInformation, organizationInfo
       MaskEdBox1.SetFocus
       Exit Sub
    Else
       Check_Whether_SalaryTransferOrNot
       If VariabletoCheckSal = True Then
            SalaryPaymentMonth = cmbmonth
            SalaryPaymentYear = cmbyear
            Salary_Preparation_for_Next_Month
            MsgBox "Salary has Prepared for the month of " & SalaryPaymentMonth & " and Year of " & SalaryPaymentYear, vbInformation, organizationInfo
        Else
            MsgBox "Salary has alredy been Transferred for this month", vbCritical, organizationInfo
        End If
    End If


Case 1
        Unload Me
End Select
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, organizationInfo
End Sub

Private Sub Form_Load()
    Load_Yr Me
    Load_MonthNm Me
End Sub
Private Sub Salary_Preparation_for_Next_Month()
On Error GoTo Errdes
With PV_Preparation
    .Connstring = strCN.Connection_String
    .SalaryPreparationYearFrom = cboYear
    .SalaryPreparationMonthFrom = cboMonth
    .SalaryPreparationYearTo = cmbyear
    .SalaryPreparationMonthTo = cmbmonth
    .SalaryPreparationDate = MaskEdBox1
    .Salary_Preparation_For_the_NextMonth_Save
End With
Exit Sub
Errdes:
MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Check_Whether_SalaryTransferOrNot()
On Error GoTo Errdes
Dim conn5 As New Connection
Dim cmd As New Command
Dim RS2 As New ADODB.Recordset
conn5.ConnectionString = strCN.Connection_String
conn5.Open
cmd.ActiveConnection = conn5
cmd.CommandType = adCmdText
cmd.CommandText = "select EMP_ID from salary_preparation where PAY_MONTH='" & cmbmonth & "' and PAY_YEAR='" & cmbyear & "' and salary_type='R'"

cmd.Properties("iRowsetChange") = True
cmd.Properties("updatability") = 7
RS2.CursorLocation = adUseClient

RS2.Open cmd.CommandText, conn5, adOpenDynamic, adLockOptimistic

If RS2.BOF Or RS2.EOF Then
   VariabletoCheckSal = True
Else
    VariabletoCheckSal = False
End If
   
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, organizationInfo
End Sub


