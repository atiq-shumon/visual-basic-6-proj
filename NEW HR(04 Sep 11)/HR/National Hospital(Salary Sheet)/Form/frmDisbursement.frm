VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form13 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Salary, Bonus and Overtime  Disbursement"
   ClientHeight    =   4815
   ClientLeft      =   2085
   ClientTop       =   2010
   ClientWidth     =   6045
   Icon            =   "frmDisbursement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form39"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6045
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   4140
      Picture         =   "frmDisbursement.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4095
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   1500
      Picture         =   "frmDisbursement.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4095
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   180
      Picture         =   "frmDisbursement.frx":3CDE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4095
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   2820
      Picture         =   "frmDisbursement.frx":5670
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4095
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3795
      Left            =   90
      TabIndex        =   14
      Top             =   135
      Width           =   5820
      Begin VB.TextBox txtEmpID 
         Height          =   315
         Left            =   3930
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   330
         Width           =   1515
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2340
         TabIndex        =   16
         Top             =   1620
         Width           =   2490
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Individually"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   3
            Top             =   10
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Collectively"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   1
            Left            =   1215
            TabIndex        =   4
            Top             =   10
            Value           =   -1  'True
            Width           =   1185
         End
      End
      Begin MSComCtl2.DTPicker dtpDisburse_Dt 
         Height          =   330
         Left            =   2295
         TabIndex        =   21
         Top             =   315
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   12582912
         CalendarTrailingForeColor=   16576
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   64290819
         CurrentDate     =   37722
      End
      Begin VB.Frame frDis_Method 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1680
         Index           =   1
         Left            =   495
         TabIndex        =   22
         Top             =   1935
         Width           =   4515
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Regardless"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   5
            Left            =   3060
            TabIndex        =   8
            Top             =   495
            Width           =   1320
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Bank"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   4
            Left            =   1890
            TabIndex        =   7
            Top             =   495
            Width           =   1320
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Unit wise"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   2
            Left            =   1890
            TabIndex        =   5
            Top             =   225
            Width           =   1185
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Job Type wise"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   3
            Left            =   3060
            TabIndex        =   6
            Top             =   225
            Width           =   1320
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Unit Name"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   25
            Top             =   825
            Width           =   750
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Criteria"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   24
            Top             =   180
            Width           =   480
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FECCC7&
            BorderColor     =   &H00FECCC7&
            Height          =   600
            Index           =   1
            Left            =   1800
            Top             =   180
            Width           =   2655
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Job type"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   23
            Top             =   1140
            Width           =   600
         End
      End
      Begin VB.Frame frDis_Method 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1365
         Index           =   0
         Left            =   495
         TabIndex        =   26
         Top             =   1935
         Width           =   4605
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Designation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   45
            TabIndex        =   32
            Top             =   1050
            Width           =   840
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   45
            TabIndex        =   31
            Top             =   600
            Width           =   405
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee ID"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   30
            Top             =   180
            Width           =   900
         End
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   0
         Left            =   2295
         Top             =   1575
         Width           =   2655
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Disbursement Date"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   675
         TabIndex        =   20
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year && Month"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   675
         TabIndex        =   19
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment type"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   675
         TabIndex        =   18
         Top             =   810
         Width           =   960
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Disbursal Method"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   675
         TabIndex        =   17
         Top             =   1665
         Width           =   1230
      End
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New Connection
Dim cmd As New Command
Dim Sal_Dis As New Salary_Monthly
Private Sub cboPayType_Click()
    If cboPayType = "Overtime" Or cboPayType = "Holiday Overtime" Then
        Option1(4).Enabled = False      ''Bank
    Else
        Option1(4).Enabled = True     ''Bank
    End If
End Sub
Private Sub cmdClear_Click()
    Clear_Screen
End Sub
Private Sub cmdClose_Click()
    Close_Msg Me
End Sub

Private Sub cmdPrint_Click()
'rptMode = 5


End Sub

Private Sub cmdSave_Click()
    
    Dim Disburse_Method As String
    Dim Disburse_To As String

    If Option1(0).Value = True Then         'Individually
        
        Disburse_Method = "Ind"
        Disburse_To = txtEmpID
        
    ElseIf Option1(2).Value = True Then     'Unit wise
        
        Disburse_Method = "Unit"
        Disburse_To = cboUnit
        
    ElseIf Option1(3).Value = True Then     'Job Type wise
        
        Disburse_Method = "JobType"
        Disburse_To = cboType
    
    ElseIf Option1(4).Value = True Then     'Bank
        
        Disburse_Method = "Bank"
        Disburse_To = "Bank"
        
    ElseIf Option1(5).Value = True Then     '"Regardless"
        
        Disburse_Method = "Regardless"
        Disburse_To = "All"
        
    End If

'If Disburse_Method = Empty Or Disburse_To = Empty Or cboPayType = "" Then Exit Sub

    With Sal_Dis
        .Connstring = strCN.Connection_String
        
        .Disburse_What = cboPayType
        .Disburse_Method = Disburse_Method
        .Disburse_To = Disburse_To
        .PAY_MONTH = Trim(Trim(cboMonth))
        .PAY_YEAR = cboYear
        .Disburse_Date = Valid_Dt(dtpDisburse_Dt.Value)
        .Disburse
    End With
    
    If Option1(0) Then txtEmpID.SetFocus
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
Public Sub Load_PayType()
On Error Resume Next
    With cboPayType
        .AddItem "Monthly Salary"
        .AddItem "Overtime"
        .AddItem "Holiday Overtime"
        
    End With
    
End Sub

Private Sub Form_Load()
    Screen_Position Me
    Load_PayType
    Load_MonthNm Me
    Load_Yr Me
    
    Option1_Click (1)
    
    dtpDisburse_Dt = Now
     
End Sub

Private Sub Option1_Click(Index As Integer)

    Select Case Index
    
        Case 0      'Disbursal Method Indivdual
        
            frDis_Method(1).Visible = False
            frDis_Method(0).Visible = True
            
            Option1(2).Value = False
            Option1(3).Value = False
            
            Option1(2).Enabled = False
            Option1(3).Enabled = False
            Option1(4).Enabled = False
            Option1(5).Enabled = False
                        
            cboUnit.Clear
            cboType.Clear
            
            cboUnit.Enabled = False
            cboType.Enabled = False
            
            txtEmpID.SetFocus
                    
        Case 1      'Collectively
        
            frDis_Method(1).Visible = True
            frDis_Method(0).Visible = False
            
            Option1(2).Enabled = True
            Option1(3).Enabled = True
            Option1(4).Enabled = True
            Option1(5).Enabled = True
            
            If cboPayType = "Overtime" Or cboPayType = "Holiday Overtime" Then
                Option1(4).Enabled = False      ''Bank
            Else
                Option1(4).Enabled = True     ''Bank
            End If
        
        Case 2      'Unit wise
            cboUnit.Enabled = True
            cboType.Enabled = False
            cboType.Clear
            'Load_UnitNm Me
            cboUnit.SetFocus
            
        Case 3      ' Job Type wise
            cboUnit.Enabled = False
            cboType.Enabled = True
            cboUnit.Clear
            Load_JbType Me
            cboType.SetFocus
            
        Case 4, 5
            cboUnit.Enabled = False
            cboType.Enabled = False
        
    End Select
        
    Set_TabIndex

End Sub

Private Sub txtEmpID_Change()

    Get_Employee txtEmpID, Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub

Public Sub Set_TabIndex()


    If Option1(0) Then
    
        cboPayType.TabIndex = 0
        cboMonth.TabIndex = 1
        cboYear.TabIndex = 2
        Option1(0).TabIndex = 3
        Option1(1).TabIndex = 4
        txtEmpID.TabIndex = 5
        cmdSave.TabIndex = 6
        cmdClear.TabIndex = 7
    Else
        
        cboPayType.TabIndex = 0
        cboMonth.TabIndex = 1
        cboYear.TabIndex = 2
        Option1(0).TabIndex = 3
        Option1(1).TabIndex = 4
        Option1(2).TabIndex = 5
        Option1(3).TabIndex = 6
        Option1(4).TabIndex = 7
        Option1(5).TabIndex = 8
        cboUnit.TabIndex = 9
        cboType.TabIndex = 10
        cmdSave.TabIndex = 11
        cmdClear.TabIndex = 12

    
    
    
    End If


End Sub
