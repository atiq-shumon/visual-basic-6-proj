VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form22 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Overtime Preparation"
   ClientHeight    =   6090
   ClientLeft      =   660
   ClientTop       =   1560
   ClientWidth     =   8055
   Icon            =   "frmOT_Preparation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   5280
      Index           =   2
      Left            =   180
      TabIndex        =   20
      Top             =   90
      Width           =   7665
      Begin VB.TextBox lblNight_Amount 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   48
         Text            =   "0"
         Top             =   2880
         Width           =   1155
      End
      Begin VB.TextBox txtRS 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   47
         Text            =   "0"
         Top             =   3600
         Width           =   1155
      End
      Begin VB.TextBox txtNight_Hr 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   46
         Text            =   "0"
         Top             =   3990
         Width           =   1155
      End
      Begin VB.TextBox lblNet_Payable 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   45
         Text            =   "0"
         Top             =   3960
         Width           =   1575
      End
      Begin VB.TextBox txtHr_Day_Deduction 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   44
         Text            =   "0"
         Top             =   3240
         Width           =   1155
      End
      Begin VB.ComboBox OvertimePeriod 
         Height          =   315
         ItemData        =   "frmOT_Preparation.frx":08CA
         Left            =   1260
         List            =   "frmOT_Preparation.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2370
         Width           =   6195
      End
      Begin VB.TextBox TextBox2 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   1080
         TabIndex        =   40
         Top             =   4650
         Width           =   6375
      End
      Begin VB.TextBox TextBox1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   39
         Text            =   "1"
         Top             =   2790
         Width           =   1575
      End
      Begin VB.TextBox TextBox3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         TabIndex        =   38
         Top             =   1950
         Width           =   6135
      End
      Begin VB.TextBox label2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   37
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox lblUnit 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         TabIndex        =   36
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox lblName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   35
         Top             =   330
         Width           =   3375
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox lblDesig 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1260
         TabIndex        =   32
         Top             =   720
         Width           =   4215
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   4455
         TabIndex        =   1
         Top             =   1125
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1335
         TabIndex        =   0
         Top             =   330
         Width           =   2085
      End
      Begin VB.OptionButton optOT_Type 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Daily Wise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Index           =   3
         Left            =   4635
         TabIndex        =   14
         Top             =   3270
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton optOT_Type 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Combindly"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   2
         Left            =   6165
         TabIndex        =   15
         Top             =   3270
         Width           =   1185
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   4500
         TabIndex        =   21
         Top             =   1620
         Width           =   2895
         Begin VB.OptionButton optOT_Type 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Overtime"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optOT_Type 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Holiday Overtime"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   1
            Left            =   1080
            TabIndex        =   13
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tk."
         Height          =   255
         Left            =   2460
         TabIndex        =   43
         Top             =   2940
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Label2"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   42
         Top             =   4320
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   2
         Top             =   4680
         Width           =   630
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Scale"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   5640
         TabIndex        =   3
         Top             =   765
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "No of Days Overtimed"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   2880
         TabIndex        =   4
         Top             =   2835
         Width           =   1560
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         Height          =   330
         Left            =   1305
         Top             =   315
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pay. Date:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   3555
         TabIndex        =   5
         Top             =   1155
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Payment Taken Type"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   2985
         TabIndex        =   6
         Top             =   3270
         Width           =   1530
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   0
         Left            =   4575
         Top             =   3225
         Width           =   2865
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scale"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   315
         TabIndex        =   7
         Top             =   1980
         Width           =   825
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Payable                                      Tk."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   4635
         TabIndex        =   8
         Top             =   4050
         Width           =   2820
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rev. Stamp                             Tk."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   315
         TabIndex        =   11
         Top             =   3645
         Width           =   2385
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount   "
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   28
         Top             =   2835
         Width           =   675
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Others                                     Tk."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   29
         Top             =   3240
         Width           =   2460
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime (/Day)"
         ForeColor       =   &H00800000&
         Height          =   420
         Index           =   1
         Left            =   315
         TabIndex        =   31
         Top             =   2340
         Width           =   1080
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year/Month"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   315
         TabIndex        =   30
         Top             =   1575
         Width           =   855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   27
         Top             =   1170
         Width           =   285
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Name"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   3555
         TabIndex        =   26
         Top             =   345
         Width           =   465
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   25
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   24
         Top             =   765
         Width           =   840
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction                                Tk."
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   23
         Top             =   4005
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   315
         Index           =   16
         Left            =   4455
         Top             =   1575
         Width           =   2985
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OT Type"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   3555
         TabIndex        =   22
         Top             =   1620
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdPreview 
      Height          =   480
      Left            =   4080
      Picture         =   "frmOT_Preparation.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5490
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   6705
      Picture         =   "frmOT_Preparation.frx":2698
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5490
      Width           =   1140
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   1515
      Picture         =   "frmOT_Preparation.frx":411A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5490
      Width           =   1140
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   240
      Picture         =   "frmOT_Preparation.frx":5AAC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5490
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   5415
      Picture         =   "frmOT_Preparation.frx":743E
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5490
      Width           =   1140
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   2805
      Picture         =   "frmOT_Preparation.frx":9028
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5490
      Width           =   1140
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objOvertime As New Overtime
Dim Month_Days As Integer
Dim BASIC As Double
Dim Monthly_Hr As Double                '' Monthly total working hour (208)
Dim OT_times As Double                  '' % of Hourly basic pay (200%)
Dim Daily_Hr As Integer                  '' Daily working hour(8)
Dim Hol_OT_times As Integer              '' % or Daily basic pay (300%)
Dim OT_Type As Integer                  '' Overtime of Holiday Overtime
Dim Track_Id As Long
Dim OvertimeVariable
Dim SSTab_Index As Integer
Private Job_Info As New clsEmp_Job_Detail
Dim NightAmountPayable, NetAmountPayable, ScaleCodeofEmployee
Private Sub cboMonth_Click()
'If txtEmpId = "" Then Exit Sub
If Trim(Combo1.Text) = "" Then Exit Sub
 Flash_Data
End Sub
Private Sub cboYear_Click()
'If txtEmpId = "" Then Exit Sub
On Error Resume Next
If Trim(Combo1.Text) = "" Then Exit Sub
    Flash_Data
    cboMonth.SetFocus
End Sub
Private Sub cmdClear_Click()
On Error GoTo Errdesc
    Clear_Screen
    Load_Yr Me
    Load_MonthNm Me
    txtHr_Day_Deduction = 0
    txtRS = 0
    txtNight_Hr = 0
    Combo1.SetFocus
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub cmdClose_Click()
    Close_Msg Me
End Sub
Private Sub cmdDelete_Click()
On Error GoTo Errdesc
With objOvertime
    If Len(Trim(Combo1)) = 0 Then
        MsgBox "Employee ID not  Avialable", vbInformation, "Daffodil Sotware Ltd"
        Combo1.SetFocus
        Exit Sub
    ElseIf MaskEdBox1.Text = "__/__/__" Then
        MsgBox "Date is not Avialable", vbInformation, "Daffodil Sotware Ltd"
        MaskEdBox1.SetFocus
        Exit Sub
    Else
        .Connstring = strCN.Connection_String
        .Emp_ID = Combo1.Text
        .PAYDATE = MaskEdBox1
        .Delete_Overtime
        
        MsgBox "Data Deleted Succesfully", vbInformation, "IT Division, DNMIH"
        cmdClear_Click
    End If
End With

Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub

Private Sub cmdPreview_Click()
rptmode = 7
Form20.Show vbModal
End Sub
Private Sub cmdPrint_Click()
Dim f As New frmOvertimePreparationRpt
f.Show 1
End Sub
Private Sub cmdSave_Click()
On Error GoTo Errdesc
With objOvertime
    If MaskEdBox1.Text = "__/__/__" Then
        MsgBox "Date is not Avialable", vbInformation, "IT Division,DNMIH"
        MaskEdBox1.SetFocus
        Exit Sub
    Else
        .Connstring = strCN.Connection_String
        .Emp_ID = Combo1.Text
        .PAYDATE = MaskEdBox1
        .YEARFORPAYMENT = cboYear
        .MONTHFORPAYMENT = cboMonth
        .PayScale = Trim(TextBox3.Text)
        .OVERTIMEHOURPERDAY = Trim(OvertimePeriod)
        .Amount = Trim(lblNight_Amount)
        .OTHERSAMOUNT = txtHr_Day_Deduction
        .REVSTAMP = txtRS
        .DEDUCTION = txtNight_Hr
        .NETPAYABLE = lblNet_Payable
        .NOOFDAYS = TextBox1.Text
        
        If optOT_Type(3).Value = True Then
            .PAYMENTTYPR = 0
        ElseIf optOT_Type(4).Value = True Then
            .PAYMENTTYPR = 1
        End If
        
        If optOT_Type(0).Value = True Then
            .OTTYPE = 0
        ElseIf optOT_Type(1).Value = True Then
            .OTTYPE = 1
        End If
        
        .Remarks = TextBox2.Text
        Dim HRS
        HRS = Mid(Trim(Me.OvertimePeriod), 1, 1)
        If HRS = "F" Then
           HRS = 8
        End If
        
        .NoofHrsOvertime = HRS
        .Save
        
        MsgBox "Data Saved Succesfully", vbInformation, "IT Division, DNMIH"
    End If
End With

Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Combo1_Change()
If Len(Trim(Combo1)) = 0 Then
    cmdClear_Click
End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
             
        With Job_Info
            .Connstring = strCN.Connection_String
            .Emp_ID = Trim(Combo1.Text)
            .Get_Employee
            lblName = .Emp_Nm
            lblDesig = .designation
            lblUnit = .Dept
        End With
        Get_Employee_Scale_Code
       cboYear.SetFocus
       MaskEdBox1.SetFocus
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next

    Screen_Position Me
      
    OT_Type = 0
    Track_Id = 0
    Load_Yr Me
    Load_MonthNm Me
    MaskEdBox1.Text = Format(Date$, "dd/mm/yy")
    Dim cmd As New Command
    Dim conn10 As New Connection
    Dim rs10 As New Recordset
    
    conn10.ConnectionString = strCN.Connection_String
    conn10.Open
    cmd.ActiveConnection = conn10
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select emp_id from emp_info order by emp_id "
    rs10.CursorLocation = adUseClient
    rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic
    
    If rs10.RecordCount > 0 Then
        Do Until rs10.EOF
            Combo1.AddItem rs10.Fields(0)
            rs10.MoveNext
        Loop
    End If
    
    rs10.Close
    conn10.Close
    
    Set_TabIndex

    OvertimePeriod.AddItem " 1(One)-Hours Overtime "
    OvertimePeriod.AddItem " 2(Two)-Hours Overtime "
    OvertimePeriod.AddItem " 3(Three)-Hours Overtime "
    OvertimePeriod.AddItem " 4(Four)-Hours Overtime "
    OvertimePeriod.AddItem " Full Shift (Single)  Overtime "
    OvertimePeriod.AddItem " Full Shift (Double)  Overtime "

End Sub

Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Get_Employee_Overtime_Info
End If
End Sub

Private Sub optOT_Type_Click(Index As Integer)

    Select Case Index
        
        Case 0
            
        Case 1
    End Select
        OT_Type = Index
        optOT_Type(Index).ForeColor = &HFF&

    Flash_Data
'    Calculate
    Set_TabIndex
    
End Sub
Private Sub OvertimePeriod_Click()
    If (Trim(OvertimePeriod.Text) = Trim("Full Shift (Single)  Overtime")) Then
       lblNight_Amount.Text = 150
    ElseIf (Trim(OvertimePeriod.Text) = Trim("Full Shift (Double)  Overtime")) Then
       lblNight_Amount.Text = 300
    Else
    
      Get_Employee_Scale_Code
    End If
    Calculate_OvertimePreparation
End Sub
Private Sub ScaleCombo_Click()
End Sub
Private Sub txtEmpID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
End Sub
Public Sub Flash_Data()
End Sub
Private Sub TextBox1_Change()
     If Len(TextBox1.Text) = 0 Then
        Get_Employee_Scale_Code
     End If
     
End Sub
Private Sub TextBox1_LostFocus()
'If KeyAscii = 13 Then

If Len(TextBox1.Text) = 0 Then
        lblNight_Amount = NightAmountPayable
        lblNet_Payable = NetAmountPayable
        Calculate_OvertimePreparation
    
   End If
   
   
   If Val(TextBox1.Text) = 0 Then
        lblNight_Amount = NightAmountPayable
        lblNet_Payable = NetAmountPayable
        Calculate_OvertimePreparation
    Else
        lblNight_Amount = lblNight_Amount * Val(TextBox1.Text)
         Calculate_OvertimePreparation
    End If
    
     If Len(TextBox1.Text) = 0 Then
        Get_Employee_Scale_Code
     End If
End Sub
Private Sub txtHr_Day_Deduction_Change()
  Calculate_OvertimePreparation
End Sub
Private Sub txtNight_Hr_Change()
Calculate_OvertimePreparation
End Sub
Private Sub txtNight_Shift_Rate_Change()
End Sub
Private Sub txtOT_Hr_Days_Change()
End Sub
Private Sub txtOT_Hr_Days_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
Private Sub txtOT_Hr_Days_LostFocus()
End Sub
Public Sub Default_Zero(txt As MSForms.TextBox)
    If Len(txt) < 1 Then txt = 0
End Sub
Private Sub txtRS_Change()
Calculate_OvertimePreparation
End Sub
Public Sub Set_TabIndex()
End Sub
Private Sub Calculate_OvertimePreparation()
    lblNet_Payable = Val(Val(lblNight_Amount) + Val(txtHr_Day_Deduction)) - Val(txtNight_Hr)
    'OvertimeVariable = lblNet_Payable - Val(txtRS)
    OvertimeVariable = lblNet_Payable + Val(txtRS)
    lblNet_Payable = OvertimeVariable
End Sub
Private Sub Get_Employee_Scale_Code()
On Error GoTo Errdesc
Dim cmd As New Command
Dim conn007 As New Connection
Dim RS007 As New ADODB.Recordset
conn007.ConnectionString = strCN.Connection_String
conn007.Open
cmd.ActiveConnection = conn007
cmd.CommandType = adCmdText
cmd.CommandText = " SELECT EMP_INFO.EMP_ID, EMP_JOB_INFO.SCALE_CODE " + _
                " From emp_info,EMP_JOB_INFO " + _
                " Where (EMP_JOB_INFO.Emp_ID = emp_info.Emp_ID ) and ( EMP_JOB_INFO.Emp_ID='" & Combo1 & "')"

RS007.CursorLocation = adUseClient
RS007.Open cmd.CommandText, conn007, adOpenDynamic, adLockOptimistic

If RS007.RecordCount > 0 Then
        Label2 = RS007.Fields(1)
        ScaleCodeofEmployee = RS007.Fields(1)
            
            If RS007.Fields(1) = "P1" Then
                TextBox3 = "Tk.23,000(Fixed)"
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 25
                    lblNet_Payable = 25
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 50
                    lblNet_Payable = 50
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 75
                    lblNet_Payable = 75
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 100
                    lblNet_Payable = 100
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If
                
            ElseIf RS007.Fields(1) = "P2" Then
                TextBox3 = "Tk.19,300-700*4-22,100"
                
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 25
                    lblNet_Payable = 25
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 50
                    lblNet_Payable = 50
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 75
                    lblNet_Payable = 75
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 100
                    lblNet_Payable = 100
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If
    
            ElseIf RS007.Fields(1) = "P3" Then
                TextBox3 = "Tk.16,800-6500*6-20,700"
                
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 25
                    lblNet_Payable = 25
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 50
                    lblNet_Payable = 50
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 75
                    lblNet_Payable = 75
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 100
                    lblNet_Payable = 100
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If

                
            ElseIf RS007.Fields(1) = "P4" Then
                TextBox3 = "Tk.15,000-600*8-19,800"
                
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 25
                    lblNet_Payable = 25
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 50
                    lblNet_Payable = 50
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 75
                    lblNet_Payable = 75
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 100
                    lblNet_Payable = 100
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If
            
            
            
            ElseIf RS007.Fields(1) = "P5" Then
                TextBox3 = "Tk.13,750- 5500*10-19,250"
                
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 25
                    lblNet_Payable = 25
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 50
                    lblNet_Payable = 50
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 75
                    lblNet_Payable = 75
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 100
                    lblNet_Payable = 100
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If

            ElseIf RS007.Fields(1) = "P6" Then
                TextBox3 = "Tk.11,000- 4750*14-17,650"
                
                
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 25
                    lblNet_Payable = 25
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 50
                    lblNet_Payable = 50
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 75
                    lblNet_Payable = 75
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 100
                    lblNet_Payable = 100
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If
            
            ElseIf RS007.Fields(1) = "P7" Then
                TextBox3 = "Tk.9,000- 405*16-15,480"
               
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 25
                    lblNet_Payable = 25
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 50
                    lblNet_Payable = 50
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 75
                    lblNet_Payable = 75
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 100
                    lblNet_Payable = 100
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If

            
            
            ElseIf RS007.Fields(1) = "P8" Then
                TextBox3 = "Tk.7,400- 350*16-13,240"
                
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 25
                    lblNet_Payable = 25
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 50
                    lblNet_Payable = 50
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 75
                    lblNet_Payable = 75
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 100
                    lblNet_Payable = 100
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If

            ElseIf RS007.Fields(1) = "P9" Then
                TextBox3 = "Tk.6,800-325*7-9,075-EB-365*11-13,090"
               
               If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 25
                    lblNet_Payable = 25
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 50
                    lblNet_Payable = 50
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 75
                    lblNet_Payable = 75
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 100
                    lblNet_Payable = 100
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If
                
            ElseIf RS007.Fields(1) = "P10" Then
            
                TextBox3 = "Tk.5,100-280*7-7,060-EB-300*11-10,360"
                
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 25
                    lblNet_Payable = 25
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 50
                    lblNet_Payable = 50
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 75
                    lblNet_Payable = 75
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 100
                    lblNet_Payable = 100
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If
                
                
            ElseIf RS007.Fields(1) = "P11" Then
                TextBox3 = "Tk.4,100-250*7-5,850-EB-270*11-8,820"
                 
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 25
                    lblNet_Payable = 25
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 50
                    lblNet_Payable = 50
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 75
                    lblNet_Payable = 75
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 100
                    lblNet_Payable = 100
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If
                        
            ElseIf RS007.Fields(1) = "P12" Then
                TextBox3 = "Tk.3,700-230*7-5310-EB-250*11-8,060"
                
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 25
                    lblNet_Payable = 25
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 50
                    lblNet_Payable = 50
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 75
                    lblNet_Payable = 75
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 100
                    lblNet_Payable = 100
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If

            '============================================ 4-hrs 75Tk. and 45Tk.
            
            ElseIf RS007.Fields(1) = "P13" Then
                   TextBox3 = "Tk.3,500-210*7-4,970-EB-230*11-7,500"
               
               If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 20
                    lblNet_Payable = 20
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 40
                    lblNet_Payable = 40
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 60
                    lblNet_Payable = 60
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 80
                    lblNet_Payable = 80
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If

                              
                
            ElseIf RS007.Fields(1) = "P14" Then
                 TextBox3 = "Tk.3,300-190*7-4,630-EB-210*11-6,940"
               
               If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 20
                    lblNet_Payable = 20
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 40
                    lblNet_Payable = 40
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 60
                    lblNet_Payable = 60
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 80
                    lblNet_Payable = 80
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If

                
            ElseIf RS007.Fields(1) = "P15" Then
               TextBox3 = "Tk.3,100-170*7-4,290-EB-190*11-6,380"
                 
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 20
                    lblNet_Payable = 20
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 40
                    lblNet_Payable = 40
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 60
                    lblNet_Payable = 60
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 80
                    lblNet_Payable = 80
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If
                
            ElseIf RS007.Fields(1) = "P16" Then
                TextBox3 = "Tk.3,000-150*7-4050-EB-170*11-5920"
                
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 20
                    lblNet_Payable = 20
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 40
                    lblNet_Payable = 40
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 60
                    lblNet_Payable = 60
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 80
                    lblNet_Payable = 80
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If
                    
             '==========================================4>65 Tk & 35 Tk.
            ElseIf RS007.Fields(1) = "P17" Then
                TextBox3 = "Tk.2,850-130*7-3,7,60-EB-150*11-5,410"
                
                If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 15
                    lblNet_Payable = 15
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 30
                    lblNet_Payable = 30
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 45
                    lblNet_Payable = 45
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 60
                    lblNet_Payable = 60
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If

            ElseIf RS007.Fields(1) = "P18" Then
                TextBox3 = "Tk.2,600-120*7-34400-EB-130*11-4,870"
                 
                 If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 15
                    lblNet_Payable = 15
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 30
                    lblNet_Payable = 30
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 45
                    lblNet_Payable = 45
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 60
                    lblNet_Payable = 60
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If
            ElseIf RS007.Fields(1) = "P19" Then
                TextBox3 = "Tk.2,500-110*7-3,270-EB-120*11-4,590"
                 
                  If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 15
                    lblNet_Payable = 15
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 30
                    lblNet_Payable = 30
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 45
                    lblNet_Payable = 45
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 60
                    lblNet_Payable = 60
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If
            ElseIf RS007.Fields(1) = "P20" Then
                TextBox3 = "Tk.2400-100*7-3,100-EB-110*11-4310"
                  If OvertimePeriod.Text = OvertimePeriod.List(0) Then
                    lblNight_Amount = 15
                    lblNet_Payable = 15
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(1) Then
                    lblNight_Amount = 30
                    lblNet_Payable = 30
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(2) Then
                    lblNight_Amount = 45
                    lblNet_Payable = 45
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                ElseIf OvertimePeriod.Text = OvertimePeriod.List(3) Then
                    lblNight_Amount = 60
                    lblNet_Payable = 60
                    NightAmountPayable = lblNight_Amount
                    NetAmountPayable = lblNet_Payable
                    
                End If

            End If
        
Else
    Label2 = "Scale N/A"
End If
RS007.Close
conn007.Close
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub Get_Employee_Overtime_Info()
Dim cmd As New Command
Dim conn7 As New Connection
Dim RS7 As New ADODB.Recordset
conn7.ConnectionString = strCN.Connection_String
conn7.Open
cmd.ActiveConnection = conn7
cmd.CommandType = adCmdText

cmd.CommandText = "select PAYDATE,YEARFORPAYMENT,MONTHFORPAYMENT," + _
                "OVERTIMEHOURPERDAY,AMOUNT,OTHERSAMOUNT,REVSTAMP,NETPAYABLE,NOOFDAYS, " + _
                " PAYMENTTYPR,OTTYPE,MONEY_TAKEN,REMARKS  from overtime_preparation where emp_id='" & Trim(Combo1) & "' and PAYDATE=TO_DATE('" & Format(MaskEdBox1, "dd/mmm/yyyy") & "','DD-MON-YYYY')"

RS7.CursorLocation = adUseClient
RS7.Open cmd.CommandText, conn7, adOpenDynamic, adLockOptimistic

If RS7.RecordCount > 0 Then

    MaskEdBox1 = Format(RS7.Fields("PAYDATE"), "dd/mm/yy")
    TextBox2 = "" & RS7.Fields("REMARKS")
    cboYear = RS7.Fields("YEARFORPAYMENT")
    cboMonth = RS7.Fields("MONTHFORPAYMENT")
    'OvertimePeriod = "" & RS7.Fields("OVERTIMEHOURPERDAY")
    lblNight_Amount = RS7.Fields("AMOUNT")
    txtHr_Day_Deduction = RS7.Fields("OTHERSAMOUNT")
    txtRS = RS7.Fields("REVSTAMP")
    cmdDelete.SetFocus
 Else
    OvertimePeriod.SetFocus
End If
End Sub

