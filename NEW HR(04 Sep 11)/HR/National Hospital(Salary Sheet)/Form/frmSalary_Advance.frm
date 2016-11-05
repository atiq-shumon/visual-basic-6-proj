VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form14 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Salary Advance                                                                            Carew & Company (Bangladesh) Limited"
   ClientHeight    =   6930
   ClientLeft      =   1575
   ClientTop       =   1650
   ClientWidth     =   10200
   Icon            =   "frmSalary_Advance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   7065
      Picture         =   "frmSalary_Advance.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6345
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   3120
      Picture         =   "frmSalary_Advance.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6345
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   1800
      Picture         =   "frmSalary_Advance.frx":3CDE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6345
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   5745
      Picture         =   "frmSalary_Advance.frx":5670
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6345
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   4455
      Picture         =   "frmSalary_Advance.frx":725A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6345
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   6225
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   0
      Width           =   10005
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3930
         Left            =   360
         TabIndex        =   17
         Top             =   2025
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   6932
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         ColumnHeaders   =   0   'False
         ForeColor       =   12582912
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1830.047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1560.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2250.142
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpIssueDt 
         Height          =   330
         Left            =   8235
         TabIndex        =   23
         Top             =   1530
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   12582912
         CalendarTrailingForeColor=   16576
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59047939
         CurrentDate     =   37722
      End
      Begin MSForms.ComboBox cboYear 
         Height          =   285
         Left            =   1485
         TabIndex        =   1
         Top             =   1530
         Width           =   1140
         VariousPropertyBits=   746604571
         ForeColor       =   255
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2011;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtAmount 
         Height          =   285
         Left            =   5580
         TabIndex        =   3
         Top             =   1575
         Width           =   1185
         VariousPropertyBits=   746604571
         ForeColor       =   16711680
         BorderStyle     =   1
         Size            =   "2090;503"
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblDesig 
         Height          =   285
         Left            =   1485
         TabIndex        =   22
         Top             =   720
         Width           =   2760
         ForeColor       =   12582912
         BackColor       =   -2147483643
         Size            =   "4868;503"
         BorderColor     =   16761024
         BorderStyle     =   1
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblCost 
         Height          =   285
         Left            =   5580
         TabIndex        =   21
         Top             =   1125
         Width           =   4065
         ForeColor       =   12582912
         BackColor       =   -2147483643
         Size            =   "7170;503"
         BorderColor     =   16761024
         BorderStyle     =   1
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblUnit 
         Height          =   285
         Left            =   1485
         TabIndex        =   20
         Top             =   1125
         Width           =   2760
         ForeColor       =   12582912
         BackColor       =   -2147483643
         Size            =   "4868;503"
         BorderColor     =   16761024
         BorderStyle     =   1
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblName 
         Height          =   285
         Left            =   5580
         TabIndex        =   19
         Top             =   315
         Width           =   4065
         ForeColor       =   12582912
         BackColor       =   -2147483643
         Size            =   "7170;503"
         BorderColor     =   16761024
         BorderStyle     =   1
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cboMonth 
         Height          =   285
         Left            =   2610
         TabIndex        =   2
         Top             =   1530
         Width           =   1590
         VariousPropertyBits=   746604571
         ForeColor       =   255
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2805;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtEmpID 
         Height          =   285
         Left            =   1485
         TabIndex        =   0
         Top             =   315
         Width           =   1500
         VariousPropertyBits=   746604571
         ForeColor       =   255
         BorderStyle     =   1
         Size            =   "2646;503"
         BorderColor     =   16761024
         SpecialEffect   =   0
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   840
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         Height          =   4020
         Index           =   5
         Left            =   315
         Top             =   1980
         Width           =   9330
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Date"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7065
         TabIndex        =   16
         Top             =   1620
         Width           =   765
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   4545
         TabIndex        =   15
         Top             =   1575
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year && Month"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   1575
         Width           =   960
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
         Left            =   360
         TabIndex        =   13
         Top             =   1125
         Width           =   750
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost Centre"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   4545
         TabIndex        =   12
         Top             =   1170
         Width           =   825
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4545
         TabIndex        =   11
         Top             =   345
         Width           =   420
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   345
         Width           =   900
      End
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sal_Adv As New Salary_Adv
Private Sal_Rs As New Recordset
Dim track_id As Long

Private Sub cboMonth_Click()
    Flash_Into_Grid
End Sub

Private Sub cmdClear_Click()
    Clear_Screen
    cboMonth = MonthName(Month(Now))
    cboYear = Year(Now)
     Flash_Into_Grid
    txtEmpId.SetFocus
End Sub

Private Sub cmdClose_Click()
    Close_Msg Me
End Sub

Private Sub cmdDelete_Click()

     With Sal_Adv
        .ConnString = strCN.Connection
        .Emp_Id = txtEmpId
        .Pay_Month = cboMonth
        .Pay_Year = cboYear
        .Issue_Dt = Valid_Dt(dtpIssueDt)
        .Delete
    End With
    
    Flash_Into_Grid
    Clear_Screen
    cboMonth = MonthName(Month(Now))
    cboYear = Year(Now)
     Flash_Into_Grid
    txtEmpId.SetFocus
    
End Sub

Private Sub cmdPrint_Click()

    
    Rpt_Month = cboMonth
    Rpt_Year = cboYear
    Rpt_Nm = "Adv1"
     
    Form20.Show vbModal



End Sub

Private Sub cmdSave_Click()
    
    With Sal_Adv
        .ConnString = strCN.Connection
        .Emp_Id = txtEmpId
        .Pay_Month = cboMonth
        .Pay_Year = cboYear
        .Issue_Dt = Valid_Dt(dtpIssueDt)
        .Amount = txtAmount
        .track_id = track_id
        .U_Id = U_Id
        .Save
    End With
    
    track_id = 0
    Clear_Screen
    cboMonth = MonthName(Month(Now))
    cboYear = Year(Now)
    Flash_Into_Grid
    txtEmpId.SetFocus
    
End Sub

Private Sub DataGrid1_Click()
   ' On Error Resume Next
    With Sal_Rs
        txtEmpId = !Emp_Id
        lblName = !EmpName
        lblUnit = !Unit
        lblCost = !Cost
        lblDesig = !Desig
        cboMonth = !Pay_Month
        cboYear = !Pay_Year
        dtpIssueDt = !Issue_Dt
        txtAmount = !Amount
        track_id = !track_id
    End With
        txtAmount.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
   Screen_Position Me
    track_id = 0
    Load_Yr Me
    Load_MonthNm Me
    Flash_Into_Grid
    txtEmpId.MaxLength = Id_Len
    dtpIssueDt = Now
End Sub


Private Sub txtAmount_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtEmpID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)

        If KeyCode = 13 Then
            Get_Employee txtEmpId, Me
            txtAmount.SetFocus
        End If
End Sub

Public Sub Flash_Into_Grid()

    On Error Resume Next
    
    With Sal_Adv
        .ConnString = strCN.Connection
        .Pay_Month = cboMonth
        .Pay_Year = cboYear
        Set Sal_Rs = .GetAll
    End With
    
     Set DataGrid1.DataSource = Sal_Rs
                    
        With DataGrid1
            .Columns(0).Width = 800
            '.Columns(0).DataField = Prod_Rs!Fields(0)

            .Columns(1).Width = 1875
            '.Columns(1).DataField = Prod_Rs!Fields(1)

            .Columns(2).Width = 750
            '.Columns(2).DataField = Prod_Rs!Fields(2)

            .Columns(3).Width = 1050
            '.Columns(3).DataField = Prod_Rs!Fields(3)
             
             .Columns(4).Width = 1050
            '.Columns(4).DataField = Prod_Rs!Fields(3)
             
             .Columns(5).Width = 650
            '.Columns(5).DataField = Prod_Rs!Fields(3)
             
             .Columns(6).Width = 1050
            '.Columns(6).DataField = Prod_Rs!Fields(3)

        End With
       
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub
