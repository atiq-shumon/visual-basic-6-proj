VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form8 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Festival Bonus / Production Bonus / Profit Bonus"
   ClientHeight    =   6015
   ClientLeft      =   1110
   ClientTop       =   1710
   ClientWidth     =   9675
   Icon            =   "frmBonus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9675
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   4500
      Picture         =   "frmBonus.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5355
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   6975
      Picture         =   "frmBonus.frx":22D4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5355
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   3255
      Picture         =   "frmBonus.frx":3D56
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5355
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   5745
      Picture         =   "frmBonus.frx":56E8
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5355
      Width           =   1185
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Regardless Unit"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   180
      TabIndex        =   10
      Top             =   5490
      Width           =   1455
   End
   Begin VB.CommandButton cmdProceess 
      Height          =   480
      Left            =   1800
      Picture         =   "frmBonus.frx":72D2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5355
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   9420
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3030
         Left            =   270
         TabIndex        =   15
         Top             =   1755
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   5345
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         ColumnHeaders   =   0   'False
         ForeColor       =   10485760
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   6
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
         ColumnCount     =   2
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1440
         TabIndex        =   6
         Top             =   480
         Width           =   3120
         Begin VB.OptionButton optFestival 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Festival"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   870
         End
         Begin VB.OptionButton optProd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Production"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1125
            TabIndex        =   8
            Top             =   0
            Width           =   1140
         End
         Begin VB.OptionButton optProfit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Profit"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2385
            TabIndex        =   7
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fiscal Year"
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
         Index           =   0
         Left            =   225
         TabIndex        =   16
         Top             =   900
         Width           =   825
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
         Left            =   4860
         TabIndex        =   5
         Top             =   1305
         Width           =   825
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
         Left            =   225
         TabIndex        =   4
         Top             =   1305
         Width           =   750
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year && Month"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   4815
         TabIndex        =   3
         Top             =   450
         Width           =   960
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   3165
         Index           =   2
         Left            =   180
         Top             =   1710
         Width           =   9060
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus Name"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   450
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   375
         Index           =   16
         Left            =   1305
         Top             =   405
         Width           =   3300
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim Bns_Type As Integer
'
'Private Bns As New Bonus
'Private Bns_Rs As New Recordset
'Private Sub cboMonth_Click()
'    Flash_Into_Grid
'End Sub
'
'Private Sub cboUnit_Click()
'    'Load_CostNm Me, cboUnit
'End Sub
'
'
'
'Private Sub cboYear_Click()
'    Flash_Into_Grid
'End Sub
'
'Private Sub Check1_Click()
'
'If Check1.Value = 1 Then
''    cboUnit = ""
''    cboUnit.Locked = True
''    cboCost = ""
''    cboCost.Locked = True
'
'Else
''    cboUnit.Locked = False
''    cboCost.Locked = False
'
'End If
'
'End Sub
'
'Private Sub cmdClear_Click()
'    Clear_Screen
'End Sub
'
'Private Sub cmdClose_Click()
'    Close_Msg Me
'End Sub
'
'Private Sub cmdDelete_Click()
'
'    Dim con As New ADODB.Connection
'    Dim cmd As New ADODB.Command
'    Dim RS As New ADODB.Recordset
'
'    If optFestival Then Bns_Type = 1
'    If optProd Then Bns_Type = 2
'    If optProfit Then Bns_Type = 3
'
'    '---------------------------------------
'        con.Open strCN.Connection_String
'
'        Set cmd.ActiveConnection = con
'
'        cmd.CommandType = adCmdStoredProc
'        cmd.CommandText = "Delete_Bonus"
'        cmd(1) = cboMonth
'        cmd(2) = cboYear
'        cmd(3) = Bns_Type
'        Set RS = cmd.Execute
'
'        MsgBox RS!Message, vbOKOnly + vbExclamation
'
'        con.Close
'
'    '---------------------------------------
'
'    Flash_Into_Grid
'End Sub
'
'Private Sub cmdPrint_Click()
'
'    If optFestival Then Rpt_Nm = "Bns13"
'    If optProd Then Rpt_Nm = "Bns14"
'    If optProfit Then Rpt_Nm = "Bns15"
'
'
'    Rpt_Month = cboMonth
'    Rpt_Year = cboYear
'    Rpt_Fiscal_Yr = cboFiscalYr
'
'
'    Form20.Show vbModal
'
'
'End Sub
'
'Private Sub cmdProceess_Click()
'
'    If optFestival Then Bns_Type = 1
'    If optProd Then Bns_Type = 2
'    If optProfit Then Bns_Type = 3
'
'    With Bns
'        .Connstring = strCN.Connection_String
'
'        .Fiscl_Year = cboFiscalYr
'        .Unit = cboUnit
'        .Cost = cboCost
'        .PAY_MONTH = cboMonth
'        .PAY_YEAR = cboYear
'        .Bonus_Type = Bns_Type
'        .U_Id = U_Id
'        .Save
'    End With
'
'    Flash_Into_Grid
'
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then Unload Me
'End Sub
'
'Private Sub Form_Load()
'   On Error Resume Next
'    Screen_Position Me
''    Load_Yr Me
''    Load_MonthNm Me
''    Load_FiscalYr Me
''    Load_UnitNm Me
'
'End Sub
'Public Sub Flash_Into_Grid()
'On Error GoTo Errdes
'    If optFestival Then Bns_Type = 1
'    If optProd Then Bns_Type = 2
'    If optProfit Then Bns_Type = 3
'
'    With Bns
'        .Connstring = strCN.Connection_String
'
'        .PAY_MONTH = cboMonth
'        .PAY_YEAR = cboYear
'        .Bonus_Type = Bns_Type
'        Set Bns_Rs = .GetAll
'    End With
'
'     Set DataGrid1.DataSource = Bns_Rs
'
'        With DataGrid1
'            .Columns(0).Width = 650
'            '.Columns(0).DataField = Prod_Rs!Fields(0)
'
'            .Columns(1).Width = 1875
'            '.Columns(1).DataField = Prod_Rs!Fields(1)
'
'            .Columns(2).Width = 600
'            '.Columns(2).DataField = Prod_Rs!Fields(2)
'
'            .Columns(3).Width = 850
'            '.Columns(3).DataField = Prod_Rs!Fields(3)
'
'             .Columns(4).Width = 650
'            '.Columns(4).DataField = Prod_Rs!Fields(3)
'
'             .Columns(5).Width = 1250
'            '.Columns(5).DataField = Prod_Rs!Fields(3)
'
'             .Columns(6).Width = 1250
'            '.Columns(6).DataField = Prod_Rs!Fields(3)
'
'        End With
'Exit Sub
'Errdes:
'    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
'
'End Sub
'
'Private Sub optFestival_Click()
'     Flash_Into_Grid
'End Sub
'
'Private Sub optProd_Click()
'     Flash_Into_Grid
'End Sub
'
'Private Sub optProfit_Click()
'     Flash_Into_Grid
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'    Destroy Me
'End Sub
'
