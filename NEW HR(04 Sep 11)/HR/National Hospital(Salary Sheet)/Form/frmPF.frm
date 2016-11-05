VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPF 
   BackColor       =   &H80000009&
   Caption         =   "Provedent Fund"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form6"
   ScaleHeight     =   6840
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   3014
      Picture         =   "frmPF.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   6204
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   4266
      Picture         =   "frmPF.frx":1A0A
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6196
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   5520
      Picture         =   "frmPF.frx":35F4
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6188
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   1762
      Picture         =   "frmPF.frx":5076
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6210
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   540
      Picture         =   "frmPF.frx":6A08
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6180
      Width           =   1185
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   9763
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   -2147483639
      TabCaption(0)   =   "PF Receive"
      TabPicture(0)   =   "frmPF.frx":839A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "PF Payment"
      TabPicture(1)   =   "frmPF.frx":83B6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Member Fund"
      TabPicture(2)   =   "frmPF.frx":83D2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
         Height          =   5235
         Left            =   -75000
         TabIndex        =   52
         Top             =   300
         Width           =   7395
         Begin MSDataGridLib.DataGrid DataGrid3 
            Height          =   2475
            Left            =   150
            TabIndex        =   65
            Top             =   2550
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4366
            _Version        =   393216
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
            Caption         =   "Member Fund"
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
                  LCID            =   2057
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
                  LCID            =   2057
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
         Begin VB.ComboBox cmbMFBank 
            Height          =   315
            Left            =   1740
            TabIndex        =   54
            Top             =   765
            Width           =   2295
         End
         Begin VB.ComboBox cmbMFAccountType 
            Height          =   315
            Left            =   1740
            TabIndex        =   53
            Top             =   1170
            Width           =   2280
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   345
            Index           =   14
            Left            =   1710
            Top             =   1170
            Width           =   2325
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   345
            Index           =   13
            Left            =   1710
            Top             =   750
            Width           =   2355
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Track Id"
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   60
            Top             =   420
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Bank Name"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   59
            Top             =   825
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Account Type"
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   58
            Top             =   1230
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Account No"
            Height          =   195
            Index           =   3
            Left            =   600
            TabIndex        =   57
            Top             =   1665
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Amount"
            Height          =   195
            Index           =   4
            Left            =   600
            TabIndex        =   56
            Top             =   2040
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000009&
         Height          =   5235
         Index           =   1
         Left            =   -75000
         TabIndex        =   11
         Top             =   300
         Width           =   7395
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   1815
            Left            =   180
            TabIndex        =   64
            Top             =   3180
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   3201
            _Version        =   393216
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
            Caption         =   "PF Payment"
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
                  LCID            =   2057
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
                  LCID            =   2057
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
         Begin VB.ComboBox cmbPaymentAccountType 
            Height          =   315
            Left            =   4800
            TabIndex        =   46
            Top             =   2160
            Width           =   1995
         End
         Begin VB.ComboBox cmbPurposeOfPayment 
            Height          =   315
            Left            =   4800
            TabIndex        =   44
            Top             =   330
            Width           =   1995
         End
         Begin VB.ComboBox cmbPaymentType 
            Height          =   315
            Left            =   4800
            TabIndex        =   42
            Top             =   1710
            Width           =   1995
         End
         Begin VB.ComboBox cmbPaymentAccountNo 
            Height          =   315
            Left            =   1500
            TabIndex        =   41
            Top             =   2565
            Width           =   1995
         End
         Begin VB.ComboBox cmbPaymentBank 
            Height          =   315
            Left            =   1500
            TabIndex        =   40
            Top             =   2100
            Width           =   1995
         End
         Begin MSComCtl2.DTPicker DTPPaymentDate 
            Height          =   315
            Left            =   4800
            TabIndex        =   43
            Top             =   1230
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   12582912
            CalendarTrailingForeColor=   16576
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   65077249
            CurrentDate     =   37722
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   12
            Left            =   4770
            Top             =   2130
            Width           =   2025
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   11
            Left            =   4770
            Top             =   1680
            Width           =   2025
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   9
            Left            =   4770
            Top             =   1200
            Width           =   2025
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   8
            Left            =   4770
            Top             =   300
            Width           =   2025
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   7
            Left            =   1470
            Top             =   2520
            Width           =   2025
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   6
            Left            =   1470
            Top             =   2070
            Width           =   2025
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Payment Id"
            Height          =   195
            Index           =   21
            Left            =   210
            TabIndex        =   22
            Top             =   390
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Payment Purpose"
            Height          =   195
            Index           =   20
            Left            =   3450
            TabIndex        =   21
            Top             =   420
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Description"
            Height          =   195
            Index           =   19
            Left            =   210
            TabIndex        =   20
            Top             =   825
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Payment Date"
            Height          =   195
            Index           =   18
            Left            =   3690
            TabIndex        =   19
            Top             =   1290
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Payment Amount"
            Height          =   195
            Index           =   17
            Left            =   210
            TabIndex        =   18
            Top             =   1260
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "PaymentType"
            Height          =   195
            Index           =   16
            Left            =   3690
            TabIndex        =   17
            Top             =   1710
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Check No"
            Height          =   195
            Index           =   15
            Left            =   210
            TabIndex        =   16
            Top             =   1680
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Account No"
            Height          =   195
            Index           =   14
            Left            =   210
            TabIndex        =   15
            Top             =   2580
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Bank Name"
            Height          =   225
            Index           =   13
            Left            =   210
            TabIndex        =   14
            Top             =   2115
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Voucher No"
            Height          =   195
            Index           =   12
            Left            =   3690
            TabIndex        =   13
            Top             =   2610
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Account Type"
            Height          =   195
            Index           =   11
            Left            =   3720
            TabIndex        =   12
            Top             =   2160
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000009&
         Height          =   5235
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   7395
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   1815
            Left            =   180
            TabIndex        =   63
            Top             =   3060
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   3201
            _Version        =   393216
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
            Caption         =   "PF Receive"
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
                  LCID            =   2057
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
                  LCID            =   2057
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
         Begin VB.ComboBox cmbReceiveAccountNo 
            Height          =   315
            Left            =   1440
            TabIndex        =   35
            Top             =   2400
            Width           =   1905
         End
         Begin VB.ComboBox cmbReceiveType 
            Height          =   315
            Left            =   4920
            TabIndex        =   31
            Top             =   1590
            Width           =   1935
         End
         Begin VB.ComboBox cmbSourceofFund 
            Height          =   315
            Left            =   4920
            TabIndex        =   30
            Top             =   330
            Width           =   1935
         End
         Begin VB.ComboBox cmbReceiveAccountType 
            Height          =   315
            Left            =   4920
            TabIndex        =   29
            Top             =   2040
            Width           =   1905
         End
         Begin VB.ComboBox cmbReceiveBank 
            Height          =   315
            Left            =   1440
            TabIndex        =   28
            Top             =   1980
            Width           =   1905
         End
         Begin MSComCtl2.DTPicker dtpReceiveDate 
            Height          =   315
            Left            =   4920
            TabIndex        =   32
            Top             =   1110
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   12582912
            CalendarTrailingForeColor=   16576
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   65077249
            CurrentDate     =   37722
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   5
            Left            =   4890
            Top             =   2010
            Width           =   1935
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   4
            Left            =   4890
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   3
            Left            =   4890
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   2
            Left            =   4890
            Top             =   300
            Width           =   1935
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   1
            Left            =   1410
            Top             =   2370
            Width           =   1935
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000003&
            Height          =   375
            Index           =   10
            Left            =   1410
            Top             =   1950
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Voucher No"
            Height          =   195
            Index           =   9
            Left            =   3810
            TabIndex        =   34
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Account Type"
            Height          =   195
            Index           =   10
            Left            =   3810
            TabIndex        =   33
            Top             =   2040
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Bank Name"
            Height          =   195
            Index           =   8
            Left            =   150
            TabIndex        =   10
            Top             =   2040
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Account No"
            Height          =   195
            Index           =   7
            Left            =   150
            TabIndex        =   9
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Check No"
            Height          =   195
            Index           =   6
            Left            =   180
            TabIndex        =   8
            Top             =   1635
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Receive Type"
            Height          =   195
            Index           =   5
            Left            =   3810
            TabIndex        =   7
            Top             =   1620
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Receive Amount"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   6
            Top             =   1215
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Receive Date"
            Height          =   195
            Index           =   3
            Left            =   3810
            TabIndex        =   5
            Top             =   1170
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Description"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   4
            Top             =   810
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Source of Fund"
            Height          =   195
            Index           =   1
            Left            =   3720
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "Receive Id"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   390
            Width           =   780
         End
      End
   End
End
Attribute VB_Name = "frmPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oPFReceive As New clsPFReceive
Private oPFPayment As New clsPFPayment
Private oMemberFund As New clsMemberFund
Dim SSTab_Index As Integer
Private Sub cmdClear_Click()

Clear_Screen

End Sub

Private Sub cmdClose_Click()
Close_Msg Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Errdesc
Select Case SSTab_Index
Case 0
        With oPFReceive
            .Connstring = strCN.Connection_String
            .PFReceiveId = txtReceiveId
            .Delete
        End With
        MsgBox "Data Deleted Successfully", vbInformation, "IT Division, DNMIH"
        Clear_Screen
        cmbSourceofFund.SetFocus
Case 1
        With oPFPayment
            .Connstring = strCN.Connection_String
            .PFPaymentId = txtPaymentId
            .Delete
        End With
        MsgBox "Data Deleted Successfully", vbInformation, "IT Division, DNMIH"
        Clear_Screen
        cmbPurposeOfPayment.SetFocus
Case 2
        With oMemberFund
            .Connstring = strCN.Connection_String
            .TrackId = txtTrackId
            .Delete
        End With

        MsgBox "Data Deleted Successfully", vbInformation, "IT Division, DNMIH"
        Clear_Screen
        cmbMFAccountType.SetFocus


End Select
Show_Data_PF_Form_Load
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub

Private Sub cmdPrint_Click()
Dim f As New frmProvidentFund
f.Show 1

End Sub

Private Sub cmdSave_Click()
  Select Case SSTab_Index
      
      Case 0
      
        With oPFReceive
        
            .Connstring = strCN.Connection_String
            .PFReceiveId = txtReceiveId
            .SourceOfFund = Get_Code(cmbSourceofFund.Text)
            .Description = txtReceiveDescription
            .PaymentReceiveDate = dtpReceiveDate
            .ReceiveAmount = txtReceiveAmount
            .PaymentReceivedType = cmbReceiveType
            .CheckNo = txtReceiveCheckNo
            .AccountNo = cmbReceiveAccountNo
            .BankCode = Get_Code(cmbReceiveBank.Text)
            .VoucherNo = txtReceiveVoucherNo
            .AccountType = Get_Code(cmbReceiveAccountType.Text)
            .Save
        End With
        
        MsgBox "Data Saved Succesfully", vbInformation, "IT Division, DNMIH"
        'TabControl_For_Form_Load
        
      Case 1

          With oPFPayment
            .Connstring = strCN.Connection_String
            .PFPaymentId = txtPaymentId
            .PaymentPurpose = Get_Code(cmbPurposeOfPayment.Text)
            .Description = txtPaymentDescription
            .PaymentDate = DTPPaymentDate
            .Amount = txtPaymentAmount
            .PaymentType = cmbPaymentType
            .CheckNo = txtPaymentCheckNo
            .AccountNo = cmbPaymentAccountNo
            .BankCode = Get_Code(cmbPaymentBank.Text)
            .VoucherNo = txtPaymentVoucherNo
            .AccountType = Get_Code(cmbPaymentAccountType.Text)
            .Save
            End With

            MsgBox "Data Saved Succesfully", vbInformation, "IT Division, DNMIH"
           ' TabControl_For_Form_Load

    Case 2
    
            With oMemberFund
            .Connstring = strCN.Connection_String
            .TrackId = txtTrackId
            .AccountType = Get_Code(cmbMFAccountType.Text)
            .AccountNo = txtMFAccountNo
            .BankCode = Get_Code(cmbMFBank.Text)
            .Amount = txtMFAmount
            .Save
            End With
        
        MsgBox "Data Saved Succesfully", vbInformation, "IT Division, DNMIH"
        

  End Select
Show_Data_PF_Form_Load

Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub

Private Sub DataGrid1_Click()
On Error GoTo Errdes
txtReceiveId = DataGrid1.Columns(0)
cmbSourceofFund.Text = DataGrid1.Columns(1) + "~" + DataGrid1.Columns(11)
txtReceiveDescription.Text = DataGrid1.Columns(2)
dtpReceiveDate.Value = DataGrid1.Columns(3)
txtReceiveAmount.Text = DataGrid1.Columns(4)
cmbReceiveType.Text = DataGrid1.Columns(5)
txtReceiveCheckNo.Text = DataGrid1.Columns(6)
cmbReceiveAccountNo.Text = DataGrid1.Columns(7)
cmbReceiveBank.Text = DataGrid1.Columns(8) + "~" + DataGrid1.Columns(13)
txtReceiveVoucherNo.Text = DataGrid1.Columns(9)
cmbReceiveAccountType.Text = DataGrid1.Columns(10) + "~" + DataGrid1.Columns(12)
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub

Private Sub DataGrid2_Click()
On Error GoTo Errdes
txtPaymentId = DataGrid2.Columns(0)
cmbPurposeOfPayment.Text = DataGrid2.Columns(1) + "~" + DataGrid2.Columns(11)
txtPaymentDescription.Text = DataGrid2.Columns(2)
DTPPaymentDate.Value = DataGrid2.Columns(3)
txtPaymentAmount.Text = DataGrid2.Columns(4)
cmbPaymentType.Text = DataGrid2.Columns(5)
txtPaymentCheckNo.Text = DataGrid2.Columns(6)
cmbPaymentAccountNo.Text = DataGrid2.Columns(7)
cmbPaymentBank.Text = DataGrid2.Columns(8) + "~" + DataGrid2.Columns(13)
txtPaymentVoucherNo.Text = DataGrid2.Columns(9)
cmbPaymentAccountType.Text = DataGrid2.Columns(10) + "~" + DataGrid2.Columns(12)
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub
Private Sub DataGrid3_Click()
On Error GoTo Errdes
txtTrackId = DataGrid3.Columns(0)
cmbMFAccountType.Text = DataGrid3.Columns(1) + "~" + DataGrid3.Columns(5)
txtMFAccountNo.Text = DataGrid3.Columns(2)
cmbMFBank.Text = DataGrid3.Columns(3) + "~" + DataGrid3.Columns(6)
txtMFAmount.Text = DataGrid3.Columns(4)
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub
Private Sub Form_Load()
On Error GoTo Errdes
    Screen_Position Me
    SSTab_Index = 0
    'Set_TabIndex
    get_Value_Into_PF_Payment_Purpose
    LOAD_PF_RECEIVE_TYPE Me
    LOAD_PF_PAYMENT_TYPE Me
    get_Value_Into_Account_type
    get_Value_Into_Bank_Name
    get_Value_Into_Account_No
    Dim cmd As New Command
    Dim conn1 As New Connection
    Dim RS As New Recordset
   
    conn1.ConnectionString = strCN.Connection_String
    conn1.Open
    cmd.ActiveConnection = conn1
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select SOURCE_ID,SOURCE_NAME from L_PF_SOURCEOF_FUND order by SOURCE_ID"
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, conn1, adOpenDynamic, adLockOptimistic
    
        If RS.RecordCount > 0 Then
            Do Until RS.EOF
            cmbSourceofFund.AddItem RS.Fields(1) & "~" & RS.Fields(0)
            RS.MoveNext
            Loop
            
        End If
        
    RS.Close
    conn1.Close
 SSTab1.Tab = 0
Show_Data_PF_Form_Load
Exit Sub
Errdes:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"

End Sub



Private Sub Show_Data_PF_Form_Load()
On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim myrs10 As New ADODB.Recordset
    
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
    
    If SSTab1.Tab = 0 Then
        cmd.CommandText = "Select A.PF_RECEIVE_ID as ReceiveId,LS.SOURCE_NAME as SourceOfFund," _
        & "A.Description,A.RECEIVE_DATE as ReceiveDate,A.RECEIVED_AMOUNT ReceiveAmount,A.RECEIVE_TYPE as ReceiveType," _
        & "A.CHECK_NO as CheckNo,A.ACCOUNT_NO as AccountNo,LB.BANK_NAME as BankName," _
        & "A.VOUCHER_NO AS VoucherNo,LA.TYPE_NAME AS TypeName,LS.SOURCE_ID AS SourceId," _
        & "LA.TYPE_ID AS TypeId,LB.BANK_ID AS BankId From PF_RECEIVE A ,L_PF_SOURCEOF_FUND LS," _
        & "L_BANK LB,L_ACCOUNT_TYPE LA Where A.SOURCE_OF_FUND = LS.SOURCE_ID AND A.ACCOUNT_TYPE=LA.TYPE_ID" _
        & " AND A.BANK_CODE=LB.BANK_ID order by A.PF_RECEIVE_ID"
        
    ElseIf SSTab1.Tab = 1 Then
         
        cmd.CommandText = "SELECT A.PF_PAYMENT_ID as PaymentID,LP.PURPOSE_NAME as PaymentPurpose," _
            & "A.Description,A.PAYMENT_DATE as PaymentDate,A.PAYMENT_AMOUNT as PaymentAmount," _
            & "A.PAYMENT_TYPE as PaymentType,A.CHECK_NO as CheckNo,A.ACCOUNT_NO as AccountNo," _
            & "LB.BANK_NAME as BankName,A.VOUCHER_NO AS VoucherNo,LA.TYPE_NAME AS TypeName," _
            & "LP.PURPOSE_ID as PurposeId,LA.TYPE_ID AS TypeId,LB.BANK_ID AS BankId" _
            & " From PF_PAYMENT A,L_PF_PAYMENT_PURPOSE LP,L_BANK LB,L_ACCOUNT_TYPE LA" _
            & " Where A.PURPOSE_OF_PAYMENT = LP.PURPOSE_ID AND A.ACCOUNT_TYPE=LA.TYPE_ID" _
            & " AND A.BANK_CODE=LB.BANK_ID order by A.PF_PAYMENT_ID"

        
    ElseIf SSTab1.Tab = 2 Then
        cmd.CommandText = "Select A.Track_Id as TrackId," _
        & "B.TYPE_NAME AS AccountType,A.ACCOUNT_NO as AccountNo," _
        & "LB.BANK_NAME as BankName,A.AMOUNT as Amount," _
        & "B.TYPE_ID as TypeId,LB.BANK_ID as BankId" _
        & " From MEMBER_FUND A,L_ACCOUNT_TYPE B,L_BANK LB" _
        & " Where A.ACCOUNT_TYPE = B.TYPE_ID AND A.BANK_CODE=LB.BANK_ID ORDER BY A.TRACK_ID"

    
    End If

    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs10.CursorLocation = adUseClient
    
    myrs10.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
    
    If SSTab1.Tab = 0 Then
        If Not (myrs10.BOF Or myrs10.EOF) Then
             Set DataGrid1.DataSource = myrs10
             
        End If
    ElseIf SSTab1.Tab = 1 Then
        If Not (myrs10.BOF Or myrs10.EOF) Then
                Set DataGrid2.DataSource = myrs10
        End If
        
    ElseIf SSTab1.Tab = 2 Then
        If Not (myrs10.BOF Or myrs10.EOF) Then
                Set DataGrid3.DataSource = myrs10
        End If
        
        
    End If
Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub get_Value_Into_PF_Payment_Purpose()
On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim RS As New ADODB.Recordset
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select PURPOSE_ID,PURPOSE_NAME from L_PF_PAYMENT_PURPOSE order by PURPOSE_ID"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
       If RS.RecordCount > 0 Then
            Do Until RS.EOF
            cmbPurposeOfPayment.AddItem RS.Fields(1) & "~" & RS.Fields(0)
            RS.MoveNext
            Loop
        End If

    
    
    Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub get_Value_Into_Account_type()
On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim RS As New ADODB.Recordset
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select TYPE_ID,TYPE_NAME from L_ACCOUNT_TYPE order by TYPE_ID"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
       If RS.RecordCount > 0 Then
            Do Until RS.EOF
            cmbMFAccountType.AddItem RS.Fields(1) & "~" & RS.Fields(0)
            cmbReceiveAccountType.AddItem RS.Fields(1) & "~" & RS.Fields(0)
            cmbPaymentAccountType.AddItem RS.Fields(1) & "~" & RS.Fields(0)
            RS.MoveNext
            Loop
            'cmbAccountType.ListIndex = 0
        End If

    
    
    Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub get_Value_Into_Bank_Name()
On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim RS As New ADODB.Recordset
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText
    
    cmd.CommandText = "select BANK_ID,BANK_NAME from L_BANK order by BANK_ID"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
       If RS.RecordCount > 0 Then
            Do Until RS.EOF
            cmbMFBank.AddItem RS.Fields(1) & "~" & RS.Fields(0)
            cmbReceiveBank.AddItem RS.Fields(1) & "~" & RS.Fields(0)
            cmbPaymentBank.AddItem RS.Fields(1) & "~" & RS.Fields(0)
            RS.MoveNext
            Loop
        End If

    
    
    Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
Private Sub get_Value_Into_Account_No()
On Error GoTo Errdes
    Dim getconnect As New Connection
    Dim cmd As New Command
    Dim RS As New ADODB.Recordset
    getconnect.ConnectionString = strCN.Connection_String
    getconnect.Open
    cmd.ActiveConnection = getconnect
    cmd.CommandType = adCmdText

    cmd.CommandText = "select ACCOUNT_NO from MEMBER_FUND"

    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, getconnect, adOpenDynamic, adLockOptimistic
       If RS.RecordCount > 0 Then
            Do Until RS.EOF
            cmbReceiveAccountNo.AddItem RS.Fields(0)
            cmbPaymentAccountNo.AddItem RS.Fields(0)
            RS.MoveNext
            Loop
        End If



    Exit Sub
Errdes:
 MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo Errdesc
    SSTab_Index = SSTab1.Tab
    'Set_TabIndex
    'get_Value_Into_Payment_Purpose
    'TabControl_For_Form_Load
    Show_Data_PF_Form_Load
    If SSTab_Index = 1 Then
    cmbMFBank.SetFocus
    End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "IT Division, DNMIH"
End Sub
