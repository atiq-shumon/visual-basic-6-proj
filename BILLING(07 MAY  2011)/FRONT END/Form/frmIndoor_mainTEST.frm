VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIndoor_mainTEST 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   -60
      TabIndex        =   50
      Top             =   6780
      Width           =   11385
      Begin VB.TextBox txtSerialNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   55
         Top             =   270
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.CommandButton CMDEXIT 
         Cancel          =   -1  'True
         Caption         =   "CLOSE"
         Height          =   375
         Left            =   9450
         TabIndex        =   54
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton CMDREPORT 
         Caption         =   "REPORT"
         Height          =   375
         Left            =   8220
         TabIndex        =   53
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdShowBed 
         Caption         =   "SHOW BED"
         Height          =   375
         Left            =   6990
         TabIndex        =   52
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdADD 
         Caption         =   "NEW"
         Height          =   375
         Left            =   5760
         TabIndex        =   51
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "SAVE"
         Height          =   375
         Left            =   4530
         TabIndex        =   15
         Top             =   210
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   -30
         Picture         =   "frmIndoor_mainTEST.frx":0000
         Stretch         =   -1  'True
         Top             =   -60
         Width           =   12675
      End
      Begin VB.Shape Shape1 
         Height          =   465
         Left            =   4440
         Top             =   150
         Width           =   6285
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1215
      Left            =   -90
      TabIndex        =   16
      Top             =   -180
      Width           =   11475
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CABIN"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   345
         Left            =   5010
         TabIndex        =   18
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PATIENT ADMISSION ENTRY"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   3420
         TabIndex        =   17
         Top             =   360
         Width           =   4665
      End
      Begin VB.Image Image1 
         Height          =   1155
         Left            =   30
         Picture         =   "frmIndoor_mainTEST.frx":5982
         Stretch         =   -1  'True
         Top             =   180
         Width           =   12180
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   5715
      Left            =   -90
      TabIndex        =   19
      Top             =   1050
      Width           =   11385
      Begin VB.ComboBox cboDMY 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmIndoor_mainTEST.frx":B304
         Left            =   1470
         List            =   "frmIndoor_mainTEST.frx":B311
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3270
         Width           =   885
      End
      Begin VB.ComboBox cboPatdept 
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
         ForeColor       =   &H000000FF&
         Height          =   315
         ItemData        =   "frmIndoor_mainTEST.frx":B31E
         Left            =   5550
         List            =   "frmIndoor_mainTEST.frx":B320
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3270
         Width           =   5235
      End
      Begin VB.ComboBox CBOYRCODE 
         Height          =   315
         ItemData        =   "frmIndoor_mainTEST.frx":B322
         Left            =   5610
         List            =   "frmIndoor_mainTEST.frx":B32C
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "YR-0708"
         Top             =   510
         Width           =   1365
      End
      Begin VB.TextBox txtAdvance 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   7350
         TabIndex        =   14
         Top             =   5190
         Width           =   3405
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   570
         TabIndex        =   13
         Top             =   5160
         Width           =   4785
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   600
         TabIndex        =   12
         Top             =   4140
         Width           =   10155
      End
      Begin VB.ComboBox cboReligion 
         Height          =   315
         ItemData        =   "frmIndoor_mainTEST.frx":B342
         Left            =   4350
         List            =   "frmIndoor_mainTEST.frx":B355
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3270
         Width           =   1005
      End
      Begin VB.ComboBox cboSex 
         Height          =   315
         ItemData        =   "frmIndoor_mainTEST.frx":B383
         Left            =   2670
         List            =   "frmIndoor_mainTEST.frx":B38D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3270
         Width           =   1695
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   570
         MaxLength       =   3
         TabIndex        =   7
         Top             =   3270
         Width           =   885
      End
      Begin VB.TextBox txtPatFatherName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5580
         TabIndex        =   6
         Top             =   2400
         Width           =   5205
      End
      Begin VB.CheckBox chkHusband 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Husband's Name"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   7620
         MaskColor       =   &H00FFFF80&
         TabIndex        =   37
         Top             =   2100
         Width           =   2715
      End
      Begin VB.CheckBox chkFather 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Father's Name"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   270
         Left            =   5580
         MaskColor       =   &H80000001&
         TabIndex        =   5
         Top             =   2100
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.TextBox txtPatName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   570
         TabIndex        =   4
         Top             =   2400
         Width           =   4785
      End
      Begin VB.ComboBox cboBedDept 
         Height          =   315
         ItemData        =   "frmIndoor_mainTEST.frx":B397
         Left            =   2670
         List            =   "frmIndoor_mainTEST.frx":B399
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1470
         Width           =   1695
      End
      Begin VB.TextBox txtServiceFee 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   9600
         TabIndex        =   33
         Top             =   1470
         Width           =   1185
      End
      Begin VB.TextBox txtAdmissionfee 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   7350
         TabIndex        =   31
         Top             =   1470
         Width           =   945
      End
      Begin VB.TextBox txtBedCharge 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   8550
         TabIndex        =   29
         Top             =   1470
         Width           =   825
      End
      Begin VB.ComboBox cboBedNo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmIndoor_mainTEST.frx":B39B
         Left            =   5610
         List            =   "frmIndoor_mainTEST.frx":B39D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1470
         Width           =   1365
      End
      Begin VB.ComboBox CboTypeNo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmIndoor_mainTEST.frx":B39F
         Left            =   4380
         List            =   "frmIndoor_mainTEST.frx":B3A1
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1470
         Width           =   975
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   7350
         TabIndex        =   25
         Top             =   510
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtRegNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2670
         TabIndex        =   23
         Top             =   510
         Width           =   2685
      End
      Begin VB.ComboBox cboBedType 
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
         ForeColor       =   &H000000FF&
         Height          =   315
         ItemData        =   "frmIndoor_mainTEST.frx":B3A3
         Left            =   570
         List            =   "frmIndoor_mainTEST.frx":B3B0
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1470
         Width           =   1845
      End
      Begin VB.TextBox txtRecNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   570
         TabIndex        =   20
         Top             =   510
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker7 
         Height          =   315
         Left            =   9540
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Delevary Time"
         Top             =   480
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   50593794
         UpDown          =   -1  'True
         CurrentDate     =   37163
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y/M/D"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   20
         Left            =   1530
         TabIndex        =   49
         Top             =   2940
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fiscal Year"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   19
         Left            =   5610
         TabIndex        =   48
         Top             =   210
         Width           =   1185
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   18
         Left            =   9540
         TabIndex        =   47
         Top             =   210
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Date"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   17
         Left            =   7380
         TabIndex        =   46
         Top             =   210
         Width           =   1710
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Department"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Index           =   16
         Left            =   5580
         TabIndex        =   45
         Top             =   2940
         Width           =   2745
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   15
         Left            =   7380
         TabIndex        =   43
         Top             =   4890
         Width           =   945
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   14
         Left            =   600
         TabIndex        =   42
         Top             =   4860
         Width           =   690
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Address"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Index           =   13
         Left            =   600
         TabIndex        =   41
         Top             =   3810
         Width           =   1785
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   12
         Left            =   4410
         TabIndex        =   40
         Top             =   2940
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   11
         Left            =   2700
         TabIndex        =   39
         Top             =   2940
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   10
         Left            =   600
         TabIndex        =   38
         Top             =   2940
         Width           =   405
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Index           =   9
         Left            =   600
         TabIndex        =   36
         Top             =   2100
         Width           =   2025
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Dept"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   8
         Left            =   2700
         TabIndex        =   35
         Top             =   1140
         Width           =   1290
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Srvc. Fee"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   7
         Left            =   9630
         TabIndex        =   34
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adm. Fee"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   6
         Left            =   7380
         TabIndex        =   32
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Fee"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   5
         Left            =   8550
         TabIndex        =   30
         Top             =   1140
         Width           =   825
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bed No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   4
         Left            =   5640
         TabIndex        =   28
         Top             =   1140
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Type No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   3
         Left            =   4350
         TabIndex        =   27
         Top             =   1140
         Width           =   945
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Index           =   2
         Left            =   2670
         TabIndex        =   24
         Top             =   210
         Width           =   1725
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Index           =   1
         Left            =   600
         TabIndex        =   22
         Top             =   210
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Type"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   0
         Left            =   630
         TabIndex        =   21
         Top             =   1140
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmIndoor_mainTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As New MyConnection
Dim Conn As New Connection
Public Connpr As New Connection
Dim Conn1 As New Connection
Dim Conn2 As New Connection
Dim Conn3 As New Connection
Dim Conn4 As New Connection
Dim Conn5 As New Connection
Dim Conn6 As New Connection
Dim cmd As New Command
Public cmdpr As New Command
Dim RS As New Recordset
Public rspr As New Recordset
Dim RS1 As New Recordset
Dim rs2 As New Recordset
Dim rs3 As New Recordset
Dim RS4 As New Recordset
Dim rs5 As New Recordset
Dim rs6 As New Recordset

Dim VoucherNumber

Public strUid As String
Public strcn        As New MyConnection
Private Sub Get_Voucher_Number()
On Error GoTo Errdesc
Dim conn10 As New Connection
Dim rs10 As New Recordset
Dim cmd As New Command
If conn10.State = 0 Then
conn10.ConnectionString = strcn.Connection_String
conn10.Open
End If
VoucherNumber = 0
cmd.ActiveConnection = conn10
cmd.CommandType = adCmdText
cmd.CommandText = "select max(acct.vou.vou_no)+1 from acct.vou where upper(acct.vou.vou_type)=upper('cr')"
rs10.CursorLocation = adUseClient
rs10.Open cmd.CommandText, conn10, adOpenDynamic, adLockOptimistic
    If rs10.RecordCount > 0 Then
        If IsNull(rs10.Fields(0)) Then
            VoucherNumber = 1
        Else
            VoucherNumber = rs10.Fields(0)
       End If
    Else
        VoucherNumber = 1
    End If
Exit Sub
If conn10.State = 1 Then
    conn10.Close
    Set conn10 = Nothing
End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "Dhaka National Medical Institute Hospital"
End Sub

Private Sub cboCabin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
cmdTypeNOCabin.SetFocus
End If

End Sub

Private Sub cboCabin_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
     If UCase(cboCabin.Text) = UCase("Common") Then
               MsgBox "Please Select a Doctor's Department", vbInformation, "Dhaka National Medical Institute Hospital"
               cboCabin.SetFocus
     End If
 End If
 
               
       
End Sub

Private Sub cboDepartmentPaying_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
cmdTypeNOPaying.SetFocus
End If

End Sub

Private Sub cboDepartmentPaying_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If UCase(cboDepartmentPaying.Text) = UCase("Common") Then
               MsgBox "Please Select a Doctor's Department", vbInformation, "Dhaka National Medical Institute Hospital"
               cboDepartmentPaying.SetFocus
     End If
 End If

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
Else
Check2.Enabled = True
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
'Check1.Enabled = False
Check1.Value = 0
Else
Check1.Enabled = True
End If

End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
Check6.Value = 0
Else
Check6.Enabled = True
End If

End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
Check5.Value = 0
Else
Check5.Enabled = True
End If

End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then
Check8.Value = 0
Else
Check8.Enabled = True
End If

End Sub

Private Sub Check8_Click()
If Check8.Value = 1 Then
Check7.Value = 0
Else
Check7.Enabled = True
End If

End Sub

Private Sub cboBedDept_Click()
  txtAdmissionfee = ""
  txtBedCharge = ""
  txtServiceFee = ""
  txtSerialNo = ""


  Call LOAD_BED_NO
  End Sub

Private Sub cboBedNo_Click()
  load_fee
End Sub
Private Sub load_fee()
   Dim Conn As New ADODB.Connection
   Dim RS As New ADODB.Recordset
   Dim cmd As New ADODB.Command
   
   If Conn.State = 0 Then
       Conn.ConnectionString = strcn.Connection_String
       Conn.Open
    End If
    cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "select bed_charge,service_charge,BED_GROUP,serial_no from bed_info  where occupy_flag='0'and upper(bed_type)=upper('" & CboBedType & "') and  BED_EXT_COL='" & Trim(CboTypeNo.Text) & "'and bed_no='" & Trim(cboBedNo.Text) & "' and doc_department='" & Trim(cboBedDept.Text) & "'"
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
   
   If RS.RecordCount > 0 Then
       txtAdmissionfee = RS!BED_GROUP
       txtBedCharge = RS!bed_CHARGE
       txtServiceFee = RS!service_charge
       txtSerialNo = RS!SERIAL_NO
    End If
      cmd.Properties("PLSQLRSet") = False
  
  
   If Conn.State = 1 Then
      Conn.Close
     Set Conn = Nothing
     Set RS = Nothing
     Set cmd = Nothing
   End If
End Sub
Private Sub CboBedType_Click()
  Label47.Caption = CboBedType.Text
  If CboBedType = "Cabin" And Len(cboBedDept) > 0 Then
     cboBedDept.Text = "Common"
     cboBedDept_Click
  End If
  CboTypeNo.clear
  cboBedNo.clear
  txtAdmissionfee = ""
  txtBedCharge = ""
  txtServiceFee = ""
  txtSerialNo = ""
End Sub

Private Sub cboPatdept_Click()
  load_service_charge_exceptional
  Gynae_validation
End Sub
Private Sub Gynae_validation()
  If UCase(cboPatdept.Text) = UCase("Gynae-1") Or UCase(cboPatdept.Text) = UCase("Gynae-2") Then
       If UCase(cboSex.Text) = UCase("M") Then
         cboSex.Text = "F"
       End If
    End If
End Sub
Private Sub load_service_charge_exceptional()
   If CboBedType = "Free-Bed" And cboBedDept = "COMMON" Then
     If cboPatdept = "Ophth." Or cboPatdept = "Gynae-1" Or cboPatdept = "Gynae-2" Or cboPatdept = "ENT" Or cboPatdept = "Surgery-1" Or cboPatdept = "Surgery-2" Then
        txtServiceFee = 250
     Else
       txtServiceFee = 0
     End If
   End If
End Sub

Private Sub chkFather_Click()
  If chkFather.Value = 1 Then
    chkHusband.Value = 0
    chkFather.ForeColor = &HFFFF80
    txtPatFatherName = "S/D/O:"
  Else
    chkHusband.Enabled = True
    chkHusband.ForeColor = &HFFFF80
    chkHusband.Value = 1
    txtPatFatherName = "W/O:"
  End If
End Sub

Private Sub chkFather_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 32 Then
     chkFather_Click
  End If
End Sub

Private Sub chkHusband_Click()
  If chkHusband.Value = 1 Then
     chkFather.Value = 0
     chkFather.ForeColor = vbWhite
  Else
    chkFather.Enabled = True
    chkFather.Value = 1
    chkHusband.ForeColor = vbWhite
    chkFather.ForeColor = &HFFFF80

   End If
End Sub
Private Sub chkHusband_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 32 Then
      chkHusband_Click
   End If
   If KeyCode = 13 Then
     txtPatName.SetFocus
   End If
End Sub
Private Sub cmdADD_Click()
  Call Clear_form
End Sub
Private Sub Clear_form()
txtPatFatherName = ""
CboBedType = CboBedType.List(0)
cboDMY = cboDMY.List(0)
cboSex = cboSex.List(0)
cboReligion = cboReligion.List(0)
'
MaskEdBox1.Text = Format(Date, "dd/mm/yyyy")
DTPicker7.Value = Time

Call show_dept
cboBedDept.Text = "Common"
txtPatFatherName = "S/D/O:"
cboPatdept = cboPatdept.List(0)

txtPatName = ""

txtAddress = ""

txtAdvance = 0
txtAge = ""

txtPhone = ""

CboBedType.SetFocus
End Sub



Private Sub cmdBedNoFree_Click()
  
    If Conn4.State = 0 Then
        Conn4.ConnectionString = strcn.Connection_String
        Conn4.Open
    End If
    cmd.ActiveConnection = Conn4
    cmd.CommandType = adCmdText
   cmd.CommandText = "select Bed_charge,serial_no from bed_info  where occupy_flag='0'and bed_type='Free-Bed' and  BED_EXT_COL='" & Trim(cmdTypeNoFree.Text) & "' and bed_no='" & Trim(cmdBedNoFree.Text) & "'and doc_department='" & Trim(comDepartmentFree.Text) & "'order by bed_no"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS4.CursorLocation = adUseClient
    
    RS4.Open cmd.CommandText, Conn4, adOpenDynamic, adLockOptimistic
   'cmdBedNoFree.clear
   cmd.Properties("iRowsetChange") = False
     
    If RS4.RecordCount > 0 Then
                  
            
           txtBedChargeFree.Text = RS4!bed_CHARGE
          txtSerial_no_free.Text = RS4!SERIAL_NO

    
    End If
    
    Set RS4 = Nothing
If Conn4.State = 1 Then
    Conn4.Close
    Set Conn4 = Nothing
 End If

End Sub


Private Sub cmdBedNOPaying_Click()

If Conn2.State = 0 Then
Conn2.ConnectionString = strcn.Connection_String
    Conn2.Open
End If
    cmd.ActiveConnection = Conn2
    cmd.CommandType = adCmdText
   cmd.CommandText = "select bed_charge,serial_no from bed_info  where occupy_flag='0'and upper(bed_type)=upper('Paying')and   bed_no='" & Trim(cmdBedNOPaying.Text) & "'and doc_department='" & Trim(cboDepartmentPaying.Text) & "'and  BED_EXT_COL='" & Trim(cmdTypeNOPaying.Text) & "'"
   
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    rs2.CursorLocation = adUseClient
    
    rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
'    cmdBedNOPaying.clear
   cmd.Properties("iRowsetChange") = False

    cmdBedNOPaying.Refresh
       
  If rs2.RecordCount > 0 Then
            
          txtChargePaying = rs2!bed_CHARGE
          txtSerial_no_paying = rs2!SERIAL_NO

 End If
     
    
    
    Set rs2 = Nothing
If Conn2.State = 1 Then
    Conn2.Close
    Set Conn2 = Nothing
End If

End Sub

Private Sub cmdBedNoSemiCabin_Change()
    
'   Conn5.ConnectionString = strcn.Connection_String
'    Conn5.Open
'    cmd.ActiveConnection = Conn5
'    cmd.CommandType = adCmdText
'   cmd.CommandText = "select bed_charge from bed_info  where occupy_flag='0'and bed_type='Semi-Cabin'and  bed_no='" & Trim(cmdTypeNOSemiCabin.Text) & "'"
'
'   cmd.Properties("iRowsetChange") = True
'   cmd.Properties("updatability") = 7
'    rs5.CursorLocation = adUseClient
'
'    rs5.Open cmd.CommandText, Conn5, adOpenDynamic, adLockOptimistic
'   'cmdBedNoSemiCabin.clear
'
'    If rs5.RecordCount > 0 Then
'       txtBedChargeSemiCabin = rs5!bed_charge
'    End If
'
'    rs5.Close
'
'    Conn5.Close

End Sub

Private Sub cmdBedNoSemiCabin_Click()
    
'   Conn6.ConnectionString = strcn.Connection_String
'    Conn6.Open
'    cmd.ActiveConnection = Conn6
'    cmd.CommandType = adCmdText
'   cmd.CommandText = "select bed_charge from bed_info  where occupy_flag='0'and bed_type='Semi-Cabin'and  bed_no='" & Trim(cmdBedNoSemiCabin.Text) & "'"
'
'   cmd.Properties("iRowsetChange") = True
'   cmd.Properties("updatability") = 7
'    rs6.CursorLocation = adUseClient
'
'    rs6.Open cmd.CommandText, Conn6, adOpenDynamic, adLockOptimistic
'   'cmdBedNoSemiCabin.clear
'
'    If rs6.RecordCount > 0 Then
'
'       txtBedChargeSemiCabin = rs6!bed_charge
'    End If
'
'    rs6.Close
'
'    Conn6.Close

End Sub

Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub
Private Sub load_bed_reg_for_showing_cabin()
 On Error GoTo ErrDes
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

       If Conn.State = 0 Then
      Conn.Open strcn.Connection_String
  End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText

   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL show_bed_no_cabin}"
    
    Debug.Print cmd.CommandText
        
        Set RS = cmd.Execute
    

 cmd.Properties("PLSQLRSet") = False
 If Conn.State = 1 Then
    Conn.Close
    Set Conn = Nothing
 End If
''      RS.Close
'
Exit Sub
ErrDes:
MsgBox Err.Description, vbCritical, "Dhaka National Medical Institute Hospital"
End Sub
Private Sub load_bed_reg_for_showing_PAYING()
       On Error GoTo ErrDes
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    If Conn.State = 0 Then
              Conn.Open strcn.Connection_String
    End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText

   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL show_bed_no_paying}"
    
    Debug.Print cmd.CommandText
        
        Set RS = cmd.Execute
    

 cmd.Properties("PLSQLRSet") = False
 If Conn.State = 1 Then
    Conn.Close
    Set Conn = Nothing
 End If
''      RS.Close
'
Exit Sub
ErrDes:
MsgBox Err.Description, vbCritical, "Dhaka National Medical Institute Hospital"
End Sub
Private Sub load_bed_reg_for_showing_FREE()
      On Error GoTo ErrDes
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    If Conn.State = 0 Then
              Conn.Open strcn.Connection_String
    End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText

   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL show_bed_no_free}"
    
    Debug.Print cmd.CommandText
        
        Set RS = cmd.Execute
    

 cmd.Properties("PLSQLRSet") = False
 If Conn.State = 1 Then
    Conn.Close
    Set Conn = Nothing
 End If
''      RS.Close
'
Exit Sub
ErrDes:
MsgBox Err.Description, vbCritical, "Dhaka National Medical Institute Hospital"
End Sub
Private Sub cmdPreview_Click()

On Error GoTo Errdesc
Dim f2 As New frmDataSelect
Dim getconnected As New Connection
Dim cmd As New Command
Dim myrs As New ADODB.Recordset

    Call load_bed_reg_for_showing_cabin
  If getconnected.State = 0 Then
    getconnected.ConnectionString = strcn.Connection_String
    getconnected.Open
  End If
    cmd.ActiveConnection = getconnected
    cmd.CommandType = adCmdText
    cmd.CommandText = " SELECT bed_no,bed_ward, in_reg_no,name from show_bed "
                    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs.CursorLocation = adUseClient
     
    myrs.Open cmd.CommandText, getconnected, adOpenDynamic, adLockOptimistic
     Set f2.adoRecordset = myrs
     Set f2.OwnerForm = Me
     f2.Width = 7500
     f2.grdDataGrid.Columns(0).Caption = "Bed No"
       
     f2.grdDataGrid.Columns(1).Caption = "Bed"
     f2.grdDataGrid.Columns(3).Caption = "Name"
     f2.grdDataGrid.Columns(2).Caption = "Registration No"
     
     f2.grdDataGrid.Columns(0).Width = 800
     f2.grdDataGrid.Columns(1).Width = 800
     f2.grdDataGrid.Columns(2).Width = 3500
     f2.grdDataGrid.Columns(3).Width = 3000
     
  cmd.Properties("iRowsetChange") = False
   
     f2.Show 1
'     Combo1(1) = myrs.Fields(0)
     'txtname = myrs.Fields(1)




'rptMode = 6
'Viewer.Show vbModal

'rptMode = 6
'Viewer.Show vbModal
''MsgBox rptMode
If getconnected.State = 1 Then
        getconnected.Close
        Set getconnected = Nothing
      'Set getconnected = Nothing
End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "Dhaka National Medical Institute Hospital"

End Sub

Private Sub cmdPrint_Click()
  If TXTPRINT_PREV_CABIN.Visible = False Then
       TXTPRINT_PREV_CABIN.Visible = True
   End If
  TXTPRINT_PREV_CABIN.ForeColor = vbBlue
  PREVIEW_VAR = Val(TXTPRINT_PREV_CABIN)
  
   If TXTPRINT_PREV_CABIN = "" Then
    TXTPRINT_PREV_CABIN.SetFocus
     Exit Sub
   Else
    PREVIEW_VAR = Val(TXTPRINT_PREV_CABIN)
        rptMode = 6
      Viewer.Show vbModal
      TXTPRINT_PREV_CABIN = ""
      TXTPRINT_PREV_CABIN.Visible = False
End If
   
End Sub

Private Sub cmdSave_Click()
              Dim validation As Variant
              Adodc1.ConnectionString = strcn.Connection_String
              Adodc1.RecordSource = "Select user_id, user_name, user_password,user_type,shift_name From Security Where (User_Id = '" & frmMAIN.lbluser_id & "')"
              Adodc1.Refresh
              validation = Adodc1.Recordset!user_id
                
                Dim Conn As New ADODB.Connection
                Dim cmd As New ADODB.Command
                Dim RS As New ADODB.Recordset

                        Dim Param1 As New Parameter
                   If Conn.State = 0 Then
                        Conn.Open strcn.Connection_String
                    End If
                      Set cmd.ActiveConnection = Conn
                    cmd.CommandType = adCmdText
    
                   Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, validation)
                    cmd.Parameters.Append Param1 'validation
                    cmd.Properties("PLSQLRSet") = True
    
                     cmd.CommandText = "{CALL shift_validation(?)}"
    
                Debug.Print cmd.CommandText
    
                    Set RS = cmd.Execute
    

                cmd.Properties("PLSQLRSet") = False
               
                
          Adodc2.ConnectionString = strcn.Connection_String
          Adodc2.RecordSource = "Select * From user_validation"
          Adodc2.Refresh
        

        
             If Adodc2.Recordset!validation = 0 Then
             MsgBox "Your Working Time has been Expired", vbInformation, "Dhaka National Medical Institute Hospital."
             Exit Sub
             End If
             
           If Len(cboBedDept) = 0 Then
              MsgBox "Bed Department Required", vbInformation, "Warning..."
              cboBedDept.SetFocus
              Exit Sub
           End If
           
           If Len(CboTypeNo) = 0 Then
              MsgBox "Bed Type Required", vbInformation, "Warning..."
              CboTypeNo.SetFocus
              Exit Sub
           End If
           
           If Len(cboBedNo) = 0 Then
              MsgBox "Bed No Required", vbInformation, "Warning..."
              cboBedNo.SetFocus
              Exit Sub
           End If
        
          If Len(txtAdmissionfee) = 0 Then
            MsgBox "Bed No Required", vbInformation, "Warning..."
              cboBedNo.SetFocus
              Exit Sub
           End If
          End If
           
          If txtPatName = "" Then
            MsgBox "Patient Name Required", vbInformation, "Warning..."
            txtPatName.SetFocus
            Exit Sub
          End If
          
          If txtPatFatherName = "" Then
            MsgBox "Guardian Name Required", vbInformation, "Warning..."
            txtPatFatherName.SetFocus
            Exit Sub
          End If
          
          If Len(txtAge) = 0 Then
            MsgBox "Age Required", vbInformation, "Warning..."
            txtAge.SetFocus
            Exit Sub
          End If
          
          If Len(cboPatdept) = 0 Then
            MsgBox "Patient Department Required", vbInformation, "Warning..."
            cboPatdept.SetFocus
            Exit Sub
          End If
          
          If Len(txtAddress) = 0 Then
            MsgBox "Patient Address Required", vbInformation, "Warning..."
            txtAddress.SetFocus
            Exit Sub
          End If
          
          If Len(txtAdvance) = 0 Then
             MsgBox "Advance Required", vbInformation, "Warning..."
             txtAdvance.SetFocus
             txtAdvance = 0
             Exit Sub
          End If
          
          
          
      If UCase(cboPatdept.Text) = UCase("Common") Then
          MsgBox "Please Select a Patient Department", vbInformation, "Dhaka National Medical Institute Hospital"
          cboPatdept.SetFocus
          Exit Sub
      End If
   
     If UCase(cboPatdept.Text) = UCase("Gynae-1") Or UCase(cboPatdept.Text) = UCase("Gynae-2") Then
       If UCase(cboSex.Text) = UCase("M") Then
           MsgBox "Please Select Sex as Female for Gynae Patient", vbInformation, "Dhaka National Medical Institute Hospital"
           cboSex.SetFocus
         Exit Sub
       End If
    End If
'' Get_Voucher_Number
   Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "SELECT occupy_flag AS REC_NO FROM bed_info  where serial_no='" & Trim(txtSerail_no_cabin.Text) & "'"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
        If Adodc3.Recordset!REC_NO = "1" Then
        
        MsgBox "This Bed Has Recently Occupied by a new Patient....select a new bed ", vbInformation, "Dhaka National Medical Institute Hospital"
       Exit Sub
       End If
    End If
   
Call save_In_door_pat_cabin
'Call acct_integration_for_cabin
'Call acct_integration_for_cabin1
'Call post_vou

MsgBox "Operation successful", vbInformation + vbOKOnly, "Save..."
print_admission
Call Clear_form

If Conn.State = 1 Then
 Conn.Close
 Set Conn = Nothing
End If
Set RS = Nothing
Set cmd = Nothing
End Sub
Private Sub acct_integration_for_cabin()
   Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    
 If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
     Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, " advance Collection in cash")
    cmd.Parameters.Append Param1 'comment
                                                                    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, "2101")
    cmd.Parameters.Append Param2 ''USER_acct
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, nbrAdv.Text)
        cmd.Parameters.Append Param3 'dr
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, 0)
        cmd.Parameters.Append Param4 'cr
   
     Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 100, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param5 'u_id
     Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 100, VoucherNumber)
    cmd.Parameters.Append Param6 'u_id
         
            
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL acct.Save_vou_bill(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
                Conn.Close
       Set Conn = Nothing
    End If
  Set RS = Nothing
End Sub
Private Sub acct_integration_for_cabin1()
   Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    
 If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
     Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, " advance Collection in cash")
    cmd.Parameters.Append Param1 'comment
                                                                    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, "6101003")
    cmd.Parameters.Append Param2 ''USER_acct
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, 0)
        cmd.Parameters.Append Param3 'dr
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, nbrAdv.Text)
        cmd.Parameters.Append Param4 'cr
   
     Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 100, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param5 'u_id
     Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 100, VoucherNumber)
    cmd.Parameters.Append Param6 'u_id
         
            
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL acct.Save_vou_bill(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
                Conn.Close
       Set Conn = Nothing
    End If
  Set RS = Nothing
End Sub

Private Sub save_In_door_pat_cabin()
On Error GoTo ErrDes
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
    Dim Param11 As New Parameter
    Dim Param12 As New Parameter
    Dim Param13 As New Parameter
    Dim Param14 As New Parameter
    Dim Param15 As New Parameter
    Dim param16 As New Parameter
    Dim Param17 As New Parameter
    Dim Param18 As New Parameter
    Dim Param19 As New Parameter
    Dim Param20 As New Parameter
    Dim Param21 As New Parameter
    Dim Param22 As New Parameter
    Dim Param23 As New Parameter
    
   If Conn.State = 0 Then
      Conn.Open strcn.Connection_String
  End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, Trim(CboBedType.Text))
    cmd.Parameters.Append Param1 'bed type
    
    Set Param2 = cmd.CreateParameter("param2", adInteger, adParamInput, 10, Trim(CboTypeNo.Text))
    cmd.Parameters.Append Param2 'Type_no
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, cboBedNo.Text)
    cmd.Parameters.Append Param3 'Bed_no

    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 45, Trim(txtPatName.Text))
    cmd.Parameters.Append Param4 'patient_name

    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 50, Trim(txtPatFatherName.Text))
    cmd.Parameters.Append Param5 'guardgian_name
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 200, Trim(txtAddress.Text))
    cmd.Parameters.Append Param6 'pat_addr1
    
    
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 5, Trim(txtAge.Text))
    cmd.Parameters.Append Param7 'age
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 2, Trim(cboDMY.Text))
    cmd.Parameters.Append Param8 'Y_M_d
   
    
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 6, Trim(cboSex.Text))
    
    cmd.Parameters.Append Param9 'Sex
    
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 10, Trim(cboReligion.Text))
    cmd.Parameters.Append Param10 'Religion
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 25, Trim(txtPhone.Text))
    cmd.Parameters.Append Param11 'phone
    
    
    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 10, Trim(txtAdvance.Text))
    cmd.Parameters.Append Param12 'Advance
    
    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 10, Trim(cboPatdept.Text))
    cmd.Parameters.Append Param13 'Department
   
    
    Set Param14 = cmd.CreateParameter("param14", adVarChar, adParamInput, 6, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param14 'u_id
    
   
    
    Set Param15 = cmd.CreateParameter("param15", adVarChar, adParamInput, 5, frmMAIN.lblBooth)
    cmd.Parameters.Append Param15 'booth
    
    
    Set param16 = cmd.CreateParameter("param16", adVarChar, adParamInput, 5, 1)
    cmd.Parameters.Append param16 'extra-flag
     
    Set Param17 = cmd.CreateParameter("param17", adDouble, adParamInput, 10, txtSerialNo.Text)
    cmd.Parameters.Append Param17 'serial_no

    
    Set Param18 = cmd.CreateParameter("param18", adInteger, adParamInput, 10, chkFather.Value)
    cmd.Parameters.Append Param18 'check_value----father's name or husband name
    
    Set Param19 = cmd.CreateParameter("param19", adVarChar, adParamInput, 10, Trim(CBOYRCODE))
    cmd.Parameters.Append Param19 'YR_CODE
    
    Set Param20 = cmd.CreateParameter("param20", adDouble, adParamInput, 10, Trim(txtAdmissionfee.Text))
    cmd.Parameters.Append Param20 'ADMISSION CHARGE
 
    Set Param21 = cmd.CreateParameter("param21", adDouble, adParamInput, 10, Trim(txtBedCharge.Text))
    cmd.Parameters.Append Param21 'BED CHARGE
 
   Set Param22 = cmd.CreateParameter("param22", adDouble, adParamInput, 10, Trim(txtServiceFee.Text))
   cmd.Parameters.Append Param22 'SERVICE CHARGE
 
   If cboBedDept = "COMMON" And CboBedType <> "Cabin" Then
      Set Param23 = cmd.CreateParameter("param23", adDouble, adParamInput, 5, 1)
      cmd.Parameters.Append Param23 'EXTRA BED FLAG INDICATOR
   Else
      Set Param23 = cmd.CreateParameter("param23", adDouble, adParamInput, 5, 0)
      cmd.Parameters.Append Param23 'EXTRA BED FLAG INDICATOR
   
   End If
   
   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Indoor_SavePatient_info(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
        
        Set RS = cmd.Execute
    

 cmd.Properties("PLSQLRSet") = False
If Conn.State = 1 Then
       Conn.Close
       Set Conn = Nothing
End If
Set RS = Nothing
'      RS.Close
''''''Exit Sub
   txtRecNo.Locked = True
   Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "SELECT MAX(REC_NO) AS REC_NO FROM RECEIPT_NO_COUNTER"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
        txtRecNo = Adodc3.Recordset!REC_NO
        PREVIEW_VAR = Val(TXTPRINT_PREV_CABIN)
    End If
  
Exit Sub
ErrDes:
MsgBox Err.Description, vbCritical, "Dhaka National Medical Institute Hospital"
End Sub
Private Sub CboTypeNo_Click()
    Call load_bed
    load_fee
End Sub
Private Sub load_bed()
    Dim Conn As New ADODB.Connection
    Dim RS   As New ADODB.Recordset
    Dim cmd  As New ADODB.Command
    
    If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
  End If
    cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "select bed_no from bed_info  where occupy_flag='0' and  BED_EXT_COL='" & Trim(CboTypeNo.Text) & "'and doc_department='" & Trim(cboBedDept.Text) & "' and Upper(bed_type)=upper('" & CboBedType & "') order by bed_no"
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    
    RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
    cboBedNo.clear
    
    If RS.RecordCount > 0 Then
         RS.MoveFirst
         Do Until RS.EOF = True
           cboBedNo.AddItem RS.Fields(0)
           RS.MoveNext
        Loop
       cboBedNo.Text = cboBedNo.List(0)
    End If
    
     cmd.Properties("PLSQLRSet") = False
    Set RS = Nothing
    If Conn.State = 1 Then
       Conn.Close
       Set Conn = Nothing
    End If
    txtSerialNo = ""
    txtAdmissionfee = ""
    txtBedCharge = ""
    txtServiceFee = ""
    
End Sub
Private Sub cmdTypeNOCabin_GotFocus()
    If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
    End If

    cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
   cmd.CommandText = "select distinct(BED_EXT_COL) from bed_info  where  bed_type= 'Cabin'and doc_department='" & Trim(cboCabin.Text) & "'"
    
    cmd.Properties("iRowsetChange") = True
   cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    
    RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
    'cmdTypeNOSemiCabin
    cmdTypeNOCabin.clear
    
    If RS.RecordCount > 0 Then
            
         
         RS.MoveFirst
         
        Do Until RS.EOF = True
            cmdTypeNOCabin.AddItem RS.Fields(0)
            'Combo4.AddItem RS.Fields(1)
            RS.MoveNext
        Loop
    
    End If
  cmd.Properties("iRowsetChange") = False
    
    'RS.Close
    
If Conn.State = 1 Then
        Conn.Close
        Set Conn = Nothing
    End If
End Sub



Private Sub cmdTypeNOCabin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
cboCabin.SetFocus
End If

End Sub

Private Sub cmdTypeNoFree_Click()
    
   
    If Conn2.State = 0 Then
         Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
    End If

    cmd.ActiveConnection = Conn2
    cmd.CommandType = adCmdText
   cmd.CommandText = "select bed_no from bed_info  where occupy_flag='0'and upper(bed_type)=upper('Free-Bed') and  BED_EXT_COL='" & Trim(cmdTypeNoFree.Text) & "'and doc_department='" & Trim(comDepartmentFree.Text) & "'order by bed_no "
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    rs2.CursorLocation = adUseClient
    
    rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
   cmdBedNoFree.clear
    
    If rs2.RecordCount > 0 Then
         cmdBedNoFree.clear
         rs2.MoveFirst
         
        Do Until rs2.EOF = True
        
            
            cmdBedNoFree.AddItem rs2!bed_no
            rs2.MoveNext
        Loop
    
    End If
    txtBedChargeFree = ""
     txtSerial_no_free = ""
    cmd.Properties("iRowsetChange") = False
    'rs2.Close

    If Conn2.State = 1 Then
        Conn2.Close
        Set Conn2 = Nothing
    End If

End Sub




Private Sub cmdTypeNoFree_GotFocus()
    If Conn1.State = 0 Then
      Conn1.ConnectionString = strcn.Connection_String
        Conn1.Open
    End If

    cmd.ActiveConnection = Conn1
    cmd.CommandType = adCmdText
   cmd.CommandText = "select distinct(BED_EXT_COL) from bed_info  where   upper(bed_type)= upper('Free-Bed')and upper(doc_department)=upper('" & Trim(comDepartmentFree.Text) & "')"
    
    cmd.Properties("iRowsetChange") = True
   cmd.Properties("updatability") = 7
    RS1.CursorLocation = adUseClient
    
    RS1.Open cmd.CommandText, Conn1, adOpenDynamic, adLockOptimistic
    
        cmdTypeNoFree.clear
    If RS1.RecordCount > 0 Then
            
          cmdTypeNoFree.clear

         RS1.MoveFirst
         
        Do Until RS1.EOF = True
        
            cmdTypeNoFree.AddItem RS1.Fields(0)
    
            RS1.MoveNext
        Loop
    
    End If
        cmd.Properties("iRowsetChange") = False
  
      'RS1.Close
    If Conn1.State = 1 Then
        Conn1.Close
        Set Conn1 = Nothing
    End If

End Sub





Private Sub cmdTypeNoFree_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
comDepartmentFree.SetFocus
End If

End Sub

Private Sub cmdTypeNOPaying_Click()
  
    If Conn2.State = 0 Then
        Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
    End If

     cmd.ActiveConnection = Conn2
     cmd.CommandType = adCmdText
      cmd.ActiveConnection = Conn2
      cmd.CommandType = adCmdText
     cmd.CommandText = "select bed_no from bed_info  where occupy_flag='0'and upper(bed_type)=upper('Paying') and  BED_EXT_COL='" & Trim(cmdTypeNOPaying.Text) & "'and doc_department='" & Trim(cboDepartmentPaying.Text) & "'"
    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    rs2.CursorLocation = adUseClient
    
    rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
    cmdBedNOPaying.Refresh
    cmdBedNOPaying.clear
    If rs2.RecordCount > 0 Then
            
          'cmdBedNOPaying.clear
          
         rs2.MoveFirst
         
        Do Until rs2.EOF = True
        
            
            cmdBedNOPaying.AddItem rs2!bed_no
            rs2.MoveNext
        Loop
    
    End If
    
    txtSerial_no_paying = ""
     txtChargePaying = ""
    cmdBedNOPaying.Refresh
     cmd.Properties("PLSQLRSet") = False
    ''rs2.Close

    If Conn2.State = 1 Then
        Conn2.Close
        Set Conn2 = Nothing
    End If


End Sub

Private Sub cmdTypeNOPaying_GotFocus()
    If Conn1.State = 0 Then
      Conn1.ConnectionString = strcn.Connection_String
        Conn1.Open
    End If

    cmd.ActiveConnection = Conn1
    cmd.CommandType = adCmdText
   cmd.CommandText = "select distinct(BED_EXT_COL) from bed_info  where   bed_type= 'Paying'and doc_department='" & Trim(cboDepartmentPaying.Text) & "'"
    
  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
    RS1.CursorLocation = adUseClient
    
    RS1.Open cmd.CommandText, Conn1, adOpenDynamic, adLockOptimistic
       cmdTypeNOPaying.clear
    
    If RS1.RecordCount > 0 Then
            
          
        cmdTypeNOPaying.clear
         RS1.MoveFirst
            
        Do Until RS1.EOF = True
        
            cmdTypeNOPaying.AddItem RS1.Fields(0)
    
            RS1.MoveNext
        Loop
    
    End If
     cmd.Properties("iRowsetChange") = False
    
    Set RS1 = Nothing
   If Conn1.State = 1 Then
        Conn1.Close
        Set Conn1 = Nothing
    End If

End Sub



Private Sub cmdTypeNOPaying_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
cboDepartmentPaying.SetFocus
End If

End Sub

Private Sub cmdTypeNOSemiCabin_Click()

'     Conn2.ConnectionString = strcn.Connection_String
'    Conn2.Open
'    cmd.ActiveConnection = Conn2
'    cmd.CommandType = adCmdText
'   cmd.CommandText = "select bed_no from bed_info  where occupy_flag='0'and bed_type='Semi-Cabin'and  BED_EXT_COL='" & Trim(cmdTypeNOSemiCabin.Text) & "'"
'
'    cmd.Properties("iRowsetChange") = True
'    cmd.Properties("updatability") = 7
'    rs2.CursorLocation = adUseClient
'
'    rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
'   cmdBedNoSemiCabin.clear
'
'    If rs2.RecordCount > 0 Then
'              rs2.MoveFirst
'               Do Until rs2.EOF = True
'
'            cmdBedNoSemiCabin.AddItem rs2!bed_no
'            rs2.MoveNext
'        Loop
'  '   cmdBedNoSemiCabin = cmdBedNoSemiCabin.List(0)
'    End If
'
'
'    rs2.Close
'
'    Conn2.Close


End Sub

Private Sub cmdTypeNOSemiCabin_GotFocus()
'     Conn.ConnectionString = strcn.Connection_String
'    Conn.Open
'    cmd.ActiveConnection = Conn
'    cmd.CommandType = adCmdText
'   cmd.CommandText = "select distinct(BED_EXT_COL) from bed_info  where  bed_type= 'Cabin'and doc_department='" & Trim(cboCabin.Text) & "'"
'
'    cmd.Properties("iRowsetChange") = True
'   cmd.Properties("updatability") = 7
'    RS.CursorLocation = adUseClient
'
'    RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
'    'cmdTypeNOSemiCabin
'    cmdTypeNOCabin.clear
'
'    If RS.RecordCount > 0 Then
'
'
'         RS.MoveFirst
'
'        Do Until RS.EOF = True
'            cmdTypeNOCabin.AddItem RS.Fields(0)
'            'Combo4.AddItem RS.Fields(1)
'            RS.MoveNext
'        Loop
'
'    End If
'
'    RS.Close
'
'    Conn.Close
End Sub



Private Sub Combo4_Click()
 
    If Conn2.State = 0 Then
       Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
    End If
    cmd.ActiveConnection = Conn2
    cmd.CommandType = adCmdText
    
     cmd.CommandText = "select bed_charge,serial_no from bed_info  where occupy_flag='0'and upper(bed_type)=upper('Cabin') and  BED_EXT_COL='" & Trim(cmdTypeNOCabin.Text) & "'and bed_no='" & Trim(Combo4.Text) & "' and doc_department='" & Trim(cboCabin.Text) & "'"
    
   cmd.Properties("iRowsetChange") = True
   cmd.Properties("updatability") = 7
    rs2.CursorLocation = adUseClient
    
    rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
    Combo4.Refresh
    'Combo4.clear
    If rs2.RecordCount > 0 Then
            
       txtBedChargeCabin = rs2!bed_CHARGE
       txtSerail_no_cabin = rs2!SERIAL_NO
    End If
    Combo4.Refresh
     cmd.Properties("PLSQLRSet") = False
    'rs2.Close

    If Conn2.State = 1 Then
        Conn2.Close
        Set Conn2 = Nothing
        End If


End Sub

Private Sub Combo4_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeySpace Then
            cmdTypeNOCabin.SetFocus
 End If
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
         If Combo4 = "" Then
           MsgBox "Please select a Bed  No", vbInformation, "Dhaka National Medical Institute Hospital"
           Combo4.SetFocus
         End If
  End If
  
End Sub

Private Sub comDepartmentFree_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
      cmdTypeNoFree.SetFocus
End If

End Sub

Private Sub comDepartmentFree_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If UCase(comDepartmentFree.Text) = UCase("Common") Then
               MsgBox "Please Select a Doctor's Department", vbInformation, "Dhaka National Medical Institute Hospital"
               comDepartmentFree.SetFocus
     End If
 End If

End Sub

Private Sub Command10_Click()
Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub



Private Sub Command12_Click()
Call clear_form_free
End Sub
Private Sub clear_form_free()
TxtAgeFree = ""
txtNameFree = ""
txtGuardNameFree = ""
TxtAddr2Free = ""
txtAddrFree = ""
txtPhoneFree = ""
txtAdvanceFree = ""

End Sub

Private Sub Command13_Click()
On Error GoTo Errdesc
Dim f2 As New frmDataSelect
Dim getconnected As New Connection
Dim cmd As New Command
Dim myrs As New ADODB.Recordset

    Call load_bed_reg_for_showing_FREE
  If getconnected.State = 0 Then
    getconnected.ConnectionString = strcn.Connection_String
    getconnected.Open
  End If
    cmd.ActiveConnection = getconnected
    cmd.CommandType = adCmdText
    cmd.CommandText = " SELECT bed_ward,doc_department,bed_no,in_reg_no,name from show_bed order by bed_ward,doc_department "
                    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs.CursorLocation = adUseClient
     
    myrs.Open cmd.CommandText, getconnected, adOpenDynamic, adLockOptimistic
     Set f2.adoRecordset = myrs
     Set f2.OwnerForm = Me
     f2.Width = 8000
     f2.grdDataGrid.Columns(0).Caption = "Ward No"
     f2.grdDataGrid.Columns(1).Caption = "Department"
     f2.grdDataGrid.Columns(2).Caption = "Bed No"
     f2.grdDataGrid.Columns(3).Caption = "Registration No"
     f2.grdDataGrid.Columns(4).Caption = "Name"
     
     f2.grdDataGrid.Columns(0).Width = 500
     f2.grdDataGrid.Columns(1).Width = 1400
     f2.grdDataGrid.Columns(2).Width = 600
     f2.grdDataGrid.Columns(3).Width = 2500
     f2.grdDataGrid.Columns(4).Width = 3500
    
  cmd.Properties("iRowsetChange") = False
   
     f2.Show 1
'     Combo1(1) = myrs.Fields(0)
     'txtname = myrs.Fields(1)
'rptMode = 6
'Viewer.Show vbModal

'rptMode = 6
'Viewer.Show vbModal
''MsgBox rptMode
If getconnected.State = 1 Then
        getconnected.Close
        Set getconnected = Nothing
      'Set getconnected = Nothing
End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "Dhaka National Medical Institute Hospital"

End Sub


Private Sub Command14_Click()

Dim validation As Variant
              Adodc1.ConnectionString = strcn.Connection_String
              Adodc1.RecordSource = "Select user_id, user_name, user_password,user_type,shift_name From Security Where (User_Id = '" & frmMAIN.lbluser_id & "')"
              Adodc1.Refresh
              validation = Adodc1.Recordset!user_id
                
                Dim Connfr As New ADODB.Connection
                Dim cmd As New ADODB.Command
                Dim RSfr As New ADODB.Recordset

                    Dim Param1 As New Parameter
                    If Connfr.State = 0 Then
                        Connfr.Open strcn.Connection_String
                    End If
    
                    Set cmd.ActiveConnection = Connfr
                    cmd.CommandType = adCmdText
    
                   Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, validation)
                    cmd.Parameters.Append Param1 'validation
                    cmd.Properties("PLSQLRSet") = True
    
                     cmd.CommandText = "{CALL shift_validation(?)}"
    
                Debug.Print cmd.CommandText
    
                    Set RSfr = cmd.Execute
    

                cmd.Properties("PLSQLRSet") = False
'                Connfr.Close
''                RSfr.Close
                
          Adodc2.ConnectionString = strcn.Connection_String
          Adodc2.RecordSource = "Select * From user_validation"
          Adodc2.Refresh
        

        
             If Adodc2.Recordset!validation = 0 Then
             MsgBox "Your Working Time has been Expired", vbInformation, "Dhaka National Medical Institute Hospital."
             Exit Sub
             End If
             





If UCase(comDepartmentFree.Text) = UCase("Common") Then
               MsgBox "Please Select a Doctor's Department", vbInformation, "Dhaka National Medical Institute Hospital"
               comDepartmentFree.SetFocus
       Exit Sub
     End If
 



If txtNameFree = "" Then
MsgBox "Patient Name Required", vbInformation, "Warning..."
txtNameFree.SetFocus
Exit Sub
End If
If TxtAddr2Free = "" Then
MsgBox "Address Required", vbInformation, "Warning..."
TxtAddr2Free.SetFocus
Exit Sub
End If
If txtAdvanceFree = "" Then
MsgBox "Advance Required", vbInformation, "Warning..."
txtAdvanceFree.SetFocus
txtAdvanceFree = 0
Exit Sub

End If

If TxtAgeFree = "" Then
MsgBox "Age Required", vbInformation, "Warning..."
TxtAgeFree.SetFocus
Exit Sub
End If
If UCase(comDepartmentFree.Text) = UCase("Gynae-1") Or UCase(comDepartmentFree.Text) = UCase("Gynae-2") Then
    If UCase(comSexFree.Text) = UCase("Male") Then
   MsgBox "Please Select Sex as Female for Gynae Patient", vbInformation, "Dhaka National Medical Institute Hospital"
        comSexFree.SetFocus
         Exit Sub
  End If
End If

'Get_Voucher_Number'
Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "SELECT occupy_flag AS REC_NO FROM bed_info  where serial_no='" & Trim(txtSerial_no_free) & "'"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
        If Adodc3.Recordset!REC_NO = "1" Then
        
        MsgBox "This Bed Has Recently Occupied by a new Patient....select a new bed ", vbInformation, "Dhaka National Medical Institute Hospital"
       Exit Sub
       End If
    End If

Call save_In_door_pat_free
'Call acct_integration_free_bed_advace
'Call acct_integration_free_bed_advace1
'Call post_vou

MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
 Call clear_form_free
' cmdTypeNoFree.Refresh
' cmdTypeNoFree.clear
' cmdBedNoFree.Refresh
' cmdBedNoFree.clear
'
' rptMode = 6
'Viewer.Show vbModal
print_admission
 cmdTypeNoFree.Refresh
 cmdTypeNoFree.clear
 cmdBedNoFree.Refresh
 cmdBedNoFree.clear
comDepartmentFree.SetFocus
If Connfr.State = 1 Then
    Connfr.Close
    Set Connfr = Nothing
End If
'RSfr.Close
 
End Sub
Private Sub post_vou()
       Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    
 
 If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
      Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, VoucherNumber)
    cmd.Parameters.Append Param1 'u_id
       Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 100, "CR")
    cmd.Parameters.Append Param2 'comment
       
            
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL acct.postvou(?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
       Conn.Close
       Set Conn = Nothing
    End If
  Set RS = Nothing
   End Sub
Private Sub acct_integration_free_bed_advace1()
     Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    
 If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
     Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, " advance Collection in cash")
    cmd.Parameters.Append Param1 'comment
                                                                    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, "6101003")
    cmd.Parameters.Append Param2 ''USER_acct
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, 0)
        cmd.Parameters.Append Param3 'dr
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, txtAdvanceFree.Text)
        cmd.Parameters.Append Param4 'cr
   
     Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 100, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param5 'u_id
     Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 100, VoucherNumber)
    cmd.Parameters.Append Param6 'u_id
         
            
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL acct.Save_vou_bill(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
                Conn.Close
       Set Conn = Nothing
    End If
  Set RS = Nothing
End Sub
Private Sub acct_integration_free_bed_advace()
     Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    
 If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
     Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, " advance Collection in cash")
    cmd.Parameters.Append Param1 'comment
                                                                    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, "2101")
    cmd.Parameters.Append Param2 ''USER_acct
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, txtAdvanceFree.Text)
        cmd.Parameters.Append Param3 'dr
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, 0)
        cmd.Parameters.Append Param4 'cr
   
     Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 100, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param5 'u_id
     Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 100, VoucherNumber)
    cmd.Parameters.Append Param6 'u_id
         
            
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL acct.Save_vou_bill(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
       Conn.Close
       Set Conn = Nothing
    End If
  Set RS = Nothing
End Sub
Private Sub print_admission()
              
              Dim Connpr As New Connection
              Dim cmdpr As New Command
              Dim Param1 As New Parameter
'            RS.Close
          If Connpr.State = 0 Then
            Connpr.Open strcn.Connection_String
          End If
            Set cmdpr.ActiveConnection = Connpr
            cmdpr.CommandType = adCmdText
            
            
            Dim Report6   As New CrystalReport6
            Set Param1 = cmdpr.CreateParameter("param1", adDouble, adParamInput, 100, PREVIEW_VAR)
            cmdpr.Parameters.Append Param1 '
     
            
            cmdpr.Properties("PLSQLRSet") = True
            cmdpr.CommandText = "{CALL Rpt_in_dr_info_adm_print(?)}"
            Set rspr = cmdpr.Execute
            cmdpr.Properties("PLSQLRSet") = False
            
            Report6.Database.SetDataSource rspr

            Report6.PrintOut
            
            'RSpr.Close
           If Connpr.State = 1 Then
            Connpr.Close
            Set Connpr = Nothing
            ''Set rspr = Nothing
           End If
           PREVIEW_VAR = 0
           txtRecNo.Locked = False
End Sub
Private Sub save_In_door_pat_free()
  Dim Connfree As New ADODB.Connection
    Dim cmdfree As New ADODB.Command
    Dim RSfree As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
    Dim Param11 As New Parameter
    Dim Param12 As New Parameter
    Dim Param13 As New Parameter
    Dim Param14 As New Parameter
    Dim Param15 As New Parameter
    Dim param16 As New Parameter
    Dim Param17 As New Parameter
    Dim Param18 As New Parameter
    Dim Param19 As New Parameter
    If Connfree.State = 0 Then
        Connfree.Open strcn.Connection_String
    End If
    
    Set cmdfree.ActiveConnection = Connfree
    cmdfree.CommandType = adCmdText
    
    Set Param1 = cmdfree.CreateParameter("param1", adVarChar, adParamInput, 10, Trim(comBedTypeFree.Text))
    cmdfree.Parameters.Append Param1 'bed type
    
    Set Param2 = cmdfree.CreateParameter("param2", adVarChar, adParamInput, 20, Trim(cmdTypeNoFree.Text))
    cmdfree.Parameters.Append Param2 'Type_no
    
    Set Param3 = cmdfree.CreateParameter("param3", adVarChar, adParamInput, 10, Trim(cmdBedNoFree.Text))
    cmdfree.Parameters.Append Param3 'Bed_no

    Set Param4 = cmdfree.CreateParameter("param4", adVarChar, adParamInput, 45, Trim(txtNameFree.Text))
    cmdfree.Parameters.Append Param4 'patient_name

    Set Param5 = cmdfree.CreateParameter("param5", adVarChar, adParamInput, 50, Trim(txtGuardNameFree.Text))
    cmdfree.Parameters.Append Param5 'guardgian_name
    
    Set Param6 = cmdfree.CreateParameter("param6", adVarChar, adParamInput, 200, Trim(TxtAddr2Free.Text))
    cmdfree.Parameters.Append Param6 'pat_addr1
    
    Set Param7 = cmdfree.CreateParameter("param7", adVarChar, adParamInput, 100, Trim(txtAddrFree.Text))
    cmdfree.Parameters.Append Param7 'imergency addr
    
    Set Param8 = cmdfree.CreateParameter("param8", adVarChar, adParamInput, 5, Trim(TxtAgeFree.Text))
    cmdfree.Parameters.Append Param8 'age
    
    Set Param9 = cmdfree.CreateParameter("param9", adVarChar, adParamInput, 6, Trim(comSexFree.Text))
    
    cmdfree.Parameters.Append Param9 'Sex
    
    Set Param10 = cmdfree.CreateParameter("param10", adVarChar, adParamInput, 10, Trim(comReligionFree.Text))
    cmdfree.Parameters.Append Param10 'Religion
    
    Set Param11 = cmdfree.CreateParameter("param11", adVarChar, adParamInput, 25, Trim(txtPhoneFree.Text))
    cmdfree.Parameters.Append Param11 'phone
    
    
    Set Param12 = cmdfree.CreateParameter("param12", adDouble, adParamInput, 10, Trim(txtAdvanceFree.Text))
    cmdfree.Parameters.Append Param12 'Advance
    
    Set Param13 = cmdfree.CreateParameter("param13", adVarChar, adParamInput, 10, Trim(comDepartmentFree.Text))
    cmdfree.Parameters.Append Param13 'Department
    
    Set Param14 = cmdfree.CreateParameter("param14", adVarChar, adParamInput, 6, frmMAIN.lbluser_id)
    cmdfree.Parameters.Append Param14 'u_id
      
    Set Param15 = cmdfree.CreateParameter("param15", adVarChar, adParamInput, 5, frmMAIN.lblBooth)
    cmdfree.Parameters.Append Param15 'booth
    
    
    Set param16 = cmdfree.CreateParameter("param16", adInteger, adParamInput, 5, 1)
    cmdfree.Parameters.Append param16 'extra-flag
    
    Set Param17 = cmdfree.CreateParameter("param17", adDouble, adParamInput, 10, Trim(txtSerial_no_free))
    cmdfree.Parameters.Append Param17 'Serial_no
    
    Set Param18 = cmdfree.CreateParameter("param18", adInteger, adParamInput, 10, Check8.Value)
    cmdfree.Parameters.Append Param18 'check_name
    
     Set Param19 = cmdfree.CreateParameter("param19", adVarChar, adParamInput, 10, Trim(CBOYRCODE.Text))
     cmdfree.Parameters.Append Param19 'check_value----father's name or husband name
 
   
    cmdfree.Properties("PLSQLRSet") = True
    
    cmdfree.CommandText = "{CALL Indoor_SavePatient_info(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    
    Debug.Print cmdfree.CommandText
    
    Set RSfree = cmdfree.Execute
    

    cmdfree.Properties("PLSQLRSet") = False
    If Connfree.State = 1 Then
        Connfree.Close
        Set Connfree = Nothing
        Set RSfree = Nothing
        Set cmdfree = Nothing
    End If
    Adodc5.ConnectionString = strcn.Connection_String
    Adodc5.RecordSource = "SELECT MAX(REC_NO) AS REC_NO FROM RECEIPT_NO_COUNTER"
    Adodc5.Refresh
    If Adodc5.Recordset.RecordCount > 0 Then
        TXTPRINT_PREV_FREE.Text = Adodc5.Recordset!REC_NO
        PREVIEW_VAR = Val(TXTPRINT_PREV_FREE)
    End If
   
'      RSfree.Close
End Sub


Private Sub Command15_Click()
Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Command17_Click()

 If TXTPRINT_PREV_PAYING.Visible = False Then
       TXTPRINT_PREV_PAYING.Width = TXTPRINT_PREV_CABIN.Width
       TXTPRINT_PREV_PAYING.Top = TXTPRINT_PREV_CABIN.Top
       TXTPRINT_PREV_PAYING.Left = TXTPRINT_PREV_CABIN.Left
       TXTPRINT_PREV_PAYING.Height = TXTPRINT_PREV_CABIN.Height
       
       TXTPRINT_PREV_PAYING.Visible = True
   End If
    PREVIEW_VAR = Val(TXTPRINT_PREV_PAYING)
  TXTPRINT_PREV_PAYING.ForeColor = vbBlue
  
   If TXTPRINT_PREV_PAYING = "" Then
    TXTPRINT_PREV_PAYING.SetFocus
     Exit Sub
   Else
         PREVIEW_VAR = Val(TXTPRINT_PREV_PAYING)
        rptMode = 6
      Viewer.Show vbModal
      TXTPRINT_PREV_PAYING = ""
      TXTPRINT_PREV_PAYING.Visible = False
End If
End Sub

Private Sub Command18_Click()
 If TXTPRINT_PREV_FREE.Visible = False Then
       TXTPRINT_PREV_FREE.Width = TXTPRINT_PREV_CABIN.Width
       TXTPRINT_PREV_FREE.Top = TXTPRINT_PREV_CABIN.Top
       TXTPRINT_PREV_FREE.Left = TXTPRINT_PREV_CABIN.Left
       TXTPRINT_PREV_FREE.Height = TXTPRINT_PREV_CABIN.Height
       
       TXTPRINT_PREV_FREE.Visible = True
   End If
    PREVIEW_VAR = Val(TXTPRINT_PREV_FREE)
  TXTPRINT_PREV_FREE.ForeColor = vbBlue
  
   If TXTPRINT_PREV_FREE = "" Then
    TXTPRINT_PREV_FREE.SetFocus
     Exit Sub
   Else
    PREVIEW_VAR = Val(TXTPRINT_PREV_FREE)
        rptMode = 6
      Viewer.Show vbModal
      TXTPRINT_PREV_FREE = ""
      TXTPRINT_PREV_FREE.Visible = False
End If
End Sub

Private Sub Command3_Click()
rptMode = 6
Viewer.Show vbModal
End Sub

Private Sub Command4_Click()
'Dim validation As Variant
'              Adodc1.ConnectionString = strcn.Connection_String
'              Adodc1.RecordSource = "Select user_id, user_name, user_password,user_type,shift_name From Security Where (User_Id = '" & frmMAIN.lbluser_id & "')"
'              Adodc1.Refresh
'              validation = Adodc1.Recordset!user_id
'
'                Dim Conn As New ADODB.Connection
'                Dim cmd As New ADODB.Command
'                Dim RS As New ADODB.Recordset
'
'                        Dim Param1 As New Parameter
'                        Conn.Open strcn.Connection_String
'
'                    Set cmd.ActiveConnection = Conn
'                    cmd.CommandType = adCmdText
'
'                   Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, validation)
'                    cmd.Parameters.Append Param1 'validation
'                    cmd.Properties("PLSQLRSet") = True
'
'                     cmd.CommandText = "{CALL shift_validation(?)}"
'
'                Debug.Print cmd.CommandText
'
'                    Set RS = cmd.Execute
'
'
'                cmd.Properties("PLSQLRSet") = False
'                Conn.Close
''                RS.Close
'
'          Adodc2.ConnectionString = strcn.Connection_String
'          Adodc2.RecordSource = "Select * From user_validation"
'          Adodc2.Refresh
'
'
'
'             If Adodc2.Recordset!validation = 0 Then
'             MsgBox "Your Working Time has been Expired", vbInformation, "Dhaka National Medical Institute Hospital."
'             Exit Sub
'             End If
'
'
'
'
'
'





If txtName_semi = "" Then
MsgBox "Patient Name Required", vbInformation, "Warning..."
txtName_semi.SetFocus
Exit Sub
End If
If txtAddr_semi = "" Then
MsgBox "Address Required", vbInformation, "Warning..."
txtAddr_semi.SetFocus
Exit Sub
End If
If txtAge_semi = "" Then
MsgBox "Age Required", vbInformation, "Warning..."
txtAge_semi.SetFocus
Exit Sub
End If

If txtAdvanceSemi = "" Then
MsgBox "Advance Required", vbInformation, "Warning..."
txtAdvanceSemi.SetFocus
Exit Sub
End If

Call save_In_door_pat_semi
MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
Call Clear_form
cmdBedNoSemiCabin.Refresh
cmdTypeNOSemiCabin.Refresh
 rptMode = 6
Viewer.Show vbModal

End Sub
Private Sub save_In_door_pat_semi()
' Dim Conn As New ADODB.Connection
'    Dim cmd As New ADODB.Command
'    Dim RS As New ADODB.Recordset
'
'    Dim Param1 As New Parameter
'    Dim Param2 As New Parameter
'    Dim Param3 As New Parameter
'    Dim Param4 As New Parameter
'    Dim Param5 As New Parameter
'    Dim Param6 As New Parameter
'    Dim Param7 As New Parameter
'    Dim Param8 As New Parameter
'    Dim Param9 As New Parameter
'    Dim Param10 As New Parameter
'    Dim Param11 As New Parameter
'    Dim Param12 As New Parameter
'    Dim Param13 As New Parameter
'    Dim Param14 As New Parameter
'    Dim Param15 As New Parameter
'    Dim Param16 As New Parameter
'
'    If Conn.State = 0 Then
'          Conn.Open strcn.Connection_String
'    End If
'
'    Set cmd.ActiveConnection = Conn
'    cmd.CommandType = adCmdText
'
'    '----------------------------------------------------------------------------------
'    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, Trim(Combo9.Text))
'    cmd.Parameters.Append Param1 'bed type
'
'    Set Param2 = cmd.CreateParameter("param2", adInteger, adParamInput, 10, Trim(cmdTypeNOSemiCabin.Text))
'    cmd.Parameters.Append Param2 'Type_no
'
'    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, Trim(cmdBedNoSemiCabin.Text))
'    cmd.Parameters.Append Param3 'Bed_no
'
'    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 45, Trim(txtName_semi.Text))
'    cmd.Parameters.Append Param4 'patient_name
'
'    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 50, Trim(Text14.Text))
'    cmd.Parameters.Append Param5 'guardgian_name
'
'    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 200, Trim(txtAddr_semi.Text))
'    cmd.Parameters.Append Param6 'pat_addr1
'
'    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 100, Trim(Text13.Text))
'    cmd.Parameters.Append Param7 'imergency addr
'
'    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 5, Trim(txtAge_semi.Text))
'    cmd.Parameters.Append Param8 'age
'
'    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 6, Trim(Combo10.Text))
'
'    cmd.Parameters.Append Param9 'Sex
'
'    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 10, Trim(Combo6.Text))
'    cmd.Parameters.Append Param10 'Religion
'
'    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 25, Trim(txtPhoneSemi.Text))
'    cmd.Parameters.Append Param11 'phone
'
'
'    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 10, Trim(txtAdvanceSemi.Text))
'    cmd.Parameters.Append Param12 'Advance
'
'    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 10, Trim(CboDepartment_semi.Text))
'    cmd.Parameters.Append Param13 'Department
'
'    Set Param14 = cmd.CreateParameter("param14", adVarChar, adParamInput, 6, frmMAIN.lbluser_id)
'    cmd.Parameters.Append Param14 'u_id
'
'
'
'    Set Param15 = cmd.CreateParameter("param15", adVarChar, adParamInput, 5, frmMAIN.lblBooth)
'    cmd.Parameters.Append Param15 'booth
'
'
'    Set Param16 = cmd.CreateParameter("param16", adVarChar, adParamInput, 5, 1)
'    cmd.Parameters.Append Param16 'extra-flag
'
'
'
'   cmd.Properties("PLSQLRSet") = True
'
'    cmd.CommandText = "{CALL Indoor_SavePatient_info(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
'
'    Debug.Print cmd.CommandText
'
'    Set RS = cmd.Execute
'
'
'    cmd.Properties("PLSQLRSet") = False
''    Close Conn
''    Close RS
'

End Sub





Private Sub Command5_Click()
Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Command7_Click()
Call clear_form_paying

End Sub
Private Sub clear_form_paying()
txtNamePaying = ""
txtGuardNamePaying = ""
txtAgePaying = ""
txtaddr2Paying = ""
txtaddrPaying = ""
txtPhonePaying = ""
txtAdvancePaying = ""

End Sub

Private Sub Command8_Click()
On Error GoTo Errdesc
Dim f2 As New frmDataSelect
Dim getconnected As New Connection
Dim cmd As New Command
Dim myrs As New ADODB.Recordset

    Call load_bed_reg_for_showing_PAYING
  If getconnected.State = 0 Then
    getconnected.ConnectionString = strcn.Connection_String
    getconnected.Open
  End If
    cmd.ActiveConnection = getconnected
    cmd.CommandType = adCmdText
    cmd.CommandText = " SELECT bed_ward,doc_department,bed_no,in_reg_no,name from show_bed order by bed_ward,doc_department "
                    
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    myrs.CursorLocation = adUseClient
     
    myrs.Open cmd.CommandText, getconnected, adOpenDynamic, adLockOptimistic
     Set f2.adoRecordset = myrs
     Set f2.OwnerForm = Me
     f2.Width = 8000
     f2.grdDataGrid.Columns(0).Caption = "Ward No"
     f2.grdDataGrid.Columns(1).Caption = "Department"
     f2.grdDataGrid.Columns(2).Caption = "Bed No"
     f2.grdDataGrid.Columns(3).Caption = "Registration No"
     f2.grdDataGrid.Columns(4).Caption = "Name"
     
     f2.grdDataGrid.Columns(0).Width = 500
     f2.grdDataGrid.Columns(1).Width = 1400
     f2.grdDataGrid.Columns(2).Width = 600
     f2.grdDataGrid.Columns(3).Width = 2500
     f2.grdDataGrid.Columns(4).Width = 3500
    
  cmd.Properties("iRowsetChange") = False
   
     f2.Show 1
'     Combo1(1) = myrs.Fields(0)
     'txtname = myrs.Fields(1)
'rptMode = 6
'Viewer.Show vbModal

'rptMode = 6
'Viewer.Show vbModal
''MsgBox rptMode
If getconnected.State = 1 Then
        getconnected.Close
        Set getconnected = Nothing
      'Set getconnected = Nothing
End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, "Dhaka National Medical Institute Hospital"

End Sub

Private Sub Command9_Click()
              
              Dim validation As Variant
              Adodc1.ConnectionString = strcn.Connection_String
              Adodc1.RecordSource = "Select user_id, user_name, user_password,user_type,shift_name From Security Where (User_Id = '" & frmMAIN.lbluser_id & "')"
              Adodc1.Refresh
              validation = Adodc1.Recordset!user_id
                
                Dim Connval As New ADODB.Connection
                Dim cmdval As New ADODB.Command
                Dim RSval As New ADODB.Recordset
                
                        Dim Param1 As New Parameter
'                        Connval.ConnectionString = strcn.Connection_String
                        Connval.Open strcn.Connection_String
'                         RSval.Open  strcn.Connection_String
                         

                    Set cmdval.ActiveConnection = Connval
                    cmdval.CommandType = adCmdText
    
                   Set Param1 = cmdval.CreateParameter("param1", adVarChar, adParamInput, 5, validation)
                    cmdval.Parameters.Append Param1 'validation
                    cmdval.Properties("PLSQLRSet") = True
    
                     cmdval.CommandText = "{CALL shift_validation(?)}"
    
                Debug.Print cmdval.CommandText
    
                    Set RSval = cmdval.Execute
    

                cmdval.Properties("PLSQLRSet") = False
'                 RSval.Close
             If Connval.State = 1 Then
                 Connval.Close
                 Set Connval = Nothing
              End If
               
                
          Adodc2.ConnectionString = strcn.Connection_String
          Adodc2.RecordSource = "Select * From user_validation"
          Adodc2.Refresh
        

        
             If Adodc2.Recordset!validation = 0 Then
             MsgBox "Your Working Time has been Expired", vbInformation, "Dhaka National Medical Institute Hospital."
             Exit Sub
             End If
             
          



If UCase(cboDepartmentPaying.Text) = UCase("Common") Then
               MsgBox "Please Select a Doctor's Department", vbInformation, "Dhaka National Medical Institute Hospital"
               cboDepartmentPaying.SetFocus

     Exit Sub
 End If


If txtNamePaying = "" Then
MsgBox "Patient Name Required", vbInformation, "Warning..."
txtNamePaying.SetFocus
Exit Sub
End If
If txtaddr2Paying = "" Then
MsgBox "Address Required", vbInformation, "Warning..."
txtaddr2Paying.SetFocus
Exit Sub
End If
 If txtAdvancePaying = "" Then
 MsgBox "Advance Required", vbInformation, "Warning..."
 txtAdvancePaying.SetFocus
 txtAdvancePaying = 0
 Exit Sub
 End If
If txtAgePaying = "" Then
MsgBox "Age Required", vbInformation, "Warning..."
txtAgePaying.SetFocus
Exit Sub
End If
If UCase(cboDepartmentPaying.Text) = UCase("Gynae-1") Or UCase(cboDepartmentPaying.Text) = UCase("Gynae-2") Then
    If UCase(CboSexPaying.Text) = UCase("Male") Then
   MsgBox "Please Select Sex as Female for Gynae Patient", vbInformation, "Dhaka National Medical Institute Hospital"
        CboSexPaying.SetFocus
         Exit Sub
  End If
End If
'Get_Voucher_Number
Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "SELECT occupy_flag AS REC_NO FROM bed_info  where serial_no='" & Trim(txtSerial_no_paying) & "'"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
        If Adodc3.Recordset!REC_NO = "1" Then
        
        MsgBox "This Bed Has Recently Occupied by a new Patient....select a new bed ", vbInformation, "Dhaka National Medical Institute Hospital"
       Exit Sub
       End If
    End If
 
Call save_In_door_pat_paying
'Call acct_integration_paying
'Call acct_integration_paying1
'Call post_vou
MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
Call clear_form_paying
cmdTypeNOPaying.Refresh
cmdBedNOPaying.Refresh

' rptMode = 6
'Viewer.Show vbModal

print_admission
cmdTypeNOPaying.clear
cmdBedNOPaying.clear
cboDepartmentPaying.SetFocus
'' RSval.Close
' Connval.Close

End Sub
Private Sub acct_integration_paying1()
   Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    
 If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
     Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, " advance Collection in cash")
    cmd.Parameters.Append Param1 'comment
                                                                    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, "6101003")
    cmd.Parameters.Append Param2 ''USER_acct
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, 0)
        cmd.Parameters.Append Param3 'dr
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, txtAdvancePaying.Text)
        cmd.Parameters.Append Param4 'cr
   
     Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 100, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param5 'u_id
     Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 100, VoucherNumber)
    cmd.Parameters.Append Param6 'u_id
         
            
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL acct.Save_vou_bill(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
                Conn.Close
       Set Conn = Nothing
    End If
  Set RS = Nothing
End Sub
Private Sub acct_integration_paying()
   Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    
 If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
     Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, " advance Collection in cash")
    cmd.Parameters.Append Param1 'comment
                                                                    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, "2101")
    cmd.Parameters.Append Param2 ''USER_acct
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, txtAdvancePaying.Text)
        cmd.Parameters.Append Param3 'dr
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, 0)
        cmd.Parameters.Append Param4 'cr
   
     Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 100, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param5 'u_id
     Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 100, VoucherNumber)
    cmd.Parameters.Append Param6 'u_id
         
            
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL acct.Save_vou_bill(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
                Conn.Close
       Set Conn = Nothing
    End If
  Set RS = Nothing
End Sub



Private Sub save_In_door_pat_paying()
'   Conn.Close
'   RS.Close
    Dim Connp As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RSp As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    Dim Param10 As New Parameter
    Dim Param11 As New Parameter
    Dim Param12 As New Parameter
    Dim Param13 As New Parameter
    Dim Param14 As New Parameter
    Dim Param15 As New Parameter
    Dim param16 As New Parameter
    Dim Param17 As New Parameter
     Dim Param18 As New Parameter
     Dim Param19 As New Parameter
    If Connp.State = 0 Then
        Connp.Open strcn.Connection_String
    End If
    
    Set cmd.ActiveConnection = Connp
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, Trim(cboBedTypePaying.Text))
    cmd.Parameters.Append Param1 'bed type
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(cmdTypeNOPaying.Text))
    cmd.Parameters.Append Param2 'Type_no
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 10, cmdBedNOPaying.Text)
    cmd.Parameters.Append Param3 'Bed_no

    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 45, Trim(txtNamePaying.Text))
    cmd.Parameters.Append Param4 'patient_name

    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 50, Trim(txtGuardNamePaying.Text))
    cmd.Parameters.Append Param5 'guardgian_name
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 200, Trim(txtaddr2Paying.Text))
    cmd.Parameters.Append Param6 'pat_addr1
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 100, Trim(txtaddrPaying.Text))
    cmd.Parameters.Append Param7 'imergency addr
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 5, Trim(txtAgePaying.Text))
    cmd.Parameters.Append Param8 'age
    
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 6, Trim(CboSexPaying.Text))
    
    cmd.Parameters.Append Param9 'Sex
    
    Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 10, Trim(cboReligionPaying.Text))
    cmd.Parameters.Append Param10 'Religion
    
    Set Param11 = cmd.CreateParameter("param11", adVarChar, adParamInput, 25, Trim(txtPhonePaying.Text))
    cmd.Parameters.Append Param11 'phone
    
    
    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 10, Trim(txtAdvancePaying.Text))
    cmd.Parameters.Append Param12 'Advance
    
    Set Param13 = cmd.CreateParameter("param13", adVarChar, adParamInput, 10, Trim(cboDepartmentPaying.Text))
    cmd.Parameters.Append Param13 'Department
    
'
    
    
    
    
'
'    Set Param10 = cmd.CreateParameter("param10", adInteger, adParamInput, 10, Get_Segment(Combo(4), False))
'    cmd.Parameters.Append Param10 'refer code
'
    
    
    
    Set Param14 = cmd.CreateParameter("param14", adVarChar, adParamInput, 6, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param14 'u_id
    
   
    
    Set Param15 = cmd.CreateParameter("param15", adVarChar, adParamInput, 5, frmMAIN.lblBooth)
    cmd.Parameters.Append Param15 'booth
    
    
    Set param16 = cmd.CreateParameter("param16", adVarChar, adParamInput, 5, 1)
    cmd.Parameters.Append param16 'extra-flag
    
    Set Param17 = cmd.CreateParameter("param17", adDouble, adParamInput, 10, txtSerial_no_paying)
   cmd.Parameters.Append Param17 'Serial no
   Set Param18 = cmd.CreateParameter("param18", adInteger, adParamInput, 10, Check6.Value)
   cmd.Parameters.Append Param18 'check_name

    Set Param19 = cmd.CreateParameter("param19", adVarChar, adParamInput, 10, Trim(CBOYRCODE))
    cmd.Parameters.Append Param19 'check_value----father's name or husband name
 

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Indoor_SavePatient_info(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    
    Set RSp = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
   If Connp.State = 1 Then
     Connp.Close
     Set Connp = Nothing
   End If
   Adodc4.ConnectionString = strcn.Connection_String
    Adodc4.RecordSource = "SELECT MAX(REC_NO) AS REC_NO FROM RECEIPT_NO_COUNTER"
    Adodc4.Refresh
    If Adodc4.Recordset.RecordCount > 0 Then
        TXTPRINT_PREV_PAYING = Adodc4.Recordset!REC_NO
          PREVIEW_VAR = Val(TXTPRINT_PREV_PAYING)
    End If

'    RSp.Close
    

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
End If
End Sub
Private Sub Form_Load()
txtRecNo.Locked = False
CboBedType = CboBedType.List(0)
cboDMY = cboDMY.List(0)
cboSex = cboSex.List(0)
cboReligion = cboReligion.List(0)
'
MaskEdBox1.Text = Format(Date, "dd/mm/yyyy")
DTPicker7.Value = Time

Call show_dept
cboBedDept.Text = "Common"
txtPatFatherName = "S/D/O:"
cboPatdept = cboPatdept.List(0)



End Sub
Private Sub show_dept()
        Dim Conn As New ADODB.Connection
        Dim RS As New ADODB.Recordset
        Dim cmd As New ADODB.Command
        
       If Conn.State = 0 Then
            Conn.ConnectionString = strcn.Connection_String
            Conn.Open
       End If
       cmd.ActiveConnection = Conn
       cmd.CommandType = adCmdText
       cmd.CommandText = "select distinct(doc_dept),refer_code from doctor_info order by refer_code"
      
         
      Set RS = cmd.Execute
      
      cboBedDept.clear
      cboPatdept.clear
      
      If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF = True
           cboBedDept.AddItem RS!doc_dept
           cboPatdept.AddItem RS!doc_dept
           RS.MoveNext
        Loop

    End If
    
    If Conn.State = 1 Then
        Conn.Close
        Set RS = Nothing
        Set cmd = Nothing
        Set Conn = Nothing
    End If

End Sub
Private Sub LOAD_BED_NO()
''On Error Resume Next
 
     If Conn.State = 0 Then
      Conn.ConnectionString = strcn.Connection_String
          Conn.Open
    End If
    cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
   cmd.CommandText = "select distinct(BED_EXT_COL) from bed_info  where  bed_type=  '" & CboBedType & "' and doc_department='" & Trim(cboBedDept.Text) & "' and occupy_flag='0' order by BED_EXT_COL"
    
    cmd.Properties("iRowsetChange") = True
   cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    
    RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
    
    CboTypeNo.clear
    cboBedNo.clear
    If RS.RecordCount > 0 Then
       RS.MoveFirst
       Do Until RS.EOF = True
          CboTypeNo.AddItem RS.Fields(0)
          RS.MoveNext
       Loop
      CboTypeNo = CboTypeNo.List(0)
   End If
     cmd.Properties("iRowsetChange") = False
   
    If Conn.State = 1 Then
        Conn.Close
       Set Conn = Nothing
       Set RS = Nothing
       Set cmd = Nothing
    End If
    
End Sub
Private Sub Load_bed_no_semi_cabin()
''On Error Resume Next
'
'      If Conn1.State = 0 Then
'            Conn1.ConnectionString = strcn.Connection_String
'            Conn1.Open
'   End If
'    cmd.ActiveConnection = Conn1
'    cmd.CommandType = adCmdText
'   cmd.CommandText = "select distinct(BED_EXT_COL) from bed_info  where   bed_type= 'Semi-Cabin'and doc_department='" & Trim(CboDepartment_semi.Text) & "'"
'
'    cmd.Properties("iRowsetChange") = True
'   cmd.Properties("updatability") = 7
'    RS1.CursorLocation = adUseClient
'
'    RS1.Open cmd.CommandText, Conn1, adOpenDynamic, adLockOptimistic
'    cmdTypeNOSemiCabin.clear
'     cmd.Properties("iRowsetChange") = False
'
'    If RS1.RecordCount > 0 Then
'
'          cmdTypeNOSemiCabin.Refresh
'         RS1.MoveFirst
'
'        Do Until RS1.EOF = True
'
'            cmdTypeNOSemiCabin.AddItem RS1.Fields(0)
'
'            RS1.MoveNext
'        Loop
'
'    End If
'
'    RS1.Close
'   If Conn1.State = 1 Then
'        Conn1.Close
'   End If
    
End Sub



Private Sub Load_bed_paying()
''On Error Resume Next

      If Conn1.State = 0 Then
       Conn1.ConnectionString = strcn.Connection_String
        Conn1.Open
   End If
    cmd.ActiveConnection = Conn1
    cmd.CommandType = adCmdText
   cmd.CommandText = "select distinct(BED_EXT_COL) from bed_info  where   bed_type= 'Paying'and doc_department='" & Trim(cboDepartmentPaying.Text) & "'"
    
  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
    RS1.CursorLocation = adUseClient
    
    RS1.Open cmd.CommandText, Conn1, adOpenDynamic, adLockOptimistic
    
    
    If RS1.RecordCount > 0 Then
            
          
        cmdTypeNOPaying.clear
         RS1.MoveFirst
            
        Do Until RS1.EOF = True
        
            cmdTypeNOPaying.AddItem RS1.Fields(0)
    
            RS1.MoveNext
        Loop
    
    End If
     cmd.Properties("iRowsetChange") = False

    Set RS1 = Nothing
  If Conn1.State = 1 Then
    Conn1.Close
    Set Conn1 = Nothing
  End If
    
End Sub
Private Sub Load_Free_Bed()
''On Error Resume Next

      If Conn1.State = 0 Then
      Conn1.ConnectionString = strcn.Connection_String
    Conn1.Open
   End If
    cmd.ActiveConnection = Conn1
    cmd.CommandType = adCmdText
   cmd.CommandText = "select distinct(BED_EXT_COL) from bed_info  where   bed_type= 'Free-Bed'and doc_department='" & Trim(comDepartmentFree.Text) & "'"
    
    cmd.Properties("iRowsetChange") = True
   cmd.Properties("updatability") = 7
    RS1.CursorLocation = adUseClient
    
    RS1.Open cmd.CommandText, Conn1, adOpenDynamic, adLockOptimistic
    
    
    If RS1.RecordCount > 0 Then
            
          cmdTypeNoFree.clear

         RS1.MoveFirst
         
        Do Until RS1.EOF = True
        
            cmdTypeNoFree.AddItem RS1.Fields(0)
    
            RS1.MoveNext
        Loop
    
    End If
      cmd.Properties("iRowsetChange") = False
    'RS1.Close
  If Conn1.State = 1 Then
    Conn1.Close
    Set Conn1 = Nothing
  End If
    
End Sub

Private Sub Frame5_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub nbrAdv_Change()
If Not IsNumeric(nbrAdv.Text) Then
             nbrAdv = ""
End If
End Sub

Private Sub nbrAdv_GotFocus()
nbrAdv.BackColor = &H80000018
End Sub

Private Sub nbrAdv_LostFocus()
nbrAdv.BackColor = vbWhite

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer) 'Cabin
If SSTab1.Caption = "Cabin" Then
Combo1.Text = "Cabin"
cboCabin.SetFocus
'cboCabin.List(0) = "COMMON"
''Call show_dept_free 'calling function for finding doctor department
Call LOAD_BED_NO
If cboCabin.ListCount <> Empty Then
cboCabin = cboCabin.List(0)
End If
'cmdTypeNOCabin.SetFocus
cboCabin.SetFocus
Dt = Date
DT_TM = Now
End If


If SSTab1.Caption = "Semi-Cabin" Then   ''semi Cabin
Combo9.Text = "Semi-Cabin"
Call show_dept_semi
If CboDepartment_semi.ListCount <> Empty Then
CboDepartment_semi = CboDepartment_semi.List(0)
End If


Combo10 = Combo10.List(0)
Combo6 = Combo6.List(0)

Call Load_bed_no_semi_cabin
If cmdTypeNOSemiCabin.ListCount <> 0 Then
cmdTypeNOSemiCabin = cmdTypeNOSemiCabin.List(0)
End If

 cmdTypeNOSemiCabin.SetFocus
End If


If SSTab1.Caption = "Paying" Then      ''Paying
DTPicker4 = Date
DTPicker3 = Now
cboBedTypePaying.Text = "Paying"
CboSexPaying = CboSexPaying.List(0)
cboReligionPaying = cboReligionPaying.List(0)
Call show_dept_paying
If cboDepartmentPaying.ListCount <> Empty Then
cboDepartmentPaying.Text = cboDepartmentPaying.List(0)
End If
''Call Load_bed_paying
'cmdTypeNOPaying.SetFocus
cboDepartmentPaying.SetFocus
If cmdTypeNOPaying.ListCount <> 0 Then
cmdTypeNOPaying = cmdTypeNOPaying.List(0)
End If

End If


If SSTab1.Caption = "Free-Bed" Then        ''Free Bed
DTPicker6 = Date
DTPicker5 = Now
comBedTypeFree.Text = "Free-Bed"
comSexFree = comSexFree.List(0)
comReligionFree = comReligionFree.List(0)
Call show_dept_free
If comDepartmentFree.ListCount <> Empty Then
comDepartmentFree.Text = comDepartmentFree.List(0)
End If
Call Load_Free_Bed
'cmdTypeNoFree.SetFocus
comDepartmentFree.SetFocus
If cmdTypeNoFree.ListCount <> 0 Then
cmdTypeNoFree = cmdTypeNoFree.List(0)
End If

End If

End Sub

Private Sub show_dept_paying()
      If Conn2.State = 0 Then
      Conn2.ConnectionString = strcn.Connection_String
         Conn2.Open
     End If
     cmd.ActiveConnection = Conn2
     cmd.CommandType = adCmdText
      cmd.CommandText = "select distinct(doc_dept),refer_code from doctor_info order by refer_code"
      
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    rs2.CursorLocation = adUseClient

    rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
      
      
     cboDepartmentPaying.clear
        
    
    If rs2.RecordCount > 0 Then


         rs2.MoveFirst

        Do Until rs2.EOF = True

            cboDepartmentPaying.AddItem rs2!doc_dept
            rs2.MoveNext
        Loop

    End If
    'Combo4.Refresh

    'rs2.Close
     cmd.Properties("iRowsetChange") = False

   If Conn2.State = 1 Then
    Conn2.Close
    Set Conn2 = Nothing
   End If

End Sub

Private Sub show_dept_semi()
'     Conn2.ConnectionString = strcn.Connection_String
'     Conn2.Open
'     cmd.ActiveConnection = Conn2
'     cmd.CommandType = adCmdText
'      cmd.CommandText = "select distinct(doc_dept),refer_code from doctor_info order by refer_code"
'
'      cmd.Properties("iRowsetChange") = True
'    cmd.Properties("updatability") = 7
'    rs2.CursorLocation = adUseClient
'
'    rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
'
'
'     CboDepartment_semi.clear
'
'
'    If rs2.RecordCount > 0 Then
'
'
'         rs2.MoveFirst
'
'        Do Until rs2.EOF = True
'
'            CboDepartment_semi.AddItem rs2!doc_dept
'            rs2.MoveNext
'        Loop
'
'    End If
'    'Combo4.Refresh
'
'    rs2.Close
'
'    Conn2.Close

End Sub

Private Sub show_dept_free()
  If Conn2.State = 0 Then
     Conn2.ConnectionString = strcn.Connection_String
         Conn2.Open
     End If
     cmd.ActiveConnection = Conn2
     cmd.CommandType = adCmdText
      cmd.CommandText = "select distinct(doc_dept),refer_code from doctor_info order by refer_code"
      
     cmd.Properties("iRowsetChange") = True
   cmd.Properties("updatability") = 7
    rs2.CursorLocation = adUseClient

    rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
      
      
     
        comDepartmentFree.clear
        
    
    If rs2.RecordCount > 0 Then


         rs2.MoveFirst

        Do Until rs2.EOF = True

            comDepartmentFree.AddItem rs2!doc_dept
            rs2.MoveNext
        Loop

    End If
    'Combo4.Refresh
cmd.Properties("iRowsetChange") = False

    'rs2.Close
If Conn2.State = 1 Then
    Conn2.Close
    Set Conn2 = Nothing
End If

End Sub

Private Sub Text4_GotFocus()
Text4.BackColor = &H80000018
End Sub

Private Sub Text4_LostFocus()
Text4.BackColor = vbWhite

End Sub

Private Sub Text5_GotFocus()
 Text5.BackColor = &H80000018
End Sub



Private Sub Text5_LostFocus()
Text5.BackColor = vbWhite

End Sub

Private Sub txtAddr_GotFocus()
txtAddr.BackColor = &H80000018
End Sub

Private Sub txtAddr_LostFocus()
txtAddr.BackColor = vbWhite

End Sub

Private Sub TxtAddr2Free_GotFocus()
TxtAddr2Free.BackColor = &H80000018
End Sub

Private Sub TxtAddr2Free_LostFocus()
TxtAddr2Free.BackColor = vbWhite
End Sub

Private Sub txtaddr2Paying_GotFocus()
txtaddr2Paying.BackColor = &H80000018
End Sub

Private Sub txtaddr2Paying_LostFocus()
txtaddr2Paying.BackColor = vbWhite
End Sub

Private Sub txtAddrFree_GotFocus()
txtAddrFree.BackColor = &H80000018
End Sub

Private Sub txtAddrFree_LostFocus()
txtAddrFree.BackColor = vbWhite
End Sub

Private Sub txtaddrPaying_GotFocus()
txtaddrPaying.BackColor = &H80000018
End Sub

Private Sub txtaddrPaying_LostFocus()
txtaddrPaying.BackColor = vbWhite
End Sub

Private Sub txtAdvanceFree_Change()
If Not IsNumeric(txtAdvanceFree.Text) Then
             txtAdvanceFree = ""
End If
End Sub

Private Sub txtAdvanceFree_GotFocus()
txtAdvanceFree.BackColor = &H80000018
End Sub

Private Sub txtAdvanceFree_LostFocus()
txtAdvanceFree.BackColor = vbWhite
End Sub

Private Sub txtAdvancePaying_Change()
If Not IsNumeric(txtAdvancePaying.Text) Then
             txtAdvancePaying = ""
End If
End Sub

Private Sub txtAdvancePaying_GotFocus()
txtAdvancePaying.BackColor = &H80000018
End Sub

Private Sub txtAdvancePaying_LostFocus()
txtAdvancePaying.BackColor = vbWhite
End Sub

Private Sub txtAdvanceSemi_Change()
If Not IsNumeric(txtAdvanceSemi.Text) Then
             txtAdvanceSemi = ""
End If
End Sub

Private Sub Text7_Change()

End Sub

Private Sub Text7_GotFocus()
  
End Sub

Private Sub txtAddress_GotFocus()
   txtAddress.BackColor = &H80000018
End Sub

Private Sub txtAddress_LostFocus()
 txtAddress.BackColor = vbWhite
End Sub

Private Sub txtAdvance_Change()
  If Not IsNumeric(txtAdvance) Then
     txtAdvance = ""
 End If
 
End Sub
Private Sub txtAdvance_GotFocus()
  txtAdvance.BackColor = &H80000018
End Sub

Private Sub txtAdvance_LostFocus()
  txtAdvance.BackColor = vbWhite
End Sub

Private Sub txtAge_Change()
If Not IsNumeric(txtAge) Then
     txtAge = ""
End If
If Val(txtAge) > 200 Then
   MsgBox "Invalid Age", vbInformation, "Dhaka National Medical Institute Hospital"
    txtAge = ""
End If

End Sub

Private Sub txtAge_GotFocus()
   txtAge.BackColor = &H80000018
End Sub

Private Sub txtAge_LostFocus()
txtAge.BackColor = vbWhite
End Sub

Private Sub txtAge_semi_Change()
If Not IsNumeric(txtAge_semi) Then
 txtAge_semi = ""
End If
If Val(txtAge_semi) > 200 Then
MsgBox "Invalid Age", vbInformation, "Dhaka National Medical Institute Hospital"
 txtAge_semi = ""
End If

End Sub

Private Sub TxtAgeFree_Change()
If Not IsNumeric(TxtAgeFree) Then
 TxtAgeFree = ""
End If
If Val(TxtAgeFree) > 200 Then
MsgBox "Invalid Age", vbInformation, "Dhaka National Medical Institute Hospital"
 TxtAgeFree = ""
End If

End Sub

Private Sub TxtAgeFree_GotFocus()
TxtAgeFree.BackColor = &H80000018
End Sub

Private Sub TxtAgeFree_LostFocus()
TxtAgeFree.BackColor = vbWhite
End Sub

Private Sub txtAgePaying_Change()
If Not IsNumeric(txtAgePaying) Then
txtAgePaying = ""
End If
If Val(txtAgePaying) > 200 Then
MsgBox "Invalid Age", vbInformation, "Dhaka National Medical Institute Hospital"
txtAgePaying = ""
End If

End Sub

Private Sub txtAgePaying_GotFocus()
txtAgePaying.BackColor = &H80000018
End Sub

Private Sub txtAgePaying_LostFocus()
txtAgePaying.BackColor = vbWhite

End Sub

Private Sub txtGuardNameFree_GotFocus()
txtGuardNameFree.BackColor = &H80000018
End Sub

Private Sub txtGuardNameFree_LostFocus()
txtGuardNameFree.BackColor = vbWhite
End Sub

Private Sub txtGuardNamePaying_GotFocus()
txtGuardNamePaying.BackColor = &H80000018
End Sub

Private Sub txtGuardNamePaying_LostFocus()
txtGuardNamePaying.BackColor = vbWhite

End Sub

Private Sub txtname_GotFocus()
txtname.BackColor = &H80000018
End Sub

Private Sub txtname_LostFocus()
txtname.BackColor = vbWhite
End Sub

Private Sub txtNameFree_GotFocus()
txtNameFree.BackColor = &H80000018
End Sub

Private Sub txtNameFree_LostFocus()
txtNameFree.BackColor = vbWhite
End Sub

Private Sub txtNamePaying_GotFocus()
txtNamePaying.BackColor = &H80000018
End Sub

Private Sub txtNamePaying_LostFocus()
txtNamePaying.BackColor = vbWhite
End Sub

Private Sub txtPhone_GotFocus()
      txtPhone.BackColor = &H80000018
End Sub

Private Sub txtPhone_LostFocus()
      txtPhone.BackColor = vbWhite
End Sub

Private Sub txtPhoneFree_GotFocus()
      txtPhoneFree.BackColor = &H80000018
End Sub

Private Sub txtPhoneFree_LostFocus()
      txtPhoneFree.BackColor = vbWhite
End Sub

Private Sub txtPhonePaying_GotFocus()
txtPhonePaying.BackColor = &H80000018
End Sub
Private Sub txtPhonePaying_LostFocus()
  txtPhonePaying.BackColor = vbWhite
End Sub

Private Sub txtPatFatherName_GotFocus()
  
  txtPatFatherName.BackColor = &H80000018
  txtPatFatherName.SelStart = 6
End Sub

Private Sub txtPatFatherName_LostFocus()
  txtPatFatherName.BackColor = vbWhite
End Sub

Private Sub txtPatName_GotFocus()
  txtPatName.BackColor = &H80000018
End Sub

Private Sub txtPatName_LostFocus()
  txtPatName.BackColor = vbWhite
End Sub
