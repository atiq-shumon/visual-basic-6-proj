VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBed_transfer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   52
      Top             =   7710
      Width           =   11385
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Software Programmer, IT Division, DNMIH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2940
         TabIndex        =   54
         Top             =   60
         Width           =   4725
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Developed && Maintenanced by:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   150
         TabIndex        =   53
         ToolTipText     =   "Developed && Maintenanced by:A.N.M. Atiqur Rahman,Software Programmer, IT Division, DNMIH"
         Top             =   90
         Width           =   2715
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   -60
      TabIndex        =   38
      Top             =   6780
      Width           =   11385
      Begin VB.TextBox txtExtraBedFlag 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1230
         TabIndex        =   49
         Top             =   510
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtPrintPreview 
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
         Height          =   285
         Left            =   7890
         TabIndex        =   43
         Top             =   30
         Visible         =   0   'False
         Width           =   1875
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   3360
         Top             =   390
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txtSerialNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   42
         Top             =   270
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.CommandButton CMDEXIT 
         Cancel          =   -1  'True
         Caption         =   "CLOSE"
         Height          =   375
         Left            =   9450
         TabIndex        =   6
         Top             =   330
         Width           =   1215
      End
      Begin VB.CommandButton CMDREPORT 
         Caption         =   "REPORT"
         Height          =   375
         Left            =   8220
         TabIndex        =   41
         Top             =   330
         Width           =   1215
      End
      Begin VB.CommandButton cmdShowBed 
         Caption         =   "SHOW BED"
         Height          =   375
         Left            =   6990
         TabIndex        =   40
         Top             =   330
         Width           =   1215
      End
      Begin VB.CommandButton cmdADD 
         Caption         =   "NEW"
         Height          =   375
         Left            =   5760
         TabIndex        =   39
         Top             =   330
         Width           =   1215
      End
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "SAVE"
         Height          =   375
         Left            =   4530
         TabIndex        =   5
         Top             =   330
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   4470
         Top             =   270
         Width           =   6255
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   -30
         Picture         =   "frmBed_transfer.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   12675
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1215
      Left            =   -90
      TabIndex        =   9
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
         Left            =   5280
         TabIndex        =   11
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PATIENT WARD/BED TRANSFER ENTRY"
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
         Left            =   3090
         TabIndex        =   10
         Top             =   360
         Width           =   6525
      End
      Begin VB.Image Image1 
         Height          =   1155
         Left            =   30
         Picture         =   "frmBed_transfer.frx":5982
         Stretch         =   -1  'True
         Top             =   180
         Width           =   12180
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   5715
      Left            =   -30
      TabIndex        =   12
      Top             =   1050
      Width           =   11385
      Begin VB.TextBox txtAddress 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   540
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2130
         Width           =   6525
      End
      Begin VB.TextBox txtPatDept 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1350
         Width           =   3495
      End
      Begin VB.TextBox TXTdmy 
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
         Left            =   6330
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   46
         Top             =   1350
         Width           =   735
      End
      Begin VB.TextBox txtBedType 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2130
         Width           =   3465
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   2550
         Top             =   5580
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.ComboBox CBOYRCODE 
         Height          =   315
         ItemData        =   "frmBed_transfer.frx":B304
         Left            =   5610
         List            =   "frmBed_transfer.frx":B314
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   510
         Width           =   1425
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
         TabIndex        =   4
         Top             =   4920
         Width           =   3405
      End
      Begin VB.TextBox txtAge 
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
         Left            =   5580
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1350
         Width           =   735
      End
      Begin VB.TextBox txtPatName 
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
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   570
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1350
         Width           =   4785
      End
      Begin VB.ComboBox cboBedDept 
         Height          =   315
         ItemData        =   "frmBed_transfer.frx":B33C
         Left            =   2670
         List            =   "frmBed_transfer.frx":B33E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   3960
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
         TabIndex        =   26
         Top             =   3960
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
         TabIndex        =   24
         Top             =   3960
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
         TabIndex        =   22
         Top             =   3960
         Width           =   825
      End
      Begin VB.ComboBox cboBedNo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmBed_transfer.frx":B340
         Left            =   5610
         List            =   "frmBed_transfer.frx":B342
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3960
         Width           =   1365
      End
      Begin VB.ComboBox CboTypeNo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmBed_transfer.frx":B344
         Left            =   4380
         List            =   "frmBed_transfer.frx":B346
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3960
         Width           =   1065
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   7350
         TabIndex        =   18
         Top             =   510
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   255
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtRegNo 
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
         Height          =   315
         Left            =   2670
         TabIndex        =   16
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
         ItemData        =   "frmBed_transfer.frx":B348
         Left            =   570
         List            =   "frmBed_transfer.frx":B355
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   3960
         Width           =   1845
      End
      Begin VB.TextBox txtRecNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   570
         TabIndex        =   13
         Top             =   510
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker7 
         Height          =   315
         Left            =   9540
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Delevary Time"
         Top             =   480
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   59965442
         UpDown          =   -1  'True
         CurrentDate     =   37163
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'Dot
         Height          =   765
         Index           =   1
         Left            =   7140
         Shape           =   4  'Rounded Rectangle
         Top             =   4590
         Width           =   3735
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'Dot
         Height          =   1005
         Index           =   0
         Left            =   480
         Shape           =   4  'Rounded Rectangle
         Top             =   3510
         Width           =   6645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   11
         Left            =   540
         TabIndex        =   51
         Top             =   1830
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "----------------------------------------TRANSFER TO-----------------------------------"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   600
         TabIndex        =   48
         Top             =   2820
         Width           =   10065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Bed No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   13
         Left            =   7350
         TabIndex        =   44
         Top             =   1860
         Width           =   1665
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   20
         Left            =   6300
         TabIndex        =   37
         Top             =   1050
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   19
         Left            =   5610
         TabIndex        =   36
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   18
         Left            =   9540
         TabIndex        =   35
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   17
         Left            =   7380
         TabIndex        =   34
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
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   16
         Left            =   7350
         TabIndex        =   33
         Top             =   1050
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
         TabIndex        =   31
         Top             =   4620
         Width           =   945
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   10
         Left            =   5610
         TabIndex        =   30
         Top             =   1050
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
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   9
         Left            =   600
         TabIndex        =   29
         Top             =   1050
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
         TabIndex        =   28
         Top             =   3630
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
         TabIndex        =   27
         Top             =   3630
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
         TabIndex        =   25
         Top             =   3630
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
         TabIndex        =   23
         Top             =   3630
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
         TabIndex        =   21
         Top             =   3630
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cab/Ward"
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
         Left            =   4380
         TabIndex        =   20
         Top             =   3630
         Width           =   1065
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
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   2
         Left            =   2670
         TabIndex        =   17
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
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   600
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   3630
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmBed_transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim var_dept_serial_no As Integer
Dim VAR_CUR_BED_SERIAL_NO As Integer
Dim UTILITY As New clsUtility
Public strUid As String
Public strcn        As New MyConnection



Private Sub cboBedDept_Click()
  txtAdmissionfee = ""
  txtBedCharge = ""
  txtServiceFee = ""
  txtSerialNo = ""
  Call LOAD_BED_NO
  If cboBedType = "Free-Bed" Then
     txtServiceFee = UTILITY.Load_Service_Charge_Exceptional(cboBedType, cboBedDept)
  End If
  
 End Sub

Private Sub cboBedDept_GotFocus()
  cboBedDept_Click
End Sub

Private Sub cboBedNo_Click()
  load_fee
  If cboBedType = "Free-Bed" Then
     txtServiceFee = UTILITY.Load_Service_Charge_Exceptional(cboBedType, cboBedDept)
  End If
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
    cmd.CommandText = "select bed_charge,service_charge,BED_GROUP,serial_no from bed_info  where occupy_flag='0'and upper(bed_type)=upper('" & cboBedType & "') and  BED_EXT_COL='" & Trim(CboTypeNo.Text) & "'and bed_no='" & Trim(cboBedNo.Text) & "' and doc_department='" & Trim(cboBedDept.Text) & "'"
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
   
   If RS.RecordCount > 0 Then
       txtAdmissionfee = RS!BED_GROUP
       txtBedCharge = RS!bed_charge
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
  Label47.Caption = cboBedType.Text
  If cboBedType = "Cabin" And Len(cboBedDept) > 0 Then
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
Private Sub cmdADD_Click()
  Call Clear_form
End Sub
Private Sub Clear_form()
MaskEdBox1.Text = Format(Date, "dd/mm/yyyy")
DTPicker7.Value = Time

txtAdvance = 0
txtAge = ""

cboBedType.SetFocus
End Sub

Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
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
MsgBox Err.Description, vbCritical, " IT, DNMIH"
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
    MsgBox Err.Description, vbInformation, " IT, DNMIH"

End Sub


Private Sub CMDREPORT_Click()
  If txtPrintPreview.Visible = False Then
      txtPrintPreview.Visible = True
   End If
  
    PREVIEW_VAR = Val(txtPrintPreview)
  
   If txtPrintPreview = "" Then
       txtPrintPreview.SetFocus
        Exit Sub
   Else
'    PREVIEW_VAR = Val(txtPrintPreview)
    rptMode = 17
    Viewer.Show vbModal
    txtPrintPreview = ""
    txtPrintPreview.Visible = False
End If
End Sub
Private Sub cmdSave_Click()
           If UTILITY.User_Shift_validation(frmMAIN.lbluser_id, frmMAIN.lblUserType.Caption) = False Then
                MsgBox "Dear. " & frmMAIN.lblName & "  Your Shift has been Expired.. " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
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
           
          If txtPatName = "" Then
            MsgBox "Patient Name Required", vbInformation, "Warning..."
            txtPatName.SetFocus
            Exit Sub
          End If
          
                       
            
          If Len(txtAdvance) = 0 Then
             MsgBox "Advance Required", vbInformation, "Warning..."
             txtAdvance.SetFocus
             txtAdvance = 0
             Exit Sub
          End If
                     
    
    Adodc3.ConnectionString = strcn.Connection_String
    Adodc3.RecordSource = "SELECT occupy_flag AS REC_NO FROM bed_info  where serial_no='" & Trim(txtSerialNo) & "'"
    Adodc3.Refresh
    If Adodc3.Recordset.RecordCount > 0 Then
        If Adodc3.Recordset!REC_NO = "1" Then
           MsgBox "This Bed Has Recently Occupied by a new Patient....select a new bed ", vbInformation, " IT, DNMIH"
       Exit Sub
       End If
    End If
    
If cboBedType = "Free-Bed" Then
     txtServiceFee = UTILITY.Load_Service_Charge_Exceptional(cboBedType, cboBedDept)
  End If
Call SAVE_PATIENT_BED_TRANSFER

MsgBox "Operation successful", vbInformation + vbOKOnly, "Save..."
print_admission (1)
Call Clear_form
LOAD_PAT_BED


Unload Me
End Sub
Private Sub SAVE_PATIENT_BED_TRANSFER()
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
      
      
    Dim Param0 As New Parameter
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
   If Conn.State = 0 Then
      Conn.Open strcn.Connection_String
  End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param0 = cmd.CreateParameter("param0", adInteger, adParamInput, 10, 1)
    cmd.Parameters.Append Param0 'mode
   
    
    Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 10, Trim(frmReg_for_Bed_Transf.txtReg_noOpr.Text))
    cmd.Parameters.Append Param1 'Registration_no
    
     Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 10, txtSerialNo.Text)
    cmd.Parameters.Append Param2 'serial_no
  
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 6, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param3 'u_id
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 5, frmMAIN.lblBooth)
    cmd.Parameters.Append Param4 'booth
    If txtAdvance.Text = "" Then
       txtAdvance.Text = 0
    End If
    
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 10, Trim(Val(txtAdvance.Text)))
    cmd.Parameters.Append Param5 'Advance
    
     Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 10, Trim(frmReg_for_Bed_Transf.CBOYRCODE.Text))
    cmd.Parameters.Append Param6 'YRCODE
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 20, Trim(cboBedType))
    cmd.Parameters.Append Param7 'BED TYPE
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 25, Trim(CboTypeNo))
    cmd.Parameters.Append Param8 'BED TYPE NO
 
   Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 15, Trim(cboBedNo))
   cmd.Parameters.Append Param9 'BED NO
 
   Set Param10 = cmd.CreateParameter("param10", adVarChar, adParamInput, 19, Trim(cboBedDept))
   cmd.Parameters.Append Param10 'BED DEPT
   
   If cboBedDept = "COMMON" And cboBedType <> "Cabin" Then
      Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 5, 1)
      cmd.Parameters.Append Param11 'EXTRA BED FLAG INDICATOR
   Else
      Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 5, 0)
      cmd.Parameters.Append Param11 'EXTRA BED FLAG INDICATOR
   
   End If
   
   Set Param12 = cmd.CreateParameter("param12", adVarChar, adParamInput, 19, Trim(txtPatDept))
   cmd.Parameters.Append Param12 'DOC DEPT
  
   Set Param13 = cmd.CreateParameter("param13", adDate, adParamInput, 8, Trim(MaskEdBox1))
   cmd.Parameters.Append Param13 'ADMISSION DATE
  
   Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 8, Trim(txtAdmissionfee))
   cmd.Parameters.Append Param14 'ADMISSION CHARGE
   
   Set Param15 = cmd.CreateParameter("param15", adDouble, adParamInput, 8, Trim(txtBedCharge))
   cmd.Parameters.Append Param15 'BED  CHARGE
 
   Set param16 = cmd.CreateParameter("param16", adDouble, adParamInput, 8, Trim(txtServiceFee))
   cmd.Parameters.Append param16 'SERVICE CHARGE
   
   Set Param17 = cmd.CreateParameter("param17", adDouble, adParamInput, 8, Trim(txtExtraBedFlag))
   cmd.Parameters.Append Param17 'EXT BED FLAG
  
   Set Param18 = cmd.CreateParameter("param18", adInteger, adParamInput, 8, var_dept_serial_no)
   cmd.Parameters.Append Param18 'Department serial
 
   Set Param19 = cmd.CreateParameter("param19", adInteger, adParamInput, 8, VAR_CUR_BED_SERIAL_NO)
   cmd.Parameters.Append Param19 'CURRENT BED SERIAL NO
   
   Set Param20 = cmd.CreateParameter("param20", adDate, adParamInput, 10, Date$)
   cmd.Parameters.Append Param20 'BED TRANSFER DATE
 
   
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_bed_transfer(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
   
    Set RS = cmd.Execute
    
 cmd.Properties("PLSQLRSet") = False
If Conn.State = 1 Then
       Conn.Close
       Set Conn = Nothing
End If
Set RS = Nothing
'      RS.Close
''''Exit Sub
Adodc3.ConnectionString = strcn.Connection_String
Adodc3.RecordSource = "SELECT MAX(REC_NO)AS REC_NO FROM RECEIPT_NO_COUNTER"
Adodc3.Refresh
If Adodc3.Recordset.RecordCount > 0 Then
   txtRecNo = Adodc3.Recordset!REC_NO
    PREVIEW_VAR = Val(txtRecNo)
    txtPrintPreview = txtRecNo
End If
Exit Sub
ErrDes:
MsgBox Err.Description, vbCritical, " IT, DNMIH"
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
    cmd.CommandText = "select bed_no from bed_info  where occupy_flag='0' and  BED_EXT_COL='" & Trim(CboTypeNo.Text) & "'and doc_department='" & Trim(cboBedDept.Text) & "' and Upper(bed_type)=upper('" & cboBedType & "') order by bed_no"
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




Private Sub print_admission(MODE As Integer)
              
              Dim Conn As New ADODB.Connection
              Dim cmd As New ADODB.Command
              Dim RS As New ADODB.Recordset
              Dim Param1 As New Parameter
              Dim RPTVIEWER As New Viewer
          If Conn.State = 0 Then
            Conn.Open strcn.Connection_String
          End If
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
            
            
            Dim Reporttran   As New CrystalReporttran
            Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 100, PREVIEW_VAR)
            cmd.Parameters.Append Param1 'comment
    
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL Rpt_in_dr_info_adm_print(?)}"
            Set RS = cmd.Execute
            cmd.Properties("PLSQLRSet") = False
            
            Reporttran.Database.SetDataSource RS
           If MODE = 1 Then
             Reporttran.PrintOut
           End If
           
           If Conn.State = 1 Then
            Conn.Close
            Set Conn = Nothing
           End If
           Set cmd = Nothing
           txtRecNo.Locked = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
End If
End Sub
Private Sub Form_Load()
txtRecNo.Locked = False
cboBedType = cboBedType.List(0)

MaskEdBox1.Text = Format(Date, "dd/mm/yyyy")
DTPicker7.Value = Time
txtRegNo = frmReg_for_Bed_Transf.txtReg_noOpr.Text
Call show_dept
cboBedDept.Text = "Common"

LOAD_PAT_INFO
LOAD_PAT_DEPT
LOAD_PAT_BED
PopulateFiscalYear

End Sub
Private Sub PopulateFiscalYear()
   Dim yearList() As String
   Dim i As Integer
   yearList = UTILITY.GetFiscalYears()
   For i = LBound(yearList) To UBound(yearList)
      CBOYRCODE.AddItem yearList(i)
   Next i
   CBOYRCODE.ListIndex = 0
End Sub
Private Sub LOAD_PAT_INFO()
  Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
    End If
        cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "select pat_name,age,y_m_d,admission_date,ADDR1  From in_door_pat_info_main Where in_reg_no ='" & Trim(frmReg_for_Bed_Transf.txtReg_noOpr.Text) & "' AND YRCODE ='" & Trim(frmReg_for_Bed_Transf.CBOYRCODE.Text) & "'"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False
          
       If RS.RecordCount > 0 Then
         txtPatName = "" & RS!pat_name
         txtAge = "" & RS!age
         TXTdmy = RS!Y_M_D
         txtAddress = RS!addr1
         MaskEdBox1.Text = IIf(RS!admission_date = Null, "__/__/____", Format(RS!admission_date, "dd/mm/yyyy"))
         
       End If
       
       If Conn.State = 1 Then
        Conn.Close
        Set Conn = Nothing
        Set RS = Nothing
        Set cmd = Nothing
     End If
End Sub
Private Sub LOAD_PAT_DEPT()
  Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
    End If
        cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "select doc_dept,serial_no  From INDOOR_PAT_DEPT_INFO Where in_reg_no ='" & Trim(frmReg_for_Bed_Transf.txtReg_noOpr.Text) & "' AND YRCODE ='" & Trim(frmReg_for_Bed_Transf.CBOYRCODE.Text) & "'  AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(frmReg_for_Bed_Transf.txtReg_noOpr.Text) & "' AND YRCODE='" & Trim(frmReg_for_Bed_Transf.CBOYRCODE) & "')"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False
          
       If RS.RecordCount > 0 Then
         txtPatDept = "" & RS!doc_dept
         var_dept_serial_no = RS!SERIAL_NO
       End If
       
       If Conn.State = 1 Then
        Conn.Close
        Set Conn = Nothing
        Set RS = Nothing
        Set cmd = Nothing
     End If

End Sub
Private Sub LOAD_PAT_BED()
  Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
    End If
        cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "select BED_TYPE,Bed_type_no,BED_NO,extra_bed_flag ,SERIAL_NO From Indoor_pat_bed_info Where in_reg_no ='" & Trim(frmReg_for_Bed_Transf.txtReg_noOpr.Text) & "' AND YRCODE ='" & Trim(frmReg_for_Bed_Transf.CBOYRCODE.Text) & "'  AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM Indoor_pat_bed_info WHERE in_reg_no ='" & Trim(frmReg_for_Bed_Transf.txtReg_noOpr.Text) & "' AND YRCODE='" & Trim(frmReg_for_Bed_Transf.CBOYRCODE) & "')"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False
          
       If RS.RecordCount > 0 Then
         txtbedType = "" & RS!Bed_type & " -  " & RS!bed_TYPE_no & " -  " & RS!bed_no
         txtExtraBedFlag = RS!Extra_bed_flag
         VAR_CUR_BED_SERIAL_NO = RS!SERIAL_NO
       End If
       
       If Conn.State = 1 Then
        Conn.Close
        Set Conn = Nothing
        Set RS = Nothing
        Set cmd = Nothing
     End If

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
     
      
      If Not RS.EOF Then
        RS.MoveFirst
        Do Until RS.EOF = True
           cboBedDept.AddItem RS!doc_dept
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
   Dim Conn As New ADODB.Connection
   Dim cmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
     If Conn.State = 0 Then
      Conn.ConnectionString = strcn.Connection_String
          Conn.Open
    End If
    cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
   cmd.CommandText = "select distinct(BED_EXT_COL) from bed_info  where  bed_type=  '" & cboBedType & "' and doc_department='" & Trim(cboBedDept.Text) & "' and occupy_flag='0' order by BED_EXT_COL"
    
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

Private Sub Frame5_DragDrop(Source As Control, X As Single, Y As Single)

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
   MsgBox "Invalid Age", vbInformation, " IT, DNMIH"
    txtAge = ""
End If

End Sub

Private Sub txtAge_GotFocus()
   txtAge.BackColor = &H80000018
End Sub

Private Sub txtAge_LostFocus()
txtAge.BackColor = vbWhite
End Sub


Private Sub txtPatName_GotFocus()
  txtPatName.BackColor = &H80000018
End Sub

Private Sub txtPatName_LostFocus()
  txtPatName.BackColor = vbWhite
End Sub

Private Sub txtPrintPreview_Change()
  If Not IsNumeric(txtPrintPreview) Then
     txtPrintPreview = ""
  End If
End Sub
