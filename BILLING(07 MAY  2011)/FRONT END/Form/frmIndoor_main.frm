VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIndoor_main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000001&
      Height          =   765
      Left            =   -30
      TabIndex        =   66
      Top             =   1020
      Width           =   11355
      Begin VB.TextBox TicketNoText 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2190
         TabIndex        =   71
         Top             =   300
         Width           =   3135
      End
      Begin VB.ComboBox YearCombo 
         Height          =   315
         Left            =   9060
         TabIndex        =   68
         Text            =   "Combo1"
         Top             =   270
         Width           =   1665
      End
      Begin VB.ComboBox MonthCombo 
         Height          =   315
         Left            =   6930
         TabIndex        =   67
         Text            =   "Combo1"
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ticket  No  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   540
         TabIndex        =   72
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   240
         Left            =   8400
         TabIndex        =   70
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   240
         Left            =   6090
         TabIndex        =   69
         Top             =   300
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   59
      Top             =   9000
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
         TabIndex        =   61
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
         TabIndex        =   60
         ToolTipText     =   "Developed && Maintenanced by:A.N.M. Atiqur Rahman,Software Programmer, IT Division, DNMIH"
         Top             =   90
         Width           =   2715
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1215
      Left            =   -90
      TabIndex        =   17
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   360
         Width           =   4665
      End
      Begin VB.Image Image1 
         Height          =   1155
         Left            =   30
         Picture         =   "frmIndoor_main.frx":0000
         Stretch         =   -1  'True
         Top             =   180
         Width           =   12180
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   -60
      TabIndex        =   51
      Top             =   8070
      Width           =   11385
      Begin VB.CommandButton cmdPrepareAdmissionForm 
         Caption         =   "PREPARE ADMISSION FORM"
         Height          =   375
         Left            =   2280
         TabIndex        =   62
         Top             =   210
         Width           =   2625
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
         TabIndex        =   57
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
         TabIndex        =   56
         Top             =   270
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.CommandButton CMDEXIT 
         Cancel          =   -1  'True
         Caption         =   "CLOSE"
         Height          =   375
         Left            =   9900
         TabIndex        =   55
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton CMDREPORT 
         Caption         =   "REPORT"
         Height          =   375
         Left            =   8670
         TabIndex        =   54
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton ResetPrinterButton 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "Reset Printer"
         Height          =   375
         Left            =   7440
         MaskColor       =   &H008080FF&
         TabIndex        =   53
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdADD 
         Caption         =   "NEW"
         Height          =   375
         Left            =   6210
         TabIndex        =   52
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "SAVE"
         Height          =   375
         Left            =   4980
         TabIndex        =   16
         Top             =   210
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
         Left            =   2190
         Top             =   150
         Width           =   8985
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   120
         Picture         =   "frmIndoor_main.frx":5982
         Stretch         =   -1  'True
         Top             =   -60
         Width           =   12675
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   -90
      TabIndex        =   20
      Top             =   1800
      Width           =   11385
      Begin VB.ComboBox cboStaff 
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
         Left            =   570
         TabIndex        =   15
         Top             =   5790
         Width           =   1845
      End
      Begin VB.ComboBox cboPrinters 
         Height          =   315
         Left            =   3210
         TabIndex        =   63
         Text            =   "Combo1"
         Top             =   4710
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.CheckBox chkCareOf 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Care Of"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   9690
         MaskColor       =   &H00FFFF80&
         TabIndex        =   58
         Top             =   2100
         Width           =   2295
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
      Begin VB.ComboBox cboDMY 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmIndoor_main.frx":B304
         Left            =   1470
         List            =   "frmIndoor_main.frx":B311
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
         ItemData        =   "frmIndoor_main.frx":B31E
         Left            =   5550
         List            =   "frmIndoor_main.frx":B320
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3270
         Width           =   5235
      End
      Begin VB.ComboBox CBOYRCODE 
         Height          =   315
         ItemData        =   "frmIndoor_main.frx":B322
         Left            =   5610
         List            =   "frmIndoor_main.frx":B324
         Locked          =   -1  'True
         TabIndex        =   45
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
         ItemData        =   "frmIndoor_main.frx":B326
         Left            =   4350
         List            =   "frmIndoor_main.frx":B339
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3270
         Width           =   1005
      End
      Begin VB.ComboBox cboSex 
         Height          =   315
         ItemData        =   "frmIndoor_main.frx":B367
         Left            =   2670
         List            =   "frmIndoor_main.frx":B371
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   7500
         MaskColor       =   &H00FFFF80&
         TabIndex        =   38
         Top             =   2100
         Width           =   2145
      End
      Begin VB.CheckBox chkFather 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Father's Name"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
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
         Width           =   1875
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
         ItemData        =   "frmIndoor_main.frx":B37B
         Left            =   2670
         List            =   "frmIndoor_main.frx":B37D
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
         TabIndex        =   34
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
         TabIndex        =   32
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
         TabIndex        =   30
         Top             =   1470
         Width           =   825
      End
      Begin VB.ComboBox cboBedNo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmIndoor_main.frx":B37F
         Left            =   5610
         List            =   "frmIndoor_main.frx":B381
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1470
         Width           =   1365
      End
      Begin VB.ComboBox CboTypeNo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmIndoor_main.frx":B383
         Left            =   4380
         List            =   "frmIndoor_main.frx":B385
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1470
         Width           =   975
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   7350
         TabIndex        =   26
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
         TabIndex        =   24
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
         ItemData        =   "frmIndoor_main.frx":B387
         Left            =   570
         List            =   "frmIndoor_main.frx":B394
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1470
         Width           =   1845
      End
      Begin VB.TextBox txtRecNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   570
         TabIndex        =   21
         Top             =   510
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker7 
         Height          =   315
         Left            =   9540
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Delevary Time"
         Top             =   480
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   60358658
         UpDown          =   -1  'True
         CurrentDate     =   37163
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   630
         TabIndex        =   65
         Top             =   5520
         Width           =   825
      End
      Begin VB.Label lblStaffName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2790
         TabIndex        =   64
         Top             =   5820
         Width           =   8385
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   33
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
         TabIndex        =   31
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
         TabIndex        =   29
         Top             =   1140
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
         Left            =   4320
         TabIndex        =   28
         Top             =   1140
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
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Index           =   2
         Left            =   2670
         TabIndex        =   25
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   1140
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmIndoor_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn As New Connection
Dim RS As New Recordset
Public rspr As New Recordset
Dim VoucherNumber
Dim UTILITY As New clsUtility
Public strUid As String
Dim checkNameIndicator As Integer
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
    MsgBox Err.Description, vbInformation, " IT, DNMIH"
End Sub

Private Sub cboBedDept_Click()
  txtAdmissionfee = ""
  txtBedCharge = ""
  txtServiceFee = ""
  txtSerialNo = ""
  Call LOAD_BED_NO
  End Sub

Private Sub cboBedDept_GotFocus()
    cboBedDept_Click
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

Private Sub cboPatdept_Click()
  If cboBedType = "Free-Bed" Then
     txtServiceFee = UTILITY.Load_Service_Charge_Exceptional(cboBedType, cboPatdept)
  End If
  Gynae_validation
End Sub
Private Sub Gynae_validation()
  If UCase(cboPatdept.Text) = UCase("Gynae-1") Or UCase(cboPatdept.Text) = UCase("Gynae-2") Then
       If UCase(cboSex.Text) = UCase("M") Then
         cboSex.Text = "F"
       End If
    End If
End Sub
Private Sub Load_Service_Charge_Exceptional()
   If cboBedType = "Free-Bed" And cboBedDept = "COMMON" Then
     If cboPatdept = "Ophth." Or cboPatdept = "Gynae-1" Or cboPatdept = "Gynae-2" Or cboPatdept = "ENT" Or cboPatdept = "Surgery-1" Or cboPatdept = "Surgery-2" Then
        txtServiceFee = 250
     Else
       txtServiceFee = 0
     End If
   End If
End Sub

Private Sub cboStaff_Change()
   GetStaffName
End Sub
Private Sub GetStaffName()
  If UTILITY.LOAD_STAFF(cboStaff.Text) = "0" Then
     lblStaffName.Caption = "INVALID STAFF ID,...PLEASE VERIFY"
     lblStaffName.ForeColor = vbRed
  Else
     cboStaff.Text = UCase(cboStaff.Text)
     lblStaffName.Caption = UTILITY.LOAD_STAFF(cboStaff.Text)
     lblStaffName.ForeColor = vbWhite
  End If
End Sub

Private Sub cboStaff_Click()
  GetStaffName
End Sub

Private Sub chkCareOf_Click()
 If chkCareOf.Value = 1 Then
    txtPatFatherName = "C/O:"
    chkFather.Value = 0
    chkHusband.Value = 0
    chkFather.ForeColor = vbWhite
  Else
    chkFather.Enabled = True
    chkFather.Value = 1
    chkHusband.ForeColor = vbWhite
    chkFather.ForeColor = &HFFFF80

   End If
End Sub

Private Sub chkFather_Click()
  If chkFather.Value = 1 Then
    chkHusband.Value = 0
    chkCareOf.Value = 0
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
     chkCareOf.Value = 0
     chkFather.ForeColor = vbWhite
  Else
    chkFather.Enabled = True
    chkFather.Value = 0
    chkCareOf.Value = 1
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
cboBedType = cboBedType.List(0)
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

txtAdvance = ""
txtAge = ""

txtPhone = ""
cboStaff.Text = ""
cboBedType.SetFocus
End Sub

Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Sure to close?", vbQuestion + vbYesNo, "Close...")
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
MsgBox Err.Description, vbCritical, " IT, DNMIH"
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
Private Sub AssignCanonPrinterName()
   If canonPrinterName = "" Then
        Dim i As Integer
        For i = 0 To cboPrinters.ListCount - 1
            If Left$(cboPrinters.List(i), 13) = "Canon LBP3300" Then
              canonPrinterName = cboPrinters.List(i)
              Exit For
            End If
            If Right$(cboPrinters.List(i), 13) = "Canon LBP3300" Then
              canonPrinterName = cboPrinters.List(i)
            End If
      
         Next i
   End If
 End Sub

Private Sub cmdPrepareAdmissionForm_Click()
  
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer
   Dim printerName As String

    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)

    If osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        Call Win95SetDefaultPrinter
    Else
    ' This assumes that future versions of Windows use the NT method
       If canonPrinterName = "" Then
                 AssignCanonPrinterName
       End If
          Call WinNTSetDefaultPrinter(canonPrinterName)
    End If

'
If txtPrintPreview = "" Then
   MsgBox "Please put a Money Receipt No first", vbInformation, "IT,DNMIH"
   CMDREPORT_Click
   Exit Sub
End If
PREVIEW_VAR = Val(txtPrintPreview)
''Dim X As Printer
''Set Printer = X
''X.DeviceName = cboPrinters.Text
''X.PaperSize = vbPRPSA4
'
'
              Dim Conn As New Connection
              Dim cmd As New Command
              Dim RS As New ADODB.Recordset
              Dim Param1 As New Parameter

          If Conn.State = 0 Then
            Conn.Open strcn.Connection_String
          End If
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText

            Dim Report6   As New CrystalReport32
            Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 100, PREVIEW_VAR)
            cmd.Parameters.Append Param1 '

            cmd.Properties("PLSQLRSet") = True

            cmd.CommandText = "{CALL Rpt_in_dr_info_adm_print(?)}"
            Set RS = cmd.Execute
            cmd.Properties("PLSQLRSet") = False
'            Report6.Text4.Width = 1650
'            Report6.Text4.SetText ("Admission")

            Report6.Database.SetDataSource RS




          Report6.PrintOut True



           If Conn.State = 1 Then
            Conn.Close
            Set Conn = Nothing
            Set Report6 = Nothing
            Set RS = Nothing
           End If
           PREVIEW_VAR = 0
           txtRecNo.Locked = False
           Call WinNTSetDefaultPrinter("EPSON LQ-300")
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
   PREVIEW_VAR = Val(txtPrintPreview)
        rptMode = 6
      Viewer.Show vbModal
      txtPrintPreview = ""
      txtPrintPreview.Visible = False
End If
End Sub


Private Sub cmdSave_Click()
          If UTILITY.User_Shift_validation(frmMAIN.lbluser_id, frmMAIN.lblUserType.Caption) = False Then
            MsgBox "Mr. " & frmMAIN.lblName & "  Your Shift has been Expired.. " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
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
          MsgBox "Please Select a Patient Department", vbInformation, " IT, DNMIH"
          cboPatdept.SetFocus
          Exit Sub
      End If
      
      If UCase(cboPatdept.Text) = UCase("CCU") Or UCase(cboPatdept.Text) = UCase("Pathology") Or UCase(cboPatdept.Text) = UCase("Radiology") Or UCase(cboPatdept.Text) = UCase("Anaes.logy") Then
          MsgBox "Please Select a Valid Patient Department", vbInformation, " IT Division, DNMIH"
          cboPatdept.SetFocus
          Exit Sub
      End If
        
     If UCase(cboPatdept.Text) = UCase("Gynae-1") Or UCase(cboPatdept.Text) = UCase("Gynae-2") Then
       If UCase(cboSex.Text) = UCase("M") Then
           MsgBox "Please Select Sex as Female for Gynae Patient", vbInformation, " IT, DNMIH"
           cboSex.SetFocus
         Exit Sub
       End If
    End If
    
    If Len(cboStaff) > 0 Then
       If UTILITY.LOAD_STAFF(cboStaff.Text) = "0" Then
             lblStaffName.Caption = "INVALID STAFF ID,...PLEASE VERIFY"
             lblStaffName.ForeColor = vbRed
             cboStaff.SetFocus
             Exit Sub
       End If
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
     txtServiceFee = UTILITY.Load_Service_Charge_Exceptional(cboBedType, cboPatdept)
  End If
  
  If chkFather.Value = 1 Then
      checkNameIndicator = 1 ''father
  ElseIf chkHusband.Value = 1 Then ''husband
     checkNameIndicator = 0  ''husband
  Else
    checkNameIndicator = 2
  End If
  
  
  
Call save_In_door_pat_ADMISSION

MsgBox "Operation successful", vbInformation + vbOKOnly, "Save..."
print_admission
cmdPrepareAdmissionForm_Click
Call Clear_form

If Conn.State = 1 Then
 Conn.Close
 Set Conn = Nothing
End If
Set UTILITY = Nothing

End Sub
Private Sub save_In_door_pat_ADMISSION()
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
    Dim Param24 As New Parameter
    
    
   If Conn.State = 0 Then
      Conn.Open strcn.Connection_String
  End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, Trim(cboBedType.Text))
    cmd.Parameters.Append Param1 'bed type
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 25, Trim(CboTypeNo.Text))
    cmd.Parameters.Append Param2 'Type_no
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 20, cboBedNo.Text)
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

    
    Set Param18 = cmd.CreateParameter("param18", adInteger, adParamInput, 10, checkNameIndicator)
    cmd.Parameters.Append Param18 'check_value----father's name or husband name
    
    Set Param19 = cmd.CreateParameter("param19", adVarChar, adParamInput, 10, Trim(CBOYRCODE))
    cmd.Parameters.Append Param19 'YR_CODE
    
    Set Param20 = cmd.CreateParameter("param20", adDouble, adParamInput, 10, Trim(txtAdmissionfee.Text))
    cmd.Parameters.Append Param20 'ADMISSION CHARGE
 
    Set Param21 = cmd.CreateParameter("param21", adDouble, adParamInput, 10, Trim(txtBedCharge.Text))
    cmd.Parameters.Append Param21 'BED CHARGE
 
   Set Param22 = cmd.CreateParameter("param22", adDouble, adParamInput, 10, Trim(txtServiceFee.Text))
   cmd.Parameters.Append Param22 'SERVICE CHARGE
 
   If cboBedDept = "COMMON" And cboBedType <> "Cabin" Then
      Set Param23 = cmd.CreateParameter("param23", adDouble, adParamInput, 5, 1)
      cmd.Parameters.Append Param23 'EXTRA BED FLAG INDICATOR
   Else
      Set Param23 = cmd.CreateParameter("param23", adDouble, adParamInput, 5, 0)
      cmd.Parameters.Append Param23 'EXTRA BED FLAG INDICATOR
   
   End If
   
   Set Param24 = cmd.CreateParameter("param24", adVarChar, adParamInput, 13, Trim(cboStaff.Text))
   cmd.Parameters.Append Param24 'STAFF ID
 
   
   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Indoor_SavePatient_info(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
      
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
        PREVIEW_VAR = Val(txtRecNo)
        txtPrintPreview = Val(txtRecNo)
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









Private Sub Command10_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
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
    MsgBox Err.Description, vbInformation, " IT, DNMIH"

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
Private Sub print_admission()
              
              Dim Conn As New Connection
              Dim cmd As New Command
              Dim RS As New ADODB.Recordset
              Dim Param1 As New Parameter
              
          If Conn.State = 0 Then
            Conn.Open strcn.Connection_String
          End If
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
                 
            Dim Report6   As New CrystalReporttran
            Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 100, PREVIEW_VAR)
            cmd.Parameters.Append Param1 '
                 
            cmd.Properties("PLSQLRSet") = True
            
            cmd.CommandText = "{CALL Rpt_in_dr_info_adm_print(?)}"
            Set RS = cmd.Execute
            cmd.Properties("PLSQLRSet") = False
            Report6.Text4.Width = 1650
            Report6.Text4.SetText ("Admission")
      
            Report6.Database.SetDataSource RS

            Report6.PrintOut
            
            
           If Conn.State = 1 Then
            Conn.Close
            Set Conn = Nothing
            Set Report6 = Nothing
            Set RS = Nothing
           End If
           PREVIEW_VAR = 0
           txtRecNo.Locked = False
End Sub
Private Sub Command15_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub
Private Sub Command3_Click()
    rptMode = 6
    Viewer.Show vbModal
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdShowBed_Click()
 
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
End If
End Sub
Private Sub Form_Load()
txtRecNo.Locked = False
cboBedType = cboBedType.List(0)
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
LoadStaffID

    Dim r As Long
    Dim Buffer As String

    ' Get the list of available printers from WIN.INI
    Buffer = Space(8192)
    r = GetProfileString("PrinterPorts", vbNullString, "", _
       Buffer, Len(Buffer))

    ' Display the list of printer in the ListBox List1
    ParseList cboPrinters, Buffer
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
Private Sub LoadStaffID()
 
  '''''''''FOR ID
     Adodc1.ConnectionString = strcn.Connection_String
      Adodc1.RecordSource = "select PAYROLL.EMP_INFO.EMP_ID  AS EMP_ID from PAYROLL.EMP_INFO ORDER BY EMP_ID"
      Adodc1.Refresh
    
    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
          cboStaff.AddItem Adodc1.Recordset!EMP_ID
          Adodc1.Recordset.MoveNext
        Wend
    End If
End Sub

Private Sub getPrinters()
'  Dim i As Integer
'  For i = 0 To cboPrinters.ListCount - 1
'      MsgBox Right$(cboPrinters.List(i), 13)
'
'  Next i
'  cboPrinters.clear
'  Dim x As Printer
'For Each x In Printers
'   cboPrinters.AddItem x.DeviceName
''Label1.Caption = "Please select a default printer for this app."
'Next x
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




Private Sub ResetPrinterButton_Click()
   Call WinNTSetDefaultPrinter("EPSON LQ-300")
End Sub

Private Sub txtAdvance_Change()
  If Not IsNumeric(txtAdvance) Then
     txtAdvance = ""
 End If
 
End Sub
Private Sub txtAdvance_GotFocus()
     With txtAdvance
       .BackColor = &H80000018
       .SelStart = 0
       .SelLength = Len(txtAdvance.Text)
     End With
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


Private Sub txtPhone_GotFocus()
      txtPhone.BackColor = &H80000018
End Sub

Private Sub txtPhone_LostFocus()
      txtPhone.BackColor = vbWhite
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

Private Sub txtPrintPreview_Change()
  If Not IsNumeric(txtPrintPreview) Then
     txtPrintPreview = ""
  End If
End Sub
Private Function PtrCtoVbString(Add As Long) As String
    Dim sTemp As String * 512, X As Long

    X = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Private Sub SetDefaultPrinter(ByVal printerName As String, _
    ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim l As Long
    DeviceLine = printerName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub

Private Sub Win95SetDefaultPrinter()
    Dim Handle As Long          'handle to printer
    Dim printerName As String
    Dim pd As PRINTER_DEFAULTS
    Dim X As Long
    Dim need As Long            ' bytes needed
    Dim pi5 As PRINTER_INFO_5   ' your PRINTER_INFO structure
    Dim LastError As Long

    ' determine which printer was selected
'    printerName = List1.List(List1.ListIndex)
    ' none - exit
    If printerName = "" Then
        Exit Sub
    End If

    ' set the PRINTER_DEFAULTS members
    pd.pDatatype = 0&
    pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess

    ' Get a handle to the printer
    X = OpenPrinter(printerName, Handle, pd)
    ' failed the open
    If X = False Then
        'error handler code goes here
        Exit Sub
    End If

    ' Make an initial call to GetPrinter, requesting Level 5
    ' (PRINTER_INFO_5) information, to determine how many bytes
    ' you need
    X = GetPrinter(Handle, 5, ByVal 0&, 0, need)
    ' don't want to check Err.LastDllError here - it's supposed
    ' to fail
    ' with a 122 - ERROR_INSUFFICIENT_BUFFER
    ' redim t as large as you need
    ReDim t((need \ 4)) As Long

    ' and call GetPrinter for keepers this time
    X = GetPrinter(Handle, 5, t(0), need, need)
    ' failed the GetPrinter
    If X = False Then
        'error handler code goes here
        Exit Sub
    End If

    ' set the members of the pi5 structure for use with SetPrinter.
    ' PtrCtoVbString copies the memory pointed at by the two string
    ' pointers contained in the t() array into a Visual Basic string.
    ' The other three elements are just DWORDS (long integers) and
    ' don't require any conversion
    pi5.pPrinterName = PtrCtoVbString(t(0))
    pi5.pPortName = PtrCtoVbString(t(1))
    pi5.Attributes = t(2)
    pi5.DeviceNotSelectedTimeout = t(3)
    pi5.TransmissionRetryTimeout = t(4)

    ' this is the critical flag that makes it the default printer
    pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT

       ' call SetPrinter to set it
       X = SetPrinter(Handle, 5, pi5, 0)

       If X = False Then   ' SetPrinter failed
           MsgBox "SetPrinter Failed. Error code: " & Err.LastDllError
           Exit Sub
       Else
'           If Printer.DeviceName <> List1.Text Then
'           ' Make sure Printer object is set to the new printer
'                SelectPrinter (List1.Text)
'           End If
       End If

    ' and close the handle
    ClosePrinter (Handle)
End Sub

Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
    String, PrinterPort As String)

    Dim iDriver As Integer
    Dim iPort As Integer
    DriverName = ""
    PrinterPort = ""

    ' The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

         ' Strip out the driver name
        DriverName = Left(Buffer, iDriver - 1)

        ' The port name is the second entry after the driver name
        ' separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            ' Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    End If
End Sub

Private Sub ParseList(lstCtl As Control, ByVal Buffer As String)
    Dim i As Integer
    Dim s As String

    Do
        i = InStr(Buffer, Chr(0))
        If i > 0 Then
            s = Left(Buffer, i - 1)
            If Len(Trim(s)) Then lstCtl.AddItem s
            Buffer = Mid(Buffer, i + 1)
        Else
            If Len(Trim(Buffer)) Then lstCtl.AddItem Buffer
            Buffer = ""
        End If
    Loop While i > 0
End Sub
Private Sub WinNTSetDefaultPrinter(printerName As String)
    Dim Buffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim r As Long
    If Len(printerName) > 0 Then
        ' Get the printer information for the currently selected
        ' printer in the list. The information is taken from the
        ' WIN.INI file.
        Buffer = Space(1024)
        
        r = GetProfileString("PrinterPorts", printerName, "", _
            Buffer, Len(Buffer))

        ' Parse the driver name and port name out of the buffer
        GetDriverAndPort Buffer, DriverName, PrinterPort

           If DriverName <> "" And PrinterPort <> "" Then
               SetDefaultPrinter printerName, DriverName, PrinterPort
               If Printer.DeviceName <> cboPrinters.Text Then
               ' Make sure Printer object is set to the new printer
                  SelectPrinter (printerName)
               End If
           End If
    End If
End Sub


