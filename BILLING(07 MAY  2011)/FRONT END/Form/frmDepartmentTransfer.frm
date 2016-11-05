VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDepartmentTransfer 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Patient Release Form"
   ClientHeight    =   10740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13770
   FillColor       =   &H00800000&
   ForeColor       =   &H80000001&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10740
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   -240
      TabIndex        =   96
      Top             =   10290
      Width           =   16845
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Developed By:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   1440
         TabIndex        =   98
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Software Programmer, IT Division,DNMIH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3360
         TabIndex        =   97
         Top             =   180
         Width           =   4965
      End
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   990
      Top             =   8310
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
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
   Begin VB.TextBox txtprintPreview 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9210
      TabIndex        =   88
      Top             =   8220
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   10470
      TabIndex        =   69
      Top             =   8550
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   9240
      TabIndex        =   68
      Top             =   8550
      Width           =   1215
   End
   Begin VB.CommandButton CMDSHOWBED 
      Caption         =   "SHOW BED"
      Height          =   375
      Left            =   8010
      TabIndex        =   67
      Top             =   8550
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "NEW"
      Height          =   375
      Left            =   6780
      TabIndex        =   66
      Top             =   8550
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   5550
      TabIndex        =   22
      Top             =   8550
      Width           =   1215
   End
   Begin VB.TextBox txtSerialNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2730
      TabIndex        =   64
      Top             =   7050
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Caption         =   "Bed Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   2280
      Left            =   120
      TabIndex        =   56
      Top             =   870
      Width           =   4380
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1995
         Left            =   60
         TabIndex        =   76
         Top             =   240
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   3519
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   14737632
         BackColorBkg    =   -2147483647
         FocusRect       =   0
         HighLight       =   2
         AllowUserResizing=   2
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000001&
      Caption         =   "Operation Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1230
      Left            =   120
      TabIndex        =   54
      Top             =   3180
      Width           =   7170
      Begin MSDataGridLib.DataGrid DataGrid5 
         Height          =   915
         Left            =   90
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   210
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   1614
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   -2147483624
         BorderStyle     =   0
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000001&
      Caption         =   "Extra Bed Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   2295
      Left            =   4500
      TabIndex        =   53
      Top             =   870
      Width           =   2670
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   1965
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   3466
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorFixed  =   14737632
         BackColorBkg    =   -2147483647
         FocusRect       =   0
         HighLight       =   2
         AllowUserResizing=   2
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000001&
      Caption         =   "  DEPARTMENT && BED TRANSFER"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   3825
      Left            =   120
      TabIndex        =   52
      Top             =   4440
      Width           =   7155
      Begin VB.TextBox txtAdvanceRelease 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4710
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   2910
         Width           =   2310
      End
      Begin VB.TextBox txtDepartmentSerial 
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
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   3420
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox ChkBedDepartmentIndicator 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Both Bed  &&  Department Transfer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   390
         Value           =   1  'Checked
         Width           =   4875
      End
      Begin VB.TextBox txtExtraBedFlag 
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   3300
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtSetupServiceFee 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5640
         TabIndex        =   71
         Text            =   "0"
         Top             =   3390
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtsetupAdmissionfee 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4650
         TabIndex        =   70
         Text            =   "0"
         Top             =   3360
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtAdvance 
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
         Height          =   285
         Left            =   210
         TabIndex        =   21
         Top             =   2910
         Width           =   2025
      End
      Begin VB.ComboBox Trans_Dept 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2010
         Width           =   6855
      End
      Begin VB.TextBox txtSetupBedCharge 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6000
         TabIndex        =   62
         Text            =   "0"
         Top             =   990
         Width           =   1035
      End
      Begin VB.ComboBox cboBedNo 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   990
         Width           =   1095
      End
      Begin VB.ComboBox CboBedTypeNo 
         Height          =   315
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   990
         Width           =   1035
      End
      Begin VB.ComboBox CboBedDept 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmDepartmentTransfer.frx":0000
         Left            =   1680
         List            =   "frmDepartmentTransfer.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   990
         Width           =   2235
      End
      Begin VB.ComboBox CboBedType 
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
         ItemData        =   "frmDepartmentTransfer.frx":002A
         Left            =   180
         List            =   "frmDepartmentTransfer.frx":0037
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   990
         Width           =   1515
      End
      Begin VB.Label Label92 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Advance"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   270
         Left            =   4740
         TabIndex        =   90
         Top             =   2580
         Width           =   2205
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEPT_SERIAL_NO"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   210
         Index           =   11
         Left            =   150
         TabIndex        =   87
         Top             =   3240
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extra Bed Flag"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   210
         Index           =   10
         Left            =   2550
         TabIndex        =   75
         Top             =   3120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Srvc. Fee"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   210
         Index           =   8
         Left            =   5700
         TabIndex        =   73
         Top             =   3210
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adm. Fee"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   210
         Index           =   7
         Left            =   4680
         TabIndex        =   72
         Top             =   3150
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   270
         Index           =   6
         Left            =   210
         TabIndex        =   65
         Top             =   2580
         Width           =   1065
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Charge"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   210
         Index           =   5
         Left            =   5970
         TabIndex        =   63
         Top             =   750
         Width           =   1080
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bed No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   210
         Index           =   4
         Left            =   4950
         TabIndex        =   61
         Top             =   750
         Width           =   645
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cab/Ward"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   210
         Index           =   3
         Left            =   3840
         TabIndex        =   60
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Dept."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   210
         Index           =   2
         Left            =   1770
         TabIndex        =   59
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---------------TRANSFER     TO     DEPARTMENT--------------"
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
         Index           =   1
         Left            =   600
         TabIndex        =   58
         Top             =   1710
         Width           =   5775
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Type"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   210
         Index           =   0
         Left            =   210
         TabIndex        =   57
         Top             =   750
         Width           =   870
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2310
      Top             =   8340
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
      Caption         =   "Adodc2"
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
   Begin VB.Frame Frame6 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   7530
      Left            =   7290
      TabIndex        =   30
      Top             =   810
      Width           =   4965
      Begin VB.TextBox txtIncubator 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   8
         Text            =   "0"
         Top             =   4395
         Width           =   1590
      End
      Begin VB.TextBox txtNebuliser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   10
         Text            =   "0"
         Top             =   5160
         Width           =   1590
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   1
         Text            =   "0"
         Top             =   1845
         Width           =   1590
      End
      Begin VB.TextBox txtCCU_Charge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   9
         Text            =   "0"
         Top             =   4770
         Width           =   1590
      End
      Begin VB.TextBox txtDeliveryCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "0"
         Top             =   2205
         Width           =   1590
      End
      Begin VB.TextBox txtdisc_percent 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3630
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "0"
         Top             =   6690
         Width           =   570
      End
      Begin VB.TextBox txtMedicine_charge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   12
         Text            =   "0"
         Top             =   5940
         Width           =   1590
      End
      Begin VB.TextBox txtP_therapyCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   6
         Text            =   "0"
         Top             =   3645
         Width           =   1590
      End
      Begin VB.TextBox txtBloodTher_charge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   7
         Text            =   "0"
         Top             =   4020
         Width           =   1590
      End
      Begin VB.TextBox txtExtTransfusion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   5
         Text            =   "0"
         Top             =   3270
         Width           =   1590
      End
      Begin VB.TextBox txtNeunetalCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   4
         Text            =   "0"
         Top             =   2910
         Width           =   1590
      End
      Begin VB.TextBox txtBabyCareCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   3
         Text            =   "0"
         Top             =   2580
         Width           =   1590
      End
      Begin VB.TextBox txtDisc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   23
         Text            =   "0"
         Top             =   6690
         Width           =   690
      End
      Begin VB.TextBox txtTotalBedCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         Height          =   300
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   720
         Width           =   1590
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   6330
         Width           =   1590
      End
      Begin VB.TextBox txtmisce 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   11
         Text            =   "0"
         Top             =   5550
         Width           =   1590
      End
      Begin VB.TextBox txtDUE_TOTAL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   7050
         Width           =   1590
      End
      Begin VB.TextBox txtExtraBedTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         Height          =   300
         Left            =   2625
         TabIndex        =   94
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1080
         Width           =   1590
      End
      Begin VB.TextBox txtTotalOpr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2625
         MaxLength       =   8
         TabIndex        =   0
         Text            =   "0"
         Top             =   1470
         Width           =   1590
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FFFFFF&
         Height          =   7305
         Left            =   30
         Top             =   150
         Width           =   4275
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTMENT TRANSFER"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   435
         Left            =   210
         TabIndex        =   95
         Top             =   270
         Width           =   4665
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Due"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   92
         Top             =   7110
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(---)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1020
         TabIndex        =   91
         Top             =   6810
         Width           =   315
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incubator  Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   150
         TabIndex        =   51
         Top             =   4425
         Width           =   1710
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nebulizer Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   50
         Top             =   5190
         Width           =   1665
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Anaesthesia Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   60
         TabIndex        =   49
         Top             =   1875
         Width           =   1965
      End
      Begin VB.Label lab10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CCU Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   48
         Top             =   4800
         Width           =   1155
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   47
         Top             =   2235
         Width           =   1530
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3420
         TabIndex        =   45
         Top             =   6750
         Width           =   210
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   42
         Top             =   5955
         Width           =   1620
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Sugar Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   41
         Top             =   4050
         Width           =   1935
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Photo Therapy Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   40
         Top             =   3675
         Width           =   2145
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ex.Transfusion Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   39
         Top             =   3300
         Width           =   2145
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Neunetal Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   38
         Top             =   2940
         Width           =   1605
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Baby-care Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   37
         Top             =   2610
         Width           =   1710
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   36
         Top             =   6750
         Width           =   810
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   35
         Top             =   750
         Width           =   1125
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Miscelleneous Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   34
         Top             =   5580
         Width           =   2100
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   33
         Top             =   6360
         Width           =   540
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Ext. BedCharge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   32
         Top             =   1110
         Width           =   2025
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Operation Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   31
         Top             =   1500
         Width           =   2235
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7485
      Top             =   3570
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
   Begin VB.Frame Frame4 
      BackColor       =   &H80000001&
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   -60
      Width           =   11925
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   8880
         TabIndex        =   84
         Top             =   180
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   -2147483647
         ForeColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "DD-MMM-YYYY"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox comDepartmentRelease 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4140
         TabIndex        =   81
         Top             =   495
         Width           =   2205
      End
      Begin VB.TextBox TXTbEDTYPE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   495
         Width           =   1545
      End
      Begin VB.TextBox txtRegNoShow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1620
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   495
         Width           =   1035
      End
      Begin VB.TextBox txtNameRelease 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   180
         Width           =   3975
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3990
         TabIndex        =   85
         Top             =   510
         Width           =   105
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   8700
         TabIndex        =   83
         Top             =   510
         Width           =   105
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   8670
         TabIndex        =   82
         Top             =   180
         Width           =   105
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Dept."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   240
         Left            =   2700
         TabIndex        =   80
         Top             =   510
         Width           =   1335
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Bed No "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   240
         Index           =   9
         Left            =   6570
         TabIndex        =   79
         Top             =   540
         Width           =   1530
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1470
         TabIndex        =   46
         Top             =   180
         Width           =   105
      End
      Begin VB.Label lblRegNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   495
         Width           =   1395
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   180
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   240
         Index           =   1
         Left            =   6540
         TabIndex        =   28
         Top             =   195
         Width           =   1500
      End
   End
   Begin VB.Shape Shape4 
      Height          =   465
      Left            =   5460
      Top             =   8490
      Width           =   6285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Total"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7080
      TabIndex        =   25
      Top             =   9480
      Visible         =   0   'False
      Width           =   660
   End
End
Attribute VB_Name = "frmDepartmentTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UTILITY As New clsUtility
Dim Conn As New Connection
Dim cmd As New Command
Dim total_bed_charge As Double
Public Extra_bed_Flag_Indicator As Integer
Public total_EXTRA_bed_charge As Double
Dim PATIENT_BED_SERIAL_NO As Integer
Dim discount_check_val As Integer
Public strUid As String
Public DEPARTMENT_MODE As Integer
Public strcn        As New MyConnection

Private Sub cboBedDept_GotFocus()
   cboBedDept_Click
End Sub
Private Sub ChkBedDepartmentIndicator_Click()
  If ChkBedDepartmentIndicator.Value = 0 Then
     ChkBedDepartmentIndicator.ForeColor = vbCyan
     ChkBedDepartmentIndicator.Caption = "Department Transfer Only"
  Else
     ChkBedDepartmentIndicator.Caption = "Both Bed  &&  Department Transfer "
     ChkBedDepartmentIndicator.ForeColor = vbWhite
     
  End If
End Sub

Private Sub ChkBedDepartmentIndicator_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If ChkBedDepartmentIndicator.Value = 0 Then
        cboBedNo.SetFocus
     Else
        ChkBedDepartmentIndicator.SetFocus
     End If
  End If
End Sub

Private Sub cmdADD_Click()
   Unload Me
End Sub
Private Sub LOAD_BED_NO()
   Dim Conn As New ADODB.Connection
   Dim cmd As New ADODB.Command
   Dim RS As New ADODB.Recordset
   
   txtSerialNo = ""
   txtSetupBedCharge = ""
    If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
    End If
  
    cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "select bed_no from bed_info  where occupy_flag='0' and  BED_EXT_COL='" & Trim(CboBedTypeNo.Text) & "'and doc_department='" & Trim(cboBedDept.Text) & "' and Upper(bed_type)=upper('" & Trim(cboBedType) & "') "
    
    
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
      cboBedNo_Click
    End If
   cmd.Properties("PLSQLRSet") = False
   If Conn.State = 1 Then
     Conn.Close
     Set Conn = Nothing
     Set RS = Nothing
     Set cmd = Nothing
   End If
        
End Sub
Private Sub cboBedNo_Click()
    LOAD_BED_CHARGE_SERIAL
End Sub
Private Sub LOAD_BED_CHARGE_SERIAL()
   Dim Conn As New ADODB.Connection
   Dim RS As New ADODB.Recordset
   Dim cmd As New ADODB.Command
   
   If Conn.State = 0 Then
       Conn.ConnectionString = strcn.Connection_String
       Conn.Open
    End If
    cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "select bed_charge,service_charge,BED_GROUP,serial_no from bed_info  where occupy_flag='0'and upper(bed_type)=upper('" & cboBedType & "') and  BED_EXT_COL='" & Trim(CboBedTypeNo.Text) & "'and bed_no='" & Trim(cboBedNo.Text) & "' and doc_department='" & Trim(cboBedDept.Text) & "'"
    cmd.Properties("iRowsetChange") = True
    cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    
    RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
    If RS.RecordCount > 0 Then
       txtSetupBedCharge = RS!bed_charge
       txtsetupAdmissionfee = RS!BED_GROUP
       txtSetupServiceFee = RS!service_charge
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
 If cboBedType = "Cabin" And Len(cboBedDept) > 0 Then
     cboBedDept.Text = "Common"
     cboBedDept_Click
  End If
  CboBedTypeNo.clear
  cboBedNo.clear
  txtsetupAdmissionfee = ""
  txtSetupBedCharge = ""
  txtSetupServiceFee = ""
  txtSerialNo = ""
End Sub

Private Sub CboBedTypeNo_Click()
    LOAD_BED_NO
End Sub

Private Sub cboBedDept_Click()
  cboBedNo.clear
  txtsetupAdmissionfee = ""
  txtSetupBedCharge = ""
  txtSetupServiceFee = ""
  txtSerialNo = ""
  Load_bed_type_no
  
End Sub
Private Sub Load_bed_type_no()
''On Error Resume Next
     Dim Conn As New ADODB.Connection
     Dim RS As New ADODB.Recordset
     Dim cmd As New ADODB.Command
     
     If Conn.State = 0 Then
      Conn.ConnectionString = strcn.Connection_String
      Conn.Open
    End If
    cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
   
   cmd.CommandText = "select distinct(BED_EXT_COL) from bed_info  where  bed_type= '" & Trim(cboBedType) & "'and doc_department='" & Trim(cboBedDept.Text) & "'  and occupy_flag='0' order by BED_EXT_COL"
    
    cmd.Properties("iRowsetChange") = True
   cmd.Properties("updatability") = 7
    RS.CursorLocation = adUseClient
    
    RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
    
    CboBedTypeNo.clear
    
    If RS.RecordCount > 0 Then
            
         
         RS.MoveFirst
         
        Do Until RS.EOF = True
            CboBedTypeNo.AddItem RS.Fields(0)
            
            RS.MoveNext
        Loop
       CboBedTypeNo.Text = CboBedTypeNo.List(0)
    End If
     cmd.Properties("iRowsetChange") = False
   '' RS.Close
If Conn.State = 1 Then
    Conn.Close
    Set Conn = Nothing
    Set RS = Nothing
    Set cmd = Nothing
End If
    
End Sub
Private Sub cmdExit_Click()
   Unload Me
End Sub
Private Sub cmdPrint_Click()
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
             
  If total_bed_charge < 0 Then
     MsgBox "Invalid Transfer Date..." & vbCrLf & " Please verify or Contact with Administrator", vbInformation, "IT DIVISION,DNMIH"
     Exit Sub
  End If
  If Len(txtAdvance) = 0 Then
    txtAdvance = 0
  End If

 If Trans_Dept.Text = comDepartmentRelease.Text Then
     MsgBox "You can't transfer to same  Department", vbInformation, " IT, DNMIH"
     Trans_Dept.SetFocus
     Exit Sub
 End If
 
 If Len(Trans_Dept) = 0 Then
     MsgBox "Please Select a Doctor's Department", vbInformation, " IT, DNMIH"
     Trans_Dept.SetFocus
     Exit Sub
 End If
  
 If Val(txtCCU_Charge.Text) > 0 And Val(txtCCU_Charge.Text) < 1000 Then
     MsgBox "Pls Verify CCU Charge", vbInformation, " IT, DNMIH"
     txtCCU_Charge.SetFocus
     Exit Sub
 End If
   
  
  If UCase(Trans_Dept.Text) = UCase("Common") Then
        MsgBox "Please Select a Doctor's Department", vbInformation, " IT, DNMIH"
        Trans_Dept.SetFocus
      Exit Sub
 End If

Dim reply As String
If ChkBedDepartmentIndicator.Value = 1 Then
   reply = MsgBox("Are you sure to Both Department & Bed Transfer?", vbQuestion + vbYesNo, "Transferring...Department+Bed")
     If reply = 6 Then
         Adodc3.ConnectionString = strcn.Connection_String
         Adodc3.RecordSource = "SELECT occupy_flag AS REC_NO FROM bed_info  where serial_no='" & Trim(txtSerialNo) & "'"
         Adodc3.Refresh
         If Adodc3.Recordset.RecordCount > 0 Then
             If Adodc3.Recordset!REC_NO = "1" Then
                MsgBox "This Bed Has Recently Occupied by a new Patient....select a new bed ", vbInformation, " IT, DNMIH"
              Exit Sub
             End If
         End If
         DEPARTMENT_MODE = 0
         
         If cboBedType = "Free-Bed" Then
             txtSetupServiceFee = UTILITY.Load_Service_Charge_Exceptional(cboBedType, Trans_Dept)
         End If
         Call save_PAT_BED_Transfer
         Call save_department_info
         MsgBox "Operation successful", vbInformation + vbOKOnly, "Save..."
         print_TRANSFER
     End If
Else
   reply = MsgBox("Are you sure to Only Department Transfer?", vbQuestion + vbYesNo, "Transferring...Department Only")
     If reply = 6 Then
        DEPARTMENT_MODE = 1
        Call save_department_info
         MsgBox "Operation successful", vbInformation + vbOKOnly, "Save..."
         print_TRANSFER
     End If
End If

Unload Me

End Sub

Private Sub save_PAT_BED_Transfer()
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
    Set Param0 = cmd.CreateParameter("param0", adInteger, adParamInput, 3, 0)
    cmd.Parameters.Append Param0 'MODE
   
    Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 10, Trim(frmDeptTransfer.txtRegNoRelease.Text))
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
    
     Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 10, Trim(frmDeptTransfer.CBOYRCODE.Text))
    cmd.Parameters.Append Param6 'YRCODE
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 15, Trim(cboBedType))
    cmd.Parameters.Append Param7 'BED TYPE
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 20, Trim(CboBedTypeNo))
    cmd.Parameters.Append Param8 'BED TYPE NO
 
   Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 20, Trim(cboBedNo))
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
   
   Set Param12 = cmd.CreateParameter("param12", adVarChar, adParamInput, 19, Trim(Trans_Dept))
   cmd.Parameters.Append Param12 'DOC DEPT
  
   Set Param13 = cmd.CreateParameter("param13", adDate, adParamInput, 8, Trim(MaskEdBox1))
   cmd.Parameters.Append Param13 'ADMISSION DATE
  
   Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 8, Trim(txtsetupAdmissionfee))
   cmd.Parameters.Append Param14 'ADMISSION CHARGE
   
   Set Param15 = cmd.CreateParameter("param15", adDouble, adParamInput, 8, Trim(txtSetupBedCharge))
   cmd.Parameters.Append Param15 'BED  CHARGE
 
   Set param16 = cmd.CreateParameter("param16", adDouble, adParamInput, 8, Trim(txtSetupServiceFee))
   cmd.Parameters.Append param16 'SERVICE CHARGE
   
    Set Param17 = cmd.CreateParameter("param17", adDouble, adParamInput, 8, Trim(txtExtraBedFlag))
   cmd.Parameters.Append Param17 'EXT BED FLAG
 
   Set Param18 = cmd.CreateParameter("param18", adInteger, adParamInput, 8, txtDepartmentSerial)
   cmd.Parameters.Append Param18 'Department Serial
   
   
   Set Param19 = cmd.CreateParameter("param19", adInteger, adParamInput, 8, PATIENT_BED_SERIAL_NO)
   cmd.Parameters.Append Param19 'CURRENT BED SERIAL NO
   
   Set Param20 = cmd.CreateParameter("param20", adDate, adParamInput, 10, Format(frmDeptTransfer.MaskEdBox1.Text, "DD/MM/YYYY"))
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
Exit Sub
ErrDes:
MsgBox Err.Description, vbCritical, " IT, DNMIH"
End Sub
Private Sub print_TRANSFER()
       Dim Connpre As New Connection
       Dim cmdpre As New Command
       Dim rspre As New ADODB.Recordset
       
       Dim Param1 As New Parameter
       Dim RPTVIEWER As New Viewer
       If Connpre.State = 0 Then
          Connpre.Open strcn.Connection_String
       End If
       Set cmdpre.ActiveConnection = Connpre
       cmdpre.CommandType = adCmdText
            
            
       Dim Reporttran   As New CrystalReporttran
       Set Param1 = cmdpre.CreateParameter("param1", adDouble, adParamInput, 100, PREVIEW_VAR)
       cmdpre.Parameters.Append Param1 'comment
    
       cmdpre.Properties("PLSQLRSet") = True
       cmdpre.CommandText = "{CALL Rpt_in_dr_info_adm_print(?)}"
       Set rspre = cmdpre.Execute
       cmdpre.Properties("PLSQLRSet") = False
       Reporttran.Text4.Width = 3000
       Reporttran.Text4.SetText ("Department Transfer")
       Reporttran.DoctorsDepartment1.Font.Bold = True
       Reporttran.Database.SetDataSource rspre
       Reporttran.PrintOut
           
       If Connpre.State = 1 Then
          Connpre.Close
          Set Connpre = Nothing
       End If
       Set cmdpre = Nothing
       
End Sub
Private Sub save_department_info()
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
    Dim Param21 As New Parameter
    Dim Param22 As New Parameter
    Dim Param23 As New Parameter
    Dim Param24 As New Parameter
    Dim Param25 As New Parameter
    Dim Param26 As New Parameter
    Dim param27 As New Parameter
    Dim param28 As New Parameter
    Dim param29 As New Parameter
    
    
    If Conn.State = 0 Then
       Conn.Open strcn.Connection_String
    End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    '----------------------------------------------------------------------------------
    Set Param0 = cmd.CreateParameter("param0", adDouble, adParamInput, 5, Val(DEPARTMENT_MODE))
    cmd.Parameters.Append Param0 'MODE=1 FOR ONLY DEPT TRANSFER
    
    Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 5, frmDeptTransfer.txtRegNoRelease)
    cmd.Parameters.Append Param1 'in_reg_no
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 8, frmDeptTransfer.CBOYRCODE)
    cmd.Parameters.Append Param2 'YR CODE
    
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, Val(txtTotalBedCharge))
    cmd.Parameters.Append Param3 'bed_sum
    

    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, Val(txtTotalOpr))
    cmd.Parameters.Append Param4 'OPERATION CHARGE
    
     
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 10, txtmisce)
    cmd.Parameters.Append Param5 'miscelleneous
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 10, Val(txtExtraBedTotal))
    cmd.Parameters.Append Param6 'total extra bed sum
   
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 10, Val(Text1))
    cmd.Parameters.Append Param7 ' anaesthesia_charge
    
    Set Param8 = cmd.CreateParameter("param8", adDouble, adParamInput, 10, Val(txtDeliveryCharge))
    cmd.Parameters.Append Param8 ' delivery charge
  
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, Val(txtBabyCareCharge))
    cmd.Parameters.Append Param9 'baby care charge
    
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, Val(txtNeunetalCharge))
    cmd.Parameters.Append Param10 'Neunetal charge
  
    Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 10, Val(txtExtTransfusion))
    cmd.Parameters.Append Param11 'txtExtTransfusion charge
    
    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 10, Val(txtP_therapyCharge))
    cmd.Parameters.Append Param12 'txtP_therapyCharge charge
    
    Set Param13 = cmd.CreateParameter("param13", adDouble, adParamInput, 10, Val(txtBloodTher_charge))
    cmd.Parameters.Append Param13 'txtBloodTher_charge charge
    
    Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 10, Val(txtMedicine_charge))
    cmd.Parameters.Append Param14 'txtMedicine_charge charge
    
    Set Param15 = cmd.CreateParameter("param15", adDouble, adParamInput, 10, Val(txtCCU_Charge))
    cmd.Parameters.Append Param15 ' CCU  charge
  
    Set param16 = cmd.CreateParameter("param16", adDouble, adParamInput, 10, Val(txtNebuliser))
    cmd.Parameters.Append param16 ' nebuliser Charge
    
    Set Param17 = cmd.CreateParameter("param17", adDouble, adParamInput, 10, Val(txtIncubator))
    cmd.Parameters.Append Param17 ' incubator charge
    
    Set Param18 = cmd.CreateParameter("param18", adDouble, adParamInput, 10, Val(txtdisc))
    cmd.Parameters.Append Param18 ' DISCOUNT
    
    Set Param19 = cmd.CreateParameter("param19", adDouble, adParamInput, 10, Val(txtTotal))
    cmd.Parameters.Append Param19 ' Net total
     
    Set Param20 = cmd.CreateParameter("param20", adDate, adParamInput, 10, frmDeptTransfer.MaskEdBox1.Text)
    cmd.Parameters.Append Param20 ' total with Admission charge
    
    
    Set Param21 = cmd.CreateParameter("param21", adVarChar, adParamInput, 10, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param21 'u_id
    
    Set Param22 = cmd.CreateParameter("param22", adVarChar, adParamInput, 10, frmMAIN.lblBooth)
    cmd.Parameters.Append Param22 'booth_no
    
    Set Param23 = cmd.CreateParameter("param23", adDate, adParamInput, 10, Format(MaskEdBox1.Text, "DD/MM/YYYY"))
    cmd.Parameters.Append Param23 'admission date
    
    Set Param24 = cmd.CreateParameter("param24", adVarChar, adParamInput, 20, Trim(Trans_Dept.Text))
    cmd.Parameters.Append Param24 'To Transfer Department
   
    Set Param25 = cmd.CreateParameter("param25", adInteger, adParamInput, 5, Trim(txtDepartmentSerial.Text))
    cmd.Parameters.Append Param25 'DEPARTMENT SERIAL
    
    Set Param26 = cmd.CreateParameter("param26", adInteger, adParamInput, 5, Val(PATIENT_BED_SERIAL_NO))
    cmd.Parameters.Append Param26 'BED SERIAL
   
   
    Set param27 = cmd.CreateParameter("param27", adDouble, adParamInput, 10, Val(txtAdvance.Text))
    cmd.Parameters.Append param27 'ADVANCE
  
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_department_transfer(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
        
    Set RS = cmd.Execute
    cmd.Properties("PLSQLRSet") = False
    
  If Conn.State = 1 Then
       Conn.Close
       Set Conn = Nothing
       Set RS = Nothing
       Set cmd = Nothing
  End If
Adodc3.ConnectionString = strcn.Connection_String
Adodc3.RecordSource = "SELECT MAX(REC_NO)AS REC_NO FROM RECEIPT_NO_COUNTER"
Adodc3.Refresh
If Adodc3.Recordset.RecordCount > 0 Then
   txtPrintPreview = Adodc3.Recordset!REC_NO
   PREVIEW_VAR = Val(txtPrintPreview)
End If

End Sub
Private Sub save_release_info_rough()
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
    Dim Param25 As New Parameter
    Dim Param26 As New Parameter
   If Conn.State = 0 Then
     Conn.Open strcn.Connection_String
   End If
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, frmRelease.txtRegNoRelease)
    cmd.Parameters.Append Param1 'in_reg_no
    
    Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 10, 0)
    cmd.Parameters.Append Param2 'test_sum
    
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, txtTotalOpr)
    cmd.Parameters.Append Param3 'total Operation sum
   
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, txtTotalBedCharge)
    cmd.Parameters.Append Param4 'bed_sum
    

    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 10, txtdisc)
    cmd.Parameters.Append Param5 'discount
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 10, txtTotal)
    cmd.Parameters.Append Param6 'total
    
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 10, txtmisce)
    cmd.Parameters.Append Param7 'miscelleneous
    

     Set Param8 = cmd.CreateParameter("param8", adDouble, adParamInput, 10, 0)
    cmd.Parameters.Append Param8 'total with miscelleneous
    
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, 0)
    cmd.Parameters.Append Param9 'net total
    
     
    
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, txtExtraBedTotal)
    cmd.Parameters.Append Param10 'total extra bed sum
    
     Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 10, txtBabyCareCharge)
    cmd.Parameters.Append Param11 'baby care charge
     Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 10, txtNeunetalCharge)
    cmd.Parameters.Append Param12 'Neunetal charge
    
    Set Param13 = cmd.CreateParameter("param13", adDouble, adParamInput, 10, txtExtTransfusion)
    cmd.Parameters.Append Param13 'txtExtTransfusion charge
    
    Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 10, txtP_therapyCharge)
    cmd.Parameters.Append Param14 'txtP_therapyCharge charge
    
    Set Param15 = cmd.CreateParameter("param15", adDouble, adParamInput, 10, txtBloodTher_charge)
    cmd.Parameters.Append Param15 'txtBloodTher_charge charge
    
    Set param16 = cmd.CreateParameter("param16", adDouble, adParamInput, 10, txtMedicine_charge)
    cmd.Parameters.Append param16 'txtMedicine_charge charge
    
    Set Param17 = cmd.CreateParameter("param17", adVarChar, adParamInput, 10, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param17 'u_id
    Set Param18 = cmd.CreateParameter("param18", adVarChar, adParamInput, 10, frmMAIN.lblBooth)
    cmd.Parameters.Append Param18 'booth_no
    Set Param19 = cmd.CreateParameter("param19", adDouble, adParamInput, 10, 0)
    cmd.Parameters.Append Param19 ' total with Admission charge
   
   Set Param20 = cmd.CreateParameter("param20", adDouble, adParamInput, 10, 0)
    cmd.Parameters.Append Param20 ' service charge
    
   Set Param21 = cmd.CreateParameter("param21", adDouble, adParamInput, 10, txtDeliveryCharge)
    cmd.Parameters.Append Param21 ' delivery charge
   
    Set Param22 = cmd.CreateParameter("param22", adDouble, adParamInput, 10, txtCCU_Charge)
    cmd.Parameters.Append Param22 ' CCU  charge
    If Text1 = "" Then
       Text1.Text = 0
    End If
    
    Set Param23 = cmd.CreateParameter("param23", adDouble, adParamInput, 10, Text1)
    cmd.Parameters.Append Param23 ' anaesthesia_charge
   
    Set Param24 = cmd.CreateParameter("param24", adInteger, adParamInput, 10, discount_check_val)
    cmd.Parameters.Append Param24 ' poor patient/staff discount
      
    Set Param25 = cmd.CreateParameter("param25", adDouble, adParamInput, 12, txtIncubator)
    cmd.Parameters.Append Param25 ' poor patient/staff discount
    
    Set Param26 = cmd.CreateParameter("param26", adDouble, adParamInput, 12, txtNebuliser)
    cmd.Parameters.Append Param26 ' poor patient/staff discount
    
   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_calculation_indoor_rough(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    cmd.Properties("PLSQLRSet") = False
  If Conn.State = 1 Then
      'Set Conn = Nothing
       Conn.Close
       Set Conn = Nothing
       Set RS = Nothing
       Set cmd = Nothing
  End If
End Sub

Private Sub cmdShowBed_Click()
  
On Error GoTo Errdesc
Dim f2 As New frmDataSelect
Dim getconnected As New Connection
Dim cmd As New Command
Dim myrs As New ADODB.Recordset
   
   If cboBedType = "Cabin" Then
      Call load_bed_reg_for_showing_cabin
   ElseIf cboBedType = "Paying" Then
      Call load_bed_reg_for_showing_PAYING
   ElseIf cboBedType = "Free-Bed" Then
      Call load_bed_reg_for_showing_FREE
   End If
   
   
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
If getconnected.State = 1 Then
        getconnected.Close
        Set getconnected = Nothing
End If
Exit Sub
Errdesc:
    MsgBox Err.Description, vbInformation, " IT, DNMIH"
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
    
    cmd.CommandText = "{CALL show_bed_no_PAYING}"
    
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
    
    cmd.CommandText = "{CALL show_bed_no_FREE}"
    
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyInsert Then
            cmdSave_Click
End If
If KeyCode = vbKeyEscape Then
  cmdExit_Click
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub
Private Sub Form_Load()
'    TOTAL_DAYS As Integer
     show_dept (1)
     show_dept (2)
     
     LOAD_CURRENT_BED
     
     cboBedType.ListIndex = 0
     cboBedDept.Text = "Common"
     
     Call Load_Patient_Info
     
     
     
     Call LOAD_BED_CHARGE
     
     Call FORMAT_GRID(1)
     Call FORMAT_GRID(2)
     
     
     Call Load_Department_Info
        
     
     discount_check_val = 0
     txtRegNoShow = frmDeptTransfer.txtRegNoRelease
    total_bed_charge = 0
    '''''''''''''''''''''''''''for bed charge''''''''''''''''''''
     Call LOAD_BED_HISTORY_FOR_DEPT
     txtTotalBedCharge = total_bed_charge
     
           
   '''''''''''''''''''''''''''for extra bed charge'''''''''''''
     If Extra_bed_Flag_Indicator = 1 Then
        Call LOAD_EXTRA_BED_HISTORY
        txtExtraBedTotal = Val(total_EXTRA_bed_charge)
     End If
        
     Call LOAD_ADVANCE
        
        
        
'        If comDepartmentRelease = "pAEDIATRIC" Then
''           Call LOAD_PAEDIATRIC_CHARGE
'        End If
'
         
         '''''''''total''''''''''''
     txtTotal.Text = Val(Text1) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtCCU_Charge) + Val(txtNebuliser)
         
   If comDepartmentRelease.Text = "Gynae-1" Or comDepartmentRelease.Text = "Gynae-2" Or comDepartmentRelease.Text = "Gynae-3" Then
      txtDeliveryCharge.Locked = False
   Else
      txtDeliveryCharge.Locked = True
   End If
   
   
   Select Case UCase(comDepartmentRelease)
          Case "SURGERY-1", "SURGERY-2", "SURGERY-3", "GYNAE-1", "GYNAE-2", "GYNAE-3", "ENT", "OPHTH.", "ORTHO."
                txtTotalOpr.Locked = False
   Case Else
          txtTotalOpr.Locked = True
   End Select

         
    End Sub
Private Sub LOAD_ADVANCE()
   Dim Conn As New ADODB.Connection
   Dim RS As New ADODB.Recordset
   Dim cmd As New ADODB.Command
  
   Conn.ConnectionString = strcn.Connection_String
   Conn.Open
   cmd.ActiveConnection = Conn
   cmd.CommandType = adCmdText
   cmd.CommandText = "select  nvl(sum(advance),0) as advance  From advance Where in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE= '" & Trim(frmDeptTransfer.CBOYRCODE.Text) & "'"
      
   cmd.Properties("iRowsetChange") = True
   cmd.Properties("updatability") = 7
   RS.CursorLocation = adUseClient
   RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
   cmd.Properties("iRowsetChange") = False

   If IsNull(RS!advance) = True Then
       txtAdvanceRelease = 0
   Else
       txtAdvanceRelease = RS!advance
   End If
   Conn.Close
   Set Conn = Nothing
   Set RS = Nothing
   Set cmd = Nothing
   
End Sub
Private Sub LOAD_PAEDIATRIC_CHARGE()
  
End Sub
Private Sub Load_Patient_Info()
  Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
  End If
  cmd.ActiveConnection = Conn
  cmd.CommandType = adCmdText
  cmd.CommandText = "select pat_name,admission_date  From in_door_pat_info_main Where in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmDeptTransfer.CBOYRCODE) & "'"
      
  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
  RS.CursorLocation = adUseClient

  RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
  cmd.Properties("iRowsetChange") = False

  If RS.RecordCount > 0 Then
     txtNameRelease = RS!pat_name
     MaskEdBox1.Text = Format(RS!admission_date, "DD/MM/yyyy")
  End If
  
  Conn.Close
  Set Conn = Nothing
  Set RS = Nothing
  Set cmd = Nothing
     
End Sub
Private Sub Load_Department_Info()
  Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
  End If
  cmd.ActiveConnection = Conn
  cmd.CommandType = adCmdText
  cmd.CommandText = "SELECT doc_dept,SERIAL_NO FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmDeptTransfer.CBOYRCODE) & "' AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmDeptTransfer.CBOYRCODE) & "')"
  
  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
  RS.CursorLocation = adUseClient

  RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
  cmd.Properties("iRowsetChange") = False

  If RS.RecordCount > 0 Then
     comDepartmentRelease = "" & RS!doc_dept
     txtDepartmentSerial = RS!SERIAL_NO
  End If
  
  Conn.Close
  Set Conn = Nothing
  Set RS = Nothing
  Set cmd = Nothing

End Sub
Private Sub LOAD_EXTRA_BED_HISTORY()
     Dim Conn As New ADODB.Connection
     Dim RS As New ADODB.Recordset
     Dim cmd As New ADODB.Command
     Dim i As Integer
     
     total_EXTRA_bed_charge = 0
     If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
     End If
     cmd.ActiveConnection = Conn
     cmd.CommandType = adCmdText
     cmd.CommandText = "select  start_date,end_date,BED_CHARGE  From Indoor_pat_Extra_bed_info,SYSDATE Where in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmDeptTransfer.CBOYRCODE) & "'"
      
     cmd.Properties("iRowsetChange") = True
     cmd.Properties("updatability") = 7
     RS.CursorLocation = adUseClient
     cmd.Properties("iRowsetChange") = False

     RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
     
    If Not RS.EOF Then
     i = 1
    With MSFlexGrid2
          Do Until RS.EOF
            MSFlexGrid2.Rows = i + 1
               .Rows = i + 1
               .ColAlignment(0) = 0
               .TextMatrix(i, 0) = Format(RS!Start_date, "DD/MM/YYYY")
               .ColAlignment(1) = 0
               .TextMatrix(i, 1) = Format(RS!End_date, "DD/MM/YYYY")
               If RS!End_date Is Null Then
                  total_EXTRA_bed_charge = total_EXTRA_bed_charge + UTILITY.calculate_date(RS!Start_date, RS!SYSDATE) * RS!bed_charge
                Else
                  total_EXTRA_bed_charge = total_EXTRA_bed_charge + UTILITY.calculate_date(RS!Start_date, RS!End_date) * RS!bed_charge
                End If
         
               i = i + 1
            RS.MoveNext
        Loop
    End With
Else
    MSFlexGrid1.Rows = 1
 End If
     cmd.Properties("iRowsetChange") = False
     
     Conn.Close
     Set Conn = Nothing
     Set RS = Nothing
     Set cmd = Nothing
     

End Sub
Private Sub FORMAT_GRID(MODE As Integer)
  If MODE = 1 Then
     With MSFlexGrid1
         .Rows = 1
         .Cols = 3
         .Col = 0: .Text = "Bed No"
         .Col = 1: .Text = " Admission Date "
         .Col = 2: .Text = " Charge "
         
         .ColWidth(0) = 1900
         .ColWidth(1) = 1500
         .ColWidth(2) = 800
         
     End With
  End If
  If MODE = 2 Then
     With MSFlexGrid2
         .Rows = 1
         .Cols = 2
         .Col = 0: .Text = "Start Date"
         .Col = 1: .Text = "End Date"
         
         .ColWidth(0) = 1500
         .ColWidth(1) = 1500
         
     End With
  End If
End Sub

Private Sub LOAD_BED_HISTORY_FOR_DEPT()

  Dim Conn As New ADODB.Connection
  Dim RS As New ADODB.Recordset
  Dim cmd As New ADODB.Command
  Dim INNER_CHARGE As Single
  Dim firstRecordAdmissionDate As Date
  total_bed_charge = 0
  
  Dim i As Integer
  
  If Conn.State = 0 Then
     Conn.ConnectionString = strcn.Connection_String
     Conn.Open
  End If
  cmd.ActiveConnection = Conn
  cmd.CommandType = adCmdText
  cmd.CommandText = "select  bed_no,bed_type,BED_TYPE_NO,bed_charge,admission_charge,admission_date,Extra_bed_flag,ed_dt  From Indoor_pat_bed_info Where " & _
  "in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmDeptTransfer.CBOYRCODE) & "' AND DOC_DEPT='" & comDepartmentRelease.Text & "' " & _
  " AND DEPT_SERIAL=(select MAX(DEPT_SERIAL) FROM Indoor_pat_bed_info WHERE in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmDeptTransfer.CBOYRCODE) & "' AND DOC_DEPT='" & comDepartmentRelease.Text & "') ORDER BY admission_date ASC"
  
  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
  RS.CursorLocation = adUseClient

  RS.Open cmd.CommandText, Conn, adOpenForwardOnly, adLockReadOnly
  cmd.Properties("iRowsetChange") = False

  If Not RS.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until RS.EOF ''' LOOP
                INNER_CHARGE = 0
                If i = 1 Then
                   firstRecordAdmissionDate = Format(RS!admission_date, "DD/MM/YYYY")
                End If
               .Rows = i + 1
                 
               .TextMatrix(i, 0) = RS!Bed_type & " (" & RS!bed_TYPE_no & " )  - " & RS!bed_no
               .ColAlignment(1) = 0
               .TextMatrix(i, 1) = Format(RS!admission_date, "DD/MM/YYYY")
               
                Extra_bed_Flag_Indicator = Val(0 & RS!Extra_bed_flag)
                
                     
                     '''MISSION FLED PAT=1  FOR ENTRY OF ABSCOND PAT
                    If RS.RecordCount = i Then ''' last or first row
                        If (firstRecordAdmissionDate = Format(frmDeptTransfer.MaskEdBox1.Text, "DD/MM/YYYY")) Then
                                 INNER_CHARGE = UTILITY.GetBedChargeForRow(RS.RecordCount, i, Format(RS!admission_date, "DD/MM/YYYY"), Format(frmDeptTransfer.MaskEdBox1.Text, "DD/MM/YYYY"), RS!bed_charge, False, True)
                        Else
                                   INNER_CHARGE = UTILITY.GetBedChargeForRow(RS.RecordCount, i, Format(RS!admission_date, "DD/MM/YYYY"), Format(frmDeptTransfer.MaskEdBox1.Text, "DD/MM/YYYY"), RS!bed_charge, False, True)
                        End If
                   Else ''' if not last record
                         INNER_CHARGE = UTILITY.GetBedChargeForRow(RS.RecordCount, i, Format(RS!admission_date, "DD/MM/YYYY"), Format(RS!ed_dt, "DD/MM/YYYY"), RS!bed_charge, False, True)
                   End If
                    
                          total_bed_charge = total_bed_charge + INNER_CHARGE
                               
               If RS!Bed_type = "Free-Bed" Then
                    .Row = i
                     .Col = 0
                    .CellForeColor = vbRed
                    .Col = 2
                    .CellForeColor = vbRed
                 End If
                .ColAlignment(2) = 0
                .TextMatrix(i, 2) = INNER_CHARGE
               
                i = i + 1
            RS.MoveNext
        Loop
    End With
 Else
    MSFlexGrid1.Rows = 1
 End If
'  MsgBox "CABIN DAYS :" & CABIN_DAYS
'  MsgBox "FREE BED DAYS  :" & FREE_BED_DAYS
'
  Conn.Close
  Set Conn = Nothing
  Set RS = Nothing
  Set cmd = Nothing
  Set UTILITY = Nothing
  
  
  
  

'  Dim Conn As New ADODB.Connection
'  Dim RS As New ADODB.Recordset
'  Dim cmd As New ADODB.Command
'  Dim INNER_CHARGE As Double
'  Dim first_bed_zero_days As Integer  '''on same date admitted and transfer to another bed
'                                      '''' in this case bed charge will be counted one days for first admitted bed
'  Dim ADM_DATE As Date
'  Dim DAYS As Integer
'  total_bed_charge = 0
'  first_bed_zero_days = 0
'
'
'  Dim i As Integer
'
'  If Conn.State = 0 Then
'     Conn.ConnectionString = strcn.Connection_String
'     Conn.Open
'  End If
'  cmd.ActiveConnection = Conn
'  cmd.CommandType = adCmdText
'  cmd.CommandText = "select  bed_no,bed_type,BED_TYPE_NO,bed_charge,admission_charge,admission_date,Extra_bed_flag,ed_dt  From Indoor_pat_bed_info Where " & _
'  "in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmDeptTransfer.CBOYRCODE) & "' AND DOC_DEPT='" & comDepartmentRelease.Text & "' " & _
'  " AND DEPT_SERIAL=(select MAX(DEPT_SERIAL) FROM Indoor_pat_bed_info WHERE in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmDeptTransfer.CBOYRCODE) & "' AND DOC_DEPT='" & comDepartmentRelease.Text & "') ORDER BY admission_date ASC"
'
'  cmd.Properties("iRowsetChange") = True
'  cmd.Properties("updatability") = 7
'  RS.CursorLocation = adUseClient
'
'  RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
'  cmd.Properties("iRowsetChange") = False
'
' If Not RS.EOF Then
'    i = 1
'    With MSFlexGrid1
'          Do Until RS.EOF
'                INNER_CHARGE = 0
'               .Rows = i + 1
'               .TextMatrix(i, 0) = RS!Bed_type & " (" & RS!bed_TYPE_no & " )  - " & RS!bed_no
'               .ColAlignment(1) = 0
'               .TextMatrix(i, 1) = Format(RS!admission_date, "DD/MM/YYYY")
'
'                Extra_bed_Flag_Indicator = Val(0 & RS!Extra_bed_flag)
'                If i = RS.RecordCount Then
'                   If first_bed_zero_days = 1 Then
'                      DAYS = UTILITY.calculate_date(Format(RS!admission_date + 1, "DD/MM/YYYY"), Format(frmDeptTransfer.MaskEdBox1.Text, "DD/MM/YYYY"))
'                   Else
'                     DAYS = UTILITY.calculate_date(Format(RS!admission_date, "DD/MM/YYYY"), Format(frmDeptTransfer.MaskEdBox1.Text, "DD/MM/YYYY"))
'                   End If
'
'                   If DAYS = 0 Then
'                      DAYS = 1
'                   End If
'                   INNER_CHARGE = DAYS * RS!bed_charge
'                   total_bed_charge = total_bed_charge + INNER_CHARGE
'                Else
'                   DAYS = UTILITY.calculate_date(Format(RS!admission_date, "DD/MM/YYYY"), Format(RS!ed_dt, "DD/MM/YYYY"))
'                   If DAYS = 0 Then
'                      DAYS = 1
'                      first_bed_zero_days = 1
'                   End If
'                   INNER_CHARGE = DAYS * RS!bed_charge
'                   total_bed_charge = total_bed_charge + INNER_CHARGE
'                End If
'
'                 If RS!Bed_type = "Free-Bed" Then
'                    .Row = i
'                     .Col = 0
'                    .CellForeColor = vbRed
'                    .Col = 2
'                    .CellForeColor = vbRed
'                 End If
'                .ColAlignment(2) = 0
'                .TextMatrix(i, 2) = INNER_CHARGE
'                i = i + 1
'            RS.MoveNext
'        Loop
'    End With
'Else
'    MSFlexGrid1.Rows = 1
' End If
'
'  Conn.Close
'  Set Conn = Nothing
'  Set RS = Nothing
'  Set cmd = Nothing
  End Sub

Private Sub LOAD_CURRENT_BED()
   Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
    End If
        cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "select BED_TYPE,Bed_type_no,BED_NO,extra_bed_flag,SERIAL_NO  From Indoor_pat_bed_info Where in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE ='" & Trim(frmDeptTransfer.CBOYRCODE.Text) & "'  AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM Indoor_pat_bed_info WHERE in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmDeptTransfer.CBOYRCODE) & "')"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False
          
       If RS.RecordCount > 0 Then
         txtbedType = "" & RS!Bed_type & "-" & RS!bed_TYPE_no & "-" & RS!bed_no
         txtExtraBedFlag = RS!Extra_bed_flag
         PATIENT_BED_SERIAL_NO = RS!SERIAL_NO
       End If
       
       If Conn.State = 1 Then
        Conn.Close
        Set Conn = Nothing
        Set RS = Nothing
        Set cmd = Nothing
     End If

End Sub
Private Sub LOAD_BED_CHARGE()
   Dim Conn As New ADODB.Connection
   Dim RS As New ADODB.Recordset
   Dim cmd As New ADODB.Command
   
   Conn.Open strcn.Connection_String
   With cmd
        .ActiveConnection = Conn
        .CommandType = adCmdText
        .CommandText = "SELECT doc_dept FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmDeptTransfer.CBOYRCODE) & "' AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(frmDeptTransfer.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmDeptTransfer.CBOYRCODE) & "')) DOC_DEPT "
         
        
   End With
   
'   Set RS = CMD.Execute
     
End Sub

Private Sub Text1_Change()
    If Not IsNumeric(Text1) Then
        Text1 = ""
 Else
            txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
     
End If
 
End Sub

Private Sub Text1_GotFocus()
        Text1.BackColor = &H96E4B1
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)
        Label10.ForeColor = vbCyan

End Sub

Private Sub Text1_LostFocus()
   If Text1 = Empty Then
       Text1 = 0
   End If
 Text1.BackColor = &H80000018
 Label10.ForeColor = vbWhite
End Sub

Private Sub txtAdvance_Change()
  If Not IsNumeric(txtAdvance) Then
     txtAdvance = ""
  End If
End Sub

Private Sub txtBabyCareCharge_Change()
  If Not IsNumeric(txtBabyCareCharge) Then
            txtBabyCareCharge = ""
  Else
         txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
   End If
  
 
End Sub
Private Sub txtBabyCareCharge_GotFocus()
  If Val(txtBabyCareCharge) = 0 Or Val(txtBabyCareCharge) > 0 Then
        txtBabyCareCharge.ForeColor = vbBlack
        txtBabyCareCharge.Locked = False
         txtBabyCareCharge.BackColor = &H96E4B1
         txtBabyCareCharge.SelStart = 0
         txtBabyCareCharge.SelLength = Len(txtBabyCareCharge)
        Label18.ForeColor = vbCyan
  End If
End Sub

Private Sub txtBabyCareCharge_LostFocus()
   If txtBabyCareCharge = Empty Then
             txtBabyCareCharge = 0
   End If
   txtBabyCareCharge.BackColor = &H80000018
   Label18.ForeColor = vbWhite
End Sub

Private Sub txtBloodTher_charge_Change()
If Not IsNumeric(txtBloodTher_charge) Then
       txtBloodTher_charge = ""
 Else
      txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
End If

 
End Sub

Private Sub txtBloodTher_charge_GotFocus()
          txtBloodTher_charge.BackColor = &H96E4B1
          txtBloodTher_charge.SelStart = 0
          txtBloodTher_charge.SelLength = Len(txtBloodTher_charge)
          Label22.ForeColor = vbCyan

End Sub
Private Sub txtBloodTher_charge_LostFocus()
    If txtBloodTher_charge = Empty Then
       txtBloodTher_charge = 0
    End If
    
     txtBloodTher_charge.BackColor = &H80000018
     Label22.ForeColor = vbWhite
End Sub
Private Sub txtCCU_Charge_Change()
If Not IsNumeric(txtCCU_Charge) Then
       txtCCU_Charge = ""
 Else
        txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
 End If
 
End Sub
Private Sub txtCCU_Charge_GotFocus()
     txtCCU_Charge.BackColor = &H96E4B1
     txtCCU_Charge.SelStart = 0
     txtCCU_Charge.SelLength = Len(txtCCU_Charge)
     lab10.ForeColor = vbCyan
End Sub
Private Sub txtCCU_Charge_LostFocus()
    If txtCCU_Charge = Empty Then
       txtCCU_Charge = 0
    End If
    txtCCU_Charge.BackColor = &H80000018
    lab10.ForeColor = vbWhite
End Sub
Private Sub txtDeliveryCharge_Change()
If Not IsNumeric(txtDeliveryCharge) Then
   txtDeliveryCharge = ""
Else
      txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
End If
 
 
End Sub
Private Sub txtDeliveryCharge_GotFocus()
  txtDeliveryCharge.BackColor = &H96E4B1
  txtDeliveryCharge.SelStart = 0
  txtDeliveryCharge.SelLength = Len(txtDeliveryCharge)
  Label3.ForeColor = vbCyan
End Sub
Private Sub txtDeliveryCharge_LostFocus()
If (txtDeliveryCharge) = Empty Then
      txtDeliveryCharge = 0
      txtDeliveryCharge.BackColor = &H80000018
End If
    txtDeliveryCharge.BackColor = &H80000018
    Label3.ForeColor = vbWhite
End Sub

Private Sub txtDisc_Change()
   If Not IsNumeric(txtdisc.Text) Then
           txtdisc = 0
   ElseIf Val(txtdisc) > Val(txtTotal) Then
         txtdisc = Val(txtTotal)
           
   End If
 
  If txtdisc_percent = 0 Or txtdisc_percent = "" Then
  
    txtDUE_TOTAL = (Val(txtTotal) - Val(txtdisc))
   
  End If

End Sub
Private Sub txtdisc_GotFocus()
txtdisc_percent = 0
txtdisc.BackColor = &H96E4B1
End Sub
Private Sub txtDisc_LostFocus()
        txtdisc.BackColor = &HF7E3B3
End Sub
Private Sub txtdisc_percent_Change()
   If Not IsNumeric(txtdisc_percent.Text) Then
        txtdisc_percent = 0
   ElseIf Val(txtdisc_percent.Text) > 100 Then
         txtdisc_percent.Text = 100
   Else
       If txtdisc_percent = 0 Or txtdisc_percent = "" Then
           txtdisc = 0
        End If
 
       If txtdisc_percent <> 0 And IsNull(txtdisc_percent) = False Then
            txtdisc = ((Val(txtTotal) * Val(txtdisc_percent)) / 100)
       End If

    txtDUE_TOTAL = (Val(txtTotal) - Val(txtdisc))
   End If
End Sub
Private Sub txtdisc_percent_GotFocus()
   txtdisc = 0
   txtdisc_percent.SelStart = 0
   txtdisc_percent.SelLength = Len(txtdisc_percent)
   txtdisc.BackColor = &H96E4B1
   
End Sub
Private Sub txtdisc_percent_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Val(txtdisc_percent) > 0 Then
        ChkBedDepartmentIndicator.SetFocus
      End If
   End If
End Sub

Private Sub txtdisc_percent_LostFocus()
txtdisc_percent.BackColor = &HF7E3B3
End Sub

Private Sub txtExtraBedTotal_Change()
   txtTotalOpr_Change
End Sub
Private Sub txtExtraBedTotal_GotFocus()
  txtExtraBedTotal.SelStart = 0
  txtExtraBedTotal.SelLength = Len(txtExtraBedTotal)
  Label12.ForeColor = vbCyan
End Sub

Private Sub txtExtraBedTotal_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtTotalOpr.SetFocus
  End If
End Sub

Private Sub txtExtraBedTotal_LostFocus()
      txtExtraBedTotal.BackColor = &H80000018
      Label12.ForeColor = vbWhite
End Sub

Private Sub txtExtTransfusion_Change()
If Not IsNumeric(txtExtTransfusion) Then
      txtExtTransfusion = ""
 Else
      txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
End If

End Sub

Private Sub txtExtTransfusion_GotFocus()
        txtExtTransfusion.BackColor = &H96E4B1
        txtExtTransfusion.SelStart = 0
        txtExtTransfusion.SelLength = Len(txtExtTransfusion)
        Label20.ForeColor = vbCyan

End Sub

Private Sub txtExtTransfusion_LostFocus()
  If txtExtTransfusion = Empty Then
     txtExtTransfusion = 0
   End If
   
     txtExtTransfusion.BackColor = &H80000018
     Label20.ForeColor = vbWhite
End Sub


Private Sub txtIncubator_Change()
  If Not IsNumeric(txtIncubator) Then
         txtIncubator = ""
   Else
 
         txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
  End If
   
End Sub

Private Sub txtIncubator_GotFocus()
     txtIncubator.SelStart = 0
     txtIncubator.SelLength = Len(txtIncubator)
     Label30.ForeColor = vbCyan

End Sub
Private Sub txtIncubator_LostFocus()
   If txtIncubator = "" Then
     txtIncubator = 0
   End If
   Label30.ForeColor = vbWhite
End Sub
Private Sub txtMedicine_charge_Change()
  If Not IsNumeric(txtMedicine_charge.Text) Then
         txtMedicine_charge = 0
   Else
         txtTotal = Val(txtMedicine_charge) + Val(txtmisce) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
   End If
End Sub
Private Sub txtMedicine_charge_GotFocus()
    txtMedicine_charge.BackColor = &H96E4B1
    txtMedicine_charge.SelStart = 0
    txtMedicine_charge.SelLength = Len(txtMedicine_charge)
    Label23.ForeColor = vbCyan

End Sub
Private Sub txtMedicine_charge_LostFocus()
 txtMedicine_charge.BackColor = &HF7E3B3
 Label23.ForeColor = vbWhite
End Sub
Private Sub txtmisce_Change()
  If Not IsNumeric(txtmisce) Then
     txtmisce = 0
  Else
     txtTotal = Val(txtmisce) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
 
  End If
  
End Sub

Private Sub txtmisce_GotFocus()
  txtmisce.BackColor = &H96E4B1
  txtmisce.SelStart = 0
  txtmisce.SelLength = Len(txtmisce)
  Label8.ForeColor = vbCyan
End Sub

Private Sub txtmisce_LostFocus()
txtmisce.BackColor = &HF7E3B3
Label8.ForeColor = vbWhite
End Sub

Private Sub txtmiscelleneous_Change()
    txtdisc_percent = 0
    txtdisc = 0
    End Sub

Private Sub txtNebuliser_Change()
  If Not IsNumeric(txtNebuliser) Then
          txtNebuliser = ""
  Else
    txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
  End If
 
End Sub

Private Sub txtNebuliser_GotFocus()
    txtNebuliser.SelStart = 0
    txtNebuliser.SelLength = Len(txtNebuliser)
    Label29.ForeColor = vbCyan

End Sub

Private Sub txtNebuliser_LostFocus()
  If txtNebuliser = "" Then
    txtNebuliser = 0
  End If
  Label29.ForeColor = vbWhite
End Sub

Private Sub txtNeunetalCharge_Change()

If Not IsNumeric(txtNeunetalCharge) Then
      txtNeunetalCharge = ""
 Else
      txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
End If


End Sub

Private Sub txtNeunetalCharge_GotFocus()
      txtNeunetalCharge.BackColor = &H96E4B1
      txtNeunetalCharge.SelStart = 0
      txtNeunetalCharge.SelLength = Len(txtNeunetalCharge)
      Label19.ForeColor = vbCyan

End Sub

Private Sub txtNeunetalCharge_LostFocus()
 If txtNeunetalCharge = Empty Then
    txtNeunetalCharge = 0
 End If
   txtNeunetalCharge.BackColor = &H80000018
   Label19.ForeColor = vbWhite
End Sub

Private Sub txtP_therapyCharge_Change()
If Not IsNumeric(txtP_therapyCharge) Then
      txtP_therapyCharge = ""
 Else
      txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
End If

 
End Sub

Private Sub txtP_therapyCharge_GotFocus()
        txtP_therapyCharge.BackColor = &H96E4B1
        txtP_therapyCharge.SelStart = 0
        txtP_therapyCharge.SelLength = Len(txtP_therapyCharge)
        Label21.ForeColor = vbCyan

End Sub

Private Sub txtP_therapyCharge_LostFocus()
If txtP_therapyCharge = Empty Then
   txtP_therapyCharge = 0
End If
     txtP_therapyCharge.BackColor = &H80000018
     Label21.ForeColor = vbWhite
End Sub



Private Sub txtServiceCharge_Change()
  txtTotalOpr_Change
End Sub


Private Sub txtServiceCharge_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    txtExtraBedTotal.SetFocus
  End If
End Sub

Private Sub txtTotal_Change()
 If Val(txtdisc_percent) > 0 Then
    txtdisc_percent_Change
  Else
    txtDUE_TOTAL = (Val(txtTotal) - Val(txtdisc))
 End If
End Sub

Private Sub txtTotalBedCharge_Change()
   If Not IsNumeric(txtTotalBedCharge) Then
       txtTotalBedCharge = ""
   Else
       txtTotalOpr_Change
   End If
End Sub

Private Sub txtTotalBedCharge_GotFocus()
  Label5.ForeColor = vbCyan
End Sub

Private Sub txtTotalBedCharge_LostFocus()
  Label5.ForeColor = vbWhite
End Sub

Private Sub txtTotalOpr_Change()
 If Not IsNumeric(txtTotalOpr) Then
        txtTotalOpr = ""
 Else
     txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
     
End If
End Sub

Private Sub txtTotalOpr_GotFocus()
       txtTotalOpr.BackColor = &H96E4B1
       txtTotalOpr.SelStart = 0
       txtTotalOpr.SelLength = Len(txtTotalOpr)
       Label13.ForeColor = vbCyan

End Sub

Private Sub txtTotalOpr_LostFocus()
   If txtTotalOpr = Empty Then
        txtTotalOpr = 0
    End If
 txtTotalOpr.BackColor = &H80000018
 Label13.ForeColor = vbWhite
End Sub

Private Sub txtWithAdmissionCharge_Change()
   txtTotalOpr_Change
End Sub

Private Sub show_dept(MODE As Integer)
       Dim Conn As New ADODB.Connection
       Dim RS As New ADODB.Recordset
       Dim cmd As New ADODB.Command
       
      Conn.Open strcn.Connection_String
      cmd.ActiveConnection = Conn
      cmd.CommandType = adCmdText
      cmd.CommandText = "select distinct(doc_dept),refer_code from doctor_info order by refer_code"
      cmd.Properties("iRowsetChange") = True
      cmd.Properties("updatability") = 7
      RS.CursorLocation = adUseClient
      RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
     If MODE = 1 Then
      Trans_Dept.clear
     If RS.RecordCount > 0 Then
         RS.MoveFirst
        Do Until RS.EOF = True
           Trans_Dept.AddItem RS!doc_dept
           RS.MoveNext
        Loop

    End If
   ElseIf MODE = 2 Then
      cboBedDept.clear
     If RS.RecordCount > 0 Then
         RS.MoveFirst
        Do Until RS.EOF = True
           cboBedDept.AddItem RS!doc_dept
            RS.MoveNext
        Loop

    End If
   
   
   End If
    cmd.Properties("iRowsetChange") = False
   If Conn.State = 1 Then
    Conn.Close
    Set Conn = Nothing
    Set RS = Nothing
    Set cmd = Nothing
   End If

End Sub

