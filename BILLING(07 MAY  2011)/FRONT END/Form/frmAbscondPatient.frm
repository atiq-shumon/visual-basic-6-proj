VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAbscondPatient 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Patient Release Form"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00800000&
   ForeColor       =   &H80000001&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtAdvanceRelease 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6EAD2&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   9900
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   7650
      Width           =   1590
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
      Left            =   3750
      TabIndex        =   59
      Top             =   8190
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   5130
      TabIndex        =   49
      Top             =   8460
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   3900
      TabIndex        =   48
      Top             =   8460
      Width           =   1215
   End
   Begin VB.CommandButton CMDSHOWBED 
      Caption         =   "SHOW BED"
      Height          =   375
      Left            =   2670
      TabIndex        =   47
      Top             =   8460
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "NEW"
      Height          =   375
      Left            =   1440
      TabIndex        =   46
      Top             =   8460
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   210
      TabIndex        =   14
      Top             =   8460
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
      TabIndex        =   45
      Top             =   7860
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
      TabIndex        =   44
      Top             =   3570
      Width           =   4170
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   1995
         Left            =   60
         TabIndex        =   50
         Top             =   240
         Width           =   4035
         _ExtentX        =   7117
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
      Height          =   1440
      Left            =   120
      TabIndex        =   42
      Top             =   5910
      Width           =   6930
      Begin MSDataGridLib.DataGrid DataGrid5 
         Height          =   1125
         Left            =   90
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   210
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   1984
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
      Left            =   4320
      TabIndex        =   41
      Top             =   3570
      Width           =   2730
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   1935
         Left            =   60
         TabIndex        =   51
         Top             =   240
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   3413
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
      Height          =   8460
      Left            =   7290
      TabIndex        =   21
      Top             =   330
      Width           =   4965
      Begin VB.TextBox txtServiceCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
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
         TabIndex        =   77
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1140
         Width           =   1590
      End
      Begin VB.TextBox txtAdmissionCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6EAD2&
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
         TabIndex        =   75
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   780
         Width           =   1590
      End
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   8
         Text            =   "0"
         Top             =   4815
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   10
         Text            =   "0"
         Top             =   5580
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   1
         Text            =   "0"
         Top             =   2265
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   9
         Text            =   "0"
         Top             =   5190
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "0"
         Top             =   2625
         Width           =   1590
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
         Height          =   300
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   12
         Text            =   "0"
         Top             =   6360
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   6
         Text            =   "0"
         Top             =   4065
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   7
         Text            =   "0"
         Top             =   4440
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   5
         Text            =   "0"
         Top             =   3690
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   4
         Text            =   "0"
         Top             =   3330
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   3
         Text            =   "0"
         Top             =   3000
         Width           =   1590
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
         Left            =   2610
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   420
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
         Top             =   6900
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   11
         Text            =   "0"
         Top             =   5970
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   8040
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
         Left            =   2610
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1500
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
         Left            =   2610
         MaxLength       =   8
         TabIndex        =   0
         Text            =   "0"
         Top             =   1890
         Width           =   1590
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Charge"
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
         TabIndex        =   78
         Top             =   1170
         Width           =   1470
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Charge"
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
         TabIndex        =   76
         Top             =   810
         Width           =   1725
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0E0FF&
         Height          =   45
         Index           =   3
         Left            =   120
         Top             =   6780
         Width           =   4095
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
         Index           =   1
         Left            =   2190
         TabIndex        =   74
         Top             =   7380
         Width           =   315
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0E0FF&
         Height          =   45
         Index           =   2
         Left            =   120
         Top             =   7860
         Width           =   4095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0E0FF&
         Height          =   45
         Index           =   1
         Left            =   120
         Top             =   7770
         Width           =   4095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0E0FF&
         Height          =   45
         Index           =   0
         Left            =   120
         Top             =   6690
         Width           =   4095
      End
      Begin VB.Label Label92 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Advance"
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
         TabIndex        =   73
         Top             =   7350
         Width           =   1710
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FFFFFF&
         Height          =   8085
         Left            =   30
         Top             =   360
         Width           =   4275
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
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   60
         Top             =   8070
         Width           =   915
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
         Left            =   105
         TabIndex        =   40
         Top             =   4845
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
         Left            =   105
         TabIndex        =   39
         Top             =   5610
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
         Left            =   75
         TabIndex        =   38
         Top             =   2295
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
         Left            =   105
         TabIndex        =   37
         Top             =   5220
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
         Left            =   105
         TabIndex        =   36
         Top             =   2655
         Width           =   1530
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
         TabIndex        =   32
         Top             =   6375
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
         TabIndex        =   31
         Top             =   4470
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
         TabIndex        =   30
         Top             =   4095
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
         TabIndex        =   29
         Top             =   3720
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
         TabIndex        =   28
         Top             =   3360
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
         TabIndex        =   27
         Top             =   3030
         Width           =   1710
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
         TabIndex        =   26
         Top             =   450
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
         TabIndex        =   25
         Top             =   6000
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
         TabIndex        =   24
         Top             =   6870
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
         TabIndex        =   23
         Top             =   1530
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
         TabIndex        =   22
         Top             =   1920
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
      Height          =   2835
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   6945
      Begin VB.TextBox TxtAddr 
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
         Height          =   405
         Left            =   1770
         Locked          =   -1  'True
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1440
         Width           =   5025
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   5310
         TabIndex        =   58
         Top             =   375
         Width           =   1485
         _ExtentX        =   2619
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
         Left            =   5250
         TabIndex        =   55
         Top             =   2250
         Width           =   1515
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   2250
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
         Left            =   1830
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   375
         Width           =   1485
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   908
         Width           =   4875
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   5070
         TabIndex        =   72
         Top             =   2220
         Width           =   105
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   71
         Top             =   2190
         Width           =   105
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   70
         Top             =   885
         Width           =   105
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   68
         Top             =   1410
         Width           =   105
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address "
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
         Left            =   210
         TabIndex        =   67
         Top             =   1470
         Width           =   870
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         Left            =   3510
         TabIndex        =   54
         Top             =   2250
         Width           =   1335
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Bed"
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
         Left            =   210
         TabIndex        =   53
         Top             =   2250
         Width           =   1170
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   35
         Top             =   330
         Width           =   105
      End
      Begin VB.Label lblRegNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No"
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
         Height          =   225
         Left            =   210
         TabIndex        =   33
         Top             =   360
         Width           =   1335
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
         Left            =   210
         TabIndex        =   20
         Top             =   930
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Date :"
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
         Left            =   3510
         TabIndex        =   19
         Top             =   375
         Width           =   1635
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   " Abscond Patient Information Entry"
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
      Left            =   3150
      TabIndex        =   66
      Top             =   0
      Width           =   5475
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
      Left            =   90
      TabIndex        =   65
      Top             =   7620
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Shape Shape4 
      Height          =   465
      Left            =   120
      Top             =   8400
      Width           =   6285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Total"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7080
      TabIndex        =   16
      Top             =   9480
      Width           =   660
   End
End
Attribute VB_Name = "frmAbscondPatient"
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
   
End Sub



Private Sub cmdADD_Click()
   Unload Me
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()
If UTILITY.User_Shift_validation(frmMAIN.lbluser_id, frmMAIN.lblUserType.Caption) = False Then
    MsgBox "Dear. " & frmMAIN.lblName & "  Your Shift has been Expired.. " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
    Exit Sub
End If
             
If total_bed_charge < 0 Then
   MsgBox "Invalid Fled Date..." & vbCrLf & " Please verify or Contact with Administrator", vbInformation, "IT DIVISION,DNMIH"
   Exit Sub
   Unload Me
   '''frmfled.Show 1
End If

 Dim reply As String
 reply = MsgBox("Are you sure to lock this Account?", vbQuestion + vbYesNo, "LOCKING...ACCOUNT")
    If reply = 6 Then
          save_fled_info
          MsgBox "Account Locked Successfully", vbInformation, "IT Division,DNMIH"
          Unload Me
          '''frmfled.Show 1
    End If
End Sub
Private Sub save_fled_info()
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
    
   If Conn.State = 0 Then
     Conn.Open strcn.Connection_String
   End If
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, Cur_reg_no)
    cmd.Parameters.Append Param1 'in_reg_no
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, cur_yr_code)
    cmd.Parameters.Append Param2 ' FISCAL YR
 
     
    Set Param3 = cmd.CreateParameter("param3", adSingle, adParamInput, 10, txtTotalOpr)
    cmd.Parameters.Append Param3 'total Operation sum
   
    Set Param4 = cmd.CreateParameter("param4", adSingle, adParamInput, 10, txtTotalBedCharge)
    cmd.Parameters.Append Param4 'bed_sum
    
    
    Set Param5 = cmd.CreateParameter("param5", adSingle, adParamInput, 10, txtAdmissionCharge)
    cmd.Parameters.Append Param5 'ADMISSION CHARGE
   
    Set Param6 = cmd.CreateParameter("param6", adSingle, adParamInput, 10, txtServiceCharge)
    cmd.Parameters.Append Param6 'SERVICE CHARGE
       
    Set Param7 = cmd.CreateParameter("param7", adSingle, adParamInput, 10, txtExtraBedTotal)
    cmd.Parameters.Append Param7 'total extra bed sum
    
     Set Param8 = cmd.CreateParameter("param8", adDouble, adParamInput, 10, txtBabyCareCharge)
    cmd.Parameters.Append Param8 'baby care charge
    
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 10, txtNeunetalCharge)
    cmd.Parameters.Append Param9 'Neunetal charge
    
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, txtExtTransfusion)
    cmd.Parameters.Append Param10 'txtExtTransfusion charge
    
    Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 10, txtP_therapyCharge)
    cmd.Parameters.Append Param11 'txtP_therapyCharge charge
    
    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 10, txtBloodTher_charge)
    cmd.Parameters.Append Param12 'txtBloodTher_charge charge
    
    Set Param13 = cmd.CreateParameter("param13", adDouble, adParamInput, 10, txtMedicine_charge)
    cmd.Parameters.Append Param13 'txtMedicine_charge charge
    
    Set Param14 = cmd.CreateParameter("param14", adDouble, adParamInput, 10, txtmisce)
    cmd.Parameters.Append Param14 'miscelleneous
  
     Set Param15 = cmd.CreateParameter("param15", adDouble, adParamInput, 10, txtDeliveryCharge)
    cmd.Parameters.Append Param15 ' delivery charge
    
     Set param16 = cmd.CreateParameter("param16", adDouble, adParamInput, 10, txtCCU_Charge)
     cmd.Parameters.Append param16 ' CCU  charge
   
    Set Param17 = cmd.CreateParameter("param17", adDouble, adParamInput, 10, Text1)
    cmd.Parameters.Append Param17 ' anaesthesia_charge
    
    Set Param18 = cmd.CreateParameter("param18", adDouble, adParamInput, 10, txtNebuliser)
    cmd.Parameters.Append Param18 ' nebuliser Charge

   Set Param19 = cmd.CreateParameter("param19", adDouble, adParamInput, 15, txtIncubator)
   cmd.Parameters.Append Param19 ' incubator charge
   
   Set Param20 = cmd.CreateParameter("param20", adDouble, adParamInput, 15, txtAdvanceRelease)
   cmd.Parameters.Append Param20 ' PREVIOUS ADVANCE
   
   Set Param21 = cmd.CreateParameter("param21", adDouble, adParamInput, 15, txtDUE_TOTAL)
   cmd.Parameters.Append Param21 ' TOTAL DUE
   
   Set Param22 = cmd.CreateParameter("param22", adDouble, adParamInput, 15, txtTotal)
   cmd.Parameters.Append Param22 ' TOTAL
   
    
   Set Param23 = cmd.CreateParameter("param23", adDate, adParamInput, 15, FLED_DATE)
   cmd.Parameters.Append Param23 ' FLED DATE
   
   Set Param24 = cmd.CreateParameter("param24", adVarChar, adParamInput, 10, frmMAIN.lbluser_id)
   cmd.Parameters.Append Param24 'u_id
    
   Set Param25 = cmd.CreateParameter("param25", adVarChar, adParamInput, 10, frmMAIN.lblBooth)
   cmd.Parameters.Append Param25 'booth_no
   
   
        
   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_fled_indoor(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
        
    Set RS = cmd.Execute
    cmd.Properties("PLSQLRSet") = False
  
  If Conn.State = 1 Then
       Set Conn = Nothing
       Set cmd = Nothing
       Set RS = Nothing
  End If
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
Private Sub print_release_rough()
 
''            Dim Connrelr As New ADODB.Connection
''            Dim cmdrelr As New ADODB.Command
''            ''Dim RSrelr As New ADODB.Recordset
''            Dim Param1r As New ADODB.Parameter
''            Dim Report14   As New CrystalReport14
''          If Connrelr.State = 0 Then
''                Connrelr.Open strcn.Connection_String
''          End If
''           Set cmdrelr.ActiveConnection = Connrelr
''           cmdrelr.CommandType = adCmdText
''
''
''
''
''            Set Param1r = cmdrelr.CreateParameter("param1r", adInteger, adParamInput, 20, frmRelease.txtRegNoRelease.Text)
''             cmdrelr.Parameters.Append Param1r 'combo
''
''            cmdrelr.Properties("PLSQLRSet") = True
''            cmdrelr.CommandText = "{CALL RptPatientRelease_rough(?)}"
''            Set rs = cmdrelr.Execute
''            cmdrelr.Properties("PLSQLRSet") = False
''
''            Report14.Database.SetDataSource rs
''            Report14.DiscardSavedData
''            Viewer.CRViewer91.ReportSource = Report14
''            Viewer.CRViewer91.ViewReport
''            Screen.MousePointer = vbDefault
''
''
''
''
''           'Report14.PrintOut
''           ' Set RSrel = Nothing
''           If Connrelr.State = 1 Then
''               Connrelr.Close
''               Set Connrelr = Nothing
''           End If
''           ''Set RSrelr = Nothing
''           Set cmdrelr = Nothing
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
     
     LOAD_CURRENT_BED
     
       
     Call Load_Patient_Info
     
     
     
     Call LOAD_BED_CHARGE
     
     Call FORMAT_GRID(1)
     Call FORMAT_GRID(2)
     
     
     Call Load_Department_Info
        
     
     discount_check_val = 0
     txtRegNoShow = Cur_reg_no
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
     Call LOAD_MAX_ADMISSION_SERVICE_FEE
        
        
'        If comDepartmentRelease = "pAEDIATRIC" Then
''           Call LOAD_PAEDIATRIC_CHARGE
'        End If
'
         
         '''''''''total''''''''''''
         txtTotal.Text = Val(Text1) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtCCU_Charge) + Val(txtNebuliser)
         
         txtDUE_TOTAL = Val(txtTotal) - Val(txtAdvanceRelease)
         
         
         
           ''''''''deduct test_total'''''''''''''''''''''''
         

         
    End Sub
Private Sub LOAD_MAX_ADMISSION_SERVICE_FEE()
  Dim Conn As New ADODB.Connection
    Dim RS As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    Conn.Open strcn.Connection_String
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "select MAX(Admission_charge)ADMISSION_CHARGE ,MAX(service_charge) SERVICE_CHARGE from INDOOR_PAT_BED_INFO where IN_REG_NO='" & Cur_reg_no & "' and YRCODE='" & cur_yr_code & "'"
    
    Set RS = cmd.Execute
    If Not RS.EOF Then
       txtAdmissionCharge = RS!ADMISSION_CHARGE
       txtServiceCharge = RS!SERVICE_CHARGE
    End If
        
    Conn.Close
    Set Conn = Nothing
    Set RS = Nothing
    Set cmd = Nothing
  
End Sub




Private Sub LOAD_ADVANCE()
   Dim Conn As New ADODB.Connection
   Dim RS As New ADODB.Recordset
   Dim cmd As New ADODB.Command
  
   Conn.ConnectionString = strcn.Connection_String
   Conn.Open
   cmd.ActiveConnection = Conn
   cmd.CommandType = adCmdText
   cmd.CommandText = "select  nvl(sum(advance),0) as advance  From advance Where in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE= '" & Trim(cur_yr_code) & "'"
      
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
  cmd.CommandText = "select pat_name,ADDR1,admission_date  From in_door_pat_info_main Where in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE='" & Trim(cur_yr_code) & "'"
      
  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
  RS.CursorLocation = adUseClient

  RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
  cmd.Properties("iRowsetChange") = False

  If RS.RecordCount > 0 Then
     txtNameRelease = RS!pat_name
     MaskEdBox1.Text = Format(RS!admission_date, "DD/MM/yyyy")
     txtAddr = RS!addr1
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
  cmd.CommandText = "SELECT doc_dept,SERIAL_NO FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE='" & Trim(cur_yr_code) & "' AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE='" & Trim(cur_yr_code) & "')"
  
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
     cmd.CommandText = "select  start_date,end_date,BED_CHARGE  From Indoor_pat_Extra_bed_info,SYSDATE Where in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE='" & Trim(cur_yr_code) & "'"
      
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
                  total_EXTRA_bed_charge = total_EXTRA_bed_charge + UTILITY.calculate_date(RS!Start_date, RS!sysdate) * RS!bed_charge
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
         
         .ColWidth(0) = 1700
         .ColWidth(1) = 1500
         .ColWidth(2) = 1000
         
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
  Dim INNER_CHARGE As Double
  
  total_bed_charge = 0
  Dim i As Integer
  
  If Conn.State = 0 Then
     Conn.ConnectionString = strcn.Connection_String
     Conn.Open
  End If
  cmd.ActiveConnection = Conn
  cmd.CommandType = adCmdText
    
   cmd.CommandText = "select  bed_no,bed_type,BED_TYPE_NO,bed_charge,admission_charge,admission_date,Extra_bed_flag,ed_dt,DOC_DEPT,DEPT_SERIAL,SYSDATE  From Indoor_pat_bed_info Where " & _
  "in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE='" & Trim(cur_yr_code) & "'  ORDER BY admission_date DESC"
   
  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
  RS.CursorLocation = adUseClient

  RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
  cmd.Properties("iRowsetChange") = False

  If Not RS.EOF Then
    i = 1
    With MSFlexGrid1
          Do Until RS.EOF
                INNER_CHARGE = 0
               .Rows = i + 1
               .TextMatrix(i, 0) = RS!Bed_type & " (" & RS!bed_TYPE_no & " )  - " & RS!bed_no
               .ColAlignment(1) = 0
               .TextMatrix(i, 1) = Format(RS!admission_date, "DD/MM/YYYY")
               
                Extra_bed_Flag_Indicator = Val(0 & RS!Extra_bed_flag)
                If i = 1 Then
                   INNER_CHARGE = UTILITY.calculate_date(Format(RS!admission_date, "DD/MM/YYYY"), Format(FLED_DATE, "DD/MM/YYYY")) * RS!bed_charge
                   total_bed_charge = total_bed_charge + INNER_CHARGE
                Else
                   INNER_CHARGE = UTILITY.calculate_date(Format(RS!admission_date, "DD/MM/YYYY"), Format(RS!ed_dt, "DD/MM/YYYY")) * RS!bed_charge
                   total_bed_charge = total_bed_charge + INNER_CHARGE
                End If
                
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
         
  Conn.Close
  Set Conn = Nothing
  Set RS = Nothing
  Set cmd = Nothing
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
        cmd.CommandText = "select BED_TYPE,Bed_type_no,BED_NO,extra_bed_flag,SERIAL_NO  From Indoor_pat_bed_info Where in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE ='" & Trim(cur_yr_code) & "'  AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM Indoor_pat_bed_info WHERE in_reg_no ='" & Trim(Cur_reg_no) & "' AND YRCODE='" & Trim(cur_yr_code) & "')"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False
          
       If RS.RecordCount > 0 Then
         txtbedType = "" & RS!Bed_type & "-" & RS!bed_TYPE_no & "-" & RS!bed_no
         
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
            txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
     
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

Private Sub txtBabyCareCharge_Change()
  If Not IsNumeric(txtBabyCareCharge) Then
            txtBabyCareCharge = ""
  Else
         txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
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
      txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
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
        txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
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
      txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
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
      txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
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
 
         txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
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
         txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtMedicine_charge) + Val(txtmisce) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
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
     txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtmisce) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
 
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


Private Sub txtNebuliser_Change()
  If Not IsNumeric(txtNebuliser) Then
          txtNebuliser = ""
  Else
    txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
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
      txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
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
      txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
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
     txtDUE_TOTAL = Val(txtTotal) - Val(txtAdvanceRelease)
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
         txtTotal = Val(txtAdmissionCharge) + Val(txtServiceCharge) + Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
     
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

