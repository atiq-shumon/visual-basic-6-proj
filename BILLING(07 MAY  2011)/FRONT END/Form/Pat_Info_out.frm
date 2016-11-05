VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Pat_Info_out 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8730
   ClientLeft      =   -105
   ClientTop       =   390
   ClientWidth     =   10860
   FillColor       =   &H007DABD0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H80000001&
      Height          =   855
      Left            =   -30
      TabIndex        =   74
      Top             =   690
      Width           =   10935
      Begin VB.TextBox TicketNoText 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1830
         TabIndex        =   0
         Top             =   360
         Width           =   3135
      End
      Begin VB.ComboBox MonthCombo 
         Height          =   315
         Left            =   6600
         TabIndex        =   77
         Text            =   "Combo1"
         Top             =   360
         Width           =   1425
      End
      Begin VB.ComboBox YearCombo 
         Height          =   315
         Left            =   8730
         TabIndex        =   75
         Text            =   "Combo1"
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label Label24 
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
         Left            =   330
         TabIndex        =   79
         Top             =   360
         Width           =   1395
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
         Left            =   5760
         TabIndex        =   78
         Top             =   390
         Width           =   750
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
         Left            =   8070
         TabIndex        =   76
         Top             =   390
         Width           =   630
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   71
      Top             =   8340
      Width           =   10995
      Begin VB.Label Label12 
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
         TabIndex        =   73
         Top             =   60
         Width           =   4725
      End
      Begin VB.Label Label8 
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
         TabIndex        =   72
         ToolTipText     =   "Developed && Maintenanced by:A.N.M. Atiqur Rahman,Software Programmer, IT Division, DNMIH"
         Top             =   90
         Width           =   2715
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   2460
      Left            =   2610
      TabIndex        =   10
      Top             =   4230
      Visible         =   0   'False
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   4339
      _Version        =   393216
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14737632
      ForeColor       =   8388608
      BackColorFixed  =   14737632
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483635
      BackColorBkg    =   16777215
      FocusRect       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtCharge 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc3"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   9900
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   3870
      Width           =   930
   End
   Begin VB.ComboBox cboDeptCode 
      DataSource      =   "Adodc2"
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
      ItemData        =   "Pat_Info_out.frx":0000
      Left            =   150
      List            =   "Pat_Info_out.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   3870
      Width           =   1155
   End
   Begin VB.TextBox txtTestCode 
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
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   4170
      TabIndex        =   9
      Top             =   3870
      Width           =   1095
   End
   Begin VB.ComboBox cboMainCode 
      DataSource      =   "Adodc2"
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
      ItemData        =   "Pat_Info_out.frx":0004
      Left            =   1320
      List            =   "Pat_Info_out.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3870
      Width           =   795
   End
   Begin VB.ComboBox cboMainName 
      DataSource      =   "Adodc3"
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
      ItemData        =   "Pat_Info_out.frx":0008
      Left            =   2145
      List            =   "Pat_Info_out.frx":000A
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Text            =   "Combo2"
      Top             =   3870
      Width           =   2040
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   0
      TabIndex        =   32
      Top             =   6810
      Width           =   10860
      Begin VB.TextBox txtServiceTotal 
         Height          =   285
         Left            =   690
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   480
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "CLOSE"
         Height          =   375
         Left            =   5100
         TabIndex        =   18
         ToolTipText     =   "CLOSE "
         Top             =   930
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "REPORT"
         Height          =   375
         Left            =   3870
         TabIndex        =   17
         ToolTipText     =   "REPORT"
         Top             =   930
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "DELETE"
         Height          =   375
         Left            =   2640
         TabIndex        =   21
         ToolTipText     =   "DELETE"
         Top             =   930
         Width           =   1215
      End
      Begin VB.CommandButton cmdADD 
         Caption         =   "NEW"
         Height          =   375
         Left            =   1410
         TabIndex        =   19
         ToolTipText     =   "NEW"
         Top             =   930
         Width           =   1215
      End
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "SAVE"
         Height          =   375
         Left            =   180
         TabIndex        =   16
         ToolTipText     =   "SAVE"
         Top             =   930
         Width           =   1215
      End
      Begin VB.ComboBox txtstaff 
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
         Height          =   315
         Left            =   7830
         TabIndex        =   48
         Top             =   570
         Width           =   1515
      End
      Begin VB.TextBox TXTRECEIPT_NO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3510
         MaxLength       =   17
         TabIndex        =   47
         Top             =   600
         Width           =   1965
      End
      Begin MSAdodcLib.Adodc Adodc7 
         Height          =   330
         Left            =   5580
         Top             =   870
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
         Caption         =   "Adodc7"
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
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7620
         TabIndex        =   44
         Top             =   630
         Width           =   225
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Check2"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7620
         TabIndex        =   43
         Top             =   420
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.TextBox txttesttype 
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc6"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
         Top             =   180
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.TextBox txtdisc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9390
         TabIndex        =   13
         Text            =   "0"
         Top             =   465
         Width           =   615
      End
      Begin VB.TextBox nbrDisc_Per 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10215
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   465
         Width           =   555
      End
      Begin VB.TextBox txtdue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9390
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   780
         Width           =   1395
      End
      Begin VB.TextBox txtchargeTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9390
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   150
         Width           =   1395
      End
      Begin VB.TextBox txtFlagOutdoor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5430
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   420
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtRegNoOutdoor 
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc6"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2580
         MaxLength       =   10
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
         Top             =   435
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   9390
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1080
         Width           =   1395
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   3525
         Top             =   90
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
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
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   3300
         Top             =   105
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
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   330
         Left            =   450
         Top             =   -75
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
      Begin MSAdodcLib.Adodc Adodc5 
         Height          =   330
         Left            =   435
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
      Begin MSAdodcLib.Adodc Adodc6 
         Height          =   330
         Left            =   3525
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   57
         Top             =   420
         Width           =   60
      End
      Begin VB.Shape Shape2 
         Height          =   465
         Left            =   90
         Top             =   870
         Width           =   6285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FISCAL YEAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   49
         Top             =   180
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   7890
         TabIndex        =   46
         Top             =   630
         Width           =   405
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Poor Patient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   7890
         TabIndex        =   45
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Due"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6825
         TabIndex        =   41
         Top             =   825
         Width           =   855
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6825
         TabIndex        =   40
         Top             =   510
         Width           =   765
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   10035
         TabIndex        =   39
         Top             =   480
         Width           =   210
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Test Charge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6825
         TabIndex        =   38
         Top             =   135
         Width           =   1545
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flag"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5010
         TabIndex        =   37
         Top             =   465
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RegNO"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2010
         TabIndex        =   36
         Top             =   450
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Total Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6825
         TabIndex        =   35
         Top             =   1095
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   0
      TabIndex        =   23
      Top             =   -75
      Width           =   10845
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OPD PATIENT DIAGNOSTIC INFORMATION ENTRY "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   465
         Index           =   0
         Left            =   420
         TabIndex        =   24
         Top             =   195
         Width           =   10125
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -210
         Picture         =   "Pat_Info_out.frx":000C
         Stretch         =   -1  'True
         Top             =   60
         Width           =   11520
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   9390
         Picture         =   "Pat_Info_out.frx":598E
         Top             =   240
         Width           =   480
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4905
      Top             =   45
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
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   -60
      TabIndex        =   22
      Top             =   1455
      Width           =   10965
      Begin VB.ComboBox CboDept 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   450
         Width           =   2955
      End
      Begin VB.ComboBox cboDMY 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Pat_Info_out.frx":6258
         Left            =   7320
         List            =   "Pat_Info_out.frx":6265
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1020
         Width           =   975
      End
      Begin VB.ComboBox CBOYRCODE 
         Height          =   315
         ItemData        =   "Pat_Info_out.frx":6272
         Left            =   4800
         List            =   "Pat_Info_out.frx":6274
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   450
         Width           =   1515
      End
      Begin VB.TextBox txtPat_ID1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc6"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   300
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   450
         Width           =   1395
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6690
         MaxLength       =   17
         TabIndex        =   4
         Top             =   1035
         Width           =   615
      End
      Begin VB.TextBox txtAddr 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   300
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1635
         Width           =   10200
      End
      Begin VB.TextBox TxtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   330
         MaxLength       =   60
         TabIndex        =   3
         Top             =   1080
         Width           =   6015
      End
      Begin VB.ComboBox ComSex 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Pat_Info_out.frx":6276
         Left            =   8760
         List            =   "Pat_Info_out.frx":6280
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1005
         Width           =   765
      End
      Begin VB.ComboBox Combo5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Pat_Info_out.frx":628A
         Left            =   9570
         List            =   "Pat_Info_out.frx":629D
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1005
         Width           =   930
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   6690
         TabIndex        =   51
         Top             =   450
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker DTPicker7 
         Height          =   315
         Left            =   8790
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "Delevary Time"
         Top             =   420
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   60096514
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
         Left            =   7380
         TabIndex        =   56
         Top             =   780
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Date"
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
         Left            =   6720
         TabIndex        =   55
         Top             =   165
         Width           =   540
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
         Left            =   8730
         TabIndex        =   54
         Top             =   165
         Width           =   585
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
         Left            =   4830
         TabIndex        =   53
         Top             =   165
         Width           =   1185
      End
      Begin VB.Label Label20 
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
         Left            =   6735
         TabIndex        =   31
         Top             =   780
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rec. No "
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
         Left            =   300
         TabIndex        =   30
         Top             =   165
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   315
         TabIndex        =   29
         Top             =   1380
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Left            =   8790
         TabIndex        =   28
         Top             =   780
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Patient Name"
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
         Left            =   300
         TabIndex        =   27
         Top             =   780
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reff. Dept."
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
         Left            =   1845
         TabIndex        =   26
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label Label21 
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
         Left            =   9600
         TabIndex        =   25
         Top             =   780
         Width           =   900
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Pat_Info_out.frx":62CB
      Height          =   2625
      Left            =   0
      TabIndex        =   59
      Top             =   4230
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   4630
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   14811135
      ForeColor       =   8388608
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
         Weight          =   700
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
   Begin VB.TextBox txtTestTitle 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   3900
      Width           =   4605
   End
   Begin VB.Shape Shape1 
      Height          =   345
      Left            =   180
      Top             =   3450
      Width           =   10665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S.Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   195
      Index           =   5
      Left            =   4350
      TabIndex        =   69
      Top             =   3480
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rec.Depart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   66
      Top             =   3480
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main.Test Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   195
      Index           =   3
      Left            =   2325
      TabIndex        =   63
      Top             =   3480
      Width           =   1395
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C9AD8F&
      BackStyle       =   0  'Transparent
      Caption         =   "Sub.Test Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   195
      Left            =   5325
      TabIndex        =   62
      Top             =   3480
      Width           =   1320
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C9AD8F&
      BackStyle       =   0  'Transparent
      Caption         =   "Charge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   330
      Left            =   10005
      TabIndex        =   61
      Top             =   3480
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   60
      Top             =   3480
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   -6480
      TabIndex        =   20
      Top             =   4770
      Width           =   270
   End
   Begin VB.Menu MNUFILE 
      Caption         =   "FILE"
      Visible         =   0   'False
      Begin VB.Menu MNUDELETE 
         Caption         =   "DELETE"
      End
      Begin VB.Menu SEP001 
         Caption         =   "-"
      End
      Begin VB.Menu MNUREFRESH 
         Caption         =   "REFRESH"
      End
      Begin VB.Menu SEP002 
         Caption         =   "-"
      End
      Begin VB.Menu MNUCLOSE 
         Caption         =   "CLOSE"
      End
   End
End
Attribute VB_Name = "Pat_Info_out"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strUid As String
Dim discount_check_val As Integer
Dim UTILITY As New clsUtility
Dim valueBeforeManualUpdate As String
Dim nationalTotal As Integer
Private Sub CboDept_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxtName.SetFocus
   End If
End Sub

Private Sub cboDeptCode_Click()
  Load_Main_Code (cboDeptCode)
  cboMainCode = cboMainCode.List(0)
End Sub

Private Sub cboDeptCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   cboMainCode.SetFocus
End If
End Sub

Private Sub cboDMY_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      ComSex.SetFocus
   End If
End Sub

Private Sub Check1_Click()
    
 If Check1.Value = 1 Then
    Label27.Visible = False
    txtstaff.Visible = True
      
      
      If Not IsNull(txtstaff) Then
        discount_check_val = 1
        Check2.Value = 0
        txtdisc = Val(nationalTotal)
      End If
  Else
    Label27.Visible = True
    txtstaff.Visible = False
    Check2.Value = 1
End If

End Sub
Private Sub Check2_Click()
  nbrDisc_Per = 0
  If Check2.Value = 1 Then
     txtdisc = 0
     Check1.Value = 0
     discount_check_val = 0
     Label7.Caption = ""
  Else
    Check1.Value = 1
  End If
    
End Sub
Private Sub cmdADD_Click()
    Call clear_main
    Call clear
End Sub
Private Sub clear_main()
    Call UTILITY.SAVE_DELETE(3, "", "", "", "", "", "", 0, frmMAIN.lblBooth)
    Call flush_grid
End Sub
Private Sub clear()

' cboDept = "Card."
 txtPat_ID1 = ""
 TxtName = ""
 txtAddr = ""
  
 txtAge = ""
 txtchargeTotal = 0
 txtdisc = 0
 txtdue = 0
 txtTotal = 0
 nbrDisc_Per = 0
 CboDept.SetFocus
 
 
End Sub
Private Sub CMDDELETE_Click()

    Dim reply As String
    reply = MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo, "Delete...")
                
              
    If reply = vbYes Then
        Call UTILITY.SAVE_DELETE(2, cboDeptCode.Text, cboMainCode.Text, txtTestCode.Text, "", "", OPD_OUT_INDICATION, 0, frmMAIN.lblBooth)
    End If
    Call flush_grid
    Call Rate_Sum
    txtdisc = 0
    nbrDisc_Per = 0
   
   
   txttesttype = ""
   txtCharge = ""
   
   
   
End Sub
   
Private Sub cmdExit_Click()
  
Dim reply As String
    reply = MsgBox("Sure to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdPreview_Click()

End Sub
Private Sub cmdPrint_Click()
  
  If TXTRECEIPT_NO.Visible = False Then
       TXTRECEIPT_NO.Visible = True
   End If
  TXTRECEIPT_NO.ForeColor = vbBlue
  
   If TXTRECEIPT_NO = "" Then
    TXTRECEIPT_NO.SetFocus
     Exit Sub
   Else
   
  Adodc7.ConnectionString = strcn.Connection_String
  Adodc7.RecordSource = "SELECT INDOOR_OUT_DOOR_TEST_FLAG AS REC_NO FROM PAT_INFO_SUB1_OUT_DOOR WHERE REG_NO='" & TXTRECEIPT_NO & "'"
  Adodc7.Refresh
    
  If Adodc7.Recordset.RecordCount > 0 Then
      If Adodc7.Recordset!REC_NO = "1" Then
         MsgBox "This receipt no is for Indoor Patient", vbCritical, " IT, DNMIH."
         Exit Sub
      End If
 End If
      
      rptMode = 4
      Viewer.Show vbModal
      TXTRECEIPT_NO = ""
      TXTRECEIPT_NO.Visible = False
End If

End Sub
Private Sub cmdSave_Click()

If UTILITY.User_Shift_validation(frmMAIN.lbluser_id, frmMAIN.lblUserType.Caption) = False Then
    MsgBox "Mr. " & frmMAIN.lblName & "  Your Shift has been Expired.. " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
    Exit Sub
End If

If TxtName = "" Then
        MsgBox "Patient Name Required", vbInformation, " IT, DNMIH"
        TxtName.SetFocus
        Exit Sub
End If

If txtAge = "" Then
    MsgBox "Age Required", vbInformation, " IT, DNMIH"
    txtAge.SetFocus
    Exit Sub
End If
If txtdisc < 0 Then
   MsgBox "Please check Discount...May be Invalid", vbCritical, " IT, DNMIH"
    txtdisc.SetFocus
    Exit Sub
End If


If txtchargeTotal = 0 And txtdisc < 0 Then
   MsgBox "Please check Discount...May be Invalid", vbCritical, " IT, DNMIH"
    txtdisc.SetFocus
    Exit Sub
End If
If txtchargeTotal = 0 And txtdisc > 0 Then
   MsgBox "Please check Discount...May be Invalid", vbCritical, " IT, DNMIH"
    txtdisc.SetFocus
    Exit Sub
End If


If Val(txtchargeTotal) < Val(txtdisc) Then
   MsgBox "Please check Discount...May be Invalid", vbCritical, " IT, DNMIH"
    txtdisc.SetFocus
    Exit Sub
End If

If Val(txtTotal) <> (Val(txtchargeTotal) - Val(txtdisc)) Then
   MsgBox "Values misMatch with the Total Amount" & vbCrLf & "Please varify", vbCritical, " IT, DNMIH"
    txtdisc.SetFocus
    Exit Sub
End If

If Check1.Value = 1 And txtstaff = "" Then
  MsgBox " Please Select a Staff ID....", vbCritical + vbOKOnly, " IT, DNMIH"
  txtstaff.SetFocus
  Exit Sub
End If

If Check1.Value = 1 And UTILITY.LOAD_STAFF(txtstaff.Text) = "0" Then
   Label7.Caption = "INVALID STAFF ID,...PLEASE VERIFY"
   Label7.ForeColor = vbRed
   txtstaff.SetFocus
  Exit Sub
End If


If DataGrid1.Row < 0 Then
  MsgBox " NO TEST SELECTED....", vbCritical + vbOKOnly, " IT, DNMIH"
  cboMainCode.SetFocus
  Exit Sub
End If

Call save_pat_info_main_out_door

MsgBox "Data Saved Successfully!!", vbInformation + vbOKOnly, "Congratulation!!!"
print_outdoor

Call flush_grid

 txtchargeTotal = 0
 txtdisc = 0
 txtdue = 0
 txtTotal = 0
 Me.nbrDisc_Per = 0
 nationalTotal = 0
 CboDept.SetFocus

End Sub
Private Sub save_pat_info_main_out_door()

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
    Dim MAX_RECEIPT_NO As Double
    
 If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    

    '----------------------------------------------------------------------------------
    Set Param0 = cmd.CreateParameter("param0", adDouble, adParamInput, 5, Trim(txtRegNoOutdoor.Text))
    cmd.Parameters.Append Param0 'regNO
    
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, frmMAIN.lblBooth)
    cmd.Parameters.Append Param1 'booth_no
                                                                    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, TxtName.Text)
    cmd.Parameters.Append Param2 'pat_name
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 40, CboDept.Text)
    cmd.Parameters.Append Param3 'pat_dept
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 25, "DOC")
    cmd.Parameters.Append Param4 'DOC ID
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 5, ComSex.Text)
    cmd.Parameters.Append Param5 'sex
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 10, txtAge.Text)
    cmd.Parameters.Append Param6 'age
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, Combo5.Text)
    cmd.Parameters.Append Param7 'religion
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 250, txtAddr.Text)
    cmd.Parameters.Append Param8 'address
    
       
    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 10, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param9 'u_id
  
    '----------------sub3 table-----------------
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, Val(txtchargeTotal.Text))
    cmd.Parameters.Append Param10 'charge total
    
    Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 10, Val(txtdisc.Text))
    cmd.Parameters.Append Param11 'disc
    
    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 10, Val(txtTotal.Text))
    cmd.Parameters.Append Param12 'net total
    
    Set Param13 = cmd.CreateParameter("param13", adInteger, adParamInput, 5, Trim(txtFlagOutdoor.Text))
    cmd.Parameters.Append Param13 'Flag
    
    Set Param14 = cmd.CreateParameter("param14", adInteger, adParamInput, 5, discount_check_val)
    cmd.Parameters.Append Param14 'DISCOUNT FLAG
   
   If discount_check_val = 0 Then
    Set Param15 = cmd.CreateParameter("param15", adVarChar, adParamInput, 10, "0")
    cmd.Parameters.Append Param15 'STAFF ID
   Else
     Set Param15 = cmd.CreateParameter("param15", adVarChar, adParamInput, 10, txtstaff)
    cmd.Parameters.Append Param15 'STAFF ID
   End If
   
    Set param16 = cmd.CreateParameter("param16", adVarChar, adParamInput, 10, CBOYRCODE.Text)
    cmd.Parameters.Append param16 'YEAR CODE
   
    Set Param17 = cmd.CreateParameter("param17", adVarChar, adParamInput, 5, cboDMY.Text)
    cmd.Parameters.Append Param17 'Y-M-D
   
   Set Param18 = cmd.CreateParameter("param18", adDouble, adParamOutput, 10, MAX_RECEIPT_NO)
   cmd.Parameters.Append Param18 'max receipt no of the specific user
    
    cmd.Properties("PLSQLRSet") = True
    
   cmd.CommandText = "{CALL SavePatient_info_out_door(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"
    
   Set RS = cmd.Execute
   
   TXTRECEIPT_NO = Param18
   txtPat_ID1 = TXTRECEIPT_NO
   
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
       Conn.Close
       Set Conn = Nothing
    End If
  Set RS = Nothing
      
End Sub
Private Sub print_outdoor()
           Dim local_con As New ADODB.Connection
           Dim local_rs As New ADODB.Recordset
           Dim local_cmd As New ADODB.Command
           
           rptMode = 4
           
           If local_con.State = 0 Then
            local_con.Open strcn.Connection_String
           End If
            Set local_cmd.ActiveConnection = local_con
            local_cmd.CommandType = adCmdText
             Dim Param1 As New Parameter
             Dim Param2 As New Parameter
             
             
              Set Param1 = local_cmd.CreateParameter("param1", adInteger, adParamInput, 3, 1)
              local_cmd.Parameters.Append Param1 'mode
              
              Set Param2 = local_cmd.CreateParameter("param2", adDouble, adParamInput, 15, Trim(TXTRECEIPT_NO))
              local_cmd.Parameters.Append Param2 'receipt no
      
                
'            Report4.DiscardSavedData
'            Screen.MousePointer = vbHourglass
'            CRViewer91.ReportSource = Report4
'            CRViewer91.ViewReport
'            Screen.MousePointer = vbDefault


'        '==========direct print==========================
'            If frmPatient_Info.txtPat_ID = "" Then
'            StrPat_ID = StPat_ID
'            Else
'            StrPat_ID = frmPatient_Info.txtPat_ID
'            End If
            
            Dim Report4   As New CrystalReport4
            local_cmd.Properties("PLSQLRSet") = True
            local_cmd.CommandText = "{CALL Rptout_door_info_print(?,?)}"
            Set local_rs = local_cmd.Execute
            local_cmd.Properties("PLSQLRSet") = False
            Report4.DiscardSavedData
            Report4.Database.SetDataSource local_rs
            Report4.PrintOut
         If local_con.State = 1 Then
            local_con.Close
            Set local_con = Nothing
         End If
         If local_rs.State = 1 Then
           Set local_rs = Nothing
         End If
         Set local_cmd = Nothing
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      cboDeptCode.SetFocus
   End If
End Sub

Private Sub cboMainCode_Click()
     Adodc3.ConnectionString = strcn.Connection_String
     Adodc3.RecordSource = "select a.m_name   from test_info_main a where to_char(a.m_code)='" & Trim(cboMainCode.Text) & "'"
     Adodc3.Refresh
        
    txtTestCode.Text = ""
        
     If Adodc3.Recordset.RecordCount > 0 Then
            txtTestCode.Text = ""
            cboMainName.clear
          
             Adodc3.Recordset.MoveFirst
        While Adodc3.Recordset.EOF = False
           cboMainName.AddItem Adodc3.Recordset!M_NAME
           Adodc3.Recordset.MoveNext
        Wend
        Else
          txtTestCode.Text = ""
       
       End If
          
    
     
     cboMainName.Text = cboMainName.List(0)
 
End Sub

Private Sub cboMainCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
     cboDeptCode.SetFocus
   End If
End Sub

Private Sub cboMainCode_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   txtTestCode.SetFocus
End If
End Sub

Private Sub ComSex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Combo5.SetFocus
   End If
End Sub
Private Sub DataGrid1_Click()
   If DataGrid1.Row >= 0 Then
          cboMainCode = DataGrid1.Columns(1)
          txtTestCode = DataGrid1.Columns(3)
          txtTestTitle.Text = DataGrid1.Columns(4)
          txttesttype.Text = DataGrid1.Columns(5)
          txtCharge.Text = DataGrid1.Columns(6)
         
   End If
End Sub
Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
           Case 2
                 PopupMenu MNUFILE
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyInsert Then
       Call cmdSave_Click
    End If
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    SendKeys Chr(9)
'End If
End Sub
Private Sub flush_grid()
    Adodc4.ConnectionString = strcn.Connection_String
    Adodc4.RecordSource = "select DEPT_CODE AS Dept,m_code AS CODE,m_name AS NAME,s_code  CODE,s_name AS NAME,test_type AS TYPE,charge from temp_test WHERE TO_NUMBER(BOOTH)=TO_NUMBER('" & frmMAIN.lblBooth & "')"
    Adodc4.Refresh
    
    With DataGrid1
        .Columns(0).Width = 480
        .Columns(1).Width = 460
        .Columns(2).Width = 1800
        .Columns(3).Width = 1000
        .Columns(4).Width = 4750
        .Columns(5).Width = 950
        .Columns(6).Width = 950
        .Columns(6).Caption = "Charge"
        
    End With
    

End Sub
Private Sub Rate_Sum()
                    Adodc5.ConnectionString = strcn.Connection_String
                    Adodc5.RecordSource = "select  nvl(sum(charge),0) charge_total from temp_test WHERE TO_NUMBER(BOOTH)=TO_NUMBER('" & frmMAIN.lblBooth & "')"
                    Adodc5.Refresh
   If Adodc5.Recordset.RecordCount > 0 Then
            txtchargeTotal.Text = Adodc5.Recordset!CHARGE_TOTAL
            
   End If
       Adodc5.Refresh
    txtchargeTotal = Val((txtchargeTotal) + Val(txtServiceTotal))
    nationalTotal = Val(txtServiceTotal)
    txtdue = txtchargeTotal - txtdisc
    
End Sub
Private Sub Form_Load()
        
        txtstaff.Visible = False
        TXTRECEIPT_NO.Visible = False
        discount_check_val = 0
        nationalTotal = 0
        Call clear_main
        MaskEdBox1.Text = Format(Date, "DD/MM/YYYY")
        DTPicker7.Value = Now
          
         Call Rate_Sum
               
         Call load_dept(0)
          
         Call LOAD_STAFF_ID
         Call Load_dept_Code
        
             
          ComSex = ComSex.List(0)
          Combo5 = Combo5.List(0)
          cboDMY = cboDMY.List(0)
          CboDept = "Card."
          
          
          cboDeptCode = "PAT"
          Call Load_Main_Code(cboDeptCode)
          cboMainCode = cboMainCode.List(0)
          
          Call flush_grid
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
Private Sub load_dept(MODE As Integer)
   If MODE = 0 Then
'       Adodc1.ConnectionString = strcn.Connection_String
'      Adodc1.RecordSource = "select distinct(doc_dept) from doctor_info"
'      Adodc1.Refresh
'
'      If Adodc1.Recordset.RecordCount > 0 Then
'          Adodc1.Recordset.MoveFirst
'          While Adodc1.Recordset.EOF = False
'             CboDept.AddItem Adodc1.Recordset!doc_dept
'              Adodc1.Recordset.MoveNext
'           Wend
'       End If

    With CboDept
         .AddItem "Card."
         .AddItem "Gynae"
         .AddItem "Surgery"
         .AddItem "Medicine"
         .AddItem "Ophth."
         .AddItem "ENT"
         .AddItem "Skin VD"
         .AddItem "Ortho."
         .AddItem "Paediatric"
         .AddItem "Dental"
         .AddItem "Physio."
         .AddItem "Emmergency"
    End With
     
     
   ElseIf MODE = 1 Then
     CboDept.AddItem OPD_OUT_INDICATION
   End If
    
End Sub
Private Sub Load_Main_Code(deptCode As String)
     
        Adodc2.ConnectionString = strcn.Connection_String
        Adodc2.RecordSource = "select distinct(m_code)from test_info_main where dept_Code='" & deptCode & "'"
        Adodc2.Refresh
      

    If Adodc2.Recordset.RecordCount > 0 Then
         cboMainCode.clear
         
        Adodc2.Recordset.MoveFirst
        While Adodc2.Recordset.EOF = False
         cboMainCode.AddItem Adodc2.Recordset!M_Code
        Adodc2.Recordset.MoveNext
        Wend
        End If
    
End Sub
Private Sub Load_dept_Code()
     
        Adodc2.ConnectionString = strcn.Connection_String
        Adodc2.RecordSource = "select distinct(dept_code)from test_info_main "
        Adodc2.Refresh
      

    If Adodc2.Recordset.RecordCount > 0 Then
         cboDeptCode.clear
         
        Adodc2.Recordset.MoveFirst
        While Adodc2.Recordset.EOF = False
         cboDeptCode.AddItem Adodc2.Recordset!dept_Code
        Adodc2.Recordset.MoveNext
        Wend
        End If
    
End Sub
Private Sub LOAD_STAFF_ID()
  '''''''''FOR ID
     Adodc1.ConnectionString = strcn.Connection_String
      Adodc1.RecordSource = "select PAYROLL.EMP_INFO.EMP_ID  AS EMP_ID from PAYROLL.EMP_INFO ORDER BY EMP_ID"
      Adodc1.Refresh

    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        While Adodc1.Recordset.EOF = False
          txtstaff.AddItem Adodc1.Recordset!EMP_ID
          
            Adodc1.Recordset.MoveNext
        Wend
    End If
End Sub
Private Sub transfer(KeyAscii As Integer)
               If KeyAscii And cboMainCode.Text = "" Then
                    txtdisc.SetFocus
                Else
                If cboMainCode.Text = "" Then
                    txtdisc.SetFocus
                End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call clear_main
End Sub

Private Sub MNUCLOSE_Click()
  cmdExit_Click
End Sub

Private Sub MNUDELETE_Click()
  CMDDELETE_Click
End Sub

Private Sub MNUREFRESH_Click()
  cmdADD_Click
End Sub

Private Sub MSFlexGrid2_DblClick()
  If Len(Trim(MSFlexGrid2.Text)) <> 0 Then
      txtTestCode.Text = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 0)
      txtTestTitle.Text = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)
      
    End If
    MSFlexGrid2.Visible = False
    txtTestCode.SetFocus
End Sub

Private Sub MSFlexGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    MSFlexGrid2.Visible = False
    txtTestCode.SetFocus
  End If
End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     MSFlexGrid2_DblClick
     txtTestCode.SetFocus
  End If
  End Sub

Private Sub nbrDisc_Per_Change()
Dim i As Double
If Not IsNumeric(nbrDisc_Per) Then
            nbrDisc_Per = 0
ElseIf Val(nbrDisc_Per) > 100 Or Val(nbrDisc_Per) < 0 Then
       MsgBox "You are allowed to discount >100 or < 100", vbCritical, " IT, DNMIH"
            nbrDisc_Per = 0
            txtdisc = 0
            nbrDisc_Per.SetFocus

Else
   If IsNull(nbrDisc_Per) = False Then
   i = (txtchargeTotal * (Val(nbrDisc_Per)) / 100)
   txtdisc = i
 txtdue = txtchargeTotal - i
 
End If
End If
End Sub

Private Sub nbrDisc_Per_GotFocus()
txtdisc = 0
End Sub

Private Sub nbrDisc_Per_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If nbrDisc_Per > 0 Then
 txtdue.SetFocus
 End If
 End If
 
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cboDMY.SetFocus
   End If
End Sub

Private Sub txtname_GotFocus()
  TxtName.BackColor = &H80000018
  TxtName.SelStart = 0
  TxtName.SelLength = Len(TxtName.Text)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      txtAge.SetFocus
   End If
End Sub

Private Sub txtname_LostFocus()
TxtName.BackColor = vbWhite

End Sub

Private Sub txtAddr_GotFocus()
  txtAddr.BackColor = &H80000018
  txtAddr.SelStart = 0
  txtAddr.SelLength = Len(txtAddr)
End Sub

Private Sub txtAddr_LostFocus()
 txtAddr.BackColor = vbWhite
End Sub

Private Sub txtAge_Change()
    If Not IsNumeric(txtAge.Text) Then
            txtAge = ""
    End If
    If Val(txtAge) > 200 Then
    MsgBox "Invalid Age", vbInformation, " IT, DNMIH"
    txtAge = ""
    End If
    
    
End Sub
Private Sub txtAge_GotFocus()
 txtAge.BackColor = &H80000018
 txtAge.SelStart = 0
 txtAge.SelLength = Len(txtAge.Text)
End Sub

Private Sub txtAge_LostFocus()
txtAge.BackColor = vbWhite
End Sub

Private Sub txtDisc_Change()
If Not IsNumeric(txtdisc.Text) Then
          txtdisc = 0
ElseIf Val(txtdisc) > Val(txtchargeTotal) Then
    MsgBox "Invalid Discount....Please Check !!!", vbCritical + vbOKOnly, " IT, DNMIH"
    txtdisc = 0
    Exit Sub
Else

 If nbrDisc_Per = "" Or nbrDisc_Per = 0 Then

txtdue = txtchargeTotal - txtdisc

End If
End If

End Sub

Private Sub txtdisc_GotFocus()
txtdisc.BackColor = &H80000018
nbrDisc_Per = 0
End Sub

Private Sub txtDisc_LostFocus()
txtdisc.BackColor = vbWhite

End Sub

Private Sub txtdue_Change()
  txtTotal = txtdue
End Sub


Private Sub TXTRECEIPT_NO_Change()
  If Not IsNumeric(TXTRECEIPT_NO) Then
      TXTRECEIPT_NO = ""
   End If

End Sub

Private Sub TXTRECEIPT_NO_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     cmdPrint_Click
  End If
End Sub

Private Sub TXTRECEIPT_NO_LostFocus()
  If TXTRECEIPT_NO = "" Then
    MsgBox "Please Enter Receipt No", vbInformation, " IT, DNMIH"
'    TXTRECEIPT_NO.SetFocus
     Exit Sub
   End If
End Sub

Private Sub txtstaff_Change()
    txtstaff_Click
End Sub
Private Sub txtstaff_Click()
    If UTILITY.LOAD_STAFF(txtstaff.Text) = "0" Then
     Label7.Caption = "INVALID STAFF ID,...PLEASE VERIFY"
     Label7.ForeColor = vbRed
  Else
     txtstaff.Text = UCase(txtstaff.Text)
     Label7.Caption = UTILITY.LOAD_STAFF(txtstaff.Text)
     Label7.ForeColor = vbWhite
  End If
End Sub

Private Sub txtTestCode_GotFocus()
      txtTestCode.SelStart = 0
      txtTestCode.SelLength = Trim(Len(txtTestCode))
End Sub
Private Sub txtTestCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeySpace Then
     cboMainCode.SetFocus
  ElseIf KeyCode = 38 Then
         txtTestCode.Text = ""
         txtTestTitle = ""
         txtCharge = ""
         cmdSAVE.SetFocus
  ElseIf KeyCode = 40 Then
      If cboMainCode.Text = "" Then
            MsgBox "Main Code Mandatory", vbInformation, "Warning: IT, DNMIH"
            cboMainCode.SetFocus
         Exit Sub
      End If
      If txtTestCode.Text = "" Then
            MsgBox "Sub Code Mandatory", vbInformation, " IT, DNMIH."
            cboMainCode.SetFocus
            Exit Sub
      End If
      
      If txtCharge.Text = " " Then
         Exit Sub
      End If
      If txtTestCode.Text <> valueBeforeManualUpdate Then
         MsgBox "Invalid Code...Press Enter to Validate", vbInformation, "IT, DNMIH"
           txtTestTitle = ""
           txtCharge = ""
           txtTestCode.SetFocus
         Exit Sub
      End If
 
         Call UTILITY.SAVE_DELETE(1, cboDeptCode.Text, cboMainCode.Text, txtTestCode.Text, cboMainName.Text, txtTestTitle.Text, OPD_OUT_INDICATION, txtCharge.Text, frmMAIN.lblBooth)

            Call flush_grid
            Call Rate_Sum

           txtTestCode.Text = ""
           txtTestTitle = ""
           txtCharge = ""
           
 End If
End Sub
Private Sub txtTestCode_KeyPress(KeyAscii As Integer)
    Dim strcn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
  
  If KeyAscii = 13 Then
     If Len(Trim(txtTestCode.Text)) = 0 Then
'        txtTestTitle = ""
'        txtCharge = ""
'        Exit Sub

   Call GetTestCode(txtTestCode)
   MSFlexGrid2.Visible = True
    MSFlexGrid2.Left = txtTestCode.Left
    MSFlexGrid2.Top = txtTestCode.Top
    MSFlexGrid2.SetFocus
  Else
    strcn.ConnectionString = Con.Connection_String
    strcn.Open
    cmd.ActiveConnection = strcn
    cmd.CommandType = adCmdText
    If OPD_OUT_INDICATION = "OPD" Then
       cmd.CommandText = "select s_name,charge_OPD charge   from test_info_sub where DEPT_code='" & Trim(cboDeptCode.Text) & "' AND  m_code='" & Trim(cboMainCode.Text) & "' and s_CODE='" & Trim(txtTestCode.Text) & "'"
    ElseIf OPD_OUT_INDICATION = "OUT" Then
        cmd.CommandText = "select s_name,charge_OUtCase charge   from test_info_sub where DEPT_code='" & Trim(cboDeptCode.Text) & "' AND  m_code='" & Trim(cboMainCode.Text) & "' and s_CODE='" & Trim(txtTestCode.Text) & "'"
    
    End If
    Set RS = cmd.Execute
    If RS.EOF = False Then
         
         txtTestTitle.Text = RS.Fields(0)
         txtCharge.Text = RS.Fields(1).Value
         
         valueBeforeManualUpdate = txtTestCode.Text
        
     Else
        Call GetTestCode(txtTestCode)
    MSFlexGrid2.Visible = True
    MSFlexGrid2.Left = txtTestCode.Left
    MSFlexGrid2.Top = txtTestCode.Top
    MSFlexGrid2.SetFocus
   End If
       
 
    End If
    Set strcn = Nothing
    Set cmd = Nothing
    Set RS = Nothing
    
  End If
  End Sub
Private Sub GetTestCode(title As String)
'  On Error GoTo err_loop
    Dim strcn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    Dim i As Integer
    MSFlexGrid2.clear
    MSFlexGrid2.Rows = 0
      
    MSFlexGrid2.ColWidth(0) = "600"
    MSFlexGrid2.ColAlignment(0) = 1
    MSFlexGrid2.ColWidth(1) = "5600"
    MSFlexGrid2.ColAlignment(2) = 1
    MSFlexGrid2.ColWidth(2) = "700"
    
    
    
    strcn.ConnectionString = Con.Connection_String
    strcn.Open
    cmd.ActiveConnection = strcn
    cmd.CommandType = adCmdText
    If OPD_OUT_INDICATION = "OPD" Then
       cmd.CommandText = "select S_CODE,s_name,charge_OPD charge   from test_info_sub where DEPT_code='" & Trim(cboDeptCode.Text) & "' AND  m_code='" & Trim(cboMainCode.Text) & "' and UPPER(s_NAME) like UPPER('" & Trim(txtTestCode.Text) & "%') ORDER BY S_CODE"
    Else
        cmd.CommandText = "select S_CODE,s_name,charge_OUtCase charge   from test_info_sub where DEPT_code='" & Trim(cboDeptCode.Text) & "' AND  m_code='" & Trim(cboMainCode.Text) & "' and UPPER(s_NAME) like UPPER('" & Trim(txtTestCode.Text) & "%') ORDER BY S_CODE"
    
    End If
      Set RS = cmd.Execute
    MSFlexGrid2.Rows = 9
    If Not RS.EOF Then
      i = 0
      With MSFlexGrid2
        Do Until RS.EOF
            MSFlexGrid2.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = Trim(RS.Fields(0))
                .TextMatrix(i, 1) = Trim(RS.Fields(1))
                .TextMatrix(i, 2) = Trim(RS.Fields(2))
                i = i + 1
            RS.MoveNext
        Loop
    End With
Else
    MSFlexGrid2.Rows = 15
         
End If
    strcn.Close
    Set cmd = Nothing
    Set RS = Nothing
   
'    MSFlexGrid2.Visible = True
'    MSFlexGrid2.Left = txtTestCode.Left
'    MSFlexGrid2.Top = txtTestCode.Top
'   ' MSFlexGrid2.TabIndex = txtTestCode.TabIndex
'    MSFlexGrid2.SetFocus
  
      Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
     
End Sub

