VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Pat_Info_out_for_indoor_test 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7725
   ClientLeft      =   -105
   ClientTop       =   390
   ClientWidth     =   11100
   FillColor       =   &H007DABD0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttesttype 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   6150
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   -30
      TabIndex        =   69
      Top             =   7380
      Width           =   11385
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
         TabIndex        =   71
         ToolTipText     =   "Developed && Maintenanced by:A.N.M. Atiqur Rahman,Software Programmer, IT Division, DNMIH"
         Top             =   90
         Width           =   2715
      End
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
         TabIndex        =   70
         Top             =   60
         Width           =   4725
      End
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   60
      TabIndex        =   15
      ToolTipText     =   "SAVE"
      Top             =   6810
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "NEW"
      Height          =   375
      Left            =   1290
      TabIndex        =   16
      ToolTipText     =   "NEW"
      Top             =   6810
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      ToolTipText     =   "DELETE"
      Top             =   6810
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   3750
      TabIndex        =   63
      ToolTipText     =   "REPORT"
      Top             =   6810
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   4980
      TabIndex        =   62
      ToolTipText     =   "CLOSE "
      Top             =   6810
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
      Left            =   7710
      TabIndex        =   53
      Top             =   6510
      Width           =   1665
   End
   Begin VB.TextBox TXTRECEIPT_NO 
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3540
      MaxLength       =   17
      TabIndex        =   52
      Top             =   6480
      Width           =   1785
   End
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   0
      Top             =   6480
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
      Caption         =   "Adodc9"
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
      Left            =   7500
      TabIndex        =   49
      Top             =   6510
      Width           =   225
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      Caption         =   "Check2"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7500
      TabIndex        =   48
      Top             =   6300
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.TextBox txt_bed_type 
      Height          =   285
      Left            =   3990
      TabIndex        =   44
      Text            =   "Text2"
      Top             =   6090
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   0
      Top             =   6480
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
      Caption         =   "Adodc8"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000001&
      Height          =   3255
      Left            =   0
      TabIndex        =   37
      Top             =   2790
      Width           =   11115
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
         ItemData        =   "Pat_Info_out_for_indoor.frx":0000
         Left            =   150
         List            =   "Pat_Info_out_for_indoor.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   330
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2460
         Left            =   2880
         TabIndex        =   67
         Top             =   720
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
      Begin VB.TextBox txtTestTitle 
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
         Height          =   315
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtTestCode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4440
         TabIndex        =   65
         Top             =   360
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Pat_Info_out_for_indoor.frx":0004
         Height          =   2505
         Left            =   120
         TabIndex        =   38
         Top             =   705
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   4419
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   -2147483624
         ForeColor       =   4210752
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
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
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
      Begin VB.TextBox txtCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   870
      End
      Begin VB.ComboBox cboMainName 
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
         ItemData        =   "Pat_Info_out_for_indoor.frx":0019
         Left            =   2130
         List            =   "Pat_Info_out_for_indoor.frx":001B
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   330
         Width           =   2340
      End
      Begin VB.ComboBox cboMainCode 
         Appearance      =   0  'Flat
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
         ItemData        =   "Pat_Info_out_for_indoor.frx":001D
         Left            =   1350
         List            =   "Pat_Info_out_for_indoor.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   330
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rec. Dept"
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
         Index           =   4
         Left            =   300
         TabIndex        =   68
         Top             =   120
         Width           =   885
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   43
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   10335
         TabIndex        =   42
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Test Name"
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
         Left            =   5700
         TabIndex        =   41
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Code"
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
         Index           =   2
         Left            =   4470
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   840
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   2250
         TabIndex        =   39
         Top             =   120
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      TabIndex        =   36
      Top             =   -90
      Width           =   11115
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "INDOOR PATIENT DIAGNOSTIC INFORMATION ENTRY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   315
         Index           =   0
         Left            =   1410
         TabIndex        =   54
         Top             =   180
         Width           =   8325
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   -570
         Picture         =   "Pat_Info_out_for_indoor.frx":0021
         Stretch         =   -1  'True
         Top             =   90
         Width           =   11820
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   2265
      Left            =   0
      TabIndex        =   27
      Top             =   510
      Width           =   11115
      Begin VB.ComboBox cboDMY 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Pat_Info_out_for_indoor.frx":59A3
         Left            =   5730
         List            =   "Pat_Info_out_for_indoor.frx":59B0
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox txtDepartment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   510
         Width           =   1875
      End
      Begin VB.TextBox txtAgeInTest 
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
         Height          =   330
         Left            =   4875
         Locked          =   -1  'True
         MaxLength       =   17
         TabIndex        =   2
         Top             =   1275
         Width           =   765
      End
      Begin VB.TextBox txtbedType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   4860
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   510
         Width           =   1875
      End
      Begin VB.TextBox txtAddrInTest 
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
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1920
         Width           =   10695
      End
      Begin VB.TextBox txtPat_ID1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc6"
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
         Height          =   330
         Left            =   300
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   510
         Width           =   2160
      End
      Begin VB.TextBox txtNameInTest 
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   330
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   1
         Top             =   1260
         Width           =   4125
      End
      Begin VB.ComboBox cboInTestSex 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Pat_Info_out_for_indoor.frx":59BD
         Left            =   7140
         List            =   "Pat_Info_out_for_indoor.frx":59C7
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1275
         Width           =   1515
      End
      Begin VB.ComboBox cboInTestReligion 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "Pat_Info_out_for_indoor.frx":59D1
         Left            =   9270
         List            =   "Pat_Info_out_for_indoor.frx":59E4
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1260
         Width           =   1680
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   315
         Left            =   7170
         TabIndex        =   55
         Top             =   510
         Width           =   1485
         _ExtentX        =   2619
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
         Left            =   9270
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "Delevary Time"
         Top             =   510
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:MM:SS"
         Format          =   22740994
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
         Left            =   5760
         TabIndex        =   61
         Top             =   1005
         Width           =   720
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
         Left            =   9300
         TabIndex        =   58
         Top             =   210
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Adm. Date"
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
         Left            =   7170
         TabIndex        =   57
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   4905
         TabIndex        =   35
         Top             =   1005
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg. No "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   330
         TabIndex        =   34
         Top             =   210
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Bed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   4860
         TabIndex        =   33
         Top             =   210
         Width           =   1230
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Indoor Patient Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   360
         TabIndex        =   32
         Top             =   1635
         Width           =   2385
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   7170
         TabIndex        =   31
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Indoor Patient Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   285
         TabIndex        =   30
         Top             =   1005
         Width           =   2190
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   2595
         TabIndex        =   29
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   9330
         TabIndex        =   28
         Top             =   1005
         Width           =   885
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   195
      Left            =   10905
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "%"
      Top             =   6390
      Width           =   150
   End
   Begin VB.TextBox txtFlagIntest 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3735
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   6180
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox txtchargeTotal_inTest 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   285
      Left            =   9465
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   6045
      Width           =   1635
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   240
      Top             =   0
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
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
   Begin VB.TextBox txtdueIntest 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9465
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   6645
      Width           =   1635
   End
   Begin VB.TextBox txtDiscInTest 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9465
      TabIndex        =   12
      Text            =   "0"
      Top             =   6345
      Width           =   780
   End
   Begin VB.TextBox txtDiscInpercent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10500
      TabIndex        =   11
      Text            =   "0"
      Top             =   6345
      Width           =   600
   End
   Begin VB.TextBox txtTotalIntest 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9465
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   6960
      Width           =   1635
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   120
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   0
      Top             =   6465
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
      Left            =   4155
      Top             =   6075
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
      Left            =   360
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   330
      Left            =   120
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
   Begin VB.Label Label12 
      BackColor       =   &H00C9AD8F&
      BackStyle       =   0  'Transparent
      Caption         =   "Test Type"
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
      Height          =   285
      Left            =   1470
      TabIndex        =   73
      Top             =   6150
      Visible         =   0   'False
      Width           =   1050
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
      Left            =   300
      TabIndex        =   64
      Top             =   6150
      Width           =   60
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   30
      Top             =   6750
      Width           =   6225
   End
   Begin VB.Label Label23 
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
      Left            =   7770
      TabIndex        =   51
      Top             =   6510
      Width           =   405
   End
   Begin VB.Label Label22 
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
      Left            =   7710
      TabIndex        =   50
      Top             =   6300
      Width           =   1095
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
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   210
      TabIndex        =   47
      Top             =   1860
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
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   210
      TabIndex        =   46
      Top             =   1650
      Width           =   1095
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flag"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3420
      TabIndex        =   26
      Top             =   6225
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   6705
      TabIndex        =   24
      Top             =   6075
      Width           =   1545
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   -6660
      TabIndex        =   19
      Top             =   4770
      Width           =   270
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   10275
      TabIndex        =   23
      Top             =   6345
      Width           =   240
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   6705
      TabIndex        =   22
      Top             =   6375
      Width           =   765
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   6705
      TabIndex        =   21
      Top             =   6660
      Width           =   855
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Left            =   6735
      TabIndex        =   20
      Top             =   6990
      Width           =   1140
   End
   Begin VB.Menu MNUFILE 
      Caption         =   "FILE"
      Visible         =   0   'False
      Begin VB.Menu MNUdl 
         Caption         =   "DELETE"
      End
      Begin VB.Menu SEP001 
         Caption         =   "-"
      End
      Begin VB.Menu MNUREF 
         Caption         =   "REFRESH"
      End
      Begin VB.Menu MNUSEP002 
         Caption         =   "-"
      End
      Begin VB.Menu MNUCLOSE 
         Caption         =   "CLOSE"
      End
   End
End
Attribute VB_Name = "Pat_Info_out_for_indoor_test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VoucherNumber
Dim discount_check_val As Integer
Public strUid As String
Dim PATIENT_BED_SERIAL_NO As Integer
Dim VAR_BED_TYPE As String
Dim UTILITY As New clsUtility
Public strcn        As New MyConnection
Dim valueBeforeManualUpdate As String
Dim nationalTotal As Integer
Private Sub load_current_dept()
   Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
  End If
  cmd.ActiveConnection = Conn
  cmd.CommandType = adCmdText
  cmd.CommandText = "SELECT doc_dept,SERIAL_NO FROM hospital_billing.INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(frmReg_no.txtReg_noInTest.Text) & "' AND YRCODE='" & Trim(frmReg_no.CBOYRCODE) & "' AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM hospital_billing.INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(frmReg_no.txtReg_noInTest.Text) & "' AND YRCODE='" & Trim(frmReg_no.CBOYRCODE) & "')"
  
  cmd.Properties("iRowsetChange") = True
  cmd.Properties("updatability") = 7
  RS.CursorLocation = adUseClient

  RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
  cmd.Properties("iRowsetChange") = False

  If RS.RecordCount > 0 Then
     txtDepartment = "" & RS!doc_dept
  End If
  
  Conn.Close
  Set Conn = Nothing
  Set RS = Nothing
  Set cmd = Nothing
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

Private Sub cboMainCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
     cboDeptCode.SetFocus
  End If
End Sub

Private Sub Check1_Click()
    ''Check1.Value = 1
 If Check1.Value = 1 Then
    Label27.Visible = False
      txtstaff.Visible = True
  If Not IsNull(txtstaff) Then
    txtDiscInTest = Val(nationalTotal)
    discount_check_val = 1
    Check2.Value = 0
 End If
 Else
  Label27.Visible = True
      txtstaff.Visible = False
End If
End Sub

Private Sub Check2_Click()
  If Check2.Value = 1 Then
     Check1.Value = 0
      txtDiscInTest = 0
     discount_check_val = 0
     txtDiscInpercent = 0
     Label7.Caption = ""
   End If
      
End Sub
Private Sub cmdADD_Click()
       Call clear_main
       Unload Me
       frmReg_no.Show
End Sub
Private Sub clear_main()
    Call UTILITY.SAVE_DELETE(3, "", "", "", "", "", "", 0, frmMAIN.lblBooth)
    Call flush_grid
  End Sub
Private Sub CMDDELETE_Click()
     Dim reply As String
    reply = MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
             Call UTILITY.SAVE_DELETE(2, cboDeptCode.Text, cboMainCode.Text, txtTestCode.Text, "", "", OPD_OUT_INDICATION, 0, frmMAIN.lblBooth)

        txtTestCode.Text = ""
         txtTestTitle = ""
         txttesttype = ""
         txtCharge = ""
        

    End If
  txtDiscInTest = 0
  txtDiscInpercent = 0
  Call flush_grid
  Call Rate_Sum
End Sub
   
Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
       
        Call clear_main
        Unload Me
    End If
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
        cmd.CommandText = "select BED_TYPE,Bed_type_no,BED_NO,SERIAL_NO  From hospital_billing.Indoor_pat_bed_info Where in_reg_no ='" & Trim(frmReg_no.txtReg_noInTest.Text) & "' AND YRCODE ='" & Trim(frmReg_no.CBOYRCODE.Text) & "'  AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM hospital_billing.Indoor_pat_bed_info WHERE in_reg_no ='" & Trim(frmReg_no.txtReg_noInTest.Text) & "' AND YRCODE='" & Trim(frmReg_no.CBOYRCODE) & "')"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False
          
       If RS.RecordCount > 0 Then
         txtbedType = "" & RS!Bed_type & "-" & RS!bed_TYPE_no & "-" & RS!bed_no
         VAR_BED_TYPE = RS!Bed_type
         PATIENT_BED_SERIAL_NO = RS!SERIAL_NO
       End If
       OPD_OUT_INDICATION = UCase(Mid(VAR_BED_TYPE, 1, 3))
       If Conn.State = 1 Then
        Conn.Close
        Set Conn = Nothing
        Set RS = Nothing
        Set cmd = Nothing
     End If

End Sub

Private Sub cmdPreview_Click()
    rptMode = 4
    Viewer.Show vbModal
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
       
            rptMode = 24
            Viewer.Show vbModal
            Me.TXTRECEIPT_NO.Visible = False
End If
TXTRECEIPT_NO = ""
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
    Dim MAX_RECEIPT_NO As Double
 
 If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
     
     Set Param0 = cmd.CreateParameter("param0", adDouble, adParamInput, 10, frmReg_no.txtReg_noInTest.Text)
    cmd.Parameters.Append Param0 'in_reg_no
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, frmMAIN.lblBooth)
    cmd.Parameters.Append Param1 'booth_no
                                                                    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, txtNameInTest.Text)
    cmd.Parameters.Append Param2 'pat_name
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 25, Trim(txtDepartment.Text))
    cmd.Parameters.Append Param3 'pat_dept.
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 10, "DOC")
    cmd.Parameters.Append Param4 'DOCTOR ID
   
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, cboInTestSex.Text)
    cmd.Parameters.Append Param5 'sex
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 25, txtAgeInTest.Text)
    cmd.Parameters.Append Param6 'age
    
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, cboInTestReligion.Text)
    cmd.Parameters.Append Param7 'religion
    
     Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 25, txtbedType.Text)
     cmd.Parameters.Append Param8 'address
    
     Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 100, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param9 'u_id
  
  '----------------sub3 table-----------------
    Set Param10 = cmd.CreateParameter("param10", adDouble, adParamInput, 10, txtchargeTotal_inTest.Text)
    cmd.Parameters.Append Param10 'charge total
    
    Set Param11 = cmd.CreateParameter("param11", adDouble, adParamInput, 10, txtDiscInTest.Text)
    cmd.Parameters.Append Param11 'disc
    
    Set Param12 = cmd.CreateParameter("param12", adDouble, adParamInput, 10, txtTotalIntest.Text)
    cmd.Parameters.Append Param12 'net total
    
    
    
    Set Param13 = cmd.CreateParameter("param13", adInteger, adParamInput, 5, 1)
    cmd.Parameters.Append Param13 'INDOOR OUT DOOR FLAG
    
    Set Param14 = cmd.CreateParameter("param14", adInteger, adParamInput, 5, discount_check_val)
    cmd.Parameters.Append Param14 'DISCOUNT FLAG
  
  
    If discount_check_val = 0 Then
    Set Param15 = cmd.CreateParameter("param15", adVarChar, adParamInput, 10, "0")
    cmd.Parameters.Append Param15 'STAFF ID
   Else
     Set Param15 = cmd.CreateParameter("param15", adVarChar, adParamInput, 10, txtstaff)
    cmd.Parameters.Append Param15 'STAFF ID
   End If
   
    Set param16 = cmd.CreateParameter("param16", adVarChar, adParamInput, 10, frmReg_no.CBOYRCODE.Text)
    cmd.Parameters.Append param16 'YEAR CODE
   
    Set Param17 = cmd.CreateParameter("param17", adVarChar, adParamInput, 5, cboDMY.Text)
    cmd.Parameters.Append Param17 'Y-M-D
       
    Set Param18 = cmd.CreateParameter("param18", adDouble, adParamOutput, 10, MAX_RECEIPT_NO)
    cmd.Parameters.Append Param18 'MAX RECEIPT NO
     
   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SavePatient_info_out_door(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}"

       
    Set RS = cmd.Execute
    
    TXTRECEIPT_NO = Param18
    
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
       Conn.Close
       Set cmd = Nothing
       Set RS = Nothing
    End If
       
End Sub

Private Sub cmdSave_Click()

If UTILITY.User_Shift_validation(frmMAIN.lbluser_id, frmMAIN.lblUserType.Caption) = False Then
    MsgBox "Mr. " & frmMAIN.lblName & "  Your Shift has been Expired.. " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
    Exit Sub
End If

If Val(txtchargeTotal_inTest) < Val(txtDiscInTest) Then
      MsgBox "Invalid Discount..Please Check", vbCritical, " IT, DNMIH."
      txtDiscInTest.SetFocus
     Exit Sub
End If

If txtDiscInTest = "" Then
   txtDiscInTest = 0
End If

If Check1.Value = 1 And txtstaff = "" Then
   MsgBox " Please Select a Staff ID....", vbCritical + vbOKOnly, " IT, DNMIH."
   txtstaff.SetFocus
   Exit Sub
End If
If txtDiscInTest < 0 Then
      MsgBox "Invalid Discount..Please Check", vbCritical, " IT, DNMIH."
      txtDiscInTest.SetFocus
     Exit Sub
End If
If txtchargeTotal_inTest = "" And txtDiscInTest > 0 Then
      MsgBox "Invalid Discount..Please Check", vbCritical, " IT, DNMIH."
      txtDiscInTest.SetFocus
     Exit Sub
End If

If Val(txtTotalIntest) <> (Val(txtchargeTotal_inTest) - Val(txtDiscInTest)) Then
   MsgBox "Values misMatch with the Total Amount" & vbCrLf & "Please varify", vbCritical, " IT, DNMIH"
    txtDiscInTest.SetFocus
    Exit Sub
End If



If DataGrid1.Row < 0 Then
  MsgBox " NO TEST SELECTED....", vbCritical + vbOKOnly, " IT, DNMIH"
  cboMainCode.SetFocus
  Exit Sub
End If

    
Call save_pat_info_main_out_door

MsgBox "Operation successful", vbInformation + vbOKOnly, "Save..."
PRINT_INDOOR

Call flush_grid
Me.TXTRECEIPT_NO.Visible = False

txtchargeTotal_inTest = 0
txtdueIntest = 0
txtTotalIntest = 0

If txtchargeTotal_inTest = 0 Then
     cboDeptCode.SetFocus
    Exit Sub
End If


End Sub
Private Sub PRINT_INDOOR()
  Dim Conn As New ADODB.Connection
  Dim RS As New ADODB.Recordset
  Dim cmd As New ADODB.Command
  
  If Conn.State = 0 Then
     Conn.Open strcn.Connection_String
  End If
  
  Set cmd.ActiveConnection = Conn
  cmd.CommandType = adCmdText
  
  Dim Param1 As New Parameter
  Dim Param2 As New Parameter
  
   Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 3, 1)
   cmd.Parameters.Append Param1 'MODE
  
   Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 100, TXTRECEIPT_NO)
   cmd.Parameters.Append Param1 'RECEIPT NO
   
   Dim Report4indoor As New CrystalReport4indoor
   cmd.Properties("PLSQLRSet") = True
   cmd.CommandText = "{CALL Rptout_door_info_print_indoor(?,?)}"
    Set RS = cmd.Execute
    cmd.Properties("PLSQLRSet") = False
          
    Report4indoor.Database.SetDataSource RS
    Report4indoor.PrintOut
    If Conn.State = 1 Then
            Conn.Close
            Set Conn = Nothing
         End If
         If RS.State = 1 Then
           Set RS = Nothing
         End If
         Set cmd = Nothing
End Sub
Private Sub cboMainCode_Click()
    Adodc3.ConnectionString = strcn.Connection_String
     Adodc3.RecordSource = "select a.m_name   from test_info_main a where to_char(a.m_code)='" & Trim(cboMainCode.Text) & "'"
     Adodc3.Refresh
        
    txtTestCode.Text = ""
    cboMainName.clear
     If Adodc3.Recordset.RecordCount > 0 Then
            txtTestCode.Text = ""
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

Private Sub cboMainCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTestCode.SetFocus
End If
End Sub



Private Sub Command2_Click()

End Sub

Private Sub comSubCode_Click()
'  Adodc2.ConnectionString = strcn.Connection_String
'        Adodc2.RecordSource = "select s_name,type,charge,service_charge from test_info_sub where m_code='" & Trim(cboMainCode.Text) & "' and s_code='" & Trim(comSubCode.Text) & "'"
'        Adodc2.Refresh
'
'If Adodc2.Recordset.RecordCount > 0 Then
'        Adodc2.Recordset.MoveFirst
''        While Adodc2.Recordset.EOF = False
'
'If Adodc2.Recordset.RecordCount > 0 Then
''      If Not IsNull(s_name) Then
'         txtSubName = Adodc2.Recordset!s_name
''         End If
'         End If
'
'If Adodc2.Recordset.RecordCount > 0 Then
'         txttesttype = Adodc2.Recordset!Type
'         End If
'
'If Adodc2.Recordset.RecordCount > 0 Then
'         txtCharge.Text = Adodc2.Recordset!charge
'         End If
'
'         If Adodc2.Recordset.RecordCount > 0 Then
'         If Not IsNull(Adodc2.Recordset!service_charge) Then
'         txtService_charge = Adodc2.Recordset!service_charge
'         End If
'
'         End If
'
'        End If
'        Call Rate_Sum
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
  If Button = 2 Then
     PopupMenu MNUFILE
  End If
End Sub

Private Sub Form_Activate()
    cboDeptCode.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyInsert Then
    cmdSave_Click
End If
If KeyCode = vbKeyEscape Then
      Call clear_main
      Unload Me
End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'          SendKeys Chr(9)
'    End If
End Sub
Private Sub flush_grid()
        Adodc4.ConnectionString = strcn.Connection_String
    Adodc4.RecordSource = "select DEPT_CODE AS Dept,m_code AS CODE,m_name AS NAME,s_code  CODE,s_name AS NAME,test_type AS TYPE,charge from temp_test WHERE TO_NUMBER(BOOTH)=TO_NUMBER('" & frmMAIN.lblBooth & "')"
    Adodc4.Refresh
    
    With DataGrid1
        .Columns(0).Width = 700
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
            txtchargeTotal_inTest.Text = Adodc5.Recordset!CHARGE_TOTAL
            
            txtchargeTotal_inTest.Text = txtchargeTotal_inTest.Text
   End If
       Adodc5.Refresh
       If Check1.Value = 1 Then
          txtDiscInpercent_Change
       End If
    txtdueIntest = Val(txtchargeTotal_inTest) - Val(txtDiscInTest)
   
    
End Sub

Private Sub Form_Load()
     txtstaff.Visible = False
     TXTRECEIPT_NO.Visible = False
     discount_check_val = 0
     DTPicker7.Value = Now
     txtPat_ID1 = frmReg_no.txtReg_noInTest
     
     Call clear_main
     
     Call LOAD_CURRENT_BED
     Call load_current_dept
     Call Load_Patient_Info
     Call Load_dept_Code
     Call Load_Main_Code(cboDeptCode.Text)
     
     If Len(txtstaff) = 0 Then
       txtstaff.Locked = False
       Call LOAD_STAFF_ID
     Else
        Check1.Value = 1
        txtDiscInpercent = 100
        txtstaff.Locked = True
     End If
     
     Call Rate_Sum
   
    
    
    
   Call flush_grid

  cboDeptCode = "PAT"
'   cboMainCode = cboMainCode.List(0)
   cboMainName = cboMainName.List(0)
  
'  cboDeptCode.SetFocus
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
        cmd.CommandText = "select pat_name,pat_guard_name,sex,age,y_m_d,religion,addr1,phone " & _
        ",admission_date,STAFF_ID  From in_door_pat_info_main Where in_reg_no ='" & Trim(frmReg_no.txtReg_noInTest.Text) & "' AND YRCODE='" & Trim(frmReg_no.CBOYRCODE) & "'"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
       
          cmd.Properties("iRowsetChange") = False
          
       If RS.RecordCount > 0 Then
          
         txtNameInTest = "" & RS!pat_name
         txtAddrInTest = "" & RS!addr1
         txtAgeInTest = "" & RS!age
         cboDMY.Text = RS!Y_M_D
         MaskEdBox1.Text = Format(RS!admission_date, "DD/MM/YYYY")
         cboInTestSex.Text = RS!sex
         cboInTestReligion = RS!religion
         txtstaff = "" & RS!STAFF_ID
       End If
      
    If Conn.State = 1 Then
       Conn.Close
       Set Conn = Nothing
       Set RS = Nothing
       Set cmd = Nothing
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

Private Sub MNUCLOSE_Click()
   cmdExit_Click
End Sub

Private Sub MNUdl_Click()
  CMDDELETE_Click
End Sub

Private Sub MNUREF_Click()
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

Private Sub txtDiscInpercent_Change()
  If Not IsNumeric(txtDiscInpercent) Then
    txtDiscInpercent = ""
  ElseIf txtDiscInpercent > 100 Or txtDiscInpercent < 0 Then
    MsgBox "Invalid Percentage...Please Check ", vbCritical, " IT, DNMIH."
           txtDiscInpercent = 0
           txtDiscInTest = 0
           
Else
If Val(txtchargeTotal_inTest) > 0 And Val(txtDiscInpercent) > 0 Then
txtDiscInTest = (txtchargeTotal_inTest * txtDiscInpercent) / 100
End If

End If

End Sub

Private Sub txtDiscInpercent_GotFocus()
txtDiscInTest = 0
End Sub

Private Sub txtDiscInpercent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Trim((txtDiscInpercent))) > 0 Then
        cmdSAVE.SetFocus
    End If
End If
End Sub

Private Sub txtDiscInpercent_LostFocus()
'If txtDiscInpercent > 0 Then
'
'cmdSAVE.SetFocus
'End If

End Sub

Private Sub txtDiscInTest_Change()
        If Not IsNumeric(txtDiscInTest.Text) Then
            txtDiscInTest = ""
'
        ElseIf Val(txtDiscInTest) > Val(txtchargeTotal_inTest) Then
             MsgBox "Invalid Discount...Please Check ", vbCritical, " IT, DNMIH."
                txtDiscInpercent = 0
                  txtDiscInTest = 0
       Else
            txtdueIntest = Val(txtchargeTotal_inTest) - Val(txtDiscInTest)
        End If
        
End Sub

Private Sub txtDiscInTest_GotFocus()
  txtDiscInpercent = 0
End Sub

Private Sub txtDiscInTest_LostFocus()
  If txtDiscInTest = "" Then
     txtDiscInTest = 0
  End If
End Sub

'Private Sub txtDiscInTest_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'If txtDiscInpercent > 0 Then
'cmdSAVE.SetFocus
'End If
'End Sub

Private Sub txtdueIntest_Change()
 txtTotalIntest = Val(txtdueIntest)
End Sub

Private Sub TXTRECEIPT_NO_Change()
  If Not IsNumeric(TXTRECEIPT_NO) Then
         TXTRECEIPT_NO = ""
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


Private Sub txtSubName_GotFocus()
'          Adodc2.ConnectionString = strcn.Connection_String
'                  Adodc2.RecordSource = "select s_code,type,charge,service_charge from test_info_sub where m_code='" & Trim(cboMainCode.Text) & "' and s_name='" & Trim(txtSubName.Text) & "'"
'                  Adodc2.Refresh
'
'          If Adodc2.Recordset.RecordCount > 0 Then
'                  Adodc2.Recordset.MoveFirst
'          '        While Adodc2.Recordset.EOF = False
'
'          If Adodc2.Recordset.RecordCount > 0 Then
'                   comSubCode = Adodc2.Recordset!s_code
'                   End If
'
'          If Adodc2.Recordset.RecordCount > 0 Then
'                   txttesttype = Adodc2.Recordset!Type
'                   End If
'
'          If Adodc2.Recordset.RecordCount > 0 Then
'                   txtCharge.Text = Adodc2.Recordset!charge
'                   End If
'
'                   If Adodc2.Recordset.RecordCount > 0 Then
'                   If Not IsNull(Adodc2.Recordset!service_charge) Then
'                   txtService_Charge = Adodc2.Recordset!service_charge
'                   End If
'
'                   End If
'
'                  End If
'                  Call Rate_Sum

End Sub

Private Sub txtSubName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
  cboMainCode.SetFocus
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
         txttesttype = ""
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
           txttesttype = ""
           txtCharge = ""
          
         txtTestCode.SetFocus
         Exit Sub
      End If
 
        Call UTILITY.SAVE_DELETE(1, cboDeptCode.Text, cboMainCode.Text, txtTestCode.Text, cboMainName.Text, txtTestTitle.Text, OPD_OUT_INDICATION, txtCharge.Text, frmMAIN.lblBooth)
        

            Call flush_grid
            
            Call Rate_Sum
           
           txtTestCode.Text = ""
           txtTestTitle = ""
           txttesttype = ""
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
'        txttesttype = ""
'        txtCharge = ""
'       Exit Sub
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
    If OPD_OUT_INDICATION = "CAB" Then
       cmd.CommandText = "select s_name,charge_Cabin charge   from test_info_sub where DEPT_code='" & Trim(cboDeptCode.Text) & "' AND  m_code='" & Trim(cboMainCode.Text) & "' and s_CODE='" & Trim(txtTestCode.Text) & "'"
    ElseIf OPD_OUT_INDICATION = "PAY" Then
        cmd.CommandText = "select s_name,charge_PAying charge   from test_info_sub where DEPT_code='" & Trim(cboDeptCode.Text) & "' AND  m_code='" & Trim(cboMainCode.Text) & "' and s_CODE='" & Trim(txtTestCode.Text) & "'"
    ElseIf OPD_OUT_INDICATION = "FRE" Then
        cmd.CommandText = "select s_name,charge_Free charge   from test_info_sub where DEPT_code='" & Trim(cboDeptCode.Text) & "' AND  m_code='" & Trim(cboMainCode.Text) & "' and s_CODE='" & Trim(txtTestCode.Text) & "'"
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
    If OPD_OUT_INDICATION = "CAB" Then
       cmd.CommandText = "select S_CODE,s_name,charge_CABIN charge   from test_info_sub where DEPT_code='" & Trim(cboDeptCode.Text) & "' AND  m_code='" & Trim(cboMainCode.Text) & "' and UPPER(s_NAME) like UPPER('" & Trim(txtTestCode.Text) & "%') ORDER BY S_CODE"
    ElseIf OPD_OUT_INDICATION = "PAY" Then
       cmd.CommandText = "select S_CODE,s_name,charge_PAYING charge   from test_info_sub where DEPT_code='" & Trim(cboDeptCode.Text) & "' AND  m_code='" & Trim(cboMainCode.Text) & "' and UPPER(s_NAME) like UPPER('" & Trim(txtTestCode.Text) & "%') ORDER BY S_CODE"
     ElseIf OPD_OUT_INDICATION = "FRE" Then
       cmd.CommandText = "select S_CODE,s_name,charge_FREE charge   from test_info_sub where DEPT_code='" & Trim(cboDeptCode.Text) & "' AND  m_code='" & Trim(cboMainCode.Text) & "' and UPPER(s_NAME) like UPPER('" & Trim(txtTestCode.Text) & "%') ORDER BY S_CODE"
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


