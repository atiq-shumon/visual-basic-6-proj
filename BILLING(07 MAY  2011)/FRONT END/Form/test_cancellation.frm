VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form test_cancellation 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7245
   ClientLeft      =   -105
   ClientTop       =   390
   ClientWidth     =   11115
   FillColor       =   &H007DABD0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   5010
      TabIndex        =   43
      ToolTipText     =   "CLOSE"
      Top             =   6780
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   3780
      TabIndex        =   42
      ToolTipText     =   "VIEW REPORT"
      Top             =   6780
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   2550
      TabIndex        =   41
      ToolTipText     =   "DELETE"
      Top             =   6780
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "NEW"
      Height          =   375
      Left            =   1320
      TabIndex        =   40
      ToolTipText     =   "NEW ENTRY"
      Top             =   6780
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   90
      TabIndex        =   39
      ToolTipText     =   "SAVE DATA"
      Top             =   6780
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc9 
      Height          =   330
      Left            =   3480
      Top             =   6600
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
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   2280
      Top             =   6240
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
      Height          =   3825
      Left            =   0
      TabIndex        =   25
      Top             =   2160
      Width           =   11115
      Begin VB.TextBox comMainCode 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   420
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   390
         Width           =   675
      End
      Begin VB.TextBox txtSubName 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   390
         Width           =   4605
      End
      Begin VB.TextBox comSubCode 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   390
         Width           =   1125
      End
      Begin VB.TextBox Combo2 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   390
         Width           =   1830
      End
      Begin VB.TextBox txtService_Charge 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   10500
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   390
         Width           =   540
      End
      Begin VB.TextBox txtCharge 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9765
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   390
         Width           =   720
      End
      Begin VB.TextBox txttesttype 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   390
         Width           =   1125
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "test_cancellation.frx":0000
         Height          =   3045
         Left            =   120
         TabIndex        =   26
         Top             =   705
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   5371
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   4
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
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "S.Charge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   10320
         TabIndex        =   33
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   1
         Left            =   435
         TabIndex        =   32
         Top             =   150
         Width           =   525
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "CHRG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   9735
         TabIndex        =   31
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "TYPE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   8640
         TabIndex        =   30
         Top             =   150
         Width           =   1050
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "S. TITLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   4920
         TabIndex        =   29
         Top             =   150
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C9AD8F&
         BackStyle       =   0  'Transparent
         Caption         =   "SUB CODE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   2
         Left            =   3045
         TabIndex        =   28
         Top             =   150
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   3
         Left            =   1470
         TabIndex        =   27
         Top             =   150
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      TabIndex        =   24
      Top             =   -120
      Width           =   11115
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TEST  CANCELLATION"
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
         Left            =   3510
         TabIndex        =   38
         Top             =   210
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -210
         Picture         =   "test_cancellation.frx":0015
         Top             =   -60
         Width           =   11820
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   0
      TabIndex        =   19
      Top             =   510
      Width           =   11115
      Begin VB.TextBox txtAgeInTest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   10395
         Locked          =   -1  'True
         MaxLength       =   17
         TabIndex        =   2
         Top             =   495
         Width           =   555
      End
      Begin VB.TextBox txtAddrInTest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1230
         Width           =   10695
      End
      Begin VB.TextBox txtPat_ID1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc6"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   330
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   510
         Width           =   2430
      End
      Begin VB.TextBox txtNameInTest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   510
         Width           =   7215
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AGE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   10395
         TabIndex        =   23
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RECEIPT NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   0
         Left            =   330
         TabIndex        =   22
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PATIENT ADDRESS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   360
         TabIndex        =   21
         Top             =   915
         Width           =   2145
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " PATIENT NAME "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   3045
         TabIndex        =   20
         Top             =   255
         Width           =   1815
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
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "%"
      Top             =   6390
      Width           =   150
   End
   Begin VB.TextBox txtchargeTotal_inTest 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc9"
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
      TabIndex        =   7
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
      TabIndex        =   10
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
      Locked          =   -1  'True
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   11
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
      Top             =   6225
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
      Left            =   4995
      Top             =   6405
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
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   30
      Top             =   6720
      Width           =   6255
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL CHARGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   7485
      TabIndex        =   18
      Top             =   6075
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   -6660
      TabIndex        =   13
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
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   10275
      TabIndex        =   17
      Top             =   6345
      Width           =   240
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DISCOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   7485
      TabIndex        =   16
      Top             =   6375
      Width           =   975
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   7485
      TabIndex        =   15
      Top             =   6660
      Width           =   1065
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL PAYABLE AMT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   7485
      TabIndex        =   14
      Top             =   6990
      Width           =   1950
   End
End
Attribute VB_Name = "test_cancellation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As New MyConnection
Dim Conn As New Connection

Dim cmd As New Command
Dim RS As New Recordset
Dim RS1 As New Recordset
Dim Conn2 As New Connection
Dim rs2 As New Recordset
Public strUid As String
Dim VoucherNumber
Public strcn        As New MyConnection
Private Sub cmdADD_Click()
     comMainCode = ""
     Combo2 = ""
     comSubCode = ""
     txtSubName = ""
     txttesttype = ""
     txtCharge = ""
     txtService_Charge = ""

End Sub
Private Sub clear_main()
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
  If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
  End If
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText

   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL delete_temp_test}"

    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
     
    Call flush_grid
  If Conn.State = 1 Then
     Conn.Close
     Set Conn = Nothing
  End If
End Sub


Private Sub CMDDELETE_Click()
                 Dim reply As String
                 reply = MsgBox("Do you want to Delete?", vbQuestion + vbYesNo, "Deleting...")
                
                
    If reply = vbYes Then
       If comMainCode.Text = "" Then
            MsgBox "Main Code Required", vbInformation, " IT, DNMIH"
            Exit Sub
       End If
  If comSubCode.Text = "" Then
            MsgBox "sub Code Required", vbInformation, " IT, DNMIH"
            Exit Sub
       End If
   If txtSubName.Text = "" Then
            MsgBox "sub Name Required", vbInformation, " IT, DNMIH"
            Exit Sub
    End If
   If txtCharge.Text = "" Then
            MsgBox "Charge Required", vbInformation, " IT, DNMIH"
            Exit Sub
    End If
    
            
'Get_Voucher_Number
txtDiscInTest = 0
Call save_test_cancellation

         ''' MsgBox "Operation successfull", vbInformation + vbOKOnly, "Success..."
flush_grid
CONFIRM_test_cancellation
'Call acct_integration_test_cancellation
'Call acct_integration_test_cancellation1
'Call post_vou

      comMainCode = ""
      Combo2 = ""
      comSubCode = ""
      txtSubName = ""
      txttesttype = ""
      txtCharge = ""
      txtService_Charge = ""

End If

  Call flush_grid
    cmdSave_Click
'Call Rate_Sum
End Sub
Private Sub acct_integration_test_cancellation()
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
    
     Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, " test Cancellation  in cash")
    cmd.Parameters.Append Param1 'comment
                                                                    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, "6101001")
    cmd.Parameters.Append Param2 ''USER_acct
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, txtCharge.Text)
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
Private Sub acct_integration_test_cancellation1()
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
    
     Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, " test Cancellation  in cash")
    cmd.Parameters.Append Param1 'comment
                                                                    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, "2101")
    cmd.Parameters.Append Param2 ''USER_acct
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, 0)
        cmd.Parameters.Append Param3 'dr
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, txtCharge.Text)
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
 
Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
       
        Call clear_main
        Unload Me
    End If
End Sub

Private Sub cmdPrint_Click()
'Viewer.Show vbModal
'
'rptMode = 4

         If Conn.State = 0 Then
            Conn.Open strcn.Connection_String
         End If
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
            
            
                
         Dim Report4   As New CrystalReport4
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL Rptout_door_info_print}"
            Set RS = cmd.Execute
            cmd.Properties("PLSQLRSet") = False
            
            Report4.Database.SetDataSource RS

            Report4.PrintOut
            Set RS = Nothing
            If Conn.State = 1 Then
               Conn.Close
           End If
'    '====================================


End Sub
Private Sub save_test_cancellation()

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
    
 If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
 End If
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
     
     Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 10, frmtest_cancel_entry.txtReg_noOpr.Text)
    cmd.Parameters.Append Param1 'in_reg_no
    
    
     Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 10, Trim(comMainCode.Text))
          cmd.Parameters.Append Param2 '''m_code3
     Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 15, Trim(comSubCode.Text))
          cmd.Parameters.Append Param3 'subcode
     
    
'    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, frmMAIN.lblBooth)
'    cmd.Parameters.Append Param1 'booth_no
'
'    Set Param9 = cmd.CreateParameter("param9", adVarChar, adParamInput, 100, frmMAIN.lbluser_id)
'    cmd.Parameters.Append Param9 'u_id
'
    
  '----------------sub3 table-----------------
  If Val(txtchargeTotal_inTest) = 0 Then
     txtDiscInTest = 0
  End If
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, Trim(txtDiscInTest.Text))
    cmd.Parameters.Append Param4 'disc
    
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 10, Trim(txtchargeTotal_inTest.Text))
    cmd.Parameters.Append Param5 'charge total
    
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 10, Trim(txtTotalIntest.Text))
    cmd.Parameters.Append Param6 'net total
        
    Set Param7 = cmd.CreateParameter("param7", adVarChar, adParamInput, 10, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param7 'net total
      
   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL test_cancellation(?,?,?,?,?,?,?)}"

    
    
    Set RS = cmd.Execute
    cmd.Properties("PLSQLRSet") = False
 If Conn.State = 1 Then
    Conn.Close
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
             MsgBox "Your Working Time has been Expired", vbInformation, " IT, DNMIH."
             Exit Sub
             End If
            


Call CONFIRM_test_cancellation
          MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
flush_grid

comMainCode = ""
Combo2 = ""
comSubCode = ""
txtSubName = ""
txttesttype = ""
txtCharge = ""
txtService_Charge = ""
If Conn.State = 1 Then
   Conn.Close
   Set Conn = Nothing
End If
End Sub
Private Sub CONFIRM_test_cancellation()
        Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
  If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
  End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
     
    Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 10, frmtest_cancel_entry.txtReg_noOpr.Text)
    cmd.Parameters.Append Param1 'in_reg_no
  '----------------sub3 table-----------------
    Set Param2 = cmd.CreateParameter("param2", adDouble, adParamInput, 10, Trim(txtDiscInTest.Text))
    cmd.Parameters.Append Param2 'disc
    
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, Trim(txtchargeTotal_inTest.Text))
    cmd.Parameters.Append Param3 'charge total
    
    
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, Trim(txtTotalIntest.Text))
    cmd.Parameters.Append Param4 'net total
        
   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SAVE_test_cancellation(?,?,?,?)}"

    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
       Conn.Close
       Set Conn = Nothing
    End If
    
End Sub
Private Sub DataGrid1_Click()
If DataGrid1.Row >= 0 Then
       comMainCode = DataGrid1.Columns(0)
       Combo2.Text = DataGrid1.Columns(1)
       comSubCode = DataGrid1.Columns(2)
       txtSubName.Text = DataGrid1.Columns(3)
       txttesttype.Text = DataGrid1.Columns(4)
       txtCharge.Text = DataGrid1.Columns(5)
       txtService_Charge = 0
       End If
End Sub

Private Sub Form_Activate()
         comMainCode.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
          Unload Me
End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys Chr(9)
End If


End Sub
Private Sub flush_grid()
    Adodc4.ConnectionString = strcn.Connection_String
    Adodc4.RecordSource = "select A.m_code,(SELECT M_NAME FROM TEST_INFO_MAIN B " & _
    "WHERE TO_NUMBER(B.M_CODE)=TO_NUMBER(A.M_CODE)) AS m_name,s_code, " & _
    "(SELECT S_NAME FROM TEST_INFO_SUB B WHERE TO_NUMBER(B.M_CODE)=TO_NUMBER(A.M_CODE) AND TO_NUMBER(A.S_CODE)=TO_NUMBER(B.S_CODE)) AS s_name,'' AS TEST_TYPE,test_charge from pat_info_sub1_out_door A Where A.reg_no ='" & Trim(frmtest_cancel_entry.txtReg_noOpr.Text) & "'"
    Adodc4.Refresh
    
    DataGrid1.Columns(0).Width = 615
    DataGrid1.Columns(1).Width = 1845
    DataGrid1.Columns(2).Width = 1070
    DataGrid1.Columns(3).Width = 4695
    DataGrid1.Columns(4).Width = 1080
    DataGrid1.Columns(5).Width = 710
''    DataGrid1.Columns(6).Width = 590
   
    Adodc9.ConnectionString = strcn.Connection_String
    Adodc9.RecordSource = "select nvl(sum(test_charge),0)as sum_total,0 as service_sum from pat_info_sub1_out_door Where reg_no ='" & Trim(frmtest_cancel_entry.txtReg_noOpr.Text) & "'"
    Adodc9.Refresh
    If Adodc9.Recordset.RecordCount > 0 Then
              txtchargeTotal_inTest = Adodc9.Recordset!sum_total
    End If


End Sub
Private Sub Rate_Sum()
                     Adodc5.ConnectionString = strcn.Connection_String
                    Adodc5.RecordSource = "select  nvl(sum(charge),0) charge_total from temp_test"
                    Adodc5.Refresh
   If Adodc5.Recordset.RecordCount > 0 Then
            txtchargeTotal_inTest.Text = Adodc5.Recordset!CHARGE_TOTAL
'            txtchargeTotal_inTest.Text = txtchargeTotal_inTest.Text + Adodc5.Recordset!service_charge
   End If
       Adodc5.Refresh
      ' txtDiscInTest = 0
'    txtdueIntest = Val(txtchargeTotal_inTest) - Val(txtDiscInTest)
   
    
End Sub

Private Sub Form_Load()
                       txtPat_ID1 = frmtest_cancel_entry.txtReg_noOpr.Text
        
      Dim temp
                   
      Call Rate_Sum
     If Conn2.State = 0 Then
        Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
     End If
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select pat_name,age  From pat_info_main_out_door Where reg_no ='" & Trim(frmtest_cancel_entry.txtReg_noOpr.Text) & "'"
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        rs2.CursorLocation = adUseClient
        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
        If rs2.RecordCount > 0 Then
              If Not IsNull(rs2!pat_name) Then
                  txtNameInTest = rs2!pat_name
              End If
         End If
         
      
       If rs2.RecordCount > 0 Then
          If Not IsNull(rs2!age) Then
                   txtAgeInTest = rs2!age
          End If
       End If
       

  
                                                                                                                                                                                                              
        
 cmd.Properties("iRowsetChange") = False
    
   Call flush_grid

Adodc8.ConnectionString = strcn.Connection_String
    Adodc8.RecordSource = "SELECT DISC AS DISCOUNT FROM PAT_INFO_SUB3_OUT_DOOR  Where reg_no ='" & Trim(frmtest_cancel_entry.txtReg_noOpr.Text) & "'"
    Adodc8.Refresh
    If Adodc8.Recordset.RecordCount > 0 Then
        txtDiscInTest.Text = Adodc8.Recordset!DISCOUNT
        ''''PREVIEW_VAR = Val(TXTPRINT_PREV_FREE)
    End If

    rptMode = 4
    
txtdueIntest = Val(txtchargeTotal_inTest) - Val(txtDiscInTest)

  If Conn2.State = 1 Then
    Conn2.Close
    Set Conn2 = Nothing
  End If
  
'   rs2.Close
   End Sub

Private Sub nbrDisc_Per_Change()

End Sub

Private Sub txtchargeTotal_inTest_Change()
'    txtDiscInTest = Val(txtDiscInTest) - Val(txtCharge)
    txtdueIntest = Val(txtchargeTotal_inTest) - Val(txtDiscInTest)

End Sub

Private Sub txtDiscInpercent_Change()
If Not IsNumeric(txtDiscInpercent) Then
txtDiscInpercent = ""
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
'         ElseIf IsNumeric(txtDiscInTest.Text) < 0 Then
'            MsgBox "Discount can't be Negative", vbInformation, " IT, DNMIH."
'
        Else
            txtdueIntest = Val(txtchargeTotal_inTest) - Val(txtDiscInTest)
        End If
        
End Sub

Private Sub txtDiscInTest_GotFocus()
txtDiscInpercent = 0
End Sub

Private Sub txtdueIntest_Change()
  txtTotalIntest = Val(txtdueIntest)
End Sub



