VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPatientfled 
   BackColor       =   &H00C0C0C0&
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
   Begin VB.Frame Frame1 
      Caption         =   "Patient History"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   8775
      Left            =   120
      TabIndex        =   64
      Top             =   810
      Width           =   7185
      Begin VB.CommandButton cmdSAVE 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         Picture         =   "frmPatientfled.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Save"
         Top             =   7665
         Width           =   495
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1125
         Picture         =   "frmPatientfled.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exit"
         Top             =   7650
         Width           =   495
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   615
         Picture         =   "frmPatientfled.frx":0F88
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Print"
         Top             =   7665
         Width           =   495
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Test Info"
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
         Height          =   2370
         Left            =   60
         TabIndex        =   71
         Top             =   5100
         Width           =   6960
         Begin MSDataGridLib.DataGrid DataGrid2 
            Height          =   2040
            Left            =   45
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   195
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   3598
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16777215
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
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   60
         TabIndex        =   69
         Top             =   1995
         Width           =   6960
         Begin MSDataGridLib.DataGrid DataGrid4 
            Height          =   870
            Left            =   75
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   210
            Width           =   6750
            _ExtentX        =   11906
            _ExtentY        =   1535
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16777215
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   1830
         Left            =   60
         TabIndex        =   67
         Top             =   3240
         Width           =   6960
         Begin MSDataGridLib.DataGrid DataGrid5 
            Bindings        =   "frmPatientfled.frx":13CA
            Height          =   1485
            Left            =   60
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   270
            Width           =   6750
            _ExtentX        =   11906
            _ExtentY        =   2619
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16777215
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   1500
         Left            =   60
         TabIndex        =   65
         Top             =   450
         Width           =   6960
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   1200
            Left            =   45
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   225
            Width           =   6750
            _ExtentX        =   11906
            _ExtentY        =   2117
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16777215
            HeadLines       =   1
            RowHeight       =   15
            WrapCellPointer =   -1  'True
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
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fled Patient Calculation Entry"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   465
         Left            =   1800
         TabIndex        =   83
         Top             =   7650
         Width           =   5010
      End
      Begin VB.Shape Shape1 
         Height          =   525
         Left            =   60
         Top             =   7620
         Width           =   1620
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
      BorderStyle     =   0  'None
      Height          =   8460
      Left            =   7245
      TabIndex        =   32
      Top             =   795
      Width           =   4275
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
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   7680
         Width           =   1185
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
         Left            =   2625
         TabIndex        =   8
         Text            =   "0"
         Top             =   4065
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
         TabIndex        =   10
         Text            =   "0"
         Top             =   4770
         Width           =   1590
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Check2"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1140
         TabIndex        =   77
         Top             =   7470
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1140
         TabIndex        =   76
         Top             =   7710
         Width           =   225
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
         TabIndex        =   1
         Text            =   "0"
         Top             =   1695
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
         TabIndex        =   9
         Text            =   "0"
         Top             =   4440
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
         TabIndex        =   2
         Text            =   "0"
         Top             =   2025
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
         TabIndex        =   18
         Text            =   "0"
         Top             =   7545
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
         TabIndex        =   16
         Text            =   "0"
         Top             =   6840
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
         TabIndex        =   6
         Top             =   3375
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
         TabIndex        =   7
         Top             =   3720
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
         TabIndex        =   5
         Top             =   3030
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
         TabIndex        =   4
         Top             =   2700
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
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2370
         Width           =   1590
      End
      Begin VB.TextBox txtNetTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   330
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   7875
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
         TabIndex        =   19
         Text            =   "0"
         Top             =   7545
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
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   0
         Width           =   1590
      End
      Begin VB.TextBox txtWithAdmissionCharge 
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
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   345
         Width           =   1590
      End
      Begin VB.TextBox txtTotal 
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   5115
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
         TabIndex        =   15
         Text            =   "0"
         Top             =   6480
         Width           =   1590
      End
      Begin VB.TextBox txtmiscelleneous 
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
         Height          =   330
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   7185
         Width           =   1590
      End
      Begin VB.TextBox txtDeductTestTotal 
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
         Height          =   330
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   -195
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txtD_tolal 
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
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   6120
         Width           =   1590
      End
      Begin VB.TextBox txtTestTotal 
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
         Height          =   285
         Left            =   -435
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   -15
         Visible         =   0   'False
         Width           =   750
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
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1020
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
         TabIndex        =   0
         Top             =   1350
         Width           =   1590
      End
      Begin VB.TextBox txtServiceCharge 
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
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   690
         Width           =   1590
      End
      Begin VB.TextBox txtAdvanceRelease 
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
         Height          =   285
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   5445
         Width           =   1590
      End
      Begin VB.TextBox txtGrandTotal 
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   5775
         Width           =   1590
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incubator  Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   150
         TabIndex        =   81
         Top             =   4095
         Width           =   1485
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nebulizer Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   90
         TabIndex        =   80
         Top             =   4815
         Width           =   1470
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
         Left            =   1410
         TabIndex        =   79
         Top             =   7470
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
         Left            =   1410
         TabIndex        =   78
         Top             =   7680
         Width           =   405
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Anaesthesia Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   60
         TabIndex        =   75
         Top             =   1725
         Width           =   1710
      End
      Begin VB.Label lab10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CCU Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   90
         TabIndex        =   74
         Top             =   4485
         Width           =   1050
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   90
         TabIndex        =   73
         Top             =   2055
         Width           =   1350
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
         TabIndex        =   61
         Top             =   7590
         Width           =   210
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   58
         Top             =   6870
         Width           =   1410
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Sugar Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   57
         Top             =   3750
         Width           =   1680
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Photo Therapy Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   56
         Top             =   3405
         Width           =   1845
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ex.Transfusion Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   55
         Top             =   3060
         Width           =   1905
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Neunetal Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   54
         Top             =   2730
         Width           =   1395
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Baby-care Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   53
         Top             =   2400
         Width           =   1485
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Paid Amount"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   52
         Top             =   7950
         Width           =   1050
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   51
         Top             =   7530
         Width           =   735
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   50
         Top             =   45
         Width           =   975
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Fee"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   49
         Top             =   375
         Width           =   1200
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   48
         Top             =   5160
         Width           =   465
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Miscelleneous Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   47
         Top             =   6525
         Width           =   1845
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Total Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   46
         Top             =   7230
         Width           =   1410
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   45
         Top             =   6180
         Width           =   465
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Ext. BedCharge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   44
         Top             =   1050
         Width           =   1740
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Operation Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   43
         Top             =   1380
         Width           =   1950
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Charge"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   42
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label92 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   41
         Top             =   5475
         Width           =   675
      End
      Begin VB.Label Label15 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2190
         TabIndex        =   40
         Top             =   5505
         Width           =   315
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   39
         Top             =   5820
         Width           =   975
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
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   120
      TabIndex        =   25
      Top             =   -60
      Width           =   11355
      Begin VB.TextBox txtRegNoShow 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1920
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   495
         Width           =   2145
      End
      Begin VB.ComboBox comDepartmentRelease 
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
         Height          =   315
         Left            =   8325
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "comDepartmentFree"
         Top             =   150
         Width           =   1965
      End
      Begin VB.TextBox txtNameRelease 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   180
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   8325
         TabIndex        =   26
         Top             =   495
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   255
         CalendarTitleForeColor=   255
         Format          =   50987009
         CurrentDate     =   38043
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
         Left            =   1710
         TabIndex        =   62
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   59
         Top             =   495
         Width           =   1395
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6480
         TabIndex        =   31
         Top             =   150
         Width           =   1170
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   360
         TabIndex        =   30
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
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   6450
         TabIndex        =   29
         Top             =   495
         Width           =   1500
      End
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "B I L L   C A L C U L A T I O N"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   6075
      Left            =   11640
      TabIndex        =   63
      Top             =   750
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Total"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7080
      TabIndex        =   24
      Top             =   9480
      Width           =   660
   End
End
Attribute VB_Name = "frmPatientfled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As New MyConnection
Dim conn As New Connection
Dim Conn2 As New Connection
Dim Conn3 As New Connection
Dim Conn4 As New Connection
Dim Conn5 As New Connection
Dim Conn6 As New Connection
Dim Conn7 As New Connection
Dim Conn8 As New Connection
Dim connext As New Connection
Dim cmd As New Command
Dim rs As New Recordset
Dim rs2 As New Recordset
Dim rs3 As New Recordset
Dim RS4 As New Recordset
Dim rs5 As New Recordset
Dim rs6 As New Recordset
Dim rs7 As New Recordset
Dim rs8 As New Recordset
Dim rsext As New Recordset
Dim VoucherNumber
Dim discount_check_val As Integer
Public strUid As String
Dim UTILITY As New clsUtility
Public strcn        As New MyConnection

Private Sub cmdADD_Click()
      Unload Me
End Sub

Private Sub Check1_Click()
      
   If Check1.Value = 1 Then
      Label27.Visible = False
      txtstaff.Visible = True
 If Not IsNull(Val(txtstaff)) Then
'    txtDisc = Val(txtmiscelleneous)
    txtdisc_percent = 100
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
     txtDisc = 0
     txtdisc_percent = 0
     Check1.Value = 0
     discount_check_val = 0
    txtNetTotal = txtmiscelleneous
   End If
   
End Sub

Private Sub CMDEXIT_Click()
Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Call delete_temp_calculation
'        rs7.Close
'        Conn7.Close
        Unload Me
    End If
End Sub
Private Sub delete_temp_calculation()
   Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    If conn.State = 0 Then
        conn.Open strcn.Connection_String
   End If
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    '----------------------------------------------------------------------------------
    
   cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL delete_temp_calculation_indoor}"
    
    Debug.Print cmd.CommandText
    
    Set rs = cmd.Execute
    

  cmd.Properties("PLSQLRSet") = False
  If conn.State = 1 Then
     conn.Close
     Set conn = Nothing
  End If
End Sub

Private Sub cmdPrint_Click()
              Dim validation As Variant
              Adodc1.ConnectionString = strcn.Connection_String
              Adodc1.RecordSource = "Select user_id, user_name, user_password,user_type,shift_name From Security Where (User_Id = '" & frmMAIN.lbluser_id & "')"
              Adodc1.Refresh
              validation = Adodc1.Recordset!user_id
                
                Dim conn As New ADODB.Connection
                Dim cmd As New ADODB.Command
                Dim rs As New ADODB.Recordset

                        Dim Param1 As New ADODB.Parameter
                    If conn.State = 0 Then
                        conn.Open strcn.Connection_String
                     End If
                    Set cmd.ActiveConnection = conn
                    cmd.CommandType = adCmdText
    
                   Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, validation)
                    cmd.Parameters.Append Param1 'validation
                    cmd.Properties("PLSQLRSet") = True
    
                     cmd.CommandText = "{CALL shift_validation(?)}"
    
                Debug.Print cmd.CommandText
    
                    Set rs = cmd.Execute
    

                cmd.Properties("PLSQLRSet") = False
                Set cmd = Nothing
                 Set rs = Nothing
                  Set conn = Nothing
          Adodc2.ConnectionString = strcn.Connection_String
          Adodc2.RecordSource = "Select * From user_validation"
          Adodc2.Refresh
        

        
             If Adodc2.Recordset!validation = 0 Then
                    MsgBox "Your Working Time has been Expired", vbInformation, " IT, DNMIH."
                    Exit Sub
             End If

Dim reply As String
    reply = MsgBox("Are you sure to print?", vbQuestion + vbYesNo, "Print...")
    If reply = vbYes Then
        ''Get_Voucher_Number
'        Call save_release_info_rough
          rptMode = 31
         '''' Viewer.Show vbModal
              '''print_release_rough
      
'        Unload Me
'        frmfled.Show
        Else

    Unload Me
     frmfled.Show
End If
Set cmd = Nothing
Set rs = Nothing
End Sub
Private Sub cmdSave_Click()
    If UTILITY.User_Shift_validation(frmMAIN.lbluser_id, frmMAIN.lblUserType) = False Then
        MsgBox "Mr. " & frmMAIN.lblName & "  Your Shift has been Expired.. " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
        Exit Sub
    End If
    
             
   '''''''''DEPARTMENT WISE PAYMENT VALIDATION''''''''''''''''''
             
'  If UPPER(comDepartmentRelease) <> UPPER("Gynae-1") Or UPPER(comDepartmentRelease) <> UPPER("Gynae-2") Then
'      If Val(txtDeliveryCharge) > 0 Then
'          MsgBox "Invalid Payment , Delivery charge is only applicable for Gynae Department", vbCritical, " IT, DNMIH."
'          Exit Sub
'          txtDeliveryCharge.SetFocus
'      End If
'  End If
'
'
'   If UPPER(comDepartmentRelease) <> UPPER("Surgery-1") Or UPPER(comDepartmentRelease) <> UPPER("Surgery-2") Or UPPER(comDepartmentRelease) <> UPPER("Gynae-1") Or UPPER(comDepartmentRelease) <> UPPER("Gynae-2") Or UPPER(comDepartmentRelease) <> UPPER("ENT") Or UPPER(comDepartmentRelease) <> UPPER("OPTH.") Then
'      If Val(txtTotalOpr) > 0 Then
'          MsgBox "Invalid Payment , Operation charge is not applicable for this Department", vbCritical, " IT, DNMIH."
'          Exit Sub
'          txtTotalOpr.SetFocus
'      End If
'  End If
'
  
  '''''''''End of DEPARTMENT WISE PAYMENT VALIDATION''''''''''''''''''
             
             
             
  Dim reply As String
          reply = MsgBox("Are you sure to LOCK this registration?", vbQuestion + vbYesNo, "Locking...")
              If reply = 6 Then
                  If Check1.Value = 1 And (txtstaff) = "" Then
                        MsgBox "Please Select a  Staff Id ", vbCritical + vbOKOnly, " IT, DNMIH."
                        txtstaff.SetFocus
                        Exit Sub
                   End If
               
                If Check1.Value = 0 And Check2.Value = 0 Then
                     MsgBox "Please select check box", vbInformation, " IT, DNMIH"
                        Exit Sub
                End If

                   Call save_fled_info
           
               MsgBox "Account Locked Successfully ", vbCritical + vbOKOnly, " IT, DNMIH."
              Unload Me
'              frmfled.Show vbModal
        Else

            Unload Me
            frmfled.Show
      End If
     Set conn = Nothing
     Set cmd = Nothing
     Set rs = Nothing
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyInsert Then
            cmdSave_Click
End If
If KeyCode = vbKeyEscape Then
  CMDEXIT_Click
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub

Private Sub Form_Load()
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

         txtstaff.Visible = False
         discount_check_val = 0
        txtRegNoShow = frmfled.txtRegNoRelease
      If Conn2.State = 0 Then
        Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
      End If
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select pat_name,doc_dept,admission_date  From in_door_pat_info_main Where in_reg_no ='" & Trim(frmfled.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmfled.CBOYRCODE) & "'"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        rs2.CursorLocation = adUseClient

        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False

        If rs2.RecordCount > 0 Then
         txtNameRelease = rs2!pat_name
         DTPicker1.Value = rs2!admission_date
         comDepartmentRelease = rs2!doc_dept
         
        Conn3.ConnectionString = strcn.Connection_String
        Conn3.Open
        cmd.ActiveConnection = Conn3
        cmd.CommandType = adCmdText
        cmd.CommandText = "select  nvl(sum(advance),0) as advance  From advance Where in_reg_no ='" & Trim(frmfled.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmfled.CBOYRCODE) & "'"
      
       cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        rs3.CursorLocation = adUseClient
'        RS3.MoveFirst
         'If RS3.RecordCount > 0 Then
        rs3.Open cmd.CommandText, Conn3, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False

        If IsNull(rs3!advance) = True Then
          txtAdvanceRelease = 0
         Else
         
         txtAdvanceRelease = rs3!advance
         End If
         
    '''''''''''''''''''''''''''for bed charge''''''''''''''''''''
   If Conn4.State = 0 Then
        Conn4.ConnectionString = strcn.Connection_String
        Conn4.Open
    End If
        cmd.ActiveConnection = Conn4
        cmd.CommandType = adCmdText
        cmd.CommandText = "select  bed_no,bed_type,bed_charge,admission_charge,admission_date  From Indoor_pat_bed_info Where in_reg_no ='" & Trim(frmfled.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmfled.CBOYRCODE) & "'"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS4.CursorLocation = adUseClient
''        RS4.MoveFirst

        RS4.Open cmd.CommandText, Conn4, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False

        If RS4.RecordCount > 0 Then
         Set DataGrid1.DataSource = RS4
    
    
         End If
         
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Conn5.State = 0 Then
       Conn5.ConnectionString = strcn.Connection_String
        Conn5.Open
     End If
        cmd.ActiveConnection = Conn5
        cmd.CommandType = adCmdText
        cmd.CommandText = "select  m_code,m_name,s_code,s_name,test_type,test_charge From Pat_Info_Sub1_out_door Where in_reg_no ='" & Trim(frmfled.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmfled.CBOYRCODE) & "'"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        rs5.CursorLocation = adUseClient
        'RS4.MoveFirst
        'If rs5.RecordCount > 0 Then
        rs5.Open cmd.CommandText, Conn5, adOpenDynamic, adLockOptimistic
         Set DataGrid2.DataSource = rs5
        cmd.Properties("iRowsetChange") = False

      'End If
    If Conn6.State = 0 Then
        Conn6.ConnectionString = strcn.Connection_String
        Conn6.Open
    End If
        cmd.ActiveConnection = Conn6
        cmd.CommandType = adCmdText
        cmd.CommandText = "select temp_nebulizer_charge, temp_anaesthesia_charge,TEMP_CCU_CHARGE,temp_operation_sum, temp_service_charge,temp_test_sum,temp_ext_bed_sum,temp_bed_sum,temp_admission_charge,temp_baby_care_charge,temp_Neunetal_bed_charge,temp_EX_transfusion_charge,photo_therapy_charge,blood_sugar_charge ,temp_total from  temp_calculation_indoor"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        rs6.CursorLocation = adUseClient
      
         rs6.Open cmd.CommandText, Conn6, adOpenDynamic, adLockOptimistic
         cmd.Properties("iRowsetChange") = False
         ''nebulizer charge
          If Not IsNull(rs6!temp_nebulizer_charge) Then
              txtNebuliser = rs6!temp_nebulizer_charge
          Else
            txtNebuliser = 0
          End If
         ''''''''''''''''''''''''''''''''''''Anaesthesia_charge
         If Not IsNull(rs6!temp_anaesthesia_charge) Then
                      Text1 = rs6!temp_anaesthesia_charge
         Else
            Text1 = 0
         End If
         
         ''''''''''''''''CCU CHARGE''''''''''''''''''''''''''
         If Not IsNull(rs6!TEMP_CCU_CHARGE) Then
                      txtCCU_Charge = rs6!TEMP_CCU_CHARGE
         Else
                    txtCCU_Charge = 0
         End If
         
         
         
         
 '''''''test sum'''''''''''''''''''
        If IsNull(rs6!temp_test_sum) = True Then
                txtTestTotal = 0
        Else
                txtTestTotal = rs6!temp_test_sum
        End If
''''''''''''''''''service sum''''''''''''''''''''''''''''

        If IsNull(rs6!temp_service_charge) = True Then
                    txtServiceCharge = 0
         Else
                txtServiceCharge = rs6!temp_service_charge
         End If
         
'''''''''''''''operation sum'''''''''''''''''''''
          
        If IsNull(rs6!temp_operation_sum) = True Then
                txtTotalOpr = 0
        Else
            txtTotalOpr = rs6!temp_operation_sum
        End If
        
        '''''''''''baby care''''''''''''
        If Not IsNull(rs6!temp_baby_care_charge) Then
           txtBabyCareCharge = rs6!temp_baby_care_charge
        Else
            txtBabyCareCharge = 0
        End If
        If Not IsNull(rs6!temp_Neunetal_bed_charge) Then
             txtNeunetalCharge = rs6!temp_Neunetal_bed_charge
        Else
        txtNeunetalCharge = 0
       End If
       If Not IsNull(rs6!temp_EX_transfusion_charge) Then
             txtExtTransfusion = rs6!temp_EX_transfusion_charge
       Else
            txtExtTransfusion = 0
       End If
       
       If Not IsNull(rs6!photo_therapy_charge) Then
             txtP_therapyCharge = rs6!photo_therapy_charge
       Else
           txtP_therapyCharge = 0
       End If
       If Not IsNull(rs6!blood_sugar_charge) Then
              txtBloodTher_charge = rs6!blood_sugar_charge
       Else
       txtBloodTher_charge = 0
       End If
       
        
'''''''''''''''''''''''''''''extra bed''''''''''''''''''''''
   If connext.State = 0 Then
        connext.ConnectionString = strcn.Connection_String
        connext.Open
   End If
        cmd.ActiveConnection = connext
        cmd.CommandType = adCmdText
        cmd.CommandText = "select  start_date,end_date,bed_charge  From Indoor_pat_Extra_bed_info Where in_reg_no ='" & Trim(frmfled.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmfled.CBOYRCODE) & "'"
      
        cmd.Properties("iRowsetChange") = True
       cmd.Properties("updatability") = 7
        rs7.CursorLocation = adUseClient
            cmd.Properties("iRowsetChange") = False

        rs7.Open cmd.CommandText, connext, adOpenDynamic, adLockOptimistic
       

        If rs7.RecordCount > 0 Then
              Set DataGrid4.DataSource = rs7
         End If
         ''''cmd.Properties("iRowsetChange") = False
'''''''''''''''''''''''''operation info'''''''''''''
       Adodc1.ConnectionString = strcn.Connection_String
       Adodc1.RecordSource = "select opr_name,opr_charge,annay_charge,service_charge from indoor_pat_operation_info where in_reg_no='" & Trim(frmfled.txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(frmfled.CBOYRCODE) & "'"
       Adodc1.Refresh
        

         If Not IsNull(rs6!temp_ext_bed_sum) Then
         txtExtraBedTotal = Val(rs6!temp_ext_bed_sum & "")
         Else
         txtExtraBedTotal = 0
         End If
         If Not IsNull(rs6!temp_bed_sum) Then
            txtTotalBedCharge = Val(rs6!temp_bed_sum)
         Else
            txtTotalBedCharge = 0
         End If
         If Not IsNull(rs6!temp_admission_charge) Then
          txtWithAdmissionCharge = Val(rs6!temp_admission_charge)
         End If
         
         
         '''''''''total''''''''''''
         txtTotal.Text = Val(Text1) + Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtServiceCharge) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtCCU_Charge) + Val(txtNebuliser)
         
         ''''''''''''grand total''''''''''''''''''
         txtGrandTotal = Val(txtTotal.Text) - Val(txtAdvanceRelease)
         
         ''''''''deduct test_total'''''''''''''''''''''''
         
         txtD_tolal = Val(txtGrandTotal.Text)
'
            txtDeductTestTotal = Val(txtTestTotal)

        txtmiscelleneous = Val(txtD_tolal)
        txtNetTotal = Val(txtD_tolal)

       
    
  Else
        MsgBox "Invalid Registration Number"
        Exit Sub
       
    
      Set rs3 = Nothing
    If Conn3.State = 1 Then
        Conn3.Close
        Set Conn3 = Nothing
    End If
    
      Set rs2 = Nothing
      If Conn2.State = 1 Then
        Conn2.Close
        Set Conn2 = Nothing
      End If
        
           
         Unload Me
    
    End If
       Set rs3 = Nothing
    If Conn3.State = 1 Then
        Conn3.Close
        Set Conn3 = Nothing
     End If
     
     Set rs2 = Nothing
     If Conn2.State = 1 Then
        Conn2.Close
        Set Conn2 = Nothing
     End If
      Set rs7 = Nothing
     If Conn7.State = 1 Then
        Conn7.Close
        Set Conn7 = Nothing
     End If
'        rs7.Close
'        txtmisce.SetFocus
    End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set RS4 = Nothing
If Conn4.State = 1 Then
    Conn4.Close
 End If

    Set rs5 = Nothing
If Conn5.State = 1 Then
    Conn5.Close
    Set Conn5 = Nothing
End If

  Set rs6 = Nothing
  If Conn6.State = 1 Then
    Conn6.Close
    Set Conn6 = Nothing
   End If
End Sub

Private Sub Text1_Change()
    If Not IsNumeric(Text1) Then
        Text1 = ""
 Else
            txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
     
End If
   txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))
   
 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtDisc))

End Sub

Private Sub Text1_GotFocus()
        Text1.BackColor = &H96E4B1
End Sub

Private Sub Text1_LostFocus()
       If Text1 = Empty Then
       Text1 = 0
    End If
 Text1.BackColor = &H80000018
 
End Sub


Private Sub Text2_Change()
 
End Sub

Private Sub txtBabyCareCharge_Change()
  If Not IsNumeric(txtBabyCareCharge) Then
            txtBabyCareCharge = ""
  Else
         txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
   End If
  txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))
   
 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtDisc))

End Sub

Private Sub txtBabyCareCharge_GotFocus()
  If Val(txtBabyCareCharge) = 0 Or Val(txtBabyCareCharge) > 0 Then
        txtBabyCareCharge.ForeColor = vbBlack
        txtBabyCareCharge.Locked = False
         txtBabyCareCharge.BackColor = &H96E4B1
  End If
End Sub

Private Sub txtBabyCareCharge_LostFocus()
   If txtBabyCareCharge = Empty Then
             txtBabyCareCharge = 0
   End If
   txtBabyCareCharge.BackColor = &H80000018
End Sub

Private Sub txtBloodTher_charge_Change()
If Not IsNumeric(txtBloodTher_charge) Then
       txtBloodTher_charge = ""
 Else
      txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
End If

 txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))
   
 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtDisc))

End Sub

Private Sub txtBloodTher_charge_GotFocus()
          txtBloodTher_charge.BackColor = &H96E4B1
End Sub

Private Sub txtBloodTher_charge_LostFocus()
    If txtBloodTher_charge = Empty Then
       txtBloodTher_charge = 0
    End If
    
     txtBloodTher_charge.BackColor = &H80000018
End Sub

Private Sub txtCCU_Charge_Change()
If Not IsNumeric(txtCCU_Charge) Then
       txtCCU_Charge = ""
 Else
        txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
 End If
 txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))
   
 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtDisc))

End Sub

Private Sub txtCCU_Charge_GotFocus()
txtCCU_Charge.BackColor = &H96E4B1
End Sub

Private Sub txtCCU_Charge_LostFocus()
    If txtCCU_Charge = Empty Then
       txtCCU_Charge = 0
    End If
    txtCCU_Charge.BackColor = &H80000018
End Sub

Private Sub txtDeliveryCharge_Change()
If Not IsNumeric(txtDeliveryCharge) Then
   txtDeliveryCharge = ""
Else
      txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
End If
 txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))
   
 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtDisc))

End Sub

Private Sub txtDeliveryCharge_GotFocus()
  txtDeliveryCharge.BackColor = &H96E4B1
End Sub

Private Sub txtDeliveryCharge_LostFocus()
If (txtDeliveryCharge) = Empty Then
      txtDeliveryCharge = 0
      txtDeliveryCharge.BackColor = &H80000018
End If
    txtDeliveryCharge.BackColor = &H80000018
End Sub

Private Sub txtDisc_Change()
If Not IsNumeric(txtDisc.Text) Then
    txtDisc = 0
End If
'If Val(txtDisc) > Val(txtmiscelleneous) And Val(txtmiscelleneous) > 0 Then
'     MsgBox "Discount can''t be greater than Total Amount....Please Check !!!", vbCritical, " IT, DNMIH."
'     txtDisc = 0
'     txtdisc_percent = 0
'End If
 
  If txtdisc_percent = 0 Or txtdisc_percent = "" Then
  
    txtNetTotal = (Val(txtmiscelleneous) - Val(txtDisc))
   
 End If
' If Val(txtDisc) > (Val(txtTotal) + Val(txtmisce) + Val(txtMedicine_charge)) Then
'    MsgBox "Discount can''t be greater than Total Amount.....Please Check !!!", vbCritical, " IT, DNMIH."
'     txtDisc = 0
'     txtdisc_percent = 0
'     Exit Sub
'End If

 

End Sub

Private Sub txtdisc_GotFocus()
txtdisc_percent = 0
txtDisc.BackColor = &H96E4B1
End Sub

Private Sub txtDisc_LostFocus()
        txtDisc.BackColor = &HF7E3B3
End Sub

Private Sub txtdisc_percent_Change()
If Not IsNumeric(txtdisc_percent.Text) Then
        txtdisc_percent = 0
Else
    If txtdisc_percent = 0 Or txtdisc_percent = "" Then
        txtDisc = 0
 End If
 
 If txtdisc_percent <> 0 And IsNull(txtdisc_percent) = False And txtmiscelleneous <= 0 Then '''IF AMOUNT IS MINUS
        txtDisc = Abs((((Val(txtTotal) + Val(txtmisce) + Val(txtMedicine_charge)) * Val(txtdisc_percent))) / 100)
ElseIf txtdisc_percent <> 0 And IsNull(txtdisc_percent) = False Then
txtDisc = Abs((((Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) * Val(txtdisc_percent))) / 100)
End If

txtNetTotal = (Val(txtmiscelleneous) - Val(txtDisc))
End If
End Sub

Private Sub txtdisc_percent_GotFocus()
txtDisc = 0
txtDisc.BackColor = &H96E4B1
End Sub

Private Sub txtdisc_percent_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(txtdisc_percent) > 0 Then
txtNetTotal.SetFocus
'cmdSAVE.SetFocus
End If
End If


End Sub

Private Sub txtdisc_percent_LostFocus()
txtdisc_percent.BackColor = &HF7E3B3
End Sub

Private Sub txtExtraBedTotal_LostFocus()
      txtExtraBedTotal.BackColor = &H80000018
End Sub

Private Sub txtExtTransfusion_Change()
If Not IsNumeric(txtExtTransfusion) Then
      txtExtTransfusion = ""
 Else
      txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
End If

 txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))
   
 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtDisc))

End Sub

Private Sub txtExtTransfusion_GotFocus()
        txtExtTransfusion.BackColor = &H96E4B1
End Sub

Private Sub txtExtTransfusion_LostFocus()
  If txtExtTransfusion = Empty Then
     txtExtTransfusion = 0
   End If
   
     txtExtTransfusion.BackColor = &H80000018
End Sub

Private Sub txtGrandTotal_Change()
txtD_tolal = Val(txtGrandTotal)
End Sub

Private Sub txtIncubator_Change()
  If Not IsNumeric(txtIncubator) Then
         txtIncubator = ""
   Else
 
         txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
  End If
 txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))
   
 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtDisc))

  
End Sub

Private Sub txtIncubator_LostFocus()
   If txtIncubator = "" Then
     txtIncubator = 0
   End If
End Sub

Private Sub txtMedicine_charge_Change()
If Not IsNumeric(txtMedicine_charge.Text) Then
    txtMedicine_charge = 0
Else
    txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + (txtMedicine_charge))
End If
'txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtAdvanceRelease) + Val(txtDisc))
 txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))

 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - Val(txtDisc)
   
End Sub

Private Sub txtMedicine_charge_GotFocus()
txtMedicine_charge.BackColor = &H96E4B1
End Sub

Private Sub txtMedicine_charge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtmiscelleneous.SetFocus
    End If
 
End Sub

Private Sub txtMedicine_charge_LostFocus()
txtMedicine_charge.BackColor = &HF7E3B3
End Sub

Private Sub txtmisce_Change()
If Not IsNumeric(txtmisce.Text) Then
    txtmisce = 0
Else
    txtmiscelleneous = Val(txtmisce) + Val(txtD_tolal) + Val(txtMedicine_charge)
    txtNetTotal = (Val(txtmiscelleneous)) - (Val(txtDisc))
End If
  txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))
   
 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtDisc))
       End Sub

Private Sub txtmisce_GotFocus()
txtmisce.BackColor = &H96E4B1
End Sub

Private Sub txtmisce_LostFocus()
txtmisce.BackColor = &HF7E3B3
End Sub

Private Sub txtmiscelleneous_Change()
    txtdisc_percent = 0
    txtDisc = 0
    txtNetTotal = (Val(txtmiscelleneous) - Val(txtDisc))
End Sub

Private Sub txtNebuliser_Change()
  If Not IsNumeric(txtNebuliser) Then
          txtNebuliser = ""
  Else
    txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
   End If
   txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))
   
 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtDisc))
 
End Sub

Private Sub txtNebuliser_LostFocus()
  If txtNebuliser = "" Then
    txtNebuliser = 0
  End If
End Sub

Private Sub txtNeunetalCharge_Change()

If Not IsNumeric(txtNeunetalCharge) Then
      txtNeunetalCharge = ""
 Else
      txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
End If

 txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))
   
 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtDisc))


End Sub

Private Sub txtNeunetalCharge_GotFocus()
      txtNeunetalCharge.BackColor = &H96E4B1
End Sub

Private Sub txtNeunetalCharge_LostFocus()
 If txtNeunetalCharge = Empty Then
    txtNeunetalCharge = 0
 End If
   txtNeunetalCharge.BackColor = &H80000018
End Sub

Private Sub txtP_therapyCharge_Change()
If Not IsNumeric(txtP_therapyCharge) Then
      txtP_therapyCharge = ""
 Else
      txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
End If

 txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))
   
 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtDisc))

End Sub

Private Sub txtP_therapyCharge_GotFocus()
        txtP_therapyCharge.BackColor = &H96E4B1
End Sub

Private Sub txtP_therapyCharge_LostFocus()
If txtP_therapyCharge = Empty Then
   txtP_therapyCharge = 0
End If
     txtP_therapyCharge.BackColor = &H80000018
End Sub



Private Sub txtTotal_Change()
txtGrandTotal = Val(txtTotal) - Val(txtAdvanceRelease)
End Sub

Private Sub txtTotalOpr_Change()
 If Not IsNumeric(txtTotalOpr) Then
        txtTotalOpr = ""
 Else
         txtTotal = Val(txtNebuliser) + Val(Text1) + Val(txtCCU_Charge) + Val(txtDeliveryCharge) + Val(txtTotalBedCharge) + Val(txtWithAdmissionCharge) + Val(txtServiceCharge) + Val(txtExtraBedTotal) + Val(txtTotalOpr) + Val(txtBabyCareCharge) + Val(txtNeunetalCharge) + Val(txtExtTransfusion) + Val(txtP_therapyCharge) + Val(txtBloodTher_charge) + Val(txtIncubator)
     
End If
   txtmiscelleneous = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge))
   
 txtNetTotal = (Val(txtD_tolal) + Val(txtmisce) + Val(txtMedicine_charge)) - (Val(txtDisc))


End Sub

Private Sub txtTotalOpr_GotFocus()
       txtTotalOpr.BackColor = &H96E4B1
End Sub

Private Sub txtTotalOpr_LostFocus()
   If txtTotalOpr = Empty Then
        txtTotalOpr = 0
    End If
 txtTotalOpr.BackColor = &H80000018
 
End Sub
