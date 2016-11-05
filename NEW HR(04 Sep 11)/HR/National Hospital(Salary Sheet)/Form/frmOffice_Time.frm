VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form17 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form17"
   ClientHeight    =   5160
   ClientLeft      =   1755
   ClientTop       =   2010
   ClientWidth     =   8160
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form17"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOffice_Time.frx":0000
   ScaleHeight     =   5160
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Special time"
      ForeColor       =   &H00800000&
      Height          =   1005
      Index           =   1
      Left            =   405
      TabIndex        =   13
      Top             =   1935
      Width           =   7305
      Begin VB.ComboBox cmbSp_End_time 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "frmOffice_Time.frx":152CE
         Left            =   5760
         List            =   "frmOffice_Time.frx":15323
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   495
         Width           =   1320
      End
      Begin VB.ComboBox cmbSp_Start_Day 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "frmOffice_Time.frx":15432
         Left            =   180
         List            =   "frmOffice_Time.frx":1544E
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   495
         Width           =   1320
      End
      Begin VB.ComboBox cmbSp_Start_Time 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "frmOffice_Time.frx":1549E
         Left            =   2070
         List            =   "frmOffice_Time.frx":154F3
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   495
         Width           =   1320
      End
      Begin VB.ComboBox cmbSp_End_Day 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "frmOffice_Time.frx":155FF
         Left            =   3960
         List            =   "frmOffice_Time.frx":1561B
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   495
         Width           =   1320
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   5
         Left            =   2025
         Top             =   450
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   6
         Left            =   135
         Top             =   450
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   7
         Left            =   3915
         Top             =   450
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   8
         Left            =   5715
         Top             =   450
         Width           =   1410
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5805
         TabIndex        =   20
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   3960
         TabIndex        =   19
         Top             =   225
         Width           =   285
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Start time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2070
         TabIndex        =   18
         Top             =   225
         Width           =   675
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   180
         TabIndex        =   17
         Top             =   225
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usual time"
      ForeColor       =   &H00800000&
      Height          =   960
      Index           =   0
      Left            =   405
      TabIndex        =   12
      Top             =   855
      Width           =   7305
      Begin VB.ComboBox cmbStart_time 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "frmOffice_Time.frx":1566E
         Left            =   180
         List            =   "frmOffice_Time.frx":156C3
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   450
         Width           =   1320
      End
      Begin VB.ComboBox cmbEnd_time 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "frmOffice_Time.frx":157D0
         Left            =   2070
         List            =   "frmOffice_Time.frx":15825
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   450
         Width           =   1320
      End
      Begin VB.ComboBox cmbAbs_Time 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         ItemData        =   "frmOffice_Time.frx":15931
         Left            =   3960
         List            =   "frmOffice_Time.frx":15986
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   450
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtpEffect_date 
         Height          =   330
         Left            =   5760
         TabIndex        =   26
         Top             =   450
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyyy"
         Format          =   58982403
         CurrentDate     =   37316
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   0
         Left            =   135
         Top             =   405
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   1
         Left            =   2025
         Top             =   405
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   3
         Left            =   3915
         Top             =   405
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   4
         Left            =   5715
         Top             =   405
         Width           =   1410
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2070
         TabIndex        =   24
         Top             =   180
         Width           =   600
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Start time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   180
         TabIndex        =   23
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Effect date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   5760
         TabIndex        =   22
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Absent After"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4005
         TabIndex        =   21
         Top             =   180
         Width           =   945
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1230
      Left            =   450
      TabIndex        =   6
      Top             =   3105
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2170
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BorderStyle     =   0
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   "Start time"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   "End time"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   ""
         Caption         =   "Absent after"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "h:nn AM/PM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   ""
         Caption         =   "Effect date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   ""
         Caption         =   "Relaxed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "M/d/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   ""
         Caption         =   "Sp.Start time"
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
      BeginProperty Column06 
         DataField       =   ""
         Caption         =   "Day"
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
      BeginProperty Column07 
         DataField       =   ""
         Caption         =   "Sp. End time"
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
      BeginProperty Column08 
         DataField       =   ""
         Caption         =   "Day"
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
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   929.764
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtRelax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1440
      TabIndex        =   9
      Top             =   4635
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4410
      Picture         =   "frmOffice_Time.frx":15A92
      ScaleHeight     =   330
      ScaleMode       =   0  'User
      ScaleWidth      =   3264.834
      TabIndex        =   7
      Top             =   4590
      Width           =   3270
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2445
         Picture         =   "frmOffice_Time.frx":15DB9
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   870
      End
      Begin VB.CommandButton cmdClear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1650
         Picture         =   "frmOffice_Time.frx":1693B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   870
      End
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   810
         Picture         =   "frmOffice_Time.frx":174BD
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   0
         Width           =   870
      End
      Begin VB.CommandButton cmdDel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   -30
         Picture         =   "frmOffice_Time.frx":1803F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton cmdEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   810
         Picture         =   "frmOffice_Time.frx":18BC1
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   870
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3240
      Top             =   4680
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3690
      Top             =   180
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
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Min."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1980
      TabIndex        =   11
      Top             =   4635
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FECCC7&
      BorderColor     =   &H00FECCC7&
      Height          =   330
      Index           =   2
      Left            =   1395
      Top             =   4590
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FECCC7&
      BorderColor     =   &H00FECCC7&
      Height          =   1320
      Index           =   11
      Left            =   405
      Top             =   3060
      Width           =   7305
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Relaxed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   720
      TabIndex        =   10
      Top             =   4635
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   6615
      TabIndex        =   5
      Top             =   450
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Office Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   855
      TabIndex        =   4
      Top             =   90
      Width           =   1500
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ref_Key As String

Private Sub cmdClear_Click()
    On Error Resume Next
    cmbStart_Time = ""
    cmbEnd_Time = ""
    txtRelax = ""
    cmbAbs_time = ""
    Ref_Key = ""
    
    cmbSp_Start_Day.ListIndex = 7
    cmbSp_End_Day.ListIndex = 7
    
    cmbSp_Start_Time.ListIndex = 26
    cmbEnd_Time.ListIndex = 26
    cmbSp_End_time.ListIndex = 26
    cmbStart_Time.ListIndex = 26
    cmbAbs_time.ListIndex = 26
    
    dtpEffect_date = Now
    
End Sub

Private Sub cmdClose_Click()
yes_no = MsgBox("Do you really want to close it?", vbYesNo + vbQuestion)
    If yes_no = vbYes Then
        Unload Me
    Else
    
        Exit Sub
    End If
End Sub

Private Sub cmdDel_Click()
    opr = "D"
    cmdSave_Click
End Sub

Private Sub cmdEdit_Click()
cmdSave_Click
End Sub

Private Sub cmdSave_Click()

    If cmbStart_Time = "" Then
    MsgBox "Starting time required", vbCritical + vbOKOnly, "Error Message"
    cmbStart_Time.SetFocus
    Exit Sub
    End If
    
    If cmbEnd_Time = "" Then
    MsgBox "Ending time required", vbCritical + vbOKOnly, "Error Message"
    cmbEnd_Time.SetFocus
    Exit Sub
    End If
    
    con.ConnectionString = strCN.Connection
    con.Open
    cmd.CommandText = "exec Office_Time_I_U_D '" + opr + "','" _
    + ChkForQuote(Trim(cmbStart_Time)) + " ',' " _
    + ChkForQuote(Trim(cmbEnd_Time)) + "','" _
    + txtRelax + "','" _
    + ChkForQuote(Trim(cmbAbs_time)) + " ',' " _
    + cmbSp_Start_Time + "','" _
    + cmbSp_Start_Day + "','" _
    + cmbSp_End_time + "','" _
    + cmbSp_End_Day + "','" _
    + Format(dtpEffect_date, "yyyy-mm-dd") + "','" _
    + U_Id + "','" _
    + CStr(Ref_Key) + "'"
        
    cmd.ActiveConnection = con
    cmd.Execute
    con.Close
    populate_grd
    '--------------------------
    Grid_Click (False), Form17
    ''-------------------------
    Label5.Visible = False      ''Relaxed time
    Label6.Visible = False
    Shape1(2).Visible = False
    txtRelax.Visible = False
    ''-------------------------
 
 

    
 
 
 
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

If Button = 2 Then Grid_Click (False), Me

If Button = 1 Then
    Grid_Click (True), Me
    cmbStart_Time = Trim(Adodc1.Recordset!Start_Time)
    cmbEnd_Time = Trim(Adodc1.Recordset!End_Time)
    txtRelax = Trim(Adodc1.Recordset!Relaxed)
    cmbAbs_time = Trim(Adodc1.Recordset!Abs_time)
    Ref_Key = Trim(Adodc1.Recordset!Ref_Key)
    cmbSp_Start_Day = Trim(Adodc1.Recordset!Sp_Start_Day)
    cmbSp_Start_Time = Trim(Adodc1.Recordset!Sp_Start_time)
    cmbSp_End_Day = Trim(Adodc1.Recordset!Sp_End_Day)
    cmbSp_End_time = Trim(Adodc1.Recordset!Sp_End_time)
    dtpEffect_date = Trim(Adodc1.Recordset!Effect_date)
    
    Grid_Click (True), Form17
    cmbStart_Time.SetFocus

End If

End Sub

Private Sub Form_Load()
On Error Resume Next
Grid_Click (False), Form17
populate_grd
dtpEffect_date.Value = Format(Now, "dd/MM/yyyy")
Timer1.Enabled = True
'cmbStart_time.Text = "09:00 AM"
'cmbEnd_time.Text = "06:30 PM"
'cmbAbs_Time.Text = "02:00 PM"
Ref_Key = 0
dtpEffect_date = Now
End Sub

Public Sub populate_grd()
Adodc1.ConnectionString = strCN.Connection
    Adodc1.RecordSource = "SP_Office_time_All 'All'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        'Adodc1.Recordset.MoveLast
        DataGrid1.Columns(0).DataField = "Start_time"
        DataGrid1.Columns(1).DataField = "End_time"
        DataGrid1.Columns(2).DataField = "Abs_time"
        DataGrid1.Columns(3).DataField = "Effect_date"
        DataGrid1.Columns(4).DataField = "Relaxed"
        DataGrid1.Columns(5).DataField = "Sp_Start_time"
        DataGrid1.Columns(6).DataField = "Sp_Start_Day"
        DataGrid1.Columns(7).DataField = "Sp_End_time"
        DataGrid1.Columns(8).DataField = "Sp_End_Day"
          
        Set DataGrid1.DataSource = Adodc1
        
        'cmbAbs_Time = Adodc1.Recordset!Abs_time
        
        txtRelax.Refresh
        DataGrid1.ReBind
        DataGrid1.Refresh
    End If
End Sub

Private Sub Label1_dblClick()
    Label5.Visible = True
    Label6.Visible = True
    Shape1(2).Visible = True
    txtRelax.Visible = True
    txtRelax.SelStart = 0
    txtRelax.SelLength = Len(txtRelax)
    txtRelax.SetFocus
End Sub

Private Sub Timer1_Timer()
lblDate = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub txtRelax_Change()
 If Val(txtRelax) > 30 Then
 txtRelax = "0"
 End If
End Sub

Private Sub txtRelax_LostFocus()
    Label5.Visible = False
    Label6.Visible = False
    Shape1(2).Visible = False
    txtRelax.Visible = False
End Sub
