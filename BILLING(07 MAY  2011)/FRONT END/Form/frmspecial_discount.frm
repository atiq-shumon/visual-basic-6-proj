VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmspcial_discount 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6945
   ClientLeft      =   -105
   ClientTop       =   390
   ClientWidth     =   9315
   FillColor       =   &H007DABD0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   30
      TabIndex        =   27
      Top             =   3300
      Width           =   9195
      Begin VB.TextBox txtCurpayment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3690
         TabIndex        =   0
         Top             =   540
         Width           =   1875
      End
      Begin VB.TextBox txtTotalPayment 
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
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   495
         Width           =   1875
      End
      Begin VB.TextBox TxtPreviousPayment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc5"
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
         Height          =   330
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label Label13 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2790
         TabIndex        =   34
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6630
         TabIndex        =   33
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payment"
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
         Left            =   7020
         TabIndex        =   32
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Payment"
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
         Left            =   3690
         TabIndex        =   31
         Top             =   240
         Width           =   1710
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Amount"
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
         TabIndex        =   30
         Top             =   240
         Width           =   1755
      End
   End
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
      Left            =   90
      Picture         =   "frmspecial_discount.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save"
      Top             =   6450
      Width           =   495
   End
   Begin VB.CommandButton cmdPreview 
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
      Left            =   1110
      Picture         =   "frmspecial_discount.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Preview"
      Top             =   6450
      Width           =   510
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
      Left            =   1620
      Picture         =   "frmspecial_discount.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   6450
      Width           =   510
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
      Left            =   600
      Picture         =   "frmspecial_discount.frx":15F2
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Print"
      Top             =   6450
      Width           =   510
   End
   Begin VB.Frame Frame3 
      Height          =   1965
      Left            =   0
      TabIndex        =   25
      Top             =   4380
      Width           =   9225
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmspecial_discount.frx":1A34
         Height          =   1335
         Left            =   150
         TabIndex        =   26
         Top             =   315
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   12632256
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
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   23
      Top             =   60
      Width           =   9225
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Special discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   480
         Left            =   2310
         TabIndex        =   24
         Top             =   90
         Width           =   3285
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   0
      TabIndex        =   14
      Top             =   690
      Width           =   9225
      Begin VB.TextBox txtAgeInTest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4875
         Locked          =   -1  'True
         MaxLength       =   17
         TabIndex        =   9
         Top             =   1275
         Width           =   555
      End
      Begin VB.TextBox txtAddrInTest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1920
         Width           =   8625
      End
      Begin VB.TextBox txtPat_ID1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc6"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   330
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   510
         Width           =   4110
      End
      Begin VB.TextBox txtNameInTest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   330
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1260
         Width           =   4125
      End
      Begin VB.ComboBox cboInTestDept 
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4875
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "cboInTestDept"
         Top             =   510
         Width           =   2040
      End
      Begin VB.ComboBox cboInTestSex 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmspecial_discount.frx":1A49
         Left            =   5520
         List            =   "frmspecial_discount.frx":1A53
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1275
         Width           =   1365
      End
      Begin VB.ComboBox cboInTestReligion 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmspecial_discount.frx":1A65
         Left            =   7020
         List            =   "frmspecial_discount.frx":1A78
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1260
         Width           =   1920
      End
      Begin MSComCtl2.DTPicker Dt_date 
         Height          =   330
         Left            =   6990
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   510
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   49152
         Format          =   60882945
         CurrentDate     =   37114
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
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   4875
         TabIndex        =   22
         Top             =   1005
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   6990
         TabIndex        =   21
         Top             =   240
         Width           =   510
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
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   0
         Left            =   330
         TabIndex        =   20
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         TabIndex        =   19
         Top             =   1605
         Width           =   885
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
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   5520
         TabIndex        =   18
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Name"
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
         Left            =   285
         TabIndex        =   17
         Top             =   1005
         Width           =   690
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
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   4875
         TabIndex        =   16
         Top             =   240
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
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   7020
         TabIndex        =   15
         Top             =   1005
         Width           =   885
      End
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
      Left            =   3420
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
      Left            =   4770
      Top             =   6510
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
      Height          =   525
      Left            =   30
      Top             =   6390
      Width           =   2175
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
End
Attribute VB_Name = "frmspcial_discount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As New MyConnection
Dim conn As New Connection
Dim cmd As New Command
Dim rs As New Recordset
Dim RS1 As New Recordset
Dim Conn2 As New Connection
Dim rs2 As New Recordset
Public strUid As String
Dim VoucherNumber
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
Private Sub CMDEXIT_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdPrint_Click()
'Viewer.Show vbModal
'
'rptMode = 4



  conn.Open strcn.Connection_String
            Set cmd.ActiveConnection = conn
            cmd.CommandType = adCmdText
            
            
                
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
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL Rptout_door_info_print}"
            Set rs = cmd.Execute
            cmd.Properties("PLSQLRSet") = False
            
            Report4.Database.SetDataSource rs

            Report4.PrintOut
            rs.Close
'    '====================================


End Sub


Private Sub cmdSave_Click()
Dim validation As Variant
              Adodc1.ConnectionString = strcn.Connection_String
              Adodc1.RecordSource = "Select user_id, user_name, user_password,user_type,shift_name From Security Where (User_Id = '" & frmMAIN.lbluser_id & "')"
              Adodc1.Refresh
              validation = Adodc1.Recordset!user_id
                
                Dim conn As New ADODB.Connection
                Dim cmd As New ADODB.Command
                Dim rs As New ADODB.Recordset

                        Dim Param1 As New Parameter
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
                
          Adodc2.ConnectionString = strcn.Connection_String
          Adodc2.RecordSource = "Select * From user_validation"
          Adodc2.Refresh
        

        
             If Adodc2.Recordset!validation = 0 Then
             MsgBox "Your Working Time has been Expired", vbInformation, " IT, DNMIH."
             Exit Sub
             End If
             


If Val(txtCurpayment) = Empty Then
MsgBox "Nothing To Save", vbInformation, " IT, DNMIH."
Exit Sub
End If
'Get_Voucher_Number
Call save_readvance
'Call acct_integration_for_readvance
'Call acct_integration_for_readvance1
'Call post_vou
MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
Call flush_grid
total_adv
txtCurpayment = ""
cmdExit.SetFocus
print_readvance
'rptMode = 9
'Viewer.Show vbModal
If conn.State = 1 Then
    conn.Close
End If
'RS.Close
End Sub
Private Sub post_vou()
 
     Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    
 
 If conn.State = 0 Then
    conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
      Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, VoucherNumber)
    cmd.Parameters.Append Param1 'u_id
       Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 100, "CR")
    cmd.Parameters.Append Param2 'comment
       
            
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL acct.postvou(?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set rs = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    If conn.State = 1 Then
       conn.Close
       Set conn = Nothing
    End If
  Set rs = Nothing
    
End Sub
Private Sub acct_integration_for_readvance1()
   Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    
 If conn.State = 0 Then
    conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
     Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, "Re advance Collection in cash")
    cmd.Parameters.Append Param1 'comment
                                                                    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, "6101019")
    cmd.Parameters.Append Param2 ''USER_acct
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, 0)
        cmd.Parameters.Append Param3 'dr
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, txtCurpayment.Text)
        cmd.Parameters.Append Param4 'cr
   
     Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 100, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param5 'u_id
     Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 100, VoucherNumber)
    cmd.Parameters.Append Param6 'u_id
         
            
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL acct.Save_vou_bill(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set rs = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    If conn.State = 1 Then
       conn.Close
       Set conn = Nothing
    End If
  Set rs = Nothing
    
End Sub

Private Sub acct_integration_for_readvance()
   Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
    Dim Param6 As New Parameter
    Dim Param7 As New Parameter
    Dim Param8 As New Parameter
    Dim Param9 As New Parameter
    
 If conn.State = 0 Then
    conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
     Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 100, "Re advance Collection in cash")
    cmd.Parameters.Append Param1 'comment
                                                                    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 200, "2101")
    cmd.Parameters.Append Param2 ''USER_acct
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, txtCurpayment.Text)
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
    
    Set rs = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    If conn.State = 1 Then
       conn.Close
       Set conn = Nothing
    End If
  Set rs = Nothing
    
End Sub

Private Sub print_readvance()
            'Public Con As New MyConnection
            Dim Connre As New Connection
            Dim cmdre As New Command
            Dim RSre As New Recordset
          If Connre.State = 0 Then
            Connre.Open strcn.Connection_String
         End If
            Set cmdre.ActiveConnection = Connre
            cmdre.CommandType = adCmdText
            
            
            Dim Report8   As New CrystalReport8
            cmdre.Properties("PLSQLRSet") = True
            cmdre.CommandText = "{CALL Rpt_in_dr_info_adm_print}"
            Set RSre = cmdre.Execute
            cmdre.Properties("PLSQLRSet") = False
            
            Report8.Database.SetDataSource RSre

            Report8.PrintOut
            RSre.Close
           If Connre.State = 1 Then
            Connre.Close
          End If
End Sub
 Private Sub save_readvance()
 Dim conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
       
If conn.State = 0 Then
    conn.Open strcn.Connection_String
End If
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 30, frmReadvancepayment.txtReg_noInTest)
    cmd.Parameters.Append Param1 'IN_REG_NO
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 30, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param2 'U_id default Sumon
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 30, frmMAIN.lblBooth)
    cmd.Parameters.Append Param3 'booth
    
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, txtCurpayment.Text)
    cmd.Parameters.Append Param4 'readvance
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_advance(?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set rs = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    If conn.State = 1 Then
       conn.Close
    End If
    
End Sub
Private Sub DataGrid1_Click()

If DataGrid1.Row > 0 Then
End If


End Sub

Private Sub Form_Activate()
     txtCurpayment.SetFocus
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
    Adodc4.RecordSource = "select advance,dt from advance where in_reg_no='" & Trim(frmReadvancepayment.txtReg_noInTest) & "'"
    Adodc4.Refresh
    
    DataGrid1.Columns(0).Width = 615
    DataGrid1.Columns(1).Width = 1845
    
End Sub

Private Sub Form_Load()
 
        
                       Dim temp
      If Conn2.State = 0 Then
        Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
      End If
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select pat_name,pat_guard_name,sex,age,religion,addr1,addr2,phone,doc_dept,admission_date  From in_door_pat_info_main Where in_reg_no ='" & Trim(frmReadvancepayment.txtReg_noInTest.Text) & "'"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        rs2.CursorLocation = adUseClient

        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
       
                  
       If rs2.RecordCount > 0 Then
       
               
         txtNameInTest = rs2!pat_name
         txtAddrInTest = rs2!addr1
       
         txtAgeInTest = rs2!age
         Dt_date.Value = rs2!admission_date
         cboInTestSex.Text = rs2!sex
         cboInTestReligion = rs2!religion
         cboInTestDept = rs2!doc_dept
         
                                                                                                                                                                                                           
        
'           comMainCode.SetFocus
       End If
       cmd.Properties("iRowsetChange") = False

       txtPat_ID1 = frmReadvancepayment.txtReg_noInTest
       

    total_adv
    
    Call flush_grid



   '' rptMode = 4
 If Conn2.State = 1 Then
    Conn2.Close
 End If
  
'   rs2.Close
   End Sub
Private Sub total_adv()

                 Adodc5.ConnectionString = strcn.Connection_String
                    Adodc5.RecordSource = "select  nvl(sum(advance),0)as advance from advance where in_reg_no ='" & Trim(frmReadvancepayment.txtReg_noInTest.Text) & "'"
                    Adodc5.Refresh
                If Adodc5.Recordset.RecordCount > 0 Then
                     TxtPreviousPayment = Adodc5.Recordset!advance

                End If
       Adodc5.Refresh

End Sub


Private Sub txtCurpayment_Change()
If Not IsNumeric(txtCurpayment) Then
     txtCurpayment = ""
Else
   If Val(txtCurpayment) < 0 Then
      MsgBox "Advance Can't be minus Fiegure", vbInformation, " IT, DNMIH."
Else
     txtTotalPayment = Val(txtCurpayment) + Val(TxtPreviousPayment)
  End If
End If

End Sub
