VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_reAdvance 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7260
   ClientLeft      =   -105
   ClientTop       =   390
   ClientWidth     =   8700
   FillColor       =   &H007DABD0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   37
      Top             =   6930
      Width           =   13365
      Begin VB.Label Label15 
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
         TabIndex        =   39
         Top             =   60
         Width           =   4725
      End
      Begin VB.Label Label14 
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
         TabIndex        =   38
         ToolTipText     =   "Developed && Maintenanced by:A.N.M. Atiqur Rahman,Software Programmer, IT Division, DNMIH"
         Top             =   90
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   7140
      TabIndex        =   2
      Top             =   6450
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   5910
      TabIndex        =   34
      Top             =   6450
      Width           =   1215
   End
   Begin VB.CommandButton CMDSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   6450
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc8 
      Height          =   330
      Left            =   2790
      Top             =   6180
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
   Begin VB.TextBox TXTREC_NO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5730
      TabIndex        =   32
      Top             =   6120
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000001&
      Height          =   1095
      Left            =   -30
      TabIndex        =   24
      Top             =   3270
      Width           =   8835
      Begin VB.TextBox txtCurpayment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   3450
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
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   525
         Width           =   2355
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   31
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000006&
         Height          =   345
         Left            =   5640
         TabIndex        =   30
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL ADV. PAYMENT"
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
         Height          =   240
         Left            =   6090
         TabIndex        =   29
         Top             =   270
         Width           =   2490
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CURR. ADVANCE"
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
         Height          =   240
         Left            =   3450
         TabIndex        =   28
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREV. TOT. ADV"
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
         Height          =   240
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   1800
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000001&
      Height          =   1965
      Left            =   -30
      TabIndex        =   22
      Top             =   4260
      Width           =   8835
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_reAdvance.frx":0000
         Height          =   1665
         Left            =   150
         TabIndex        =   23
         Top             =   225
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2937
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
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
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   21
      Top             =   -60
      Width           =   8835
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RE-ADVANCE PAYMANET ENTRY"
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
         Left            =   1800
         TabIndex        =   33
         Top             =   270
         Width           =   5175
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -60
         Picture         =   "frm_reAdvance.frx":0015
         Stretch         =   -1  'True
         Top             =   30
         Width           =   9780
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   -60
      TabIndex        =   12
      Top             =   690
      Width           =   8895
      Begin VB.TextBox txtCurrentBed 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   510
         Width           =   2280
      End
      Begin VB.TextBox txtAgeInTest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4635
         Locked          =   -1  'True
         MaxLength       =   17
         TabIndex        =   7
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
         TabIndex        =   10
         Top             =   2040
         Width           =   8175
      End
      Begin VB.TextBox txtPat_ID1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc6"
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
         Height          =   330
         Left            =   330
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   510
         Width           =   1470
      End
      Begin VB.TextBox txtNameInTest 
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1260
         Width           =   4125
      End
      Begin VB.ComboBox cboInTestDept 
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4635
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "cboInTestDept"
         Top             =   510
         Width           =   2040
      End
      Begin VB.ComboBox cboInTestSex 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frm_reAdvance.frx":5997
         Left            =   5280
         List            =   "frm_reAdvance.frx":59A1
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1275
         Width           =   1365
      End
      Begin VB.ComboBox cboInTestReligion 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frm_reAdvance.frx":59AB
         Left            =   6855
         List            =   "frm_reAdvance.frx":59BE
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1260
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker Dt_date 
         Height          =   330
         Left            =   6810
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   510
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   49152
         Format          =   58916865
         CurrentDate     =   37114
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   2160
         TabIndex        =   36
         Top             =   240
         Width           =   1230
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
         Left            =   4635
         TabIndex        =   20
         Top             =   1005
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Date"
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
         Left            =   6870
         TabIndex        =   19
         Top             =   240
         Width           =   1650
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
         TabIndex        =   18
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
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   360
         TabIndex        =   17
         Top             =   1725
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
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   5310
         TabIndex        =   16
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
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   285
         TabIndex        =   15
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
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   4635
         TabIndex        =   14
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
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   6870
         TabIndex        =   13
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
      Left            =   2790
      Top             =   6180
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
      Left            =   2790
      Top             =   6120
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
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   4620
      Top             =   6390
      Width           =   3795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   -6660
      TabIndex        =   11
      Top             =   4770
      Width           =   270
   End
End
Attribute VB_Name = "frm_reAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VAR_CUR_BED_SERIAL_NO As Integer
Dim UTILITY As New clsUtility
Public strUid As String
Dim VoucherNumber
Public strcn        As New MyConnection
Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub
Private Sub cmdPrint_Click()
If TXTREC_NO.Visible = False Then
       
       TXTREC_NO.Visible = True
   End If
   
  TXTREC_NO.ForeColor = vbBlue
  
   If TXTREC_NO = "" Then
    TXTREC_NO.SetFocus
     Exit Sub
   Else
       
      rptMode = 9
      Viewer.Show vbModal
      TXTREC_NO = ""
      TXTREC_NO.Visible = False
End If
End Sub

Private Sub cmdSave_Click()
If UTILITY.User_Shift_validation(frmMAIN.lbluser_id, frmMAIN.lblUserType.Caption) = False Then
    MsgBox "Mr. " & frmMAIN.lblName & "  Your Shift has been Expired.. " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
    Exit Sub
End If


If Val(txtCurpayment) = Empty Then
MsgBox "Nothing To Save", vbInformation, " IT, DNMIH."
Exit Sub
End If

Call save_readvance
MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
Call flush_grid
total_adv
txtCurpayment = ""
CMDEXIT.SetFocus
print_readvance
 
End Sub
Private Sub print_readvance()
            Dim Connre As New Connection
            Dim cmdre As New Command
            Dim RSre As New Recordset
          If Connre.State = 0 Then
            Connre.Open strcn.Connection_String
         End If
            Set cmdre.ActiveConnection = Connre
            cmdre.CommandType = adCmdText
            
            
            Dim Report8   As New CrystalReporttran
            Dim Param1 As New Parameter
            
            Set Param1 = cmdre.CreateParameter("param1", adDouble, adParamInput, 30, frm_reAdvance.TXTREC_NO)
            cmdre.Parameters.Append Param1 'IN_REG_NO

            cmdre.Properties("PLSQLRSet") = True
            cmdre.CommandText = "{CALL Rpt_in_dr_info_adm_print(?)}"
            Set RSre = cmdre.Execute
            cmdre.Properties("PLSQLRSet") = False
            
            Report8.Text4.Width = 1650
            Report8.Text4.SetText ("Re-Advance")
            
            Report8.Text2.SetText ("Re-Advance")
            Report8.Text2.Font.Bold = True
            
            Report8.Database.SetDataSource RSre

            Report8.PrintOut
            RSre.Close
           If Connre.State = 1 Then
              Connre.Close
              Set Connre = Nothing
              Set Report8 = Nothing
              Set RSre = Nothing
              Set cmdre = Nothing
          End If
End Sub
 Private Sub save_readvance()
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
    
    Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 30, frmReadvancepayment.txtReg_noInTest)
    cmd.Parameters.Append Param1 'IN_REG_NO
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 30, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param2 'U_id default Sumon
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 30, frmMAIN.lblBooth)
    cmd.Parameters.Append Param3 'booth
    
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 10, txtCurpayment.Text)
    cmd.Parameters.Append Param4 'readvance
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 10, Trim(frmReadvancepayment.CBOYRCODE))
    cmd.Parameters.Append Param5 'readvance
    
    Set Param6 = cmd.CreateParameter("param6", adVarChar, adParamInput, 20, cboInTestDept.Text)
    cmd.Parameters.Append Param6 'DOC DEPT
    
    Set Param7 = cmd.CreateParameter("param7", adInteger, adParamInput, 10, Val(VAR_CUR_BED_SERIAL_NO))
    cmd.Parameters.Append Param7 'CURRENT BED SERIAL
    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_advance(?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
       Conn.Close
    End If
    Adodc8.ConnectionString = strcn.Connection_String
    Adodc8.RecordSource = "SELECT MAX(REC_NO) AS REC_NO FROM RECEIPT_NO_COUNTER"
    Adodc8.Refresh
    If Adodc8.Recordset.RecordCount > 0 Then
        TXTREC_NO.Text = Adodc8.Recordset!REC_NO
    End If
End Sub
Private Sub Form_Activate()
     TXTREC_NO.Visible = False
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
    Adodc4.RecordSource = "select advance,dt from advance where in_reg_no='" & Trim(frmReadvancepayment.txtReg_noInTest) & "' AND YRCODE='" & Trim(frmReadvancepayment.CBOYRCODE) & "'"
    Adodc4.Refresh
    
    DataGrid1.Columns(0).Width = 1000
    DataGrid1.Columns(1).Width = 2000

End Sub
Private Sub Form_Load()
      Dim Conn As New ADODB.Connection
      Dim cmd As New ADODB.Command
      Dim RS As New ADODB.Recordset

      If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
      End If
      cmd.ActiveConnection = Conn
      cmd.CommandType = adCmdText
      cmd.CommandText = "select pat_name,pat_guard_name,sex,age,religion,addr1,phone,(SELECT DOC_DEPT FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(frmReadvancepayment.txtReg_noInTest.Text) & "' AND YRCODE='" & Trim(frmReadvancepayment.CBOYRCODE.Text) & "' AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM INDOOR_PAT_DEPT_INFO WHERE in_reg_no ='" & Trim(frmReadvancepayment.txtReg_noInTest.Text) & "' AND YRCODE='" & Trim(frmReadvancepayment.CBOYRCODE.Text) & "' )) doc_dept,admission_date  From in_door_pat_info_main Where in_reg_no ='" & Trim(frmReadvancepayment.txtReg_noInTest.Text) & "' AND YRCODE='" & Trim(frmReadvancepayment.CBOYRCODE.Text) & "'"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
       
                  
       If RS.RecordCount > 0 Then
          
         txtNameInTest = RS!pat_name
         txtAddrInTest = RS!addr1
       
         txtAgeInTest = RS!age
         Dt_date.Value = RS!admission_date
         cboInTestSex.Text = RS!sex
         cboInTestReligion = RS!religion
         cboInTestDept = "" & RS!doc_dept
         
       End If
       cmd.Properties("iRowsetChange") = False

       txtPat_ID1 = frmReadvancepayment.txtReg_noInTest
       

     Call total_adv
    
    Call flush_grid
    Call load_current_Bed_no

 If Conn.State = 1 Then
    Conn.Close
    Set Conn = Nothing
    Set RS = Nothing
    Set cmd = Nothing
 End If
  
 End Sub
Private Sub load_current_Bed_no()
  Dim Conn As New ADODB.Connection
  Dim cmd As New ADODB.Command
  Dim RS As New ADODB.Recordset
  
  If Conn.State = 0 Then
        Conn.ConnectionString = strcn.Connection_String
        Conn.Open
    End If
        cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
        cmd.CommandText = "select BED_TYPE,Bed_type_no,BED_NO,extra_bed_flag,SERIAL_NO  From Indoor_pat_bed_info Where in_reg_no ='" & Trim(frmReadvancepayment.txtReg_noInTest) & "' AND YRCODE ='" & Trim(frmReadvancepayment.CBOYRCODE.Text) & "'  AND SERIAL_NO=(SELECT MAX(SERIAL_NO) FROM Indoor_pat_bed_info WHERE in_reg_no ='" & Trim(frmReadvancepayment.txtReg_noInTest.Text) & "' AND YRCODE='" & Trim(frmReadvancepayment.CBOYRCODE) & "')"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        RS.CursorLocation = adUseClient

        RS.Open cmd.CommandText, Conn, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False
          
       If RS.RecordCount > 0 Then
         txtCurrentBed = "" & RS!Bed_type & " -  " & RS!bed_TYPE_no & " -  " & RS!bed_no
         VAR_CUR_BED_SERIAL_NO = RS!SERIAL_NO
       End If
       
       If Conn.State = 1 Then
        Conn.Close
        Set Conn = Nothing
        Set RS = Nothing
        Set cmd = Nothing
     End If


End Sub
Private Sub total_adv()

    Adodc5.ConnectionString = strcn.Connection_String
    Adodc5.RecordSource = "select  nvl(sum(advance),0)as advance from advance where in_reg_no ='" & Trim(frmReadvancepayment.txtReg_noInTest.Text) & "' AND YRCODE='" & Trim(frmReadvancepayment.CBOYRCODE.Text) & "'"
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
Private Sub txtCurpayment_GotFocus()
  TXTREC_NO.Visible = False
End Sub
