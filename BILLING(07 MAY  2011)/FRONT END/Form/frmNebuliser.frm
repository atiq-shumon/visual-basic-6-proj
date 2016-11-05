VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmNebuliser 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Operation information Entry"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   0
      TabIndex        =   13
      Top             =   840
      Width           =   9735
      Begin VB.ComboBox cboOprBed 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc2"
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmNebuliser.frx":0000
         Left            =   7620
         List            =   "frmNebuliser.frx":0010
         TabIndex        =   22
         Top             =   600
         Width           =   1650
      End
      Begin VB.ComboBox cboOprDept 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc2"
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmNebuliser.frx":0028
         Left            =   7620
         List            =   "frmNebuliser.frx":0032
         TabIndex        =   20
         Top             =   180
         Width           =   1650
      End
      Begin VB.TextBox txtInregOpr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2400
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
         Top             =   660
         Width           =   3975
      End
      Begin VB.TextBox txtNameOpr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2370
         MultiLine       =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "If you want to edit previous patient information then put here Patient ID and press Enter"
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   7500
         TabIndex        =   25
         Top             =   540
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   7500
         TabIndex        =   24
         Top             =   150
         Width           =   75
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6390
         TabIndex        =   23
         Top             =   660
         Width           =   885
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6390
         TabIndex        =   21
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No"
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
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   2250
         TabIndex        =   17
         Top             =   180
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   2250
         TabIndex        =   16
         Top             =   600
         Width           =   105
      End
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
      Height          =   465
      Left            =   555
      Picture         =   "frmNebuliser.frx":0047
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Preview"
      Top             =   4980
      Width           =   510
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
      Height          =   465
      Left            =   60
      Picture         =   "frmNebuliser.frx":06B1
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save"
      Top             =   4980
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1080
      Picture         =   "frmNebuliser.frx":0D1B
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   4980
      Width           =   510
   End
   Begin VB.TextBox txtOprtotal 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc4"
      Height          =   330
      Left            =   8400
      TabIndex        =   11
      Top             =   4950
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3360
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmNebuliser.frx":1639
      Height          =   2295
      Left            =   0
      TabIndex        =   10
      Top             =   2640
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   16777215
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1920
      Top             =   5040
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   0
      TabIndex        =   5
      Top             =   -90
      Width           =   9735
      Begin VB.ComboBox cboOprcode 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmNebuliser.frx":164E
         Left            =   360
         List            =   "frmNebuliser.frx":1650
         TabIndex        =   31
         Top             =   2310
         Width           =   1005
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1065
         Left            =   -30
         TabIndex        =   29
         Top             =   0
         Width           =   9765
         Begin VB.Image Image1 
            Height          =   480
            Left            =   4860
            Picture         =   "frmNebuliser.frx":1652
            Top             =   210
            Width           =   480
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Nebulizer Charge Entry"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   795
            Left            =   5490
            TabIndex        =   30
            Top             =   180
            Width           =   4155
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   8220
         TabIndex        =   26
         Top             =   2310
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38203
      End
      Begin VB.ComboBox cboOprName 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc2"
         Height          =   315
         ItemData        =   "frmNebuliser.frx":1F1C
         Left            =   1365
         List            =   "frmNebuliser.frx":1F1E
         TabIndex        =   0
         Text            =   "Nebuliser Charge"
         Top             =   2310
         Width           =   5760
      End
      Begin VB.TextBox TxtOPrCharge 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc2"
         Height          =   300
         Left            =   7140
         TabIndex        =   1
         Top             =   2310
         Width           =   1065
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   8520
         TabIndex        =   28
         Top             =   2070
         Width           =   735
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Case"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7920
         TabIndex        =   9
         Top             =   4710
         Width           =   585
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7320
         TabIndex        =   8
         Top             =   2070
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Serial No"
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
         Index           =   0
         Left            =   390
         TabIndex        =   7
         Top             =   2070
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  Name"
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
         Left            =   1680
         TabIndex        =   6
         Top             =   2070
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5400
      Top             =   5040
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   8595
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   4650
      TabIndex        =   27
      Top             =   2640
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   585
      Left            =   30
      Top             =   4920
      Width           =   1605
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Charge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7170
      TabIndex        =   12
      Top             =   5010
      Width           =   1125
   End
End
Attribute VB_Name = "frmNebuliser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As New MyConnection
Dim Conn2 As New Connection
Dim Conn3 As New Connection
Dim rs2 As New Recordset
Dim rs3 As New Recordset
Dim cmd As New Command
Private Sub flush_grid()
  Dim opr_total, annay_total
  
         Adodc3.ConnectionString = strcn.Connection_String
        Adodc3.RecordSource = "select serial_no,name,nebuliser_charge,neb_date from indoor_pat_nebuliser_info Where in_reg_no ='" & Trim(frmReg_nebuliser.txtReg_noOpr.Text) & "'"

         Adodc3.Refresh
         
       Adodc4.ConnectionString = strcn.Connection_String
        Adodc4.RecordSource = "select sum(nebuliser_charge)as t_s  from indoor_pat_nebuliser_info Where in_reg_no ='" & Trim(frmReg_nebuliser.txtReg_noOpr.Text) & "'"
          
           Adodc4.Refresh
         If IsNull(opr_total) Then
                opr_total = 0
          End If
         
         
         txtOprtotal = Val("" & Adodc4.Recordset.Fields(0))
         
End Sub
'Private Sub cboOprcode_Click()
'       Adodc2.ConnectionString = strcn.Connection_String
'         Adodc2.RecordSource = "select opr_name,opr_type,opr_department,opr_bed,opr_charge,annay_charge from Operation_info where opr_code='" & Trim(cboOprcode.Text) & "'"
'         Adodc2.Refresh
'
'         cboOprName.clear
'         cboOprType.clear
'         cboOprDept.clear
'         cboOprBed.clear
'
'     If Adodc2.Recordset.RecordCount > 0 Then
'
'        Adodc2.Recordset.MoveFirst
'
'        While Adodc2.Recordset.EOF = False
'            cboOprName.AddItem Adodc2.Recordset!opr_name
'            cboOprType.AddItem Adodc2.Recordset!opr_type
'            cboOprDept.AddItem Adodc2.Recordset!opr_department
'            cboOprBed.AddItem Adodc2.Recordset!opr_bed
'            TxtOPrCharge = Adodc2.Recordset!opr_charge
'            txtannayCharge = Adodc2.Recordset!annay_charge
'            Adodc2.Recordset.MoveNext
'        Wend
'        cboOprName = cboOprName.List(0)
'        cboOprType = cboOprType.List(0)
'        cboOprDept = cboOprDept.List(0)
'        cboOprBed = cboOprBed.List(0)
'
'
'    End If
'
'End Sub

Private Sub CMDEXIT_Click()

    Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

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
              If conn.State = 1 Then
                  conn.Close
                  Set conn = Nothing
              End If
          Adodc2.ConnectionString = strcn.Connection_String
          Adodc2.RecordSource = "Select * From user_validation"
          Adodc2.Refresh
        

        
             If Adodc2.Recordset!validation = 0 Then
             MsgBox "Your Working Time has been Expired", vbInformation, " IT, DNMIH."
             Exit Sub
             End If
             

Call saveNebuliser_info
MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
Call flush_grid
End Sub
Private Sub saveNebuliser_info()
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
 If conn.State = 0 Then
     conn.Open strcn.Connection_String
 End If
    Set cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adDouble, adParamInput, 10, frmReg_nebuliser.txtReg_noOpr.Text)
    cmd.Parameters.Append Param1 'in_reg_no
           
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 50, Trim(cboOprName.Text))
    cmd.Parameters.Append Param2 'Operation_name
    
    Set Param3 = cmd.CreateParameter("param3", adDouble, adParamInput, 10, TxtOPrCharge)
    cmd.Parameters.Append Param3 'Operation charge
    
    
   Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 5, frmMAIN.lblBooth)
    cmd.Parameters.Append Param4 'booth
    
   Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 5, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param5 'u_id
    
    Set Param6 = cmd.CreateParameter("param6", adDate, adParamInput, 10, DTPicker1.Value)
    cmd.Parameters.Append Param6 'date
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL Save_Nebulizer_indoor(?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set rs2 = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    If conn.State = 1 Then
       conn.Close
       Set conn = Nothing
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
            SendKeys Chr(9)
End If

End Sub

Private Sub Form_Load()

      txtInregOpr = frmReg_nebuliser.txtReg_noOpr
      
'         Adodc1.ConnectionString = strcn.Connection_String
'         Adodc1.RecordSource = "select opr_code from Operation_info"
'         Adodc1.Refresh
        
'     If Adodc1.Recordset.RecordCount > 0 Then
'        'cboOprcode.clear
'
'        Adodc1.Recordset.MoveFirst
'        While Adodc1.Recordset.EOF = False
'            cboOprcode.AddItem Adodc1.Recordset!opr_code
'            Adodc1.Recordset.MoveNext
'        Wend
'    End If
    '''''''''''''''data grid ''''''''''''''''''''''
            
        Call flush_grid
        
'         Set DataGrid1 = Adodc1.Recordset
         
          
    
    If Conn2.State = 0 Then
             Conn2.ConnectionString = strcn.Connection_String
             Conn2.Open
    End If
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select pat_name,age,sex,DOC_DEPT From in_door_pat_info_main Where in_reg_no ='" & Trim(frmReg_nebuliser.txtReg_noOpr.Text) & "'"
        cmd.Properties("iRowsetChange") = True
       cmd.Properties("updatability") = 7
        rs2.CursorLocation = adUseClient

        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
       
                  
       If rs2.RecordCount > 0 Then
                 txtNameOpr = rs2!pat_name
      If Not IsNull(rs2!doc_dept) Then
                cboOprDept = rs2!doc_dept
       End If
'         TxtAgeOpr = rs2!age
'         comSexOpr.Text = rs2!sex
        End If
        
      cmd.Properties("iRowsetChange") = False
     
        Set rs2 = Nothing
   If Conn2.State = 1 Then
            Conn2.Close
            Set Conn2 = Nothing
   End If
End Sub

Private Sub txtOprCharge_Change()
        If Not IsNumeric(TxtOPrCharge.Text) Then
                    TxtOPrCharge = ""
        End If

End Sub
