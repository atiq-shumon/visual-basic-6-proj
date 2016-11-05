VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Information"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Acct.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSalvageValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11160
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   1260
      Width           =   645
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   405
      Left            =   1440
      Top             =   2550
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   714
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
      Caption         =   "Adodc4"
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
   Begin VB.TextBox txtBangla_Name 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5850
      TabIndex        =   3
      Top             =   1260
      Width           =   3930
   End
   Begin VB.TextBox nbrDepRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10740
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   1260
      Width           =   405
   End
   Begin VB.ListBox lstCheckAccName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1200
      Left            =   9300
      Sorted          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2850
      Visible         =   0   'False
      Width           =   4065
   End
   Begin VB.CommandButton cmdEdit 
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
      Left            =   9300
      Picture         =   "Acct.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Edit"
      Top             =   360
      Width           =   510
   End
   Begin VB.CommandButton cmdDelete 
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
      Left            =   9810
      Picture         =   "Acct.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Delete"
      Top             =   360
      Width           =   510
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Acct.frx":1286
      Height          =   2490
      Left            =   105
      TabIndex        =   22
      Top             =   1575
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   4392
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
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
      ColumnCount     =   5
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4094.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   480.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3780
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
   Begin VB.TextBox txtAccHead 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   8700
      TabIndex        =   21
      Text            =   "AccHead"
      Top             =   -30
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.ComboBox cboUserHead 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   195
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   465
      Width           =   1770
   End
   Begin VB.TextBox txtUserAcc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   105
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   13
      Top             =   1260
      Width           =   1785
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
      Left            =   10320
      Picture         =   "Acct.frx":129B
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Preview"
      Top             =   360
      Width           =   510
   End
   Begin VB.ComboBox cboHeadName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1950
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   465
      Width           =   4080
   End
   Begin VB.TextBox txtAccName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1260
      Width           =   3960
   End
   Begin VB.TextBox nbrAccBudg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9780
      MaxLength       =   15
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   1260
      Width           =   945
   End
   Begin VB.CommandButton cmdADD 
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
      Left            =   8790
      Picture         =   "Acct.frx":1905
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "New"
      Top             =   360
      Width           =   510
   End
   Begin VB.CommandButton cmdEXIT 
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
      Height          =   420
      Left            =   10830
      Picture         =   "Acct.frx":1F6F
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Exit"
      Top             =   360
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
      Height          =   420
      Left            =   8280
      Picture         =   "Acct.frx":288D
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Save"
      Top             =   360
      Width           =   510
   End
   Begin VB.TextBox nbrTrack_id 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   6720
      MaxLength       =   10
      TabIndex        =   14
      Text            =   "Tack_Id"
      Top             =   30
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   4950
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
      Left            =   -120
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "S.Value"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   11160
      TabIndex        =   25
      Top             =   990
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts Name(In Bengali)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   5850
      TabIndex        =   24
      Top             =   990
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "D. Rate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10515
      TabIndex        =   23
      Top             =   990
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   525
      Left            =   8220
      Top             =   300
      Width           =   3210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Control Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2115
      TabIndex        =   20
      Top             =   195
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts Name(In English)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   1905
      TabIndex        =   19
      Top             =   990
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   17
      Top             =   990
      Width           =   690
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Budget"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   9840
      TabIndex        =   16
      Top             =   990
      Width           =   510
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Control Head"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   315
      TabIndex        =   15
      Top             =   195
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      Height          =   735
      Left            =   105
      Top             =   180
      Width           =   6000
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objmyConAcc As New ADODB.Connection
Dim strAccConn As String

Private Sub SaveAcct()
'On Error GoTo err_loop
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
  
    
    Dim userid As String
    userid = Form1.txtUserID.Text
    Dim DepAcc As Double
    DepAcc = 100
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 40, txtAccHead.Text)
    cmd.Parameters.Append Param1

    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 40, txtUserAcc.Text)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 40, txtAccName.Text)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 40, txtBangla_Name.Text)
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 9, Val(nbrAccBudg.Text))
    cmd.Parameters.Append Param5
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 9, Val(nbrDepRate.Text))
    cmd.Parameters.Append Param6
    
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 9, Val(DepAcc))
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 50, userid)
    cmd.Parameters.Append Param8
    Set Param9 = cmd.CreateParameter("param9", adDouble, adParamInput, 50, Val(txtSalvageValue))
    cmd.Parameters.Append Param9
    
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL SaveAcct(?, ?, ?, ?, ?, ?, ?, ?, ?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub


Private Sub ClearScreen()

    Me.txtUserAcc.Text = ""
    txtBangla_Name.Text = ""
    Me.txtAccName.Text = ""
    Me.nbrAccBudg.Text = "0.00"
    Me.nbrDepRate.Text = "0.00"
    Me.nbrTrack_id.Text = ""
    
End Sub

Private Sub cboHeadName_Click()

'    Dim Conn As New Connection
'    Dim cmd As New Command
'    Dim RS As New Recordset
'
'    Conn.Open strcn.Connection_String
'    Set cmd.ActiveConnection = Conn
'
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "Select user_acc,acc_name from Acct where acc_name= '" + Trim(cboHeadName.Text) + "'"
''    cmd.Properties("IRowsetChange") = True
''    cmd.Properties("Updatability") = 7
'
'    RS.CursorLocation = adUseClient
'    RS.Open cmd.CommandText, Conn, adOpenStatic, adLockOptimistic
'
'    If Not (RS.EOF And RS.BOF) Then
'        cboUserHead.Text = RS("User_acc")
'    Else
'        cboUserHead.Text = ""
'    End If
End Sub

Private Sub cboHeadName_GotFocus()
'    Call ShowControl
End Sub

Private Sub cboHeadName_LostFocus()

'    If Len(Trim(cboHeadName.Text)) = 0 Then Exit Sub
'    Call GetControlCode(Me, Trim(Me.cboHeadName.Text))
    
End Sub
Private Sub cboUserHead_Click()
    
'    Call GetControlName(Me, Trim(Me.cboUserHead.Text))
    
    Adodc2.ConnectionString = strcn.Connection_String
    Adodc2.RecordSource = "select acc_code,user_acc,acc_name from Acct where user_acc='" + cboUserHead.Text + "'"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
            cboHeadName.Text = Adodc2.Recordset!acc_name
            txtAccHead.Text = Adodc2.Recordset!acc_code
    End If
    Call AutoAccCode
'    Adodc4.ConnectionString = strcn.Connection_String
'    Adodc4.RecordSource = "select max(acc_code)as code from Acct where acc_head='" + txtAccHead.Text + "'"
'    Adodc4.Refresh
'    If Adodc4.Recordset.RecordCount > 0 Then
'
'        txtUserAcc.Text = Adodc4.Recordset!code + 1

'    End If
    txtAccName.Text = ""
    txtBangla_Name = ""
    Call GetGrdData
End Sub

Private Sub cboUserHead_GotFocus()
    Call ShowControl
End Sub

Private Sub cmdADD_Click()
    Call ClearScreen
End Sub
Private Sub AutoAccCode()
    On Error GoTo err_loop

    Adodc4.ConnectionString = strcn.Connection_String
    Adodc4.RecordSource = "select max(acc_code) as new_acct from acct where acc_head='" & Trim(txtAccHead.Text) & "'"
    Adodc4.Refresh
    

    If IsNull(Adodc4.Recordset!new_acct) = True Then
       If Len(Trim(txtAccHead.Text)) <= 2 Then
          txtUserAcc.Text = (Val(txtAccHead.Text) * 100) + 1
       Else
          txtUserAcc.Text = (Val(txtAccHead.Text) * 1000) + 1
       End If
    Else
       If Val(txtAccHead.Text) = Val(Adodc4.Recordset!new_acct) Then
          If Len(Trim(txtAccHead.Text)) <= 2 Then
             txtUserAcc.Text = (Val(txtAccHead.Text) * 100) + 1
          Else
             txtUserAcc.Text = (Val(txtAccHead.Text) * 1000) + 1
          End If
       Else
            txtUserAcc.Text = Val(Adodc4.Recordset!new_acct) + 1
       End If
    End If
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub
Private Sub cmdDELETE_Click()

    On Error GoTo err_loop
        If Len(Trim(Me.txtUserAcc.Text)) = 0 Then
           MsgBox "Accounts code required", vbCritical, "Warning..."
           Me.txtUserAcc.Text = ""
           Me.txtUserAcc.SetFocus
           Exit Sub
        End If

        If Val(Me.nbrTrack_id.Text) <= 0 Then
           MsgBox "Item not selected", vbCritical, "Warning..."
           Exit Sub
        End If
        '''Checking control head''''''''''''''
           Adodc2.ConnectionString = strcn.Connection_String
            Adodc2.RecordSource = "select acc_head from acct where acc_head in(select acc_code from acct where track_id=" & Val(nbrTrack_id.Text) & ")"
            Adodc2.Refresh
            If Adodc2.Recordset.RecordCount > 0 Then
                MsgBox ("You can not delete Control Head"), vbCritical, "Warning..."
            Exit Sub
            End If
            
'
'            Adodc3.ConnectionString = strcn.Connection_String
'            Adodc3.RecordSource = "select acc_code from ledger where acc_code in(select acc_code from acct where track_id=" & Val(nbrTrack_id.Text) & ")"
'            Adodc3.Refresh
'            If Adodc3.Recordset.RecordCount > 0 Then
'                MsgBox ("Code in use"), vbCritical, "Warning..."
'            Exit Sub
'            End If
        '''---------------------------------------------------------------
        
    
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
        
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adInteger, adParamInput, 4, Val(nbrTrack_id.Text))
    cmd.Parameters.Append Param1
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL DeleteAcct(?)}"
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    
    
    Call GetGrdData
    Call ClearScreen
 
     Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical
    Resume Next
End Sub

Private Sub cmdEdit_Click()
    
    On Error GoTo err_loop

    If Len(Trim(Me.txtUserAcc.Text)) = 0 Then
        MsgBox "Accounts code required", vbCritical, "IT Division,DNMIH"
        txtUserAcc.SetFocus
        Exit Sub
    End If

    If Len(Trim(Me.txtAccName.Text)) = 0 Then
        MsgBox "Accounts name required", vbCritical, "IT Division,DNMIH"
        txtAccName.SetFocus
        Exit Sub
    End If

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

    
    Dim userid As String
    userid = "Emdad"
    Dim DepAcc As Double
    DepAcc = 100
    Conn.Open strcn.Connection_String
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 40, txtUserAcc.Text)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 40, txtAccName.Text)
    cmd.Parameters.Append Param2
    
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 40, txtBangla_Name.Text)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adDouble, adParamInput, 9, Val(nbrAccBudg.Text))
    cmd.Parameters.Append Param4
    
    Set Param5 = cmd.CreateParameter("param5", adDouble, adParamInput, 9, Val(nbrDepRate.Text))
    cmd.Parameters.Append Param5
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 9, Val(txtSalvageValue.Text))
    cmd.Parameters.Append Param6
   
    
    Set Param7 = cmd.CreateParameter("param7", adInteger, adParamInput, 9, Val(nbrTrack_id.Text))
    cmd.Parameters.Append Param7
    
    Set Param8 = cmd.CreateParameter("param8", adVarChar, adParamInput, 50, userid)
    cmd.Parameters.Append Param8
    '----------------------------------------------------------------------------------

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL EditAcct(?,?,?, ?, ?, ?, ?,?)}"
Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
    
    
    Call GetGrdData
    MsgBox "Edited Successfully", vbInformation, "IT Division,DNMIH"
    Exit Sub
err_loop:
    MsgBox Err.Description, vbCritical, "IT Division,DNMIH"
    Resume Next

    
End Sub

Private Sub cmdEXIT_Click()

    Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdPREVIEW_Click()

'     rptMode = 1
'    CRViewer1.Show vbModal
    
End Sub

Private Sub cmdSAVE_Click()
     
     If Len(Trim(Me.cboUserHead.Text)) = 0 Then
        MsgBox "Control head required", vbCritical, "IT Division,DNMIH"
        cboUserHead.SetFocus
        Exit Sub
     End If
     
     If Len(Trim(Me.cboHeadName.Text)) = 0 Then
        MsgBox "Control name required", vbCritical, "IT Division,DNMIH"
        cboHeadName.SetFocus
        Exit Sub
     End If
     
     If Len(Trim(Me.txtUserAcc.Text)) = 0 Then
        MsgBox "Accounts code required", vbCritical, "IT Division,DNMIH"
        txtUserAcc.SetFocus
        Exit Sub
     End If
     
     If Len(Trim(Me.txtAccName.Text)) = 0 Then
        MsgBox "Accounts name required", vbCritical, "IT Division,DNMIH"
        txtAccName.SetFocus
        Exit Sub
     End If
     
     If Len(Trim(txtBangla_Name.Text)) = 0 Then
        MsgBox "Bangla Accounts name required", vbCritical, "IT Division,DNMIH"
        txtBangla_Name.SetFocus
        Exit Sub
     End If
     Call SaveAcct
     MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
     Call GetGrdData
     txtUserAcc.Text = ""
     txtAccName.Text = ""
'    txtBangla_Name.Text = ""
     txtAccName.SetFocus
'     Call ShowControl
     Call AutoAccCode
End Sub

Private Sub DataGrid1_Click()
  If DataGrid1.Row >= 0 Then
   Me.txtUserAcc.Text = Me.DataGrid1.Columns(0).Text
   Me.txtAccName.Text = Me.DataGrid1.Columns(1).Text
   Me.txtBangla_Name.Text = Me.DataGrid1.Columns(2).Text
    Me.nbrAccBudg.Text = Me.DataGrid1.Columns(3).Text
    Me.nbrTrack_id.Text = Me.DataGrid1.Columns(5).Text
    Me.lstCheckAccName.Visible = False
  End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
    
End Sub

Private Sub Form_Load()
'     strAccConn = "Provider=OraOLEDB.Oracle.1;Persist Security Info=true;User ID=acct;Password=dn_acct;Initial Catalog=bank;Data Source=bank;"
'     DBConnectionString = strAccConn
'     objmyConAcc.ConnectionString = strAccConn
'     objmyConAcc.Open
     Call ShowControl
     Call GetGrdData
     
End Sub

Private Sub GetGrdData()
'    On Error GoTo err_loop
        Adodc1.ConnectionString = strcn.Connection_String
        Adodc1.RecordSource = "select user_acc,acc_name,acc_name_beng,acc_budg,(select distinct case A.acc_group " & _
                          " WHEN 0 THEN 'Assets'" & _
                          " WHEN 1 THEN 'Assets'" & _
                          " WHEN 2 THEN  'Liabilities'" & _
                          " WHEN 3 THEN 'Equity'" & _
                          " WHEN 4 THEN  'Liabilities'" & _
                          " WHEN 5 THEN  'Income'" & _
                          " WHEN 6 THEN 'Expenses'" & _
                          " WHEN 7 THEN  'Expenses'" & _
                          " WHEN 8 THEN  'Expenses'" & _
                          " WHEN 9 THEN  'Income'" & _
                          " WHEN 10 THEN  'Expenses' END from Acct A where A.user_acc=Acct.user_acc) as category,track_id from acct " & _
                          "where acc_head='" & Trim(Me.txtAccHead.Text) & "' and acc_code <>'" & Trim(Me.txtAccHead.Text) & "'"
        Adodc1.Refresh
        
        DataGrid1.Columns(0).Width = 1470.047
        DataGrid1.Columns(0).Caption = "Code"
        DataGrid1.Columns(0).Locked = True

        DataGrid1.Columns(1).Width = 3950 + 3950
        DataGrid1.Columns(1).Caption = "Accounts Name(English)"
        
        DataGrid1.Columns(2).Width = 0
        DataGrid1.Columns(2).Caption = "Accounts Name(Bangla)"
        
        
'        DataGrid1.Columns(2).CellText = "Accounts Name(Bangla)"


        DataGrid1.Columns(3).Width = 1275.024
        DataGrid1.Columns(3).Caption = "Budget"
        DataGrid1.Columns(3).Alignment = dbgRight

        DataGrid1.Columns(4).Width = 2500
        DataGrid1.Columns(4).Caption = "Category"

       DataGrid1.Columns(5).Width = 100
        DataGrid1.Columns(5).Visible = False
'
'        DataGrid1.Columns(6).Width = 100
'        DataGrid1.Columns(6).Visible = False

'    Exit Sub
'err_loop:
'    MsgBox Err.Description, vbCritical
'    Resume Next
    
End Sub
Private Sub ShowControl()
  
    Adodc2.ConnectionString = strcn.Connection_String
    Adodc2.RecordSource = "select user_acc,acc_name from Acct"
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount > 0 Then
        cboUserHead.Clear
        cboHeadName.Clear
        Adodc2.Recordset.MoveFirst
       While Adodc2.Recordset.EOF = False
           cboUserHead.AddItem Adodc2.Recordset!user_acc
            cboHeadName.AddItem Adodc2.Recordset!acc_name
            Adodc2.Recordset.MoveNext

       Wend
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set objmyConAcc = Nothing
End Sub

Private Sub nbrAccBudg_GotFocus()

    nbrAccBudg.SelLength = Len(nbrAccBudg.Text)
    
End Sub

Private Sub nbrAccBudg_KeyPress(KeyAscii As Integer)

    If KeyAscii > 26 Then
       If InStr("0123456789.+-", Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    
End Sub

Private Sub nbrDepRate_GotFocus()

    nbrDepRate.SelLength = Len(nbrDepRate.Text)
    
End Sub

Private Sub nbrDepRate_KeyPress(KeyAscii As Integer)

    If KeyAscii > 26 Then
       If InStr("0123456789.+-", Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    
End Sub

Private Sub txtAccName_Change()

'    On Error GoTo err_loop
'    If Len(Trim(txtAccName.Text)) = 0 Then
'       lstCheckAccName.Visible = False
'       Exit Sub
'    Else
'       Me.lstCheckAccName.Left = Me.txtAccName.Left
'       Me.lstCheckAccName.Top = Me.DataGrid1.Top
'
'       lstCheckAccName.Visible = True
'    End If
'
'    lstCheckAccName.Clear
'    Con.ConnectionString = strcn
'    Con.Open
'    RS.Open "select acc_name from acct where acc_name like '" & Trim(txtAccName.Text) & "%'", Con
'    If RS.EOF = False Then
'        Do Until RS.EOF
'            lstCheckAccName.AddItem RS!acc_name
'            RS.MoveNext
'        Loop
'    End If
'    RS.Close
'    Con.Close
'    Exit Sub
'err_loop:
'    MsgBox Err.Description, vbCritical
'    Resume Next
    
End Sub

Private Sub txtAccName_KeyPress(KeyAscii As Integer)

    If KeyAscii = 39 Then
       KeyAscii = Asc(Chr(96))
    End If
    
End Sub

Private Sub txtAccName_LostFocus()

    lstCheckAccName.Visible = False
    
End Sub

