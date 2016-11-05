VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form14 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Subscription Rate Setup"
   ClientHeight    =   5430
   ClientLeft      =   1740
   ClientTop       =   1740
   ClientWidth     =   8640
   Icon            =   "frmSubscription.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Height          =   480
      Left            =   6300
      Picture         =   "frmSubscription.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4770
      Width           =   1185
   End
   Begin VB.CommandButton cmdClear 
      Height          =   480
      Left            =   2355
      Picture         =   "frmSubscription.frx":234C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4770
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Height          =   480
      Left            =   1035
      Picture         =   "frmSubscription.frx":3CDE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4770
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   480
      Left            =   4980
      Picture         =   "frmSubscription.frx":5670
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4770
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   480
      Left            =   3690
      Picture         =   "frmSubscription.frx":725A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4770
      Width           =   1185
   End
   Begin VB.TextBox txtWF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   240
      Left            =   4995
      MaxLength       =   5
      TabIndex        =   4
      Top             =   855
      Width           =   915
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4425
      Left            =   135
      TabIndex        =   12
      Top             =   180
      Width           =   8385
      Begin VB.TextBox txtUnion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   3780
         MaxLength       =   5
         TabIndex        =   3
         Top             =   675
         Width           =   870
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3120
         Left            =   135
         TabIndex        =   23
         Top             =   1035
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   5503
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         ColumnHeaders   =   0   'False
         ForeColor       =   12582912
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
      Begin VB.TextBox txtMos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   7155
         MaxLength       =   5
         TabIndex        =   6
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   7200
         TabIndex        =   21
         Top             =   675
         Width           =   825
      End
      Begin VB.TextBox txtBF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   6030
         MaxLength       =   5
         TabIndex        =   5
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox txtClub 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   2790
         MaxLength       =   5
         TabIndex        =   2
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox txtCoop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   1755
         MaxLength       =   5
         TabIndex        =   1
         Top             =   675
         Width           =   870
      End
      Begin VB.TextBox txtPercentage 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   6075
         TabIndex        =   13
         Top             =   675
         Width           =   825
      End
      Begin VB.TextBox txtSubs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   315
         MaxLength       =   5
         TabIndex        =   0
         Top             =   675
         Width           =   1275
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mosque"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   7290
         TabIndex        =   22
         Top             =   315
         Width           =   570
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   14
         Left            =   7065
         Top             =   270
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   5
         Left            =   7065
         Top             =   585
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   3210
         Index           =   2
         Left            =   135
         Top             =   990
         Width           =   8070
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   11
         Left            =   5940
         Top             =   585
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   1
         Left            =   1710
         Top             =   585
         Width           =   1050
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Subscription Code"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   270
         TabIndex        =   19
         Top             =   315
         Width           =   1290
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   4
         Left            =   135
         Top             =   585
         Width           =   1590
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   3
         Left            =   135
         Top             =   270
         Width           =   1590
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   8
         Left            =   1710
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cooperative"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   18
         Top             =   315
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   0
         Left            =   2745
         Top             =   585
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   6
         Left            =   2745
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Club"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   17
         Top             =   315
         Width           =   315
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   7
         Left            =   3735
         Top             =   585
         Width           =   1050
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   9
         Left            =   3735
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Union"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   4005
         TabIndex        =   16
         Top             =   315
         Width           =   420
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   420
         Index           =   10
         Left            =   4770
         Top             =   585
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   12
         Left            =   4770
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Welfare Fund"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   4860
         TabIndex        =   15
         Top             =   315
         Width           =   960
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FECCC7&
         BorderColor     =   &H00FECCC7&
         Height          =   330
         Index           =   13
         Left            =   5940
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "B Fund"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   6210
         TabIndex        =   14
         Top             =   315
         Width           =   510
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Index           =   0
      Left            =   1440
      Top             =   3555
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
      Caption         =   ""
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
      Index           =   2
      Left            =   2655
      Top             =   3555
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
      Caption         =   ""
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
      Index           =   3
      Left            =   2655
      Top             =   3825
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
      Caption         =   ""
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
      Index           =   4
      Left            =   225
      Top             =   3825
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
      Caption         =   ""
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
      Index           =   1
      Left            =   1440
      Top             =   3825
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
      Caption         =   ""
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
      Index           =   5
      Left            =   225
      Top             =   3555
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
      Caption         =   ""
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
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   855
      TabIndex        =   20
      Top             =   585
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Subs As New St_Subs
Private Subs_Rs As New Recordset
Dim Track_Id As Long

Private Sub cmdClear_Click()
    Clear_Screen
    txtSubs.SetFocus
End Sub

Private Sub cmdClose_Click()
    Close_Msg Me
End Sub

Private Sub cmdDelete_Click()
    Delete_Record "16", Track_Id
    Flash_Into_Grid
    Clear_Screen
    txtSubs.SetFocus
End Sub

Private Sub cmdSave_Click()
    With Subs
        .ConnString = strCN.Connection
        .Subs_Code = txtSubs
        .Uni = txtUnion
        .Cop = txtCoop
        .Clb = txtClub
        .WF = txtWF
        .BF = txtBF
        .Mos = txtMos
        .Track_Id = Track_Id
        .Save
    End With
    
    Flash_Into_Grid
    Clear_Screen
    Track_Id = 0
    txtSubs.SetFocus
End Sub

Private Sub DataGrid1_Click()

On Error Resume Next

    txtSubs = Subs_Rs!Subs_Code
    txtCoop = Subs_Rs!Cop
    txtUnion = Subs_Rs!Uni
    txtClub = Subs_Rs!Clb
    txtWF = Subs_Rs!WF
    txtBF = Subs_Rs!BF
    txtMos = Subs_Rs!Mos
    Track_Id = Subs_Rs!Track_Id
    
    txtCoop.SetFocus
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Track_Id = 0
    Flash_Into_Grid
   
End Sub

Public Sub Flash_Into_Grid()

    With Subs
        .ConnString = strCN.Connection
        Set Subs_Rs = .GetAll
    End With
    
     Set DataGrid1.DataSource = Subs_Rs
                    
        With DataGrid1
            .Columns(0).Width = 1230
            '.Columns(0).DataField = Subs_Rs!Fields(0)

            .Columns(1).Width = 1035
            '.Columns(1).DataField = Subs_Rs!Fields(1)

            .Columns(2).Width = 990
            '.Columns(2).DataField = Subs_Rs!Fields(2)

            .Columns(3).Width = 1055
            '.Columns(3).DataField = Subs_Rs!Fields(3)

            .Columns(4).Width = 1170
            '.Columns(4).DataField = Subs_Rs!Fields(4)

            .Columns(5).Width = 1140
            '.Columns(4).DataField = Subs_Rs!Fields(5)

            .Columns(6).Width = 1180
            '.Columns(4).DataField = Subs_Rs!Fields(6)

        End With
       
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Destroy Me
End Sub


Private Sub txtBF_KeyPress(KeyAscii As Integer)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only

End Sub

Private Sub txtClub_KeyPress(KeyAscii As Integer)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only

End Sub

Private Sub txtCoop_KeyPress(KeyAscii As Integer)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtMos_KeyPress(KeyAscii As Integer)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtUnion_KeyPress(KeyAscii As Integer)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub

Private Sub txtWF_KeyPress(KeyAscii As Integer)
      KeyAscii = IsNum(KeyAscii)         'Accept numeric only
End Sub
