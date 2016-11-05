VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_otherchargereceipt 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6615
   ClientLeft      =   -105
   ClientTop       =   390
   ClientWidth     =   9315
   FillColor       =   &H007DABD0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   6240
      Width           =   13365
      Begin VB.Label Label9 
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
         TabIndex        =   28
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
         TabIndex        =   27
         ToolTipText     =   "Developed && Maintenanced by:A.N.M. Atiqur Rahman,Software Programmer, IT Division, DNMIH"
         Top             =   90
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   8010
      TabIndex        =   25
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   6780
      TabIndex        =   24
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "NEW"
      Height          =   375
      Left            =   5550
      TabIndex        =   23
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox TXTPRINT_Others 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc3"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   270
      TabIndex        =   21
      Top             =   5340
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000001&
      Height          =   3495
      Left            =   30
      TabIndex        =   13
      Top             =   2160
      Width           =   9285
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   270
         TabIndex        =   29
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   270
         TabIndex        =   5
         Top             =   2280
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   270
         TabIndex        =   4
         Top             =   1380
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   270
         TabIndex        =   3
         Top             =   900
         Width           =   255
      End
      Begin VB.TextBox Text1 
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
         ForeColor       =   &H00FF0000&
         Height          =   600
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2550
         Visible         =   0   'False
         Width           =   6375
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
         Height          =   360
         Left            =   7020
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   570
         Width           =   1875
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Consultant Ticket  Collection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   630
         TabIndex        =   30
         Top             =   1800
         Width           =   2955
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   240
         Left            =   7440
         TabIndex        =   20
         Top             =   330
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Purpose Of  Collection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   240
         Left            =   3300
         TabIndex        =   19
         Top             =   2250
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Others Collection"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   240
         Left            =   630
         TabIndex        =   16
         Top             =   2280
         Width           =   1785
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emergency Ticket Collection"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   240
         Left            =   660
         TabIndex        =   15
         Top             =   1365
         Width           =   2985
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OPD Ticket Collection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   240
         Left            =   720
         TabIndex        =   14
         Top             =   900
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   0
      TabIndex        =   12
      Top             =   -90
      Width           =   9315
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OTHERS INCOME ENTRY"
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
         Left            =   2250
         TabIndex        =   22
         Top             =   300
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   0
         Picture         =   "frm_other_receipt.frx":0000
         Stretch         =   -1  'True
         Top             =   60
         Width           =   10290
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   0
      TabIndex        =   10
      Top             =   690
      Width           =   9315
      Begin VB.TextBox txtNameInTest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   330
         TabIndex        =   0
         Top             =   450
         Width           =   4125
      End
      Begin VB.TextBox txtAddrInTest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   330
         TabIndex        =   2
         Top             =   1110
         Width           =   8625
      End
      Begin MSComCtl2.DTPicker Dt_date 
         Height          =   330
         Left            =   6990
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   450
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   16711680
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   49152
         Format          =   63242241
         CurrentDate     =   37114
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
         TabIndex        =   18
         Top             =   195
         Width           =   690
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
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Left            =   6990
         TabIndex        =   17
         Top             =   195
         Width           =   510
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
         TabIndex        =   11
         Top             =   795
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
      Left            =   3570
      Top             =   4800
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
      Left            =   3570
      Top             =   4770
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
      Left            =   4260
      Top             =   5700
      Width           =   4995
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   -6660
      TabIndex        =   9
      Top             =   4770
      Width           =   270
   End
End
Attribute VB_Name = "frm_otherchargereceipt"
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
Dim UTILITY As New clsUtility
Dim VoucherNumber
Dim OptionBtnvalue As Integer
Public strcn        As New MyConnection
Private Sub Check1_Click()
  TxtPreviousPayment.Top = 780
   If Check1.Value = 0 Then
      Label4.Enabled = False
      Label13.Enabled = False
   End If
 
  If Check1.Value = 1 Then
      TxtPreviousPayment.Top = 780
      Label3.Top = TxtPreviousPayment.Top - 200
      OptionBtnvalue = 1
      Label4.Enabled = True
      Check2.Value = 0
      Check3.Value = 0
      Check4.Value = 0
  End If
End Sub

Private Sub Check2_Click()
  TxtPreviousPayment.Top = 1110
  If Check2.Value = 1 Then
    Label6.Enabled = True
     TxtPreviousPayment.Top = 1110
    OptionBtnvalue = 2
  Else
    Label6.Enabled = False
    Label13.Enabled = False
  End If
  Check1.Value = 0
  Check3.Value = 0
  Check4.Value = 0
End Sub

Private Sub Check3_Click()
TxtPreviousPayment.Top = 1560
   If Check3.Value = 1 Then
      OptionBtnvalue = 3
      Label1.Visible = True
      Text1.Visible = True
      Label12.Enabled = True
      TxtPreviousPayment.Top = 1560
      Text1.SetFocus
      
   Else
      Label1.Visible = False
      Text1.Visible = False
      Label12.Enabled = False
      Label13.Enabled = False
  End If
  Check1.Value = 0
  Check2.Value = 0
  Check4.Value = 0
End Sub
Private Sub Check4_Click()
TxtPreviousPayment.Top = 1710
Label3.Top = 1380
Label13.Enabled = True
   If Check4.Value = 1 Then
      OptionBtnvalue = 4
      Label1.Visible = False
      Text1.Visible = False
      Label12.Enabled = False
      TxtPreviousPayment.Top = 1560
  Else
      Label1.Visible = False
      Text1.Visible = False
      Label12.Enabled = False
      Label4.Enabled = False
  End If
  Check1.Value = 0
  Check2.Value = 0
  Check3.Value = 0
End Sub
Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdPreview_Click()

End Sub

Private Sub cmdPrint_Click()
If TXTPRINT_Others.Visible = False Then
       TXTPRINT_Others.Visible = True
   End If
  TXTPRINT_Others.ForeColor = vbBlue
  PREVIEW_VAR = Val(TXTPRINT_Others)
  
   If TXTPRINT_Others = "" Then
    TXTPRINT_Others.SetFocus
     Exit Sub
   Else
    PREVIEW_VAR = Val(TXTPRINT_Others)
        rptMode = 42
      Viewer.Show vbModal
      TXTPRINT_Others = ""
      TXTPRINT_Others.Visible = False
End If

End Sub



Private Sub cmdSave_Click()
 MsgBox (OptionBtnvalue)
 Exit Sub
 
If UTILITY.User_Shift_validation(frmMAIN.lbluser_id, frmMAIN.lblUserType.Caption) = False Then
    MsgBox "Mr. " & frmMAIN.lblName & "  Your Shift has been Expired.. " & vbCrLf & " " & vbCrLf & "Please Contact With Administrator", vbInformation, " IT, DNMIH."
    Exit Sub
End If

If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 Then
   MsgBox "Please select an check box", vbInformation, " IT, DNMIH."
   Exit Sub
End If
   
If TxtPreviousPayment = "" Then
    MsgBox "Please enter an Amount ", vbInformation, " IT, DNMIH."
   Exit Sub
End If
   
Call save_othercoll
MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
CMDEXIT.SetFocus

''''''''''''''printing''''''''''''
    rptMode = 42
    Viewer.Show vbModal
txtNameInTest = ""
txtAddrInTest = ""
TxtPreviousPayment = ""
Check1.Value = 1
CMDEXIT.SetFocus

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
            
            
            Dim Reportoth   As New CrystalReport13
            cmdre.Properties("PLSQLRSet") = True
            cmdre.CommandText = "{CALL Rpt_Others_money}"
            Set RSre = cmdre.Execute
            cmdre.Properties("PLSQLRSet") = False
            
            Reportoth.Database.SetDataSource RSre

            Reportoth.PrintOut
           
          
            Set RSre = Nothing
           If Connre.State = 1 Then
            Connre.Close
          End If
End Sub
 Private Sub save_othercoll()
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
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 30, txtNameInTest)
    cmd.Parameters.Append Param1 'IN_REG_NO
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 100, txtAddrInTest)
    cmd.Parameters.Append Param2 'IN_REG_NO
     
     Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 200, Text1)
    cmd.Parameters.Append Param3 'IN_REG_NO
  
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 30, frmMAIN.lbluser_id)
    cmd.Parameters.Append Param4 'U_id default Sumon
    
    Set Param5 = cmd.CreateParameter("param5", adVarChar, adParamInput, 30, frmMAIN.lblBooth)
    cmd.Parameters.Append Param5 'booth
    
    Set Param6 = cmd.CreateParameter("param6", adDouble, adParamInput, 10, TxtPreviousPayment)
    cmd.Parameters.Append Param6 'readvance
    
    Set Param7 = cmd.CreateParameter("param7", adDouble, adParamInput, 5, OptionBtnvalue)
    cmd.Parameters.Append Param7 'option button
    
    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL save_othersmoney(?,?,?,?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    If Conn.State = 1 Then
       Conn.Close
    End If
    ''''''''''''
    Adodc4.ConnectionString = strcn.Connection_String
      Adodc4.RecordSource = "SELECT MAX(REC_NO) AS REC_NO FROM RECEIPT_NO_COUNTER"
      Adodc4.Refresh
    If Adodc4.Recordset.RecordCount > 0 Then
        TXTPRINT_Others = Adodc4.Recordset!REC_NO
          PREVIEW_VAR = Val(TXTPRINT_Others)
    End If
  
    End Sub


Private Sub Form_Activate()
     '''''.SetFocus
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

Private Sub Form_Load()

    OptionBtnvalue = 1
    Me.Check1.Value = 1
    Dt_date = Date
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

Private Sub TxtPreviousPayment_Change()
  If Not IsNumeric(TxtPreviousPayment) Then
     TxtPreviousPayment = ""
  End If
End Sub
