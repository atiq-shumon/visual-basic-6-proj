VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form13 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income Statement"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "rpt_income_statement.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      ForeColor       =   &H80000001&
      Height          =   795
      Left            =   -30
      TabIndex        =   2
      Top             =   -120
      Width           =   5775
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Income/Expenditure  Statement"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   210
         TabIndex        =   6
         Top             =   210
         Width           =   5325
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1185
      Left            =   -30
      TabIndex        =   9
      Top             =   510
      Width           =   5715
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   285
         Left            =   3765
         TabIndex        =   1
         Top             =   555
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         _Version        =   393216
         Format          =   54067201
         CurrentDate     =   37637
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   285
         Left            =   315
         TabIndex        =   0
         Top             =   555
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   54067201
         CurrentDate     =   37637
      End
      Begin VB.Shape Shape3 
         Height          =   375
         Left            =   3720
         Top             =   510
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   270
         Top             =   510
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Range"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   285
         Index           =   1
         Left            =   2190
         TabIndex        =   12
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   11
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   3780
         TabIndex        =   10
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      Height          =   915
      Left            =   -60
      TabIndex        =   7
      Top             =   1560
      Width           =   5775
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
         Left            =   5100
         Picture         =   "rpt_income_statement.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exit"
         Top             =   300
         Width           =   510
      End
      Begin VB.CommandButton cmdPREVIEW 
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
         Left            =   4590
         Picture         =   "rpt_income_statement.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Preview"
         Top             =   300
         Width           =   510
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000001&
         Caption         =   "Bengali"
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
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Caption         =   "English"
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
         Height          =   345
         Index           =   1
         Left            =   1140
         TabIndex        =   3
         Top             =   330
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         Height          =   525
         Left            =   4530
         Top             =   240
         Width           =   1125
      End
      Begin VB.Shape Shape6 
         Height          =   465
         Left            =   90
         Top             =   270
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3240
      Top             =   2490
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
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEXIT_Click()
    Dim reply As String
    reply = MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdPREVIEW_Click()
  If Option1(0).Value = True Then
        Screen.MousePointer = vbHourglass
        rptMode = 13
        CRViewer1.Show vbModal
  ElseIf Option1(1).Value = True Then
  
      Screen.MousePointer = vbHourglass
        rptMode = 20
        CRViewer1.Show vbModal
  End If
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub dtpStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub txtCompName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys Chr(9)
    End If
End Sub
Private Sub Form_Load()
   dtpStart.Value = Date
    dtpEnd.Value = Date
End Sub
