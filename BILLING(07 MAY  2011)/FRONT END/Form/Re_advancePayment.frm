VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmReadvancepayment 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   -30
      TabIndex        =   6
      Top             =   1770
      Width           =   5145
      Begin VB.Image Image2 
         Height          =   855
         Left            =   -2400
         Picture         =   "Re_advancePayment.frx":0000
         Top             =   -30
         Width           =   11820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   -90
      TabIndex        =   5
      Top             =   -120
      Width           =   5145
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RE-ADVANCE PAYMENT"
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
         Left            =   150
         TabIndex        =   7
         Top             =   210
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -450
         Picture         =   "Re_advancePayment.frx":5982
         Top             =   0
         Width           =   11820
      End
   End
   Begin VB.ComboBox CBOYRCODE 
      Height          =   315
      ItemData        =   "Re_advancePayment.frx":B304
      Left            =   1380
      List            =   "Re_advancePayment.frx":B317
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   750
      Width           =   2835
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1380
      Top             =   1380
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
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4275
      MaskColor       =   &H00FF0000&
      TabIndex        =   2
      Top             =   1110
      Width           =   555
   End
   Begin VB.TextBox txtReg_noInTest 
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
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   1380
      TabIndex        =   0
      Top             =   1170
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FISCAL YEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   810
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REG NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "frmReadvancepayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Con As New MyConnection
Dim Conn As New Connection
Dim Conn1 As New Connection
Dim Conn2 As New Connection
Dim cmd As New Command
Dim RS As New Recordset
Public UTILITY As New clsUtility
Public strUid As String
Public strcn        As New MyConnection

Private Sub Form_Activate()
txtReg_noInTest = ""
txtReg_noInTest.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
End If

End Sub

Private Sub Form_Load()
 txtReg_noInTest = ""
 CBOYRCODE.ListIndex = 0
 Select Case formIndicator
        Case 0
            Label7.Caption = "Re-Advance Entry"
        Case 1
            Label7.Caption = "CCU BED Entry"
        Case Else
 End Select
End Sub

Private Sub Command1_Click()
  On Error GoTo ERR_DESC
 Dim MSG As String
 MSG = UTILITY.GetPatientCurrentStatusInStringValue(txtReg_noInTest, CBOYRCODE)
 Select Case MSG
        Case 0
           If formIndicator = 0 Then
              frm_reAdvance.Show 1
          ElseIf formIndicator = 1 Then
              frmCCU_BED_Main.Show 1
          End If
              txtReg_noInTest.Text = ""
              Exit Sub
        Case Else
           MsgBox MSG, vbInformation, "IT DIVISION,DNMIH"
           txtReg_noInTest.Text = ""
           Exit Sub
 End Select

Exit Sub
ERR_DESC:
        MsgBox Err.Description, vbInformation, "IT DIVISION,DHMIH"


End Sub

Private Sub txtReg_noInTest_Change()
If Not IsNumeric(txtReg_noInTest.Text) Then
            txtReg_noInTest = ""
End If

End Sub


Private Sub txtReg_noInTest_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
End If
End Sub

Private Sub txtReg_noInTest_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         Command1_Click
    End If
    
End Sub
