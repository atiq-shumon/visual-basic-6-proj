VERSION 5.00
Begin VB.Form frmReg_for_EDIT_PAT 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   5685
      Begin VB.Image Image2 
         Height          =   855
         Left            =   -810
         Picture         =   "frmReg_for_EDIT_PAT.frx":0000
         Top             =   -210
         Width           =   11820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   -120
      TabIndex        =   5
      Top             =   -60
      Width           =   5685
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EDIT PAT. INFORMATION"
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
         Left            =   510
         TabIndex        =   7
         Top             =   180
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -720
         Picture         =   "frmReg_for_EDIT_PAT.frx":5982
         Top             =   0
         Width           =   11820
      End
   End
   Begin VB.ComboBox CBOYRCODE 
      Height          =   315
      ItemData        =   "frmReg_for_EDIT_PAT.frx":B304
      Left            =   1320
      List            =   "frmReg_for_EDIT_PAT.frx":B306
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   2835
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4140
      MaskColor       =   &H00FF0000&
      TabIndex        =   2
      Top             =   1410
      Width           =   525
   End
   Begin VB.TextBox txtReg_noOpr 
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
      Height          =   360
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Enter Registration No to Edit Patient Information"
      Top             =   1410
      Width           =   2820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FISCAL YEAR"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REG NO :"
      ForeColor       =   &H00004080&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "frmReg_for_EDIT_PAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Con As New MyConnection
Dim Conn As New Connection
Dim cmd As New Command
Dim RS As New Recordset
Public UTILITY As New clsUtility
Public strUid As String
Public strcn        As New MyConnection


Private Sub Form_Activate()
txtReg_noOpr = ""
txtReg_noOpr.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
End If

End Sub

Private Sub Form_Load()
   PopulateFiscalYear
End Sub
Private Sub PopulateFiscalYear()
   Dim yearList() As String
   Dim i As Integer
   yearList = UTILITY.GetFiscalYears()
   For i = LBound(yearList) To UBound(yearList)
       CBOYRCODE.AddItem yearList(i)
   Next i
   CBOYRCODE.ListIndex = 0
End Sub
Private Sub Command1_Click()

   On Error GoTo ERR_DESC
    Dim MSG As String
    MSG = UTILITY.GetPatientCurrentStatusInStringValue(txtReg_noOpr, CBOYRCODE)
    Select Case MSG
        Case 0
            Form1.Show 1
            txtReg_noOpr.Text = ""
            Exit Sub
        Case Else
           MsgBox MSG, vbInformation, "IT DIVISION,DNMIH"
           txtReg_noOpr.Text = ""
           Exit Sub
    End Select
  
  Exit Sub
ERR_DESC:
        MsgBox Err.Description, vbInformation, "IT DIVISION,DHMIH"
End Sub

Private Sub txtReg_noInTest_Change()

End Sub

Private Sub txtReg_noInTest_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub txtReg_noOpr_Change()
If Not IsNumeric(txtReg_noOpr.Text) Then
            txtReg_noOpr = ""
End If
End Sub

Private Sub txtReg_noOpr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
End If

End Sub

Private Sub txtReg_noOpr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Command1_Click
    End If
End Sub
