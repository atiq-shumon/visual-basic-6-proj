VERSION 5.00
Begin VB.Form frmtest_cancel_entry 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   -180
      TabIndex        =   3
      Top             =   -120
      Width           =   5175
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "REC. NO TO  CANCEL TEST"
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
         Left            =   60
         TabIndex        =   4
         Top             =   150
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -330
         Picture         =   "frmtest_cancel_entry.frx":0000
         Top             =   0
         Width           =   11820
      End
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
      Height          =   345
      Left            =   3870
      MaskColor       =   &H00FF0000&
      TabIndex        =   2
      Top             =   900
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
      Left            =   1530
      TabIndex        =   0
      Top             =   870
      Width           =   2340
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   -2580
      Picture         =   "frmtest_cancel_entry.frx":5982
      Top             =   1620
      Width           =   11820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RECEIPT  NO : "
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
      Left            =   150
      TabIndex        =   1
      Top             =   930
      Width           =   1380
   End
End
Attribute VB_Name = "frmtest_cancel_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Con As New MyConnection
Dim conn As New Connection
Dim Conn1 As New Connection
Dim Conn2 As New Connection
Dim cmd As New Command
Dim rs As New Recordset
Dim RS1 As New Recordset
Dim rs2 As New Recordset
'Public rptMode As Integer
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
'txtReg_noOpr = ""
End Sub

Private Sub Command1_Click()
If Conn2.State = 0 Then
 Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
 End If
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select reg_no From pat_info_main_out_door Where reg_no ='" & Trim(txtReg_noOpr.Text) & "'"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        rs2.CursorLocation = adUseClient

        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
        cmd.Properties("iRowsetChange") = False
         If rs2.RecordCount > 0 Then
                    test_cancellation.Show vbModal
                    Set rs2 = Nothing
                   If Conn2.State = 1 Then
                       Conn2.Close
                       Set Conn2 = Nothing
                    End If
        Else
        MsgBox "Invalid Receipt  No", vbInformation, " IT, DNMIH"
                Set rs2 = Nothing
                   If Conn2.State = 1 Then
                       Conn2.Close
                       Set Conn2 = Nothing
                    End If

                txtReg_noOpr = ""
                txtReg_noOpr.SetFocus

              Exit Sub
        End If
  
  txtReg_noOpr = ""
  txtReg_noOpr.SetFocus


'frmOperation.Show
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
