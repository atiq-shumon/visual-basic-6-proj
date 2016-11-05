VERSION 5.00
Begin VB.Form frmREPRINT_ADVANCE 
   Appearance      =   0  'Flat
   BackColor       =   &H00DDD6DE&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "RECEIPT RE-PRINT FOR RE-ADVANCE"
   ClientHeight    =   930
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "OK"
      Height          =   315
      Left            =   3810
      MaskColor       =   &H00FF0000&
      TabIndex        =   2
      Top             =   270
      Width           =   525
   End
   Begin VB.TextBox txtReg_noOpr 
      Appearance      =   0  'Flat
      BackColor       =   &H00DDEAEE&
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
      Left            =   900
      TabIndex        =   0
      Top             =   270
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REC  NO"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   150
      TabIndex        =   1
      Top             =   270
      Width           =   660
   End
End
Attribute VB_Name = "frmREPRINT_ADVANCE"
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
        
           
           Dim Report4   As New CrystalReport4
            cmd.Properties("PLSQLRSet") = True
            cmd.CommandText = "{CALL Rptout_door_info_print}"
            Set RS = cmd.Execute
            cmd.Properties("PLSQLRSet") = False
            
            Report4.Database.SetDataSource RS

            Report4.PrintOut
            RS.Close
    Else
        MsgBox "Invalid Registration No", vbInformation, "Warning:Daffodil Software Ltd"
                     Set rs2 = Nothing
                      If Conn2.State = 1 Then
                            Conn2.Close
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
