VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOpr_no 
   Appearance      =   0  'Flat
   BackColor       =   &H0095A392&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Patient Reg. No. Entry for Operation"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   300
      Top             =   600
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
      Caption         =   "Reg  No "
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   270
      TabIndex        =   1
      Top             =   270
      Width           =   645
   End
End
Attribute VB_Name = "frmOpr_no"
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
       Adodc1.ConnectionString = strcn.Connection_String
        Adodc1.RecordSource = "select a.release_flag as release_flag From in_door_pat_info_main a  Where  a.in_reg_no ='" & Trim(txtReg_noOpr.Text) & "'"
        Adodc1.Refresh
      
       If Adodc1.Recordset.RecordCount > 0 Then
            If Adodc1.Recordset!release_flag = "2" Then
               MsgBox "This Patient has been fled     ", vbInformation, "Warning: IT, DNMIH"
               txtReg_noOpr = ""
               txtReg_noOpr.SetFocus
               Exit Sub
            End If
       End If
   If Conn2.State = 0 Then
        Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
   End If
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select release_flag From in_door_pat_info_main Where in_reg_no ='" & Trim(txtReg_noOpr.Text) & "'"
      
        cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
        rs2.CursorLocation = adUseClient

        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
       
        cmd.Properties("iRowsetChange") = False

         If rs2.RecordCount > 0 Then
                If (rs2!release_flag) = "1" Or (rs2!release_flag) = "3" Then
                    MsgBox "This Patient has been Released", vbInformation, "Warning: IT, DNMIH"
                        Set rs2 = Nothing
                      If Conn2.State = 1 Then
                        Conn2.Close
                      End If
         
                        txtReg_noOpr = ""
                        txtReg_noOpr.SetFocus
                        Exit Sub
        
                 Else
                    frmOperation.Show vbModal
                    Set rs2 = Nothing
                      If Conn2.State = 1 Then
                        Conn2.Close
                      End If

                    End If
    Else
        MsgBox "Invalid Registration No", vbInformation, "Warning: IT, DNMIH"
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
