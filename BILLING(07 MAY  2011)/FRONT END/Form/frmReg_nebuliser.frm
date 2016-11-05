VERSION 5.00
Begin VB.Form frmReg_nebuliser 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Patient Release"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4620
   FillColor       =   &H008080FF&
   ForeColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      TabIndex        =   2
      Top             =   270
      Width           =   465
   End
   Begin VB.TextBox txtRegNoRelease 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   765
      TabIndex        =   1
      Top             =   270
      Width           =   3165
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Reg No:"
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   315
      Width           =   735
   End
End
Attribute VB_Name = "frmReg_nebuliser"
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


Private Sub Command1_Click()
   If Conn2.State = 0 Then
        Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
    End If
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select release_flag From in_door_pat_info_main Where in_reg_no ='" & Trim(txtRegNoRelease.Text) & "'"
      
        cmd.Properties("iRowsetChange") = True
       cmd.Properties("updatability") = 7
        rs2.CursorLocation = adUseClient

        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
       cmd.Properties("iRowsetChange") = False

         If rs2.RecordCount > 0 Then
         If (rs2!release_flag) = "1" Then
         MsgBox "This Patient has been Released", vbInformation, "Warning:Daffodil Software Ltd"
          txtRegNoRelease = ""
    
          Set rs2 = Nothing
        If Conn2.State = 1 Then
            Conn2.Close
            Set Conn2 = Nothing
         End If
          txtRegNoRelease = ""
         txtRegNoRelease.SetFocus
         Exit Sub
        
         Else
         Call temp_calculation
                frmPatientRelease.Show
               
   
     End If
          
    Else
        MsgBox "Invalid Registration No", vbInformation, "Warning:Daffodil Software Ltd"
      Set rs2 = Nothing
        If Conn2.State = 1 Then
            Conn2.Close
            Set Conn2 = Nothing
         End If
      Exit Sub
           txtRegNoRelease.SetFocus
          txtRegNoRelease = ""
          txtRegNoRelease.SetFocus
       
    End If
           
           
           Set rs2 = Nothing
        If Conn2.State = 1 Then
            Conn2.Close
            Set Conn2 = Nothing
         End If
' frmPatientRelease.Conn7.Close
'          rs7.Close
          
End Sub

Private Sub temp_calculation()



    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

     Dim Param1 As New Parameter
  If Conn.State = 0 Then
     Conn.Open strcn.Connection_String
  End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, txtRegNoRelease)
    cmd.Parameters.Append Param1 'in_reg_no

    
    

    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL temp_indoor_calculation(?)}"
    
    Debug.Print cmd.CommandText
    
    Set RS = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
If Conn.State = 1 Then
   Conn.Close
   Set Conn = Nothing
End If
 
End Sub



Private Sub Form_Activate()
      txtRegNoRelease = ""
      txtRegNoRelease.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
       Unload Me
  End If

End Sub
Private Sub txtRegNoRelease_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
            Command1_Click
      End If
      If KeyAscii = 27 Then
         Unload Me
      End If
      
End Sub
