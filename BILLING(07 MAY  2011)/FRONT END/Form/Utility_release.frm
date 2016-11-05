VERSION 5.00
Begin VB.Form frmUlitity_release 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5265
   FillColor       =   &H8000000B&
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "SHOW"
      Height          =   375
      Left            =   4020
      TabIndex        =   2
      ToolTipText     =   "PRESS TO SHOW REPORT"
      Top             =   2010
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Height          =   825
      Left            =   -30
      TabIndex        =   8
      Top             =   2520
      Width           =   5385
      Begin VB.Image Image2 
         Height          =   855
         Left            =   0
         Picture         =   "Utility_release.frx":0000
         Top             =   90
         Width           =   11820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   -180
      TabIndex        =   7
      Top             =   -120
      Width           =   5565
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RELEASE PAT. REC. REVIEW"
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
         Left            =   450
         TabIndex        =   9
         Top             =   210
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   885
         Left            =   180
         Picture         =   "Utility_release.frx":5982
         Stretch         =   -1  'True
         Top             =   120
         Width           =   11820
      End
   End
   Begin VB.TextBox txtRecNo 
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
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1335
      TabIndex        =   5
      Top             =   1470
      Width           =   2655
   End
   Begin VB.ComboBox CBOYRCODE 
      Height          =   315
      ItemData        =   "Utility_release.frx":B304
      Left            =   1320
      List            =   "Utility_release.frx":B306
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   930
      Width           =   2625
   End
   Begin VB.TextBox txtRegNoRelease 
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
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1335
      TabIndex        =   1
      Top             =   2010
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rec No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   90
      TabIndex        =   6
      Top             =   1545
      Width           =   1095
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
      Left            =   90
      TabIndex        =   4
      Top             =   990
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   2085
      Width           =   1095
   End
End
Attribute VB_Name = "frmUlitity_release"
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
Public Utility As New clsUtility
Private Sub cmdPrint_Click()
On Error GoTo ERR_DESC
If Conn2.State = 0 Then
        Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
End If
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select release_flag From in_door_pat_info_main Where in_reg_no ='" & Trim(txtRegNoRelease.Text) & "' AND YRCODE='" & Trim(CBOYRCODE.Text) & "'"
      
        cmd.Properties("iRowsetChange") = True
       cmd.Properties("updatability") = 7
        rs2.CursorLocation = adUseClient

        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
       cmd.Properties("iRowsetChange") = False
         If rs2.RecordCount > 0 Then
         If (rs2!release_flag) = "1" Or (rs2!release_flag) = "3" Then
         
        rptMode = 8
        Viewer.Show 1
      Set rs2 = Nothing
                   If Conn2.State = 1 Then
                       Conn2.Close
                    End If
         
          txtRegNoRelease = ""
         txtRegNoRelease.SetFocus
         Exit Sub
        
         Else
    MsgBox "This Utility is for the Patient who has Released But Want to see his Information", vbCritical, " IT, DNMIH"
            txtRegNoRelease = ""
          txtRegNoRelease.SetFocus
       
               
     End If
          
    Else
        MsgBox "Invalid Registration No", vbInformation, " IT, DNMIH"
         txtRegNoRelease = ""
          txtRegNoRelease.SetFocus
       
        Set rs2 = Nothing
                   If Conn2.State = 1 Then
                       Conn2.Close
                    End If
           Exit Sub
          txtRegNoRelease = ""
          txtRegNoRelease.SetFocus
       
    End If
           
           
     Set rs2 = Nothing
     If Conn2.State = 1 Then
        Conn2.Close
     End If
    Exit Sub
ERR_DESC:
         MsgBox Err.Description, vbCritical, " IT, DNMIH"

End Sub

Private Sub Command1_Click()
   On Error Resume Next
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
         
        rptMode = 7
        Viewer.Show 1
     Set rs2 = Nothing
                   If Conn2.State = 1 Then
                       Conn2.Close
                    End If
         
          txtRegNoRelease = ""
         txtRegNoRelease.SetFocus
         Exit Sub
        
         Else
    MsgBox "This Utility is for the Patient who has Released But Want to see his Information", vbCritical, " IT, DNMIH"

               
     End If
          
    Else
        MsgBox "Invalid Registration No", vbInformation, " IT, DNMIH"
         txtRegNoRelease = ""
          txtRegNoRelease.SetFocus
       
     Set rs2 = Nothing
                   If Conn2.State = 1 Then
                       Conn2.Close
                    End If

           Exit Sub
          txtRegNoRelease = ""
          txtRegNoRelease.SetFocus
       
    End If
           
           
          Set rs2 = Nothing
                   If Conn2.State = 1 Then
                       Conn2.Close
                    End If
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
  End If
End Sub

Private Sub Form_Activate()
txtRegNoRelease = ""
txtRegNoRelease.SetFocus
End Sub
Private Sub get_reg_NO()
  Dim Con As New ADODB.Connection
 Dim RS As New ADODB.Recordset
 Dim cmd As New ADODB.Command
 Con.ConnectionString = strcn.Connection_String
 Con.Open
 cmd.ActiveConnection = Con
 cmd.CommandType = adCmdText
 cmd.CommandText = "select in_reg_no from indoor_pat_money where pat_id=" & txtRecNo & ""
 Set RS = cmd.Execute
 
 If Not RS.EOF Or Not RS.BOF Then
    txtRegNoRelease.Text = RS!in_reg_no
 End If
 
 Set RS = Nothing
 Set cmd = Nothing
 Set Con = Nothing
End Sub


Private Sub Form_Load()
    PopulateFiscalYear
End Sub
Private Sub PopulateFiscalYear()
   Dim yearList() As String
   Dim i As Integer
   yearList = Utility.GetFiscalYears()
   For i = LBound(yearList) To UBound(yearList)
      CBOYRCODE.AddItem yearList(i)
   Next i
   CBOYRCODE.ListIndex = 0
End Sub
Private Sub txtRecNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Len(txtRecNo) = 0 Then
        MsgBox "Please put a valid Receipt no of Released patient", vbInformation
        txtRecNo.SetFocus
     Else
        get_reg_NO
        txtRegNoRelease.SetFocus
     End If
  End If
End Sub

Private Sub txtRegNoRelease_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
If KeyCode = 13 Then
   cmdPrint.SetFocus
End If
End Sub


