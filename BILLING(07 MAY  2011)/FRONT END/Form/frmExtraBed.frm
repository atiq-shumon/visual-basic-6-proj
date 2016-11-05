VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmExtraBed 
   Appearance      =   0  'Flat
   BackColor       =   &H80000010&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extra Bed Booking"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4800
   FillColor       =   &H008080FF&
   ForeColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmExtraBed.frx":0000
   ScaleHeight     =   1095
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CBOYRCODE 
      Height          =   315
      ItemData        =   "frmExtraBed.frx":7229
      Left            =   1380
      List            =   "frmExtraBed.frx":7233
      TabIndex        =   3
      Text            =   "YR-0910"
      Top             =   150
      Width           =   2835
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -690
      Top             =   1080
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
      Height          =   360
      Left            =   4290
      TabIndex        =   2
      Top             =   600
      Width           =   465
   End
   Begin VB.TextBox txtRegNoExtraBed 
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
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   1350
      TabIndex        =   1
      Top             =   600
      Width           =   2925
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
      Top             =   210
      Width           =   1200
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " REG NO :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   0
      Top             =   675
      Width           =   885
   End
End
Attribute VB_Name = "frmExtraBed"
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
      Adodc1.ConnectionString = strcn.Connection_String
        Adodc1.RecordSource = "select a.release_flag as release_flag From in_door_pat_info_main a  Where  a.in_reg_no ='" & Trim(txtRegNoExtraBed.Text) & "' AND A.YRCODE='" & Trim(CBOYRCODE) & "'"
        Adodc1.Refresh
      
       If Adodc1.Recordset.RecordCount > 0 Then
            If Adodc1.Recordset!release_flag = "2" Then
               MsgBox "This Patient has been fled     ", vbInformation, "Warning: IT, DNMIH"
               txtRegNoExtraBed = ""
               txtRegNoExtraBed.SetFocus
               Exit Sub
            End If
     End If
     
     If Conn2.State = 0 Then
        Conn2.ConnectionString = strcn.Connection_String
        Conn2.Open
     End If
        cmd.ActiveConnection = Conn2
        cmd.CommandType = adCmdText
        cmd.CommandText = "select release_flag From in_door_pat_info_main Where in_reg_no ='" & Trim(txtRegNoExtraBed.Text) & "' AND YRCODE='" & Trim(CBOYRCODE) & "'"
      
       cmd.Properties("iRowsetChange") = True
        cmd.Properties("updatability") = 7
       If rs2.State = 0 Then
        rs2.CursorLocation = adUseClient

        rs2.Open cmd.CommandText, Conn2, adOpenDynamic, adLockOptimistic
      End If
      cmd.Properties("iRowsetChange") = False
         If rs2.RecordCount > 0 Then
                If (rs2!release_flag) = "1" Or (rs2!release_flag) = "3" Then
                     MsgBox "This Patient has been Released", vbInformation, "Warning: IT, DNMIH"
                     ' rs2.Close
                     If Conn2.State = 1 Then
                          Conn2.Close
                      End If
                       Exit Sub
         Else
               Adodc1.ConnectionString = strcn.Connection_String
                  Adodc1.RecordSource = "select bed_type from bed_info  where upper(in_reg_no)='" & Trim(frmExtraBed.txtRegNoExtraBed.Text) & "' and temp_serial_no=(select max(temp_serial_no) from bed_info where upper(in_reg_no)='" & Trim(frmExtraBed.txtRegNoExtraBed.Text) & "') "
                    Adodc1.Refresh
             If Adodc1.Recordset.RecordCount > 0 Then
              If UCase(Adodc1.Recordset!Bed_type) = UCase("Cabin") Then
                    frmExtraBedMain.Show vbModal
                        'rs2.Close
                 If Conn2.State = 1 Then
                    Conn2.Close
                End If
           End If
        Else
        MsgBox "This Id is Not a Cabin Patient", vbInformation, " IT, DNMIH."
        End If
       End If
    Else
        MsgBox "Invalid Registration No", vbInformation, "Warning: IT, DNMIH"
         Set rs2 = Nothing
        If Conn2.State = 1 Then
            Conn2.Close
        End If
        txtRegNoExtraBed = ""
         txtRegNoExtraBed.SetFocus

        Exit Sub
        
    End If
'  rs2.Close
'  Conn2.Close
'  txtRegNoExtraBed = ""
'  txtRegNoExtraBed.SetFocus

'Call temp_calculation
   'cmd.Properties("iRowsetChange") = False


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
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, txtRegNoExtraBed)
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
    txtRegNoExtraBed = ""
     txtRegNoExtraBed.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       Unload Me
     End If

End Sub

Private Sub txtRegNoExtraBed_Change()
If Not IsNumeric(txtRegNoExtraBed.Text) Then
            txtRegNoExtraBed = ""
End If
End Sub

Private Sub txtRegNoExtraBed_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then
        Unload Me
   End If
End Sub

Private Sub txtRegNoExtraBed_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        Command1_Click
   End If
End Sub
