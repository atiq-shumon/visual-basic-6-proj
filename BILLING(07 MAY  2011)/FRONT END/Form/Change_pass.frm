VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Change_pass 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   -90
      TabIndex        =   18
      Top             =   4740
      Width           =   7875
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Powered by:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00915411&
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   150
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Software Programmer, IT Division, DNMIH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   2190
         TabIndex        =   19
         Top             =   120
         Width           =   5040
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   6120
      TabIndex        =   17
      Top             =   4110
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   4890
      TabIndex        =   16
      Top             =   4110
      Width           =   1215
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "NEW"
      Height          =   375
      Left            =   3660
      TabIndex        =   11
      Top             =   4110
      Width           =   1215
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   2430
      TabIndex        =   10
      Top             =   4110
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   480
      Top             =   3810
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   390
      Top             =   4020
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
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
   Begin VB.Frame Change_password 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   0
      TabIndex        =   2
      Top             =   2580
      Width           =   7785
      Begin VB.TextBox txtCpass 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2025
         TabIndex        =   7
         Top             =   780
         Width           =   4455
      End
      Begin VB.TextBox txtpass 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2010
         TabIndex        =   6
         Top             =   360
         Width           =   4485
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   203
         TabIndex        =   4
         Top             =   450
         Width           =   1635
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   225
         TabIndex        =   3
         Top             =   930
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   -30
      TabIndex        =   0
      Top             =   -30
      Width           =   7545
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   675
         Left            =   0
         TabIndex        =   14
         Top             =   30
         Width           =   7755
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CHANGE PASSWORD UTILITY"
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
            Left            =   1320
            TabIndex        =   15
            Top             =   150
            Width           =   4755
         End
         Begin VB.Image Image1 
            Height          =   855
            Left            =   -3180
            Picture         =   "Change_pass.frx":0000
            Top             =   -150
            Width           =   11820
         End
      End
      Begin VB.TextBox txtpassword 
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
         ForeColor       =   &H000000FF&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2070
         PasswordChar    =   "?"
         TabIndex        =   12
         Top             =   2040
         Width           =   4425
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   2070
         TabIndex        =   9
         Top             =   1050
         Width           =   4425
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         ForeColor       =   &H80000001&
         Height          =   345
         Left            =   2070
         TabIndex        =   5
         Top             =   1530
         Width           =   4425
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Password:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   13
         Top             =   2100
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Id:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   1110
         Width           =   705
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   1590
         Width           =   705
      End
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Left            =   2370
      Top             =   4050
      Width           =   4995
   End
End
Attribute VB_Name = "Change_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Con As New MyConnection
Dim Conn2 As New Connection
Dim rs2 As New Recordset
Dim cmd As New Command


Private Sub cmdADD_Click()
txtpass = ""
txtCpass = ""
txtpass.SetFocus
End Sub

Private Sub cmdExit_Click()
Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

End Sub


Private Sub cmdSave_Click()
Dim validation As Variant
              Adodc1.ConnectionString = strcn.Connection_String
              Adodc1.RecordSource = "Select user_id, user_name, user_password,user_type,shift_name From Security Where (User_Id = '" & frmMAIN.lbluser_id & "')"
              Adodc1.Refresh
              validation = Adodc1.Recordset!user_id
                
                Dim Conn As New ADODB.Connection
                Dim cmd As New ADODB.Command
                Dim RS As New ADODB.Recordset

                        Dim Param1 As New Parameter
                 If Conn.State = 0 Then
                        Conn.Open strcn.Connection_String
                 End If
    
                    Set cmd.ActiveConnection = Conn
                    cmd.CommandType = adCmdText
    
                   Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 5, validation)
                    cmd.Parameters.Append Param1 'validation
                    cmd.Properties("PLSQLRSet") = True
    
                     cmd.CommandText = "{CALL shift_validation(?)}"
    
                Debug.Print cmd.CommandText
    
                    Set RS = cmd.Execute
    

                cmd.Properties("PLSQLRSet") = False
                
          Adodc2.ConnectionString = strcn.Connection_String
          Adodc2.RecordSource = "Select * From user_validation"
          Adodc2.Refresh
        

        
             If Adodc2.Recordset!validation = 0 Then
             MsgBox "Your Working Time has been Expired", vbInformation, " IT, DNMIH."
             Exit Sub
             End If
             




If Text1 = "" Then
MsgBox "User id Required", vbInformation, " IT, DNMIH"
Text1.SetFocus
Exit Sub
End If
If txtpassword = "" Then
MsgBox "User Previous password Required", vbInformation, " IT, DNMIH"
txtpassword.SetFocus
Exit Sub
End If
If txtpass = "" Then
MsgBox "User New password Required", vbInformation, " IT, DNMIH"
txtpass.SetFocus
Exit Sub
End If

If txtCpass = "" Then
MsgBox "User New Confirm password Required", vbInformation, " IT, DNMIH"
txtCpass.SetFocus
Exit Sub
End If


Call save_change_password

MsgBox "Operation successfull", vbInformation + vbOKOnly, "Save..."
If Conn.State = 1 Then
   Conn.Close
End If
End Sub

Private Sub save_change_password()
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset

    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim Param5 As New Parameter
 If Conn.State = 0 Then
    Conn.Open strcn.Connection_String
 End If
    
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 30, Me.Text1)
    cmd.Parameters.Append Param1 'user_id name
    
   
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 50, txtpassword)
    cmd.Parameters.Append Param2 'previous_password
    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 30, Me.txtpass)
    cmd.Parameters.Append Param3 'pass name
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 30, Me.txtCpass)
    cmd.Parameters.Append Param4 'user_c pass

    
    cmd.Properties("PLSQLRSet") = True
    
    cmd.CommandText = "{CALL change_password(?,?,?,?)}"
    
    Debug.Print cmd.CommandText
    
    Set rs2 = cmd.Execute
    

    cmd.Properties("PLSQLRSet") = False
   If Conn.State = 1 Then
      Conn.Close
   End If
End Sub

Private Sub Form_Load()
txtpass.BackColor = &H808080
txtCpass.BackColor = &H808080
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(Text1) <> 0 Then
            txtpassword.SetFocus
        Else
            MsgBox "Please enter any User Id.", vbInformation, " IT, DNMIH"
        End If
    End If
End Sub

Private Sub txtCpass_GotFocus()
txtCpass.BackColor = &HFFFFFF
End Sub

Private Sub txtCpass_LostFocus()
txtCpass.BackColor = &H808080
End Sub

Private Sub txtpass_GotFocus()
txtpass.BackColor = &HFFFFFF
End Sub

Private Sub txtpass_LostFocus()
txtpass.BackColor = &H808080


End Sub

Private Sub txtPassWord_KeyPress(KeyAscii As Integer)
If Text1 = "" Then
    MsgBox "User Id Required", vbCritical, "Warning"
    Text1.SetFocus
    Exit Sub
Else
    If KeyAscii = 13 Then
        Adodc1.ConnectionString = strcn.Connection_String
        Adodc1.RecordSource = "Select user_id, user_name, user_password From Security Where (User_Id = '" & Text1 & "')"
        Adodc1.Refresh

        
        If Adodc1.Recordset.EOF = True Then
            MsgBox "No such ID exists.", vbCritical, "Warning"
            Text1 = ""
            txtpassword = ""
            Text1.SetFocus
            Exit Sub
        Else
            If txtpassword = Adodc1.Recordset!user_password Then
            Text3 = Adodc1.Recordset!user_name
            txtpass.BackColor = &HFFFFFF
            
            txtCpass.BackColor = &HFFFFFF
            
            txtpass.Enabled = True
            txtCpass.Enabled = True
            txtpass.SetFocus
            Else
                MsgBox "Incorrect Password", vbCritical, "Warning"
            End If


        End If
    End If
End If

End Sub
