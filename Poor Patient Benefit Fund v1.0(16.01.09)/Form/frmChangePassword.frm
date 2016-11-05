VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form13"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   Icon            =   "frmChangePassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form13"
   LockControls    =   -1  'True
   ScaleHeight     =   2055
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Height          =   405
      Left            =   4590
      Picture         =   "frmChangePassword.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   630
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   4590
      Picture         =   "frmChangePassword.frx":1DD4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3195
      Top             =   45
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
   Begin VB.TextBox txtNew_Pass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1350
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1575
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.TextBox txtu_name 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1350
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   915
      Width           =   3060
   End
   Begin VB.TextBox txtu_id 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1350
      MaxLength       =   50
      TabIndex        =   0
      Top             =   585
      Width           =   3060
   End
   Begin VB.TextBox txtOld_pass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   1350
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1245
      Width           =   3060
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   285
      Index           =   4
      Left            =   1320
      Top             =   1560
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   285
      Index           =   3
      Left            =   1320
      Top             =   1230
      Width           =   3120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   285
      Index           =   2
      Left            =   1320
      Top             =   900
      Width           =   3120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   285
      Index           =   1
      Left            =   1320
      Top             =   570
      Width           =   3120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      Height          =   945
      Index           =   0
      Left            =   4530
      Top             =   585
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   405
      TabIndex        =   10
      Top             =   900
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   225
      TabIndex        =   9
      Top             =   1260
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   90
      TabIndex        =   8
      Top             =   135
      Width           =   2145
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   630
      TabIndex        =   7
      Top             =   585
      Width           =   615
   End
   Begin VB.Label lblNew_Pass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   150
      TabIndex        =   6
      Top             =   1575
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strPass As String
Private Sub cmdCANCEL_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    If Len(Trim(txtNew_Pass.Text)) = 0 Then Exit Sub
    Dim Conn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim RS As New ADODB.Recordset
    
    Dim Param1 As New Parameter
    Dim Param2 As New Parameter
    Dim Param3 As New Parameter
    Dim Param4 As New Parameter
    Dim mode As String
    mode = "1"
     
    Conn.Open strcn.Connection_String
    Set cmd.ActiveConnection = Conn
    cmd.CommandType = adCmdText
    '----------------------------------------------------------------------------------
    Set Param1 = cmd.CreateParameter("param1", adVarChar, adParamInput, 10, mode)
    cmd.Parameters.Append Param1
    
    Set Param2 = cmd.CreateParameter("param2", adVarChar, adParamInput, 15, Trim(txtu_id.Text))
    cmd.Parameters.Append Param2

    Set Param3 = cmd.CreateParameter("param3", adVarChar, adParamInput, 20, txtu_name.Text)
    cmd.Parameters.Append Param3
    
    Set Param4 = cmd.CreateParameter("param4", adVarChar, adParamInput, 15, txtNew_Pass.Text)
    cmd.Parameters.Append Param4
    '----------------------------------------------------------------------------------
    cmd.Properties("PLSQLRSet") = True
    cmd.CommandText = "{CALL SaveSecurity(?, ?, ?,?)}"
    Set RS = cmd.Execute
    
    cmd.Properties("PLSQLRSet") = False
    MsgBox "Password changed successfully", vbInformation, "IT Division,DNMIH"
    '==============
    txtu_id.Text = ""
    txtu_name.Text = ""
    txtNew_Pass.Text = ""
    txtOld_pass.Text = ""
    txtu_id.SetFocus
    
    txtNew_Pass.Visible = False
    lblNew_Pass.Visible = False
    Shape1(4).Visible = False
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub
Private Sub txtNew_Pass_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtOld_pass_KeyPress(KeyAscii As Integer)

KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
  If Len(Trim(txtOld_pass.Text)) = 0 Then Exit Sub
    If strPass = Trim(txtOld_pass.Text) Then
       lblNew_Pass.Visible = True
       txtNew_Pass.Visible = True
       Shape1(4).Visible = True
       txtNew_Pass.SetFocus
    Else
       MsgBox "Invalid Password  !!!", vbCritical + vbDefaultButton1, "IT Division,DNMIH"
       lblNew_Pass.Visible = False
       txtNew_Pass.Visible = False
       Shape1(4).Visible = False
       txtOld_pass = ""
       txtOld_pass.SetFocus
    End If
 
End If
End Sub
Private Sub txtu_id_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtu_id_LostFocus()
 On Error GoTo err_desc
    If Len(Trim(txtu_id.Text)) = 0 Then Exit Sub
    Adodc1.ConnectionString = strcn.Connection_String
    Adodc1.RecordSource = "select user_id,pass_word,user_name from security where user_id='" & Trim(txtu_id.Text) & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
       txtu_name.Text = "" & Adodc1.Recordset!user_name
       strPass = Adodc1.Recordset!Pass_word
'       txtOld_pass.SetFocus
    Else
       MsgBox "Invalid User  !!!", vbCritical
       txtu_id.SetFocus
       Exit Sub
    End If
    
   Exit Sub
   
err_desc:
     MsgBox Err.Description, vbInformation, "IT Division,DNMIH"
    
End Sub
