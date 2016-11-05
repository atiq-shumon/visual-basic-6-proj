VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form92 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2625
   ClientLeft      =   2205
   ClientTop       =   4815
   ClientWidth     =   8460
   Icon            =   "frmChange_Prev_Password.frx":0000
   LinkTopic       =   "Form36"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   8460
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   6435
      ScaleHeight     =   195
      ScaleWidth      =   780
      TabIndex        =   14
      Top             =   1395
      Width           =   780
      Begin VB.CommandButton cmdApply 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   -45
         Picture         =   "frmChange_Prev_Password.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   -45
         Width           =   870
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   5625
      ScaleHeight     =   195
      ScaleWidth      =   780
      TabIndex        =   11
      Top             =   1395
      Width           =   780
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   -45
         Picture         =   "frmChange_Prev_Password.frx":162C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   -45
         Width           =   870
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   7245
      ScaleHeight     =   225
      ScaleWidth      =   750
      TabIndex        =   10
      Top             =   1395
      Width           =   750
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   -180
         Picture         =   "frmChange_Prev_Password.frx":238E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   -45
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7695
      Top             =   -45
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6435
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
   Begin VB.TextBox txtConf_Pass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   5670
      PasswordChar    =   "#"
      TabIndex        =   3
      Top             =   990
      Width           =   2340
   End
   Begin VB.TextBox txtNew_Pass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   5670
      PasswordChar    =   "#"
      TabIndex        =   2
      Top             =   585
      Width           =   2340
   End
   Begin VB.TextBox txtOld_Pass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   1395
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   1365
      Width           =   2430
   End
   Begin VB.TextBox txtU_Id 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1395
      TabIndex        =   0
      Top             =   585
      Width           =   2430
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00E9C2C2&
      Height          =   285
      Index           =   7
      Left            =   5580
      Top             =   1350
      Width           =   2445
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1395
      TabIndex        =   9
      Top             =   990
      Width           =   2400
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FAEDF1&
      BorderColor     =   &H00FDD9E8&
      Height          =   285
      Index           =   4
      Left            =   5580
      Top             =   945
      Width           =   2445
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4095
      TabIndex        =   8
      Top             =   990
      Width           =   1350
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4095
      TabIndex        =   7
      Top             =   540
      Width           =   1140
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   405
      TabIndex        =   6
      Top             =   1395
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FAEDF1&
      BorderColor     =   &H00FDD9E8&
      Height          =   285
      Index           =   3
      Left            =   1305
      Top             =   1335
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FAEDF1&
      BorderColor     =   &H00FDD9E8&
      Height          =   285
      Index           =   2
      Left            =   5580
      Top             =   540
      Width           =   2445
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   405
      TabIndex        =   5
      Top             =   990
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   405
      TabIndex        =   4
      Top             =   540
      Width           =   525
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FAEDF1&
      BorderColor     =   &H00FDD9E8&
      Height          =   285
      Index           =   0
      Left            =   1305
      Top             =   540
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FAEDF1&
      BorderColor     =   &H00FDD9E8&
      Height          =   285
      Index           =   1
      Left            =   1305
      Top             =   945
      Width           =   2535
   End
End
Attribute VB_Name = "Form92"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''
''Dim PW_Len_Min_Flag   As Integer
''Dim PW_Len_Max_Flag   As Integer
''
''Dim PW_Len_Min   As Integer
''Dim PW_Len_Max   As Integer
'''-------------------------------------
''Dim Password     As String
''Dim Passlen      As Integer
''Dim Passtot      As Integer
''Dim Passnum      As Double
''Dim Finalpass    As String
''Dim Flen         As Integer
''Dim mssg         As String
''Dim emp          As String
''
''''**************************
''
''Dim newlen       As Integer
''Dim npass        As String
''Dim Conflen      As Integer
''Dim cpass        As String
''Dim newtot       As Integer
''Dim newpass      As Integer
''Dim conftot      As Integer
''Dim Confpass     As Integer
''Dim myrst        As New ADODB.Recordset
''Dim Rs           As New Recordset
''
''
''Private Sub cmdCancel_Click()
''    txtConf_Pass = ""
''    txtNew_Pass = ""
''    txtOld_Pass = ""
''    txtU_Id = ""
''    lblName = ""
''    txtOld_Pass.Locked = False
''    txtU_Id.Locked = False
''    txtU_Id.SetFocus
''End Sub
''
''Private Sub cmdClose_Click()
''    txtConf_Pass = ""
''    txtNew_Pass = ""
''    txtOld_Pass = ""
''    txtU_Id = ""
''    lblName = ""
''    txtOld_Pass.Locked = False
''    txtU_Id.Locked = False
''Unload Me
''End Sub
''
''Private Sub Image2_Click()
''Unload Me
''End Sub
''
''Private Sub lblName_Change()
''txtOld_Pass = ""
''If lblName <> "" Then
''    Timer1.Enabled = True
''End If
''End Sub
''
''
''Private Sub Timer1_Timer()
''txtOld_Pass.SetFocus
''Timer1.Enabled = False
''End Sub
''
''Private Sub txtOld_Pass_KeyPress(KeyAscii As Integer)
''If KeyAscii = 13 Then
'''----------------------------------------------------------------
''''Validate Old Password Prior to Change it.
'''----------------------------------------------------------------
''    If txtU_Id = Empty Or txtOld_Pass = Empty Then
''        mssg = MsgBox("Must enter current password.", _
''        vbOKOnly + vbExclamation, "Confirmation")
''    Else
''        Password = LTrim(RTrim(txtOld_Pass))
''        Passlen = Len(Password)
''        Passtot = 0
''
''        Select Case Passlen
''
''                Case 1
''                        Passtot = Passtot + Asc(Password)
''                        Passnum = Val(Mid("12345678901234123456789123456789", 15, 9)) + Passtot
''                        Finalpass = "12345678901234" + LTrim(RTrim(CStr(Passnum))) + "123456789"
''                Case 2
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Mid("1234567812345", 9, 5)) + Passtot
''                        Finalpass = "12345678" + LTrim(RTrim(CStr(Passnum)))
''                Case 3
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Left("12312123456", 3)) + Passtot
''                        Finalpass = LTrim(RTrim(CStr(Passnum))) + "12123456"
''                Case 4
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Mid("123123456112345", 4, 6)) + Passtot
''                        Finalpass = "123" + LTrim(RTrim(CStr(Passnum))) + "112345"
''                Case 5
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Mid("123123456781234567890", 4, 8)) + Passtot
''                        Finalpass = "123" + LTrim(RTrim(CStr(Passnum))) + "1234567890"
''                Case 6
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Left("123456123456789", 6)) + Passtot
''                        Finalpass = LTrim(RTrim(CStr(Passnum))) + "123456789"
''                Case 7
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Mid("1234561212345", 9, 5)) + Passtot
''                        Finalpass = "12345612" + LTrim(RTrim(CStr(Passnum)))
''                Case 8
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Mid("12345612345123456", 7, 5)) + Passtot
''                        Finalpass = "123456" + LTrim(RTrim(CStr(Passnum))) + "123456"
''                Case 9
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Mid("123456781231234561234", 12, 6)) + Passtot
''                        Finalpass = "12345678123" + LTrim(RTrim(CStr(Passnum))) + "1234"
''                Case 10
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Mid("12345678912312345678", 13, 8)) + Passtot
''                        Finalpass = "123456789123" + LTrim(RTrim(CStr(Passnum)))
''                Case 11
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Left("123456123456712345", 6)) + Passtot
''                        Finalpass = LTrim(RTrim(CStr(Passnum))) + "123456712345"
''                Case 12
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Mid("1234512345612345", 12, 5)) + Passtot
''                        Finalpass = "12345123456" + LTrim(RTrim(CStr(Passnum)))
''                Case 13
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Left("12345123456123456", 5)) + Passtot
''                        Finalpass = LTrim(RTrim(CStr(Passnum))) + "123456123456"
''                Case 14
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Left("123451234", 5)) + Passtot
''                        Finalpass = LTrim(RTrim(CStr(Passnum))) + "1234"
''                Case 15
''                        For Flen = 1 To Passlen
''                            Passtot = Passtot + Asc(Mid(Password, Flen, 1))
''                        Next
''                        Passnum = Val(Mid("123123456789012312", 4, 10)) + Passtot
''                        Finalpass = "123" + LTrim(RTrim(CStr(Passnum))) + "12312"
''        End Select
''
''
''        Adodc1.ConnectionString = strCN.Connection
''        emp = "select user_pass from Soft_Pass where u_id='" + RTrim(txtU_Id) + "' and cancel='0'"
''        Adodc1.RecordSource = emp
''        Adodc1.Refresh
''        If Adodc1.Recordset.EOF = False Then
''            If IsNull(Adodc1.Recordset!User_Pass) Then
''                mssg = MsgBox("Password not found, please call <Database administrator>.", _
''                vbOKOnly + vbExclamation, "Security Manager")
''                If mssg = vbOK Then
''                End If
''            Else
''                If RTrim(Adodc1.Recordset!User_Pass) = RTrim(Finalpass) Then
''                    'u_id = LTrim(RTrim(txtU_Id))
''                   ' upass = RTrim(Finalpass)
''                ''*************************************************
''                ''If Old Password validtion done successfully then
''                ''allow user to set New password
''                    txtU_Id.Locked = True
''                    txtOld_Pass.Locked = True
''                    txtNew_Pass.SetFocus
''                ''*************************************************
''                Else
''
''                    MsgBox "The password that you typed is not correct,Try typing it again.", vbOKOnly + vbCritical, "Security Manager"
''                    txtOld_Pass.SetFocus
''                    txtOld_Pass.SelStart = 0
''                    txtOld_Pass.SelLength = Len(txtOld_Pass)
''
''                End If
''            End If
''        Else
''
''           MsgBox "Invalid employee ID and Password.", vbOKOnly + vbExclamation, "Security Manager"
''           txtU_Id = ""
''           txtOld_Pass = ""
''           txtU_Id.SetFocus
''        End If
''    End If
''End If
''End Sub
''Private Sub txtU_Id_KeyPress(KeyAscii As Integer)
''If txtU_Id <> "" And lblName <> "" And KeyAscii = 13 Then
''    txtOld_Pass.SetFocus
''End If
''End Sub
''Private Sub Form_Load()
''
'''PW_Len_Min_Flag = Get_Param_Flag(14)
'''PW_Len_Max_Flag = Get_Param_Flag(15)
'''
'''PW_Len_Min = Get_Param_Value(14)
'''PW_Len_Max = Get_Param_Value(15)
'''
'''txtU_Id = U_Id
''End Sub
''Private Sub txtU_Id_Change()
''
''Adodc1.ConnectionString = strCN.Connection
''Adodc1.RecordSource = "exec POP_UserName_SPrivilege '" + txtU_Id + "'"
''Adodc1.Refresh
''If Adodc1.Recordset.RecordCount > 0 Then
''    lblName = Adodc1.Recordset!U_Name
''    Adodc1.Refresh
''Else
''    lblName = ""
''End If
''
''End Sub
''
''Private Sub txtConf_Pass_KeyPress(KeyAscii As Integer)
''If KeyAscii = 13 Then
''    If Trim(txtConf_Pass) <> Trim(txtNew_Pass) Then
''        MsgBox "Password confirmation does not match.", vbOKOnly + vbExclamation, "Confirmatiom"
''        txtConf_Pass.SetFocus
''        txtConf_Pass.SelStart = 0
''        txtConf_Pass.SelLength = Len(txtConf_Pass)
''        Exit Sub
''    Else
''        cmdApply.SetFocus
''    End If
''End If
''End Sub
''
''Private Sub cmdApply_Click()
''
''    If txtConf_Pass = Empty Then
''        MsgBox "Information Incomplete.", vbOKOnly + vbExclamation, "Confirmation"
''    Else
''        Conflen = Len(LTrim(RTrim(txtConf_Pass)))
''        conftot = 0
''        Select Case Conflen
''                Case 1
''                    conftot = conftot + Asc(txtConf_Pass)
''                    cpass = "12345678901234" + _
''                    LTrim(RTrim(CStr(123456789 + conftot))) + "123456789"
''                Case 2
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = "12345678" + LTrim(RTrim(CStr(12345 + conftot)))
''                Case 3
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = LTrim(RTrim(CStr(123 + conftot))) + "12123456"
''                 Case 4
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = "123" + LTrim(RTrim(CStr(123456 + conftot))) + "112345"
''                 Case 5
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = "123" + LTrim(RTrim(CStr(12345678 + conftot))) + "1234567890"
''                 Case 6
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = LTrim(RTrim(CStr(123456 + conftot))) + "123456789"
''                Case 7
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = "12345612" + LTrim(RTrim(CStr(12345 + conftot)))
''                Case 8
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = "123456" + LTrim(RTrim(CStr(12345 + conftot))) + "123456"
''                Case 9
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = "12345678123" + LTrim(RTrim(CStr(123456 + conftot))) + "1234"
''                Case 10
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = "123456789123" + LTrim(RTrim(CStr(12345678 + conftot)))
''                Case 11
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = LTrim(RTrim(CStr(123456 + conftot))) + "123456712345"
''                Case 12
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = "12345123456" + LTrim(RTrim(CStr(12345 + conftot)))
''                Case 13
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = LTrim(RTrim(CStr(12345 + conftot))) + "123456123456"
''                Case 14
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = LTrim(RTrim(CStr(12345 + conftot))) + "1234"
''                Case 15
''                    For Confpass = 1 To Conflen
''                        conftot = conftot + _
''                        Asc(Mid(LTrim(RTrim(txtConf_Pass)), Confpass, 1))
''                    Next
''                    cpass = "123" + LTrim(RTrim(CStr(1234567890 + conftot))) + "12312"
''        End Select
''
''            con.ConnectionString = strCN.Connection
''            con.Open
''            Set cmd.ActiveConnection = con
''
''                cmd.CommandText = "exec pro_soft_pass'" _
''                + Trim(txtU_Id) + "','" _
''                + Trim(lblName) + "','" + cpass + "','" _
''                + U_Id + "',0,'C'"
''
''
''
''
''              Set Rs = cmd.Execute
''
''            MsgBox Rs!Message, vbExclamation + vbOKOnly, "Password"
''
''            con.Close
''
''            txtNew_Pass = ""
''            txtConf_Pass = ""
''
''            cmdClose.SetFocus
''
''    End If
''End Sub
''
''
''Private Sub txtU_IdChange()
''
'''txtU_Id = Carry_Emp_ID
''
''Adodc1(0).ConnectionString = strCN.Connection
''Adodc1(0).RecordSource = "exec sp_check_id '" + Trim(txtU_Id) + "'"
''
''Adodc1(0).Refresh
''If Adodc1(0).Recordset.RecordCount > 0 Then
''    lblName = Adodc1(0).Recordset!nm
''
''End If
''
''End Sub
''
''Private Sub txtNew_Pass_KeyPress(KeyAscii As Integer)
''If KeyAscii = 13 Then
''    If txtNew_Pass = Empty Then
''
''        MsgBox "Blank not allowed.", vbOKOnly + vbExclamation, "Confirmation"
''
''    ElseIf PW_Len_Min_Flag = 1 And Len(Trim(txtNew_Pass)) < PW_Len_Min Then
''
''        MsgBox "Password must be larger than or equal to " & CStr(PW_Len_Min) & " characters !", vbOKOnly + vbExclamation, "Confirmation"
''
''    ElseIf PW_Len_Max_Flag = 1 And Len(Trim(txtNew_Pass)) > PW_Len_Max Then
''
''        MsgBox "Password must be smaller than or equal to " & CStr(PW_Len_Max) & " characters !", vbOKOnly + vbExclamation, "Confirmation"
''
''    Else
''        txtConf_Pass.Enabled = True
''        txtConf_Pass.SetFocus
''    End If
''End If
''End Sub
''
''
