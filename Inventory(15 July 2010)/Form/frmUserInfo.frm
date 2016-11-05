VERSION 5.00
Begin VB.Form frmUserInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Creation"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H80000001&
      Height          =   675
      Left            =   -30
      TabIndex        =   27
      Top             =   5730
      Width           =   8145
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   375
         Left            =   6870
         TabIndex        =   31
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "New"
         Height          =   375
         Left            =   5760
         TabIndex        =   30
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   4680
         TabIndex        =   29
         Top             =   180
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   4620
         Top             =   120
         Width           =   3345
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000001&
      Height          =   795
      Left            =   -30
      TabIndex        =   26
      Top             =   -60
      Width           =   8235
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User Creation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   390
         Left            =   3180
         TabIndex        =   28
         Top             =   210
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "User Status"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   3930
      TabIndex        =   24
      Top             =   3540
      Width           =   4170
      Begin VB.OptionButton opt 
         Caption         =   "Continue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   465
         TabIndex        =   9
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opt 
         Caption         =   "Closed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox txtFields 
      Height          =   675
      Index           =   8
      Left            =   1680
      MaxLength       =   75
      TabIndex        =   11
      Top             =   4680
      Width           =   5925
   End
   Begin VB.Frame Frame2 
      Caption         =   "User Password"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   0
      TabIndex        =   20
      Top             =   3540
      Width           =   3915
      Begin VB.TextBox txtFields 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   1680
         MaxLength       =   75
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   540
         Width           =   1815
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   1680
         MaxLength       =   75
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   540
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Information"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2775
      Left            =   -30
      TabIndex        =   12
      Top             =   750
      Width           =   8130
      Begin VB.CommandButton cmdSearch 
         Height          =   300
         Left            =   2700
         Picture         =   "frmUserInfo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   420
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1620
         Width           =   2595
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   5
         Left            =   1440
         MaxLength       =   75
         TabIndex        =   6
         Top             =   2130
         Width           =   2565
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   0
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   4
         Left            =   5250
         MaxLength       =   75
         TabIndex        =   5
         Top             =   2100
         Width           =   2385
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   3
         Left            =   5280
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1620
         Width           =   2355
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1200
         Width           =   6165
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   1
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   1
         Top             =   750
         Width           =   6165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   255
         TabIndex        =   19
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4785
         TabIndex        =   18
         Top             =   2100
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4605
         TabIndex        =   17
         Top             =   1620
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   225
         TabIndex        =   16
         Top             =   1590
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   225
         TabIndex        =   15
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   225
         TabIndex        =   14
         Top             =   750
         Width           =   405
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   1
         Left            =   225
         TabIndex        =   13
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   270
      TabIndex        =   23
      Top             =   4740
      Width           =   630
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon

Private Sub cmbCategory_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtfields(3).SetFocus
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
For i = 0 To 8
    txtfields(i) = ""
Next
txtfields(1).SetFocus
End Sub

Private Sub cmdSAVE_Click()
Dim cmd As New Command
Dim RS As New Recordset
Dim userid As String
Set RS = objcom.Get_RS("Select Max(UserID)+1 From UserInfo where UserCategory<>1", objmyCon)
If Not RS.EOF Then
    userid = RS(0)
Else
    userid = "00001"
End If

    If txtfields(1) = "" Then
        MsgBox "User Name Required", vbInformation, "IT Division, DNMIH."
        txtfields(1).SetFocus
        Exit Sub
    End If
    If txtfields(6) = "" Then
        MsgBox "User Password Required", vbInformation, "IT Division, DNMIH."
        txtfields(6).SetFocus
        Exit Sub
    End If
    
    If txtfields(6) <> txtfields(7) Then
        MsgBox "Pls. confirm password", vbInformation, "IT Division, DNMIH."
        txtfields(7).SetFocus
        Exit Sub
    End If
    
Set cmd.ActiveConnection = objmyCon
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "UserInfo_Save"
    If Trim(txtfields(0)) = "" Then
        cmd(1) = Format(userid, "00000")
    Else
        cmd(1) = Format(Trim(txtfields(0)), "00000")
    End If
    cmd(2) = Trim(txtfields(1))
    cmd(3) = Trim(txtfields(2))
    cmd(4) = Mid(Trim(cmbCategory), 1, 1)
    cmd(5) = Trim(txtfields(3))
    cmd(6) = Trim(txtfields(4))
    cmd(7) = Trim(txtfields(5))
    cmd(8) = Trim(txtfields(6))
    If opt(0).value = True Then
        cmd(9) = 1
    Else
        cmd(9) = 0
    End If
    cmd(10) = Trim(txtfields(8))
    
    cmd.Execute
    
    If Trim(txtfields(0)) = "" Then
        MsgBox "Your created user ID is " & Format(userid, "00000"), vbInformation, strmsgtitle
        txtfields(0) = Format(userid, "00000")
    Else
        MsgBox "Your edited user ID is " & Trim(txtfields(0)), vbInformation, strmsgtitle
    End If
End Sub

Private Sub cmdSearch_Click()
'Dim objRs As New ADODB.Recordset
'
'txtfields(0).Enabled = True
'Set objRs = objcom.Get_RS("select UserID as [User ID],UserName as [User Name] from UserInfo", objmyCon)
'Dim f As New frmFind
'Set f.OwnerForm = Me
'Set f.objFindRS = objRs
'    f.Caption = "User Search"
'    f.intInputsel = 0
'    f.Show 1
'    txtfields(0).SetFocus
'    Set objRs = Nothing
End Sub

Private Sub Command3_Click()
 Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys (Chr(9))
  End If
  If KeyAscii = 27 Then
     Unload Me
  End If
End Sub

Private Sub Form_Load()
Dim RS As New Recordset
Set RS = objcom.Get_RS("SELECT UserCategoryID, UserCategoryName From UserCategory " + _
        "WHERE (UserCategoryID <> 1)", objmyCon)
If Not RS.EOF Then
    Do Until RS.EOF
        cmbCategory.AddItem RS(0) & " - " & RS(1)
        RS.MoveNext
    Loop
    If cmbCategory.ListCount > 0 Then cmbCategory.ListIndex = 0
End If
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 0
            txtfields(8).SetFocus
        Case 1
            txtfields(8).SetFocus
    End Select
End If
End Sub

Private Sub txtfields_GotFocus(Index As Integer)
txtfields(Index).SelStart = 0
txtfields(Index).SelLength = Len(Trim(txtfields(Index)))
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 8
            cmdsave.SetFocus
            
        Case Index
            If Index <> 2 Then
                txtfields(Index + 1).SetFocus
            Else
                cmbCategory.SetFocus
            End If

        Case 3
            If Index <> 6 Then
                txtfields(Index + 1).SetFocus
            Else
                opt(0).SetFocus
            End If
    End Select
End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Dim RS As New Recordset
Select Case Index
    Case 0
        If Trim(txtfields(0)) = "" Then
            Exit Sub
        Else
            txtfields(0) = Format(Trim(txtfields(0)), "00000")
        End If
        Set RS = objcom.Get_RS("Select *from UserInfo Where UserID = '" & Trim(txtfields(0)) & "'", objmyCon)
        If Not RS.EOF Then
            txtfields(1) = "" & RS.Fields("UserName")
            txtfields(2) = "" & RS.Fields("Address")
            
            Dim CategoryDesc As New Recordset
            Set CategoryDesc = objcom.Get_RS("SELECT UserCategoryName From UserCategory " + _
                    "WHERE (UserCategoryID = " & RS.Fields("UserCategory") & " )", objmyCon)
            If Not CategoryDesc.EOF Then
                cmbCategory.Text = RS.Fields("UserCategory") & " - " & CategoryDesc.Fields("UserCategoryName")
            End If
            
            txtfields(3) = "" & RS.Fields("Phone")
            txtfields(4) = "" & RS.Fields("Fax")
            txtfields(5) = "" & RS.Fields("EMail")
            txtfields(6) = "" & RS.Fields("UserPass")
            txtfields(7) = "" & RS.Fields("UserPass")
            
            If RS.Fields("UserStatus") = True Then
                opt(0).value = True
                opt(1).value = False
            Else
                opt(0).value = False
                opt(1).value = True
            End If
            
            txtfields(8) = "" & RS.Fields("Remarks")
        End If
        Set RS = objcom.Get_RS("Select *from UserInfo where UserID = '" & Trim(txtfields(0)) & "'", objmyCon)
End Select
End Sub
