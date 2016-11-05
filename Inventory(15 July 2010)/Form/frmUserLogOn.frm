VERSION 5.00
Begin VB.Form frmUserLogOn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log on : DNMIH Inventory Management System"
   ClientHeight    =   2925
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   5400
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmUserLogOn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStep3 
      Height          =   1665
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   750
      Width           =   5400
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   900
         Width           =   2715
      End
      Begin VB.CommandButton cmdView 
         Height          =   285
         Left            =   3840
         Picture         =   "frmUserLogOn.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   375
      End
      Begin VB.TextBox Password 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1560
         MaxLength       =   25
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   540
         Width           =   2685
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   0
         Top             =   210
         Width           =   2265
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Work Area"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   12
         Top             =   930
         Width           =   870
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   420
         TabIndex        =   11
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "User's ID"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   435
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   400
      Left            =   4560
      Picture         =   "frmUserLogOn.frx":0724
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Press to Exit"
      Top             =   2460
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Height          =   400
      Left            =   3720
      Picture         =   "frmUserLogOn.frx":23EE
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Press to Enter"
      Top             =   2460
      Width           =   795
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   8
      Top             =   4680
      Width           =   1425
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "System Date :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   4920
      Width           =   1080
   End
   Begin VB.Label BankName 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Software Programmer, IT Division,DNMIH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   -15
      TabIndex        =   6
      Top             =   495
      Width           =   5430
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   -15
      Picture         =   "frmUserLogOn.frx":40B8
      Stretch         =   -1  'True
      Top             =   -300
      Width           =   5490
   End
End
Attribute VB_Name = "frmUserLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objcom As New DSLComFram.clsCommon
Dim objRs As New ADODB.Recordset
'Private Sub cmdCancel_Click()
'Dim f As New frmBranchLogOn
'Unload Me
'Load f
'f.Show 1
'End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'On Error GoTo Err_Des:
If Len(txtfields(0)) = 0 Then
   MsgBox "Please Put a valid User id", vbInformation, cmp
   txtfields(0).SetFocus
   Exit Sub
 End If

If Len(Password) = 0 Then
   MsgBox "Please Put a valid Password", vbInformation, cmp
   Password.SetFocus
   Exit Sub
End If

If Len(Combo1) = 0 Then
   MsgBox "Category Required..", vbInformation, cmp
   Combo1.SetFocus
   Exit Sub
End If


Set objRs = Nothing
Set objRs = objcom.Get_RS("SELECT UserID, UserPass From UserInfo WHERE UserID ='" & Trim(txtfields(0)) & "' AND UserPass ='" & Trim(Password) & "'and UserStatus =1 ", objmyCon)
If Not objRs.EOF Then
  
    'check menu permission
    '======================
    strAppUser = Trim(txtfields(0))
    'MenuEnabled (strAppUser)
    'Category = "ABC"
    CategoryTitle = "( " + Get_Description(Combo1.Text) + " )"
    CategoryCode = Get_Code(Combo1.Text)
    userid = txtfields(0).Text
    get_userName (userid)
    
    Unload Me
     Load frmmainApp
     frmmainApp.Show
   
Else
    MsgBox "The System Could Not Log On. Make Sure Your UserName and Password are correct.", vbExclamation + vbOKOnly, "Logon Message"
    txtfields(0) = ""
    Password = ""
    txtfields(0).SetFocus
End If
Set objRs = Nothing
Exit Sub
Err_Des:
    MsgBox Err.Description, vbCritical + vbOKOnly, strmsgtitle
End Sub
Private Sub get_userName(id As String)
  Set objRs = objcom.Get_RS("SELECT UserName  from  UserInfo where UserID='" & id & "'", objmyCon)
 If Not objRs.EOF Then
    userName = objRs!userName
 End If
End Sub
Private Sub cmdView_Click()
'On Error GoTo Err_Des:
'Dim frmfindform As New frmFind
'
'Set objRs = objcom.Get_RS("SELECT UserInfo.UserID, EmpInfo.EmpName FROM UserInfo " _
'                        & "INNER JOIN EmpInfo ON UserInfo.EmpCode = EmpInfo.EmpCode " _
'                        & "ORDER BY UserInfo.EmpCode", objmyCon)
'
'Set frmfindform.objFindRS = objRs
'frmfindform.intInputsel = 0
'Set frmfindform.OwnerForm = Me
'frmfindform.Show 1
'txtfields(0).SetFocus
'
'Set objRs = Nothing
'Exit Sub
'Err_Des:
'    MsgBox Err.Description, vbCritical + vbOKOnly, strmsgtitle

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    cmdOK.SetFocus
End If
End Sub

Private Sub Form_Load()
  loag_category

Exit Sub
Err_Des:
    MsgBox Err.Description, vbCritical + vbOKOnly, strmsgtitle

End Sub
Private Sub loag_category()
 Set objRs = objcom.Get_RS("SELECT cate_code,cate_name from item_cate_info order by cate_code ", objmyCon)
 Combo1.Clear
 If Not objRs.EOF Then
    objRs.MoveFirst
    Do Until objRs.EOF
       Combo1.AddItem objRs(1) + "~" + objRs(0)
       objRs.MoveNext
    Loop
 End If
  
End Sub

Private Sub Password_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo Err_Des:

If KeyCode = vbKeyReturn Then
    Combo1.SetFocus
End If

Exit Sub
Err_Des:
    MsgBox Err.Description, vbCritical + vbOKOnly, strmsgtitle

End Sub

Private Sub txtfields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo Err_Des:
If KeyCode = vbKeyReturn Then
    txtfields(0) = Trim(txtfields(0))
    Password.SetFocus
End If

Exit Sub
Err_Des:
    MsgBox Err.Description, vbCritical + vbOKOnly, strmsgtitle
End Sub
