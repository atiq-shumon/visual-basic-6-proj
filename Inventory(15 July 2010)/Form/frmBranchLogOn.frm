VERSION 5.00
Begin VB.Form frmBranchLogOn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect to Database"
   ClientHeight    =   2010
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   4590
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmBranchLogOn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "User DBA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1700
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   400
      Left            =   3600
      Picture         =   "frmBranchLogOn.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1170
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Height          =   400
      Left            =   3600
      Picture         =   "frmBranchLogOn.frx":210C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   800
   End
   Begin VB.Frame fraStep3 
      Height          =   1000
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   615
      Width           =   3330
      Begin VB.TextBox txtFields 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
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
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1200
         MaxLength       =   25
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "dn_inventory"
         Top             =   565
         Width           =   1935
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch ID"
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
         Left            =   195
         TabIndex        =   0
         Top             =   285
         Width           =   825
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
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
         Index           =   3
         Left            =   195
         TabIndex        =   2
         Top             =   615
         Width           =   750
      End
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   60
      Picture         =   "frmBranchLogOn.frx":3DD6
      Top             =   0
      Width           =   4470
   End
End
Attribute VB_Name = "frmBranchLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
'On Error GoTo Err_Des
If Check1.Value = 1 Then
    txtfields(0) = "national_inventory"
    'txtFields(0).Enabled = False
    txtfields(1).Enabled = True
    txtfields(1).SetFocus
Else
    txtfields(0).Enabled = False
    txtfields(1).Enabled = False
    txtfields(0) = strDUPass
    txtfields(1) = strDUPass
    cmdOK.SetFocus
End If
Exit Sub
Err_Des:
    MsgBox Err.Description, vbCritical + vbOKOnly, strmsgtitle
End Sub

Private Sub cmdCancel_Click()
blnBRLogon = False
Unload Me
End
End Sub

Private Sub cmdOK_Click()
strDUser = "national_inventory"
strDUPass = "dn_inventory"
blnBRLogon = True
Unload Me
End Sub

Private Sub Form_Load()

'txtfields(0) = "national_inventory"
'txtfields(1) = "dn_inventory"
  
'    txtfields(0) = strDUser
'    txtfields(1) = strDUPass
'
'    txtfields(0).Enabled = False
'    txtfields(1).Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmBranchLogOn = Nothing
End Sub

Private Sub txtfields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdOK.SetFocus
End Sub
