VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   465
      Left            =   4110
      TabIndex        =   13
      Top             =   4830
      Width           =   1995
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   465
      Left            =   2010
      TabIndex        =   11
      Top             =   4800
      Width           =   1995
   End
   Begin VB.ComboBox cmdAccessLevel 
      Height          =   315
      ItemData        =   "FrmUserSecurity.frx":0000
      Left            =   1770
      List            =   "FrmUserSecurity.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4110
      Width           =   2325
   End
   Begin VB.ComboBox cmdDept 
      Height          =   315
      ItemData        =   "FrmUserSecurity.frx":0004
      Left            =   1770
      List            =   "FrmUserSecurity.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3300
      Width           =   2325
   End
   Begin VB.ComboBox cmdAcitveorNot 
      Height          =   315
      ItemData        =   "FrmUserSecurity.frx":0008
      Left            =   1770
      List            =   "FrmUserSecurity.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2520
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   405
      Left            =   1770
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1740
      Width           =   2205
   End
   Begin VB.TextBox txtUserName 
      Height          =   405
      Left            =   1770
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1020
      Width           =   4725
   End
   Begin VB.TextBox txtUserID 
      Height          =   405
      Left            =   1770
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2205
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Access Level"
      Height          =   285
      Left            =   150
      TabIndex        =   12
      Top             =   4140
      Width           =   1305
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   150
      TabIndex        =   9
      Top             =   3390
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Active or Not"
      Height          =   435
      Left            =   60
      TabIndex        =   8
      Top             =   2550
      Width           =   1665
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   405
      Left            =   240
      TabIndex        =   7
      Top             =   1830
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   345
      Left            =   360
      TabIndex        =   6
      Top             =   1170
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Id:"
      Height          =   405
      Left            =   330
      TabIndex        =   1
      Top             =   570
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private classSecurity As New clsSecurity

Private Sub CmdSave_Click()
    classSecurity.UserID = txtUserID
    classSecurity.UserName = txtUserName
    classSecurity.UserActiveOrNot = cmdAcitveorNot
    classSecurity.UserDepartment = cmdDept
    classSecurity.UserPassword = txtPassword
    msg = classSecurity.Save(classSecurity)
    MsgBox msg
End Sub

Private Sub cmdShow_Click()
     mode = 1
     rtpViewer.Show
End Sub

Private Sub Form_Load()

cmdDept.AddItem ("Acc")
cmdDept.AddItem ("Per")

cmdAccessLevel.AddItem ("1")
cmdAccessLevel.AddItem ("2")
cmdAccessLevel.AddItem (3)

End Sub
