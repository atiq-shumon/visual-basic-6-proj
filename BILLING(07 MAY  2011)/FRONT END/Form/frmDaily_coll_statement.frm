VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDaily_collection_statement 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   2970
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   -30
      TabIndex        =   3
      Top             =   840
      Width           =   5715
      Begin VB.CheckBox Chk_date_to_date 
         Appearance      =   0  'Flat
         Caption         =   "DAY TO DAY COLLECTION"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2910
         TabIndex        =   13
         Top             =   210
         Width           =   2535
      End
      Begin VB.CheckBox Chk_date 
         Appearance      =   0  'Flat
         Caption         =   "SINGLE DAY  COLLECTION"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   210
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   3270
         TabIndex        =   8
         Top             =   1050
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60424193
         CurrentDate     =   38195
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1050
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60424193
         CurrentDate     =   38195
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1050
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60424193
         CurrentDate     =   38195
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   1050
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   1485
         Left            =   0
         Top             =   630
         Width           =   5745
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00E0E0E0&
         Height          =   465
         Left            =   -30
         Top             =   90
         Width           =   5685
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   435
         Left            =   630
         TabIndex        =   4
         Top             =   2010
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Option"
      Height          =   1755
      Left            =   0
      TabIndex        =   0
      Top             =   930
      Width           =   1755
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Shift Specific"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmDaily_coll_statement.frx":0000
         Left            =   2280
         List            =   "frmDaily_coll_statement.frx":000A
         TabIndex        =   1
         Top             =   2160
         Width           =   2085
      End
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   2910
      Top             =   3180
      Width           =   2565
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   0
      Picture         =   "frmDaily_coll_statement.frx":0020
      Top             =   2910
      Width           =   11820
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COLLECTION STATEMENT"
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
      Top             =   180
      Width           =   4755
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Picture         =   "frmDaily_coll_statement.frx":59A2
      Top             =   0
      Width           =   11820
   End
End
Attribute VB_Name = "frmDaily_collection_statement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Check1_Click()

    frmpatient_search.Chk_date.Value = 0
    frmpatient_search.Chk_date_to_date.Value = 0
    'Check1.Value = 1
     Label1.Caption = "Enter a Name"
    txtName.Visible = True
    DTPicker1.Visible = False
    DTPicker2.Visible = False
    DTPicker3.Visible = False
End Sub

Private Sub Chk_date_Click()
     Check1.Value = 0
    Chk_date_to_date.Value = 0
    
     txtName.Visible = False
     DTPicker2.Visible = False
    DTPicker3.Visible = False
    DTPicker1.Visible = True
    DTPicker1.Height = DTPicker2.Height
    DTPicker1.Top = DTPicker2.Top
    'frmpatient_search.Chk_date.Value = 1
    'Chk_date_to_date.Value = 0
'     Check1.Value = 0
 ' Else
   'Label1.Caption = ""
  'End If
End Sub

Private Sub Chk_date_to_date_Click()
         Label1.Caption = ""
       Check1.Value = 0
       Chk_date.Value = 0
      txtName.Visible = False
      DTPicker1.Visible = False
      'Chk_date.Visible = False
     DTPicker2.Visible = True
    DTPicker3.Visible = True
    
   
  
End Sub

Private Sub CMDEXIT_Click()

        Unload Me
End Sub

Private Sub cmdPreview_Click()

   If Chk_date.Value = 0 And Me.Chk_date_to_date.Value = 0 Then

  MsgBox "Choose Search Opition", vbInformation, " IT, DNMIH."

   Exit Sub
  End If
  
   If frmDaily_collection_statement.Chk_date.Value = 1 Then
        optionbuttonval = 1
           End If
    If frmDaily_collection_statement.Chk_date_to_date.Value = 1 Then
       optionbuttonval = 2
    End If
   Screen.MousePointer = vbHourglass
    rptMode = 16
     Viewer.Show vbModal
   
 End Sub
    
    
    
Private Sub Form_Load()
'    DTPicker1.Visible = False
    DTPicker2.Visible = False
    DTPicker3.Visible = False
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker3.Value = Date
 'rptMode = 1
' Check1.Value = 1
' Me.Chk_date.Value = 0
' Me.Chk_date_to_date.Value = 0
' 'Combo1.Text = "Medicine"
End Sub

'Private Sub Option1_Click(Index As Integer)
'Select Case Index
'    Case 0
'        If Option1(0).Value = True Then
'              IntOption = 1
'
''            Option1(1).Enabled = False
'            Combo1.Enabled = False
'        Else
''            Option1(1).Enabled = True
'            Combo1.Enabled = True
'
'        End If
'    Case 1
'        If Option1(1).Value = True Then
'             IntOption = 2
'
''            Option1(1).Enabled = True
'            Combo1.Enabled = True
'        Else
''            Option1(1).Enabled = False
'            Combo1.Enabled = False
'
'        End If
'End Select
'End Sub
Private Sub txtName_Change()
  
End Sub

Private Sub txtname_GotFocus()
    txtName.BackColor = vbCyan
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
           cmdPreview_Click
      End If
End Sub

Private Sub txtname_LostFocus()
    txtName.BackColor = vbWhite
End Sub
