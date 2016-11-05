VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmgroup_statement 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gruop Wise Collection Statement"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1755
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   3495
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60096513
         CurrentDate     =   38195
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60096513
         CurrentDate     =   38195
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Format          =   60096513
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
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   3135
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
         Height          =   435
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Option"
      Height          =   1755
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1755
      Begin VB.CheckBox Chk_date 
         Appearance      =   0  'Flat
         Caption         =   "Date Specific"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox Chk_date_to_date 
         Appearance      =   0  'Flat
         Caption         =   "Date to date"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Shift Specific"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frm_group_statement.frx":0000
         Left            =   2280
         List            =   "frm_group_statement.frx":000A
         TabIndex        =   3
         Top             =   2160
         Width           =   2085
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   600
      Picture         =   "frm_group_statement.frx":0020
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   1950
      Width           =   510
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   60
      Picture         =   "frm_group_statement.frx":093E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Preview"
      Top             =   1950
      Width           =   510
   End
   Begin VB.Shape Shape1 
      Height          =   525
      Left            =   0
      Top             =   1890
      Width           =   1215
   End
End
Attribute VB_Name = "frmgroup_statement"
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
    Label1.Caption = "Select a Date "
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
    Label1.Caption = "Select from date To date "

   
  
End Sub

Private Sub CMDEXIT_Click()

Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdPreview_Click()

   If Chk_date.Value = 0 And Me.Chk_date_to_date.Value = 0 Then

  MsgBox "Please Select an Search Opition", vbInformation, " IT, DNMIH."

   Exit Sub
  End If
  
   If frmgroup_statement.Chk_date.Value = 1 Then
        optionbuttonval = 1
           End If
    If frmgroup_statement.Chk_date_to_date.Value = 1 Then
       optionbuttonval = 2
    End If
   Screen.MousePointer = vbHourglass
    rptMode = 47
     Viewer.Show vbModal
   
 End Sub
    
    
    
Private Sub Form_Load()
    DTPicker1.Visible = False
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
