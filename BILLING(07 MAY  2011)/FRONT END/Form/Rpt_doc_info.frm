VERSION 5.00
Begin VB.Form Rpt_doc_info 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Doctors' Report"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   0
      TabIndex        =   2
      Top             =   -60
      Width           =   4515
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "All"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   330
         Width           =   645
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "Specific"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   885
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Rpt_doc_info.frx":0000
         Left            =   1110
         List            =   "Rpt_doc_info.frx":000A
         TabIndex        =   3
         Top             =   720
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor's Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Top             =   210
         Width           =   2385
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
      Left            =   570
      Picture         =   "Rpt_doc_info.frx":0020
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   1380
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
      Picture         =   "Rpt_doc_info.frx":093E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Preview"
      Top             =   1380
      Width           =   510
   End
   Begin VB.Shape Shape1 
      Height          =   525
      Left            =   0
      Top             =   1320
      Width           =   1155
   End
End
Attribute VB_Name = "Rpt_doc_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDEXIT_Click()

Dim reply As String
    reply = MsgBox("Sure to Close?", vbQuestion + vbYesNo, "Close...")
    If reply = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdPreview_Click()
  Screen.MousePointer = vbHourglass
       Viewer.Show vbModal
       
End Sub

Private Sub Form_Load()
 rptMode = 1
 Option1(0).Value = True
 Combo1.Text = "Medicine"
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
    Case 0
        If Option1(0).Value = True Then
              IntOption = 1

'            Option1(1).Enabled = False
            Combo1.Enabled = False
        Else
'            Option1(1).Enabled = True
            Combo1.Enabled = True

        End If
    Case 1
        If Option1(1).Value = True Then
             IntOption = 2
             
'            Option1(1).Enabled = True
            Combo1.Enabled = True
        Else
'            Option1(1).Enabled = False
            Combo1.Enabled = False

        End If
End Select
End Sub
