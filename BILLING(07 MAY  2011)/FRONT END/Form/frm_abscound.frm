VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_abscond 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C8AC59&
      Height          =   825
      Left            =   -30
      TabIndex        =   7
      Top             =   -150
      Width           =   5805
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ABSCONDED PATIENT REPORT"
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
         Left            =   390
         TabIndex        =   9
         Top             =   300
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   -270
         Picture         =   "frm_abscound.frx":0000
         Top             =   30
         Width           =   11820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   -30
      TabIndex        =   0
      Top             =   540
      Width           =   5775
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   540
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38197
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   3210
         TabIndex        =   2
         Top             =   540
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60882945
         CurrentDate     =   38197
      End
      Begin VB.Shape Shape3 
         Height          =   375
         Left            =   3150
         Top             =   510
         Width           =   2295
      End
      Begin VB.Shape Shape2 
         Height          =   375
         Left            =   210
         Top             =   510
         Width           =   2085
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3210
         TabIndex        =   6
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000001&
      Height          =   825
      Left            =   -120
      TabIndex        =   8
      Top             =   1710
      Width           =   5805
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "CLOSE"
         Height          =   375
         Left            =   4290
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "VIEW"
         Height          =   375
         Left            =   3060
         TabIndex        =   3
         Top             =   300
         Width           =   1215
      End
      Begin VB.Shape Shape4 
         Height          =   465
         Left            =   3000
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frm_abscond"
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
       rptMode = 48
       Viewer.Show vbModal
    
End Sub

Private Sub Form_Load()
 DTPicker1.Value = Date
 DTPicker2.Value = Date
End Sub


