VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Rpt_discount 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   3690
      TabIndex        =   5
      ToolTipText     =   "CLOSE"
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   2460
      TabIndex        =   4
      ToolTipText     =   "VIEW REPORT"
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   5115
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   60
         TabIndex        =   1
         Top             =   1110
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         _Version        =   393216
         Format          =   22740993
         CurrentDate     =   38197
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   3150
         TabIndex        =   2
         Top             =   1080
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   609
         _Version        =   393216
         Format          =   22740993
         CurrentDate     =   38197
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DATE RANGE"
         Height          =   195
         Left            =   1830
         TabIndex        =   6
         Top             =   1140
         Width           =   1050
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DISCOUNT SUMMARY"
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
         Left            =   -90
         TabIndex        =   3
         Top             =   240
         Width           =   4755
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   0
         Picture         =   "Rpt_discount.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11700
      End
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   2400
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   0
      Picture         =   "Rpt_discount.frx":5982
      Stretch         =   -1  'True
      Top             =   1770
      Width           =   11610
   End
End
Attribute VB_Name = "Rpt_discount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDEXIT_Click()
        Unload Me
End Sub

Private Sub cmdPreview_Click()
  Screen.MousePointer = vbHourglass
   rptMode = 13
   Viewer.Show vbModal
       
End Sub

Private Sub Form_Load()
     DTPicker1.Value = Date
     DTPicker2.Value = Date
End Sub
