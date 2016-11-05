VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Rpt_SHIFTWISE_PAT_ADMISSION 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMDEXIT 
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   5970
      TabIndex        =   3
      ToolTipText     =   "CLOSE"
      Top             =   2340
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   4740
      TabIndex        =   2
      ToolTipText     =   "VIEW REPORT"
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   1575
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   7845
      Begin VB.ComboBox CBOYRCODE 
         Height          =   315
         ItemData        =   "Rpt_SHIFTWISE_PAT_ADMISSION.frx":0000
         Left            =   330
         List            =   "Rpt_SHIFTWISE_PAT_ADMISSION.frx":0002
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "YR-0809"
         Top             =   1020
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.ComboBox cboShift 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   420
         ItemData        =   "Rpt_SHIFTWISE_PAT_ADMISSION.frx":0004
         Left            =   4650
         List            =   "Rpt_SHIFTWISE_PAT_ADMISSION.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   570
         Width           =   2235
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   405
         Left            =   330
         TabIndex        =   0
         Top             =   570
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61997057
         CurrentDate     =   38197
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Shift :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4650
         TabIndex        =   7
         Top             =   270
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   330
         TabIndex        =   6
         Top             =   270
         Width           =   1785
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0FF&
         Height          =   1395
         Left            =   -540
         Top             =   -30
         Width           =   8565
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SHIFT WISE PATIENT ADMISSION"
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
      Left            =   1440
      TabIndex        =   5
      Top             =   120
      Width           =   5115
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   -600
      Picture         =   "Rpt_SHIFTWISE_PAT_ADMISSION.frx":0008
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11610
   End
   Begin VB.Shape Shape1 
      Height          =   435
      Left            =   4680
      Top             =   2310
      Width           =   2535
   End
End
Attribute VB_Name = "Rpt_SHIFTWISE_PAT_ADMISSION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
        Unload Me
End Sub
Private Sub cmdPreview_Click()
 Screen.MousePointer = vbHourglass
 rptMode = 507
 Viewer.Show vbModal
 
       
End Sub

Private Sub Form_Load()
 DTPicker1.Value = Date
 populateCombo
End Sub
Private Sub populateCombo()
    cboShift.AddItem ("Morning")
    cboShift.AddItem ("Evening")
    cboShift.AddItem ("Night")
    cboShift.ListIndex = 0
End Sub
