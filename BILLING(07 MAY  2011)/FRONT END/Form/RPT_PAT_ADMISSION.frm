VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RPT_PAT_ADMISSION 
   Appearance      =   0  'Flat
   BackColor       =   &H80000014&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3270
      TabIndex        =   9
      Top             =   960
      Width           =   1245
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   8
      Top             =   960
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton CMDEXIT 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   4050
      TabIndex        =   4
      ToolTipText     =   "CLOSE"
      Top             =   2910
      Width           =   1215
   End
   Begin VB.CommandButton CMDREPORT 
      Caption         =   "REPORT"
      Height          =   375
      Left            =   2820
      TabIndex        =   3
      ToolTipText     =   "VIEW REPORT"
      Top             =   2910
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   1770
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   12582912
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mm-YYYY"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   345
      Left            =   3180
      TabIndex        =   2
      Top             =   1770
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   12582912
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd-mm-YYYY"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0FF&
      FillColor       =   &H00FFFF00&
      Height          =   585
      Left            =   -30
      Top             =   780
      Width           =   5505
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   2370
      Width           =   75
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   2760
      Top             =   2850
      Width           =   2565
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TO  DATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   3240
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FROM DATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   270
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PATIENT ADMISSION  REPORT"
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
      Left            =   330
      TabIndex        =   0
      Top             =   150
      Width           =   4755
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   1
      Left            =   -390
      Picture         =   "RPT_PAT_ADMISSION.frx":0000
      Stretch         =   -1  'True
      Top             =   2730
      Width           =   11610
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   0
      Left            =   0
      Picture         =   "RPT_PAT_ADMISSION.frx":5982
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11610
   End
End
Attribute VB_Name = "RPT_PAT_ADMISSION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UTILITY As New clsUtility
Private Sub CMDEXIT_Click()
 Unload Me
End Sub
Private Sub CMDREPORT_Click()
    If UTILITY.START_END_VALIDATION(MaskEdBox1, MaskEdBox2) = False Then
      Label2.Caption = "Start Date can't be greater(>) than End date..Verify"
      MaskEdBox1.SetFocus
      Exit Sub
   End If
  Screen.MousePointer = vbHourglass
  Label2.Caption = "Please wait while processing...."
  If Option1(0).Value = True Then
      rptMode = 504
  ElseIf Option1(1).Value = True Then
      ' rptMode = 501
      frmpatient_history.Show 1
  End If
  
  Viewer.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys (Chr(9))
  End If
End Sub
Private Sub Form_Load()
   MaskEdBox1.Text = Format(Date, "DD/MM/YY")
   MaskEdBox2.Text = Format(Date, "DD/MM/YY")
  
End Sub

Private Sub MaskEdBox1_Change()
  Label2.Caption = ""
End Sub

Private Sub MaskEdBox2_GotFocus()
  With MaskEdBox2
       .SelStart = 0
       .SelLength = Len(MaskEdBox2)
       
  End With
End Sub

Private Sub MaskEdBox1_GotFocus()
  With MaskEdBox1
       .SelStart = 0
       .SelLength = Len(MaskEdBox1)
  End With
  
End Sub


