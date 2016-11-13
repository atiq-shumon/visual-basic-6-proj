VERSION 4.00
Begin VB.Form frm_help 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "HELP"
   ClientHeight    =   6855
   ClientLeft      =   75
   ClientTop       =   1695
   ClientWidth     =   7890
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Height          =   7260
   Left            =   15
   LinkTopic       =   "Form2"
   ScaleHeight     =   6855
   ScaleWidth      =   7890
   Top             =   1350
   Width           =   8010
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7470
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3000
      TabIndex        =   1
      Top             =   600
      Width           =   15
   End
End
Attribute VB_Name = "frm_help"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Form_Load()
  On Error GoTo sos33
     Open "c:\ria\help.prt" For Input As #21
     Do
       Line Input #21, b$
       list1.AddItem b$
     Loop While Not EOF(21)
       Close #21
       Exit Sub
sos33:
    MsgBox "HELP.PRT missing. Copy it from installation diskette into the directory with RIA.EXE"
         RIA_BASE.Show
         missing% = 1
         
       Exit Sub
End Sub

