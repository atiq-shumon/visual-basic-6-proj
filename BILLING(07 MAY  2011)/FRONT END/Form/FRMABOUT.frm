VERSION 5.00
Begin VB.Form FRMABOUT 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6960
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   525
      Left            =   0
      TabIndex        =   3
      Top             =   4080
      Width           =   6975
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   120
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   30
      Left            =   3480
      TabIndex        =   1
      Top             =   4560
      Width           =   30
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   -30
      TabIndex        =   0
      Top             =   -90
      Width           =   7005
      Begin VB.Frame Frame3 
         Height          =   2685
         Left            =   60
         TabIndex        =   2
         Top             =   1500
         Width           =   6915
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   0
         X2              =   6930
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line2 
         X1              =   900
         X2              =   900
         Y1              =   1350
         Y2              =   1380
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   60
         X2              =   6990
         Y1              =   750
         Y2              =   750
      End
   End
End
Attribute VB_Name = "FRMABOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Unload Me
End Sub

