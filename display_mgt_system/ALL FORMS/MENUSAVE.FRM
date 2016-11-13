VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "MENU and FILE save operation using CMD"
   ClientHeight    =   4755
   ClientLeft      =   735
   ClientTop       =   2205
   ClientWidth     =   6240
   Height          =   5160
   Left            =   675
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6240
   Top             =   1860
   Width           =   6360
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   2025
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "menusave.frx":0000
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2025
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   2325
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE AS..."
      Height          =   495
      Left            =   810
      TabIndex        =   1
      Top             =   1200
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OPEN"
      Height          =   495
      Left            =   810
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   810
      Top             =   2040
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      FilterIndex     =   1
      InitDir         =   "c:\word"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'Public my%(100)



Private Sub Command1_Click()
 
     CommonDialog1.ShowOpen
     filename$ = CommonDialog1.filename
     Text1.Text = filename$
     Open filename$ For Output As #1
 For jj% = 0 To 10
     Print #1, jj%
 Next jj%
 
 Close #1
End Sub


Private Sub Command2_Click()
 CommonDialog1.ShowSave
 filename$ = CommonDialog1.filename$
 Text1.Text = filename$
 Open filename$ For Output As #1
 Print #1, Text2.Text
 Close #1
End Sub


Private Sub Form_Load()
  For j = 1 To 50
    my(j) = j
  Next j
  
 CommonDialog1.Filter = "TEXT ( *.doc )| *.doc |picture (*.bmp)|*.bmp |data (*.dat)|*.dat|HPGL (*.plt)| *.plt|"
 
End Sub


