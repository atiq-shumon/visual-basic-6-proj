VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "535 digital ports, 8-bit analog inputs"
   ClientHeight    =   5940
   ClientLeft      =   1320
   ClientTop       =   1470
   ClientWidth     =   6690
   ForeColor       =   &H00C0E0FF&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4380
      Top             =   4650
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command6 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   2295
      TabIndex        =   25
      Top             =   4920
      Width           =   690
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   3105
      TabIndex        =   23
      Text            =   "Text5"
      Top             =   3600
      Width           =   825
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   0
      Top             =   3600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   810
      TabIndex        =   22
      Top             =   3120
      Width           =   420
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "8"
      Height          =   375
      Index           =   7
      Left            =   1350
      TabIndex        =   21
      Top             =   4200
      Width           =   420
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "7"
      Height          =   375
      Index           =   6
      Left            =   1350
      TabIndex        =   20
      Top             =   3840
      Width           =   420
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "6"
      Height          =   375
      Index           =   5
      Left            =   1350
      TabIndex        =   19
      Top             =   3480
      Width           =   420
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "5"
      Height          =   375
      Index           =   4
      Left            =   1350
      TabIndex        =   18
      Top             =   3120
      Width           =   420
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "4"
      Height          =   375
      Index           =   3
      Left            =   810
      TabIndex        =   17
      Top             =   4200
      Width           =   420
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "3"
      Height          =   375
      Index           =   2
      Left            =   810
      TabIndex        =   16
      Top             =   3840
      Width           =   420
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "2"
      Height          =   375
      Index           =   1
      Left            =   810
      TabIndex        =   15
      Top             =   3480
      Width           =   420
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1755
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1755
      Top             =   1200
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3105
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   2400
      Width           =   825
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1350
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   1800
      Width           =   825
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3105
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   1200
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1350
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   600
      Width           =   825
   End
   Begin VB.CommandButton Command5 
      Caption         =   "READ"
      Height          =   495
      Left            =   2295
      TabIndex        =   4
      Top             =   3600
      Width           =   690
   End
   Begin VB.CommandButton Command4 
      Caption         =   "READ"
      Height          =   495
      Left            =   2295
      TabIndex        =   3
      Top             =   2400
      Width           =   690
   End
   Begin VB.CommandButton Command3 
      Caption         =   "WRITE"
      Height          =   495
      Left            =   2295
      TabIndex        =   2
      Top             =   1800
      Width           =   690
   End
   Begin VB.CommandButton Command2 
      Caption         =   "READ"
      Height          =   495
      Left            =   2295
      TabIndex        =   1
      Top             =   1200
      Width           =   690
   End
   Begin VB.CommandButton Command1 
      Caption         =   "WRITE"
      Height          =   495
      Left            =   2295
      TabIndex        =   0
      Top             =   600
      Width           =   690
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ANALOG 8-BIT ADC"
      Height          =   255
      Left            =   2295
      TabIndex        =   24
      Top             =   3240
      Width           =   1635
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter 0-255"
      Height          =   255
      Left            =   1350
      TabIndex        =   14
      Top             =   360
      Width           =   825
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select channel"
      Height          =   255
      Left            =   675
      TabIndex        =   9
      Top             =   2880
      Width           =   1230
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PORT4,write"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   960
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PORT4,read"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PORT1,read"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PORT1,write"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myindex%

Private Sub Command1_Click()
 MSComm1.Output = "O"
 MSComm1.Output = Chr$(Val(Text1.Text))
 Text1.Text = ""
End Sub


Private Sub Command2_Click()
   MSComm1.Output = "M"
   MSComm1.Output = "I"
   Timer1.Enabled = True
End Sub


Private Sub Command3_Click()
 MSComm1.Output = "X"
 MSComm1.Output = Chr$(Val(Text3.Text))
 Text3.Text = ""
End Sub

Private Sub Command4_Click()
   MSComm1.Output = "N"
   MSComm1.Output = "Y"
   Timer2.Enabled = True
End Sub

Private Sub Command5_Click()
     channel$ = Chr$(104 + myindex)
     'Text1.Text = channel$
     
     MSComm1.Output = channel$
     Timer3.Enabled = True
End Sub

Private Sub Command6_Click()
   End
End Sub

Private Sub Form_Load()
   MSComm1.CommPort = 1
   MSComm1.InputLen = 0
   MSComm1.PortOpen = True
WindowState = 2
End Sub

Private Sub Option1_Click(Index As Integer)
     myindex% = Index
     
End Sub

Private Sub Timer1_Timer()
   a$ = MSComm1.Input
   Text2.Text = Asc(a$)
   Timer1.Enabled = False
   
End Sub


Private Sub Timer2_Timer()
   a$ = MSComm1.Input
   Print a$
   Text4.Text = Asc(a$)
   Timer2.Enabled = False
End Sub


Private Sub Timer3_Timer()
   Timer3.Enabled = False
   ana$ = MSComm1.Input
   Text5.Text = Str$(Asc(ana$))
End Sub


