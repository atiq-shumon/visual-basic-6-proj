VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   2640
   ClientTop       =   3570
   ClientWidth     =   6570
   Height          =   2610
   Left            =   2580
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   6570
   Top             =   2940
   Width           =   6690
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      Height          =   255
      Left            =   2565
      TabIndex        =   6
      Top             =   1080
      Width           =   3390
   End
   Begin VB.TextBox Text5 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3645
      TabIndex        =   3
      Text            =   "Text5"
      Top             =   480
      Width           =   2295
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2025
      Top             =   480
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2565
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TEST"
      Height          =   495
      Left            =   675
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "TRANSLATED:"
      Height          =   255
      Left            =   4185
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "AS TEXT:"
      Height          =   255
      Left            =   2700
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "T"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   135
      Top             =   1080
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      CDTimeout       =   0
      CommPort        =   1
      CTSTimeout      =   0
      DSRTimeout      =   0
      DTREnable       =   -1  'True
      Handshaking     =   0
      InBufferSize    =   1024
      InputLen        =   0
      Interval        =   1000
      NullDiscard     =   0   'False
      OutBufferSize   =   512
      ParityReplace   =   "?"
      RThreshold      =   0
      RTSEnable       =   0   'False
      Settings        =   "9600,n,8,1"
      SThreshold      =   0
   End
   Begin VB.Menu exit 
      Caption         =   "EXIT"
      Begin VB.Menu con 
         Caption         =   "...confirm"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
a$ = "T"
    mscomm1.Output = a$
    Timer3.Enabled = True
End Sub



Private Sub Command2_Click()
   Text1.Text = ""
   Text5.Text = ""
End Sub












Private Sub con_Click()
 End
End Sub

Private Sub Form_Load()
    
    ' Use COM1.
    mscomm1.CommPort = 1
    ' 9600 baud, no parity, 8 data, and 1 stop bit.
    mscomm1.Settings = "9600,N,8,1"
    ' Tell the control to read entire buffer when Input is used.
    mscomm1.InputLen = 0
    ' Open the port.
    mscomm1.PortOpen = True
    ' Send the attention command to the modem.
    mscomm1.Output = "AT" + Chr$(13)
    ' Wait for data to come back to the serial port.
    
    
    
End Sub








Private Sub Timer3_Timer()
     Timer3.Enabled = False
     b$ = mscomm1.Input
     Text1.Text = b$
   
   'translation:
      word$ = ""
    For gg% = 1 To 7
      x$ = Mid$(b$, gg%, 1)
      y$ = Asc(x$)
      word$ = word$ + " " + y$
    Next gg%
      'result is:  1 2 3 13 10 4 5
     Text5.Text = word$
End Sub





