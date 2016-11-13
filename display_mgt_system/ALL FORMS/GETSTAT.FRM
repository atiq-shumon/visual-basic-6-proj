VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "GET STATUS"
   ClientHeight    =   2250
   ClientLeft      =   2160
   ClientTop       =   1890
   ClientWidth     =   4140
   Height          =   2940
   Left            =   2100
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4140
   Top             =   1260
   Width           =   4260
   Begin VB.CommandButton Command3 
      Caption         =   "STOP"
      Height          =   495
      Left            =   945
      TabIndex        =   4
      Top             =   1440
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      Height          =   495
      Left            =   945
      TabIndex        =   3
      Top             =   840
      Width           =   1230
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2295
      Top             =   240
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2835
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GET STATUS"
      Height          =   495
      Left            =   945
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "P"
      Height          =   435
      Left            =   360
      TabIndex        =   6
      Top             =   1575
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "S"
      Height          =   435
      Left            =   360
      TabIndex        =   5
      Top             =   945
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "G"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   420
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3480
      Top             =   1785
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
   mscomm1.Output = "S"
   
End Sub

Private Sub Command2_Click()
   mscomm1.Output = "G"
   Timer1.Enabled = True
End Sub


Private Sub Command3_Click()
  mscomm1.Output = "P"
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




Private Sub Timer1_Timer()
       b$ = mscomm1.Input
       Timer1.Enabled = False
       Text2.Text = b$
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


Private Sub Timer7_Timer()
  Timer7.Enabled = False
  res$ = mscomm1.Input
  Text4.Text = res$
   
   'translation:
      word$ = ""
    For gg% = 1 To 7
      x$ = Mid$(res$, gg%, 1)
      y$ = Asc(x$)
      word$ = word$ + " " + y$
    Next gg%
      'result is:
    Text9.Text = word$
    
    
   ' structure of result:
   ' low_count midle_count * 256 + high_count * 256 ^ 2
   ' carriage return, line feed
   ' seconds_low + 256* seconds_high
    
End Sub


