VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "TEST, TIMER and COUNTER for 535 or RIA"
   ClientHeight    =   5310
   ClientLeft      =   1665
   ClientTop       =   1800
   ClientWidth     =   9075
   Height          =   6000
   Left            =   1605
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   9075
   Top             =   1170
   Width           =   9195
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   1755
      TabIndex        =   29
      Text            =   "Insert preset time in seconds"
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox Text9 
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
      Left            =   6210
      TabIndex        =   22
      Text            =   "Text9"
      Top             =   4440
      Width           =   2655
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
      Left            =   5880
      TabIndex        =   21
      Text            =   "Text5"
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4725
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4185
      Top             =   4440
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4200
      Top             =   360
   End
   Begin VB.CommandButton Command10 
      Caption         =   "STOP"
      Height          =   495
      Left            =   2835
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "START"
      Height          =   495
      Left            =   2835
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4200
      Top             =   2160
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "READ"
      Height          =   495
      Left            =   2835
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PRESET"
      Height          =   495
      Left            =   2835
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   2835
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GET STATUS"
      Height          =   495
      Left            =   2835
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TEST"
      Height          =   495
      Left            =   2835
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "E +"
      Height          =   375
      Left            =   1215
      TabIndex        =   28
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME(s)"
      Height          =   255
      Left            =   7965
      TabIndex        =   27
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LSB   MSB    CR+LF   LSB    MSB"
      Height          =   255
      Left            =   6240
      TabIndex        =   26
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "COUNTS"
      Height          =   255
      Left            =   6345
      TabIndex        =   25
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSLATED:"
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "AS TEXT:"
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   4560
      Width           =   315
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   495
      Left            =   2400
      TabIndex        =   18
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      Height          =   495
      Left            =   2400
      TabIndex        =   17
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " No return"
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " No return"
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   12
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " No return"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " No return"
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   240
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

Private Sub Command10_Click()
 a$ = "P"
    mscomm1.Output = a$
   
End Sub

Private Sub Command2_Click()
    a$ = "G"
    mscomm1.Output = a$
    Timer1.Enabled = True
End Sub


Private Sub Command3_Click()
a$ = "C"
    mscomm1.Output = a$
End Sub

Private Sub Command4_Click()
    Number% = Val(Text6.Text)
    'preset = 65536 - Number%
    preset = Number%
    pre_high = preset \ 256
    pre_low = preset Mod 256
    
    Print pre_high
    Print pre_low
   'Exit Sub
    mscomm1.Output = "E"
    mscomm1.Output = Chr$(pre_low)
    mscomm1.Output = Chr$(pre_high)
   
    
    
End Sub

Private Sub Command5_Click()
   x$ = "A"
   xx% = Val(Text7.Text)
     
  If xx% > 255 Then xx% = 255
  If xx% < 0 Then xx% = 0
   
   mscomm1.Output = x$
   mscomm1.Output = Chr$(xx%)
End Sub

Private Sub Command6_Click()
 x$ = "B"
   xx% = Val(Text8.Text)
   
  If xx% > 255 Then xx% = 255
  If xx% < 0 Then xx% = 0
   
   mscomm1.Output = x$
   mscomm1.Output = Chr$(xx%)
End Sub


Private Sub Command7_Click()
a$ = "R"
    mscomm1.Output = a$
    Timer7.Enabled = True
End Sub

Private Sub Command8_Click()
x$ = "H"
   
   
   xx% = Val(Text3.Text)
  If xx% > 255 Then xx% = 255
  If xx% < 0 Then xx% = 0
   
   
   mscomm1.Output = x$
   mscomm1.Output = Chr$(xx%)
End Sub

Private Sub Command9_Click()
 a$ = "S"
    mscomm1.Output = a$
   
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
   ' low_count+midle_count * 256 + high_count * 256 ^ 2
   ' carriage return, line feed
   ' seconds_low + 256* seconds_high
    
End Sub


