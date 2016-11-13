VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "LEVEL, WINDOW, HIGH VOLTAGE SETTING"
   ClientHeight    =   3045
   ClientLeft      =   495
   ClientTop       =   2415
   ClientWidth     =   6345
   Height          =   3735
   Left            =   435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   6345
   Top             =   1785
   Width           =   6465
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   5280
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   4680
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TEST"
      Height          =   435
      Left            =   1440
      TabIndex        =   15
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1620
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   1620
      TabIndex        =   4
      Text            =   "Text8"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   1620
      TabIndex        =   3
      Text            =   "Text7"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "SET HV"
      Height          =   495
      Left            =   2700
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SET WINDOW"
      Height          =   495
      Left            =   2700
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SET LEVEL"
      Height          =   495
      Left            =   2700
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "No return"
      Height          =   255
      Left            =   4455
      TabIndex        =   14
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label Label19 
      Caption         =   "0 to 1023"
      Height          =   255
      Left            =   135
      TabIndex        =   13
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label18 
      Caption         =   "0 to 1023"
      Height          =   255
      Left            =   135
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label17 
      Caption         =   "0 to 1023"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label16 
      Caption         =   "H +"
      Height          =   375
      Left            =   945
      TabIndex        =   10
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "B +"
      Height          =   375
      Left            =   945
      TabIndex        =   9
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   "A +"
      Height          =   375
      Left            =   945
      TabIndex        =   8
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   " No return"
      Height          =   255
      Index           =   5
      Left            =   4185
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   " No return"
      Height          =   255
      Index           =   4
      Left            =   4185
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   135
      Top             =   2280
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
mscomm1.Output = "T"
Timer1.Enabled = 1
End Sub


Private Sub Command5_Click()
x$ = "0"
   xx = Val(Text7.Text)
     
  If xx > 1023 Then xx = 1023
  If xx < 0 Then xx = 0
   
   mscomm1.Output = x$
   first = (xx * 64) \ 256
   yy% = Int(first)
   mscomm1.Output = Chr$(yy%)
   Alter = (xx Mod 256)
   zz% = Int(Alter)
   mscomm1.Output = Chr$(zz%)


End Sub

Private Sub Command6_Click()
 x$ = "1"
 xx = Val(Text8.Text)
     
  If xx > 1023 Then xx = 1023
  If xx < 0 Then xx = 0
   
   mscomm1.Output = x$
   first = (xx * 64) \ 256
   yy% = Int(first)
   mscomm1.Output = Chr$(yy%)
   Alter = (xx Mod 256)
   zz% = Int(Alter)
   mscomm1.Output = Chr$(zz%)
  
End Sub



Private Sub Command8_Click()
x$ = "2"
 

 xx = Val(Text3.Text)
     
  If xx > 1023 Then xx = 1023
  If xx < 0 Then xx = 0
   
   mscomm1.Output = x$
   first = (xx * 64) \ 256
   yy% = Int(first)
   mscomm1.Output = Chr$(yy%)
   Alter = (xx Mod 256)
   zz% = Int(Alter)
   mscomm1.Output = Chr$(zz%)

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
 Timer1.Enabled = 0
 r$ = mscomm1.Input
 text1.Text = r$
End Sub


