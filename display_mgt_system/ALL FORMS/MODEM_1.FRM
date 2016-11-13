VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "COMMUNICATION WITH ATMEL"
   ClientHeight    =   6210
   ClientLeft      =   915
   ClientTop       =   1770
   ClientWidth     =   8175
   Height          =   6900
   Left            =   855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8175
   Top             =   1140
   Width           =   8295
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   480
      Top             =   2400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      Height          =   330
      Left            =   5520
      TabIndex        =   5
      Top             =   5670
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   480
      Top             =   1050
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3240
      TabIndex        =   4
      Text            =   "633"
      Top             =   2100
      Width           =   1695
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   1050
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   5520
      TabIndex        =   3
      Top             =   105
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "HOOK ON"
      Height          =   750
      Left            =   1440
      TabIndex        =   2
      Top             =   3150
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   3570
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CALL"
      Height          =   750
      Left            =   1440
      TabIndex        =   1
      Top             =   1995
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TEST"
      Height          =   645
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1800
      Top             =   4320
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      CDTimeout       =   0
      CommPort        =   1
      CTSTimeout      =   0
      DSRTimeout      =   0
      DTREnable       =   -1  'True
      Handshaking     =   2
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "short delay"
      Height          =   330
      Left            =   3480
      TabIndex        =   10
      Top             =   735
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "delay 2s for data"
      Height          =   435
      Left            =   120
      TabIndex        =   9
      Top             =   1890
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "reading messages"
      Height          =   435
      Left            =   120
      TabIndex        =   8
      Top             =   630
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "delay for +++"
      Height          =   330
      Left            =   120
      TabIndex        =   7
      Top             =   3150
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Type desired number"
      Height          =   330
      Left            =   3240
      TabIndex        =   6
      Top             =   1785
      Width           =   1695
   End
   Begin VB.Menu exit 
      Caption         =   "EXIT"
      Begin VB.Menu confirm 
         Caption         =   "confirm"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public mytag%, mytag1%
Sub delay_1s()
 Timer2.Enabled = True
  Do
    DoEvents
  Loop Until mytag1% = 1
   mytag1% = 0
  Timer2.Enabled = 0
End Sub

Sub mydelay()
  timer4.Enabled = True
  Do
    DoEvents
  Loop Until mytag% = 1
   mytag% = 0
  timer4.Enabled = 0
End Sub

Private Sub Command1_Click()
 'TEST
    MSComm1.Output = "AT"
    MSComm1.Output = Chr$(13) 'confirm

End Sub

Private Sub Command2_Click()
   List1.Clear
End Sub

Private Sub Command3_Click()
 'DIAL NUMBER
     
MSComm1.Output = "ATL3" 'high speaker volume
MSComm1.Output = Chr$(13)
  Call mydelay
  x$ = MSComm1.Input
  List1.AddItem x$
 
MSComm1.Output = "ATM1" 'speaker on until connect
MSComm1.Output = Chr$(13)
   Call mydelay
 x$ = MSComm1.Input
 List1.AddItem x$
  numb$ = Text4.Text
  mycall$ = "ATDT" + numb$
  MSComm1.Output = mycall$ '
  MSComm1.Output = Chr$(13) ' carriage return, to confirm
   Call mydelay
List1.AddItem x$
 
End Sub

Private Sub Command6_Click()
'HOOK ON
 ' We have to wait at least 1 second

Call delay_1s
   MSComm1.Output = "+++"
Call delay_1s
   text1.Text = "You can talk with modem again"
   MSComm1.Output = "ATH0"
   MSComm1.Output = Chr$(13)
End Sub



Private Sub confirm_Click()
 
    MSComm1.Output = "H0"   'put down
    MSComm1.Output = Chr$(13) 'confirm
  End
End Sub

Private Sub Form_Load()
 MSComm1.Handshaking = comRTS
 MSComm1.DTREnable = True
  MSComm1.PortOpen = True
   MSComm1.Output = "ATM1" 'speaker on until connect
   MSComm1.Output = Chr$(13)
End Sub


Private Sub Timer1_Timer()
 'reading messages from COMM port.Running permanently
      x$ = MSComm1.InputcomRTS
   If x$ <> "" Then
      List1.AddItem x$
   End If
      
   
  'recognize message "connect":
   mypos = InStr(1, x$, "CONNECT", 0)
       If mypos > 0 Then
          MSComm1.Output = "U"
          MSComm1.Output = Chr$(10)
          Timer1.Enabled = False
          List1.Clear
          List1.AddItem "CONNECTED"
          Timer3.Enabled = True
       End If

End Sub


Private Sub Timer2_Timer()
  mytag1% = 1
End Sub


Private Sub Timer3_Timer()
  'waiting 2 sec for message
  'message put on the list
 x$ = MSComm1.Input
   List1.AddItem x$
   Timer3.Enabled = 0
 
End Sub

Private Sub timer4_Timer()
   mytag% = 1
End Sub


