VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H0000C0C0&
   Caption         =   "CALL ATMEL"
   ClientHeight    =   1890
   ClientLeft      =   2835
   ClientTop       =   1905
   ClientWidth     =   4395
   Height          =   2295
   Left            =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   4395
   Top             =   1560
   Width           =   4515
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2760
      Top             =   525
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2565
      TabIndex        =   1
      Top             =   1050
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...here"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1050
      Width           =   690
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ATMEL, can you hear me?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   105
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Press.."
      Height          =   225
      Left            =   1800
      TabIndex        =   2
      Top             =   735
      Width           =   495
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   810
      Top             =   1050
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
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub Command1_Click()
 MSComm1.Output = "U" ' question to ATMEL
 Timer1.Enabled = True
 Text1.Text = ""
End Sub


Private Sub Form_Load()
  
  MSComm1.CommPort = 1
  MSComm1.PortOpen = True
 
End Sub


Private Sub Timer1_Timer()
 'ATMEL answer arrived...
   Timer1.Enabled = False
   a$ = MSComm1.Input
   Text1.Text = a$ 'Reply from ATMEL printed
   
End Sub


