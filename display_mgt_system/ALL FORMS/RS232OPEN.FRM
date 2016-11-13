VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "RS232 open"
   ClientHeight    =   1740
   ClientLeft      =   2640
   ClientTop       =   1545
   ClientWidth     =   3075
   Height          =   2145
   Left            =   2580
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   3075
   Top             =   1200
   Width           =   3195
   Begin VB.TextBox Text1 
      Height          =   1590
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "rs232open.frx":0000
      Top             =   105
      Width           =   2295
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   420
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



Private Sub Form_Load()
   On Error GoTo erhandler
    ' Use COM1.
    MSComm1.CommPort = 2
    'Input command read complete buffer content
    MSComm1.InputLen = 0
    ' Open the port.
    MSComm1.PortOpen = True
 Exit Sub
    
erhandler:
  MsgBox "COM2  not available . Trying to open COM1"
    MSComm1.CommPort = 1
    MSComm1.InputLen = 0
    MSComm1.PortOpen = True
  MsgBox "Communication through COM1"
 
End Sub






