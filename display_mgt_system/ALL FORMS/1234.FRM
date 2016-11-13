VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "TRAIN OF SAMPLES IN SHORT INTERVALS"
   ClientHeight    =   6690
   ClientLeft      =   570
   ClientTop       =   1770
   ClientWidth     =   9570
   Height          =   7380
   Left            =   510
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9570
   Top             =   1140
   Width           =   9690
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4590
      TabIndex        =   13
      Top             =   1320
      Width           =   1230
   End
   Begin VB.OptionButton Option1 
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   1350
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   1350
      TabIndex        =   10
      Top             =   3600
      Width           =   960
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   1350
      TabIndex        =   9
      Top             =   3240
      Width           =   960
   End
   Begin VB.OptionButton Option1 
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   1350
      TabIndex        =   8
      Top             =   2880
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   6105
      Left            =   7830
      TabIndex        =   4
      Top             =   120
      Width           =   1230
   End
   Begin VB.ListBox List1 
      Height          =   6105
      Left            =   6345
      TabIndex        =   3
      Top             =   120
      Width           =   1230
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   4080
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1350
      TabIndex        =   2
      Text            =   "512"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1350
      TabIndex        =   1
      Text            =   "1"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RUN"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1350
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Spent (sec):"
      Height          =   255
      Left            =   3645
      TabIndex        =   14
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Select channel"
      Height          =   495
      Left            =   270
      TabIndex        =   12
      Top             =   2880
      Width           =   960
   End
   Begin VB.Label Label3 
      Caption         =   "Number:"
      Height          =   375
      Left            =   405
      TabIndex        =   7
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "X 100 usec"
      Height          =   255
      Left            =   2565
      TabIndex        =   6
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Time interval"
      Height          =   255
      Left            =   270
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   270
      Top             =   4680
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
   Begin VB.Menu quit 
      Caption         =   "QUIT"
      Begin VB.Menu ok 
         Caption         =   "OK?"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub Command1_Click()
  
 Static yy%
 Command1.Enabled = False

a$ = Chr$(&H70 + train_select%) 'acquire activation
  'Print train_select%
  For kk = 1 To 7000: Next kk 'wait loop
  'TIME: N* 100 useconds
      
         cc% = Val(Text1.Text)
      If cc% > 255 Then
         Text1.Text = 255
         MsgBox "Maximum number is 255"
         cc% = 255
      End If
        b$ = Chr$(cc%)
     
  'SAMPLE NUMBER (0 = max  to 127 =min)
    numb% = Val(Text2.Text)
    numb1% = numb% \ 128 'number of blocks of 256 doubled
    ' there are two readings for one result
    numb% = numb1% * 128 ' corrected number
    c$ = Chr$(numb1%)
    Text2.Text = numb% 'corrected number
    For kk = 1 To 7000: Next kk
         
         MSComm1.InputLen = 1
 'instruction prepared: channel, spacing and number of samples
      a$ = Chr$(&H70 + train_select%) + b$ + c$
         MSComm1.Output = a$
         mytimer = Timer
     'waiting for "J":
nazaj:
        D$ = MSComm1.Input
           If D$ = "J" Then GoTo ven
        GoTo nazaj
ven:
   text3.Text = Timer - mytimer
   'zdaj beremo
    a$ = "D"
   'to je komanda:
       MSComm1.Output = a$
    'cakamo odgovore, ki niso prazni.
    'beremo po enega
         
       For GG% = 0 To numb% - 2
back:
        mem$(GG%) = MSComm1.Input
        If mem$(GG%) = "" Then GoTo back
      Next GG%
     
      
      For jj% = 0 To numb% - 4
         high% = Asc(mem$(jj%))
         jj% = jj% + 1
         low% = Asc(mem$(jj%))
         List1.AddItem low% + 256 * high%
         List2.AddItem Str$(high%) + "   " + Str$(low%)
         yy% = yy% + 1
         ww% = ww% + 1
      Next jj%
         If yy% > 30 Then List1.ListIndex = yy% - 1
         If yy% > 30 Then List2.ListIndex = yy% - 1
         'Print ww%
         MSComm1.InputLen = 0
         Command1.Enabled = True
 Exit Sub

 
End Sub


Private Sub Form_Load()
      WindowState = 2
      Form1.MSComm1.CommPort = 1
      Form1.MSComm1.PortOpen = True
End Sub


Private Sub ok_Click()
  End
End Sub


Private Sub Option1_Click(index As Integer)
     train_select% = Option1(index).Caption
     'Print train_select%
End Sub


Private Sub Option2_Click()
End Sub


