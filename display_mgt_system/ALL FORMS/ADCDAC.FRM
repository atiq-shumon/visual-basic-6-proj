VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Testing ADC 12-bit and DAC 12-bit"
   ClientHeight    =   5580
   ClientLeft      =   2010
   ClientTop       =   1665
   ClientWidth     =   5760
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   6270
   Left            =   1950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   5760
   Top             =   1035
   Width           =   5880
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2970
      Top             =   4320
   End
   Begin VB.TextBox Text2 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      Height          =   375
      Left            =   2295
      TabIndex        =   17
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "SINGLE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   945
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   900
      ScaleHeight     =   1395
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   3360
      Width           =   615
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         Caption         =   "D"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         Caption         =   "C"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         Caption         =   "B"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   0
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FFFF&
         Caption         =   "A"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "OUT"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1620
      TabIndex        =   10
      Top             =   4320
      Width           =   990
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1620
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CH.4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   960
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CH.3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   960
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CH.2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CH.1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   945
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   135
      Top             =   120
   End
   Begin VB.ListBox List1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      Left            =   4185
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   135
      Top             =   1080
   End
   Begin VB.TextBox text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "data in ""ADC.DAT"""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      BorderWidth     =   5
      Height          =   495
      Left            =   2970
      Shape           =   3  'Circle
      Top             =   3720
      Visible         =   0   'False
      Width           =   420
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   135
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "DAC 12-bit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1620
      TabIndex        =   8
      Top             =   3360
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Read 12 bit: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   945
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "measured:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2430
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Menu test 
      Caption         =   "TEST"
      Begin VB.Menu do_it 
         Caption         =   "Do it!"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "EXIT"
   End
   Begin VB.Menu go_dos 
      Caption         =   "DOS"
      Begin VB.Menu do_dos 
         Caption         =   "GO"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public pp% ' number of samples in "ADC.dat"
Public myyy% ' handler for list1
Public numb%
Public mycall% 'index of channel
Public mytag% ' periodic=1, noperiodic=0
Public mytime
Public out_ch% ' DAC





Private Sub Check1_Click()
mytag% = mytag% Xor 1
          
      Select Case mytag%
       Case Is = 0
           Check1.Caption = "SINGLE"
       Case Is = 1
           Check1.Caption = "PERIODIC"
      End Select
      'Print mytag%
End Sub

Private Sub Command1_Click(Index As Integer)
      mycall% = Index
      a$ = Chr$(128 + mycall%)
      MSComm1.Output = a$ 'activate channel INDEX
      If mytag% = 1 Then Open "ADC.dat" For Output As 1
      Timer2.Enabled = True
      
End Sub









Private Sub Command2_Click()
text1.Text = ""
'how many items are in list?
  mm% = list1.ListCount
  ww% = mm% - 1
  myyy% = 0
  For jj% = ww% To 0 Step -1
     list1.RemoveItem jj%
  Next jj%
   pp% = 0 ' event counter reset
   Text2.Visible = 0
   
End Sub



Private Sub Command6_Click()
     first$ = Str$(out_ch%) 'we run 12 bit DAC
 MSComm1.Output = first$
     'what to output?
 dac% = Val(Text4.Text)
    If dac% > 4095 Then dac% = 4095
    If dac% < 0 Then dac% = 0
    high% = dac% \ 16
 MSComm1.Output = Chr$(high%)
 
    low% = (dac% Mod 16) * 16
    
 MSComm1.Output = Chr$(low%)
 'Print high%
 'Print low%
    Shape1.Visible = True
    Timer1.Enabled = True
End Sub

Private Sub do_it_Click()
     a$ = "T"
      MSComm1.Output = a$
      Timer3.Enabled = True
End Sub

Private Sub exit_Click()
     End
End Sub

Private Sub Form_Load()
 
 Beep
      Form1.MSComm1.CommPort = 1
      Form1.MSComm1.PortOpen = True
    
     
     
End Sub

Private Sub Option2_Click(Index As Integer)
   out_ch% = Index%
End Sub














Private Sub Timer1_Timer()
   Timer1.Enabled = False
   Shape1.Visible = False
End Sub

Private Sub Timer2_Timer()

 
        Select Case mytag%
         Case Is = 0   'single reading
          For jj% = 0 To 3
            Command1(jj%).Enabled = 1
          Next jj%
              
          Timer2.Enabled = False
         Case Is = 1   'periodic
          For jj% = 0 To 3
            Command1(jj%).Enabled = 0
          Next jj%
            Text2.Visible = 1
            Label3.Visible = True
        End Select
        
next_time:
         a$ = MSComm1.Input
         nn% = 1 'take first
         b$ = Mid$(a$, nn%, 1)
         H$ = Asc(b$)
         nn% = nn% + 1 'take next
         c$ = Mid$(a$, nn%, 1)
         L$ = Asc(c$)
         nn% = nn% + 1 'take next
         low% = Val(H$)
         high% = Val(L$)
         
         res% = high% * 256 + low%
         text1.Text = res%
         list1.AddItem Str$(high%) + "  " + Str$(low%)
  Select Case mytag%
    Case Is = 0 ' single reading
          GoTo ven
    Case Is = 1 'periodic
          Text2.Text = Str$(200 - pp%) + " readings left"
          pp% = pp% + 1
          Print #1, res%
      'handling List1 scroll:
       
          If myyy% > 21 Then list1.ListIndex = myyy%
          myyy% = myyy% + 1
      If pp% > 200 Then 'is measurement over?
          Timer2.Enabled = False
          pp% = 0
          Close #1
         For kk% = 0 To 3
           Command1(kk%).Enabled = 1
         Next kk%
          Text2.Visible = False
          Label3.Visible = False
          yy% = 0
          GoTo ven
      End If
          MSComm1.Output = Chr$(128 + mycall%)
   End Select
ven:
End Sub


Private Sub Timer3_Timer()
'timer interval for TEST
       answer$ = Form1.MSComm1.Input 'buffer emptied
       Timer3.Enabled = False
       rr$ = ""
     For tt% = 1 To 3  'counts
       testres$ = Mid$(answer$, tt%, 1)
       If testres$ = "" Then
          MsgBox "FAILURE"
          GoTo down
       End If

       Number% = Asc(testres$)
       rr$ = rr$ + Str$(Number%)
     Next tt%
       'CR and LF not taken
     For gg% = 6 To 7
       testres$ = Mid$(answer$, gg%, 1)
       Number% = Asc(testres$)
       rr$ = rr$ + Str$(Number%)
     Next gg%
        
        If rr$ = " 1 2 3 4 5" Then
         MsgBox "SYSTEM OK!"
        End If
        If rr$ <> " 1 2 3 4 5" Then
        MsgBox "FAILURE"
        End If
     
down:
     
End Sub





        
       

   


