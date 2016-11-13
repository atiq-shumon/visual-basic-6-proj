VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "535 DATA COLLECTION AND SAVING"
   ClientHeight    =   5940
   ClientLeft      =   1140
   ClientTop       =   2820
   ClientWidth     =   6690
   Height          =   6630
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Top             =   2190
   Width           =   6810
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6210
      Top             =   0
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   3510
      TabIndex        =   5
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   540
      TabIndex        =   2
      Text            =   "Text7"
      Top             =   2520
      Width           =   1635
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   540
      TabIndex        =   1
      Text            =   "Text6"
      Top             =   1560
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RUN"
      Height          =   615
      Left            =   540
      TabIndex        =   0
      Top             =   3120
      Width           =   960
   End
   Begin VB.Label Label4 
      Caption         =   "Results are in blocks of 128"
      Height          =   345
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   2085
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6210
      Top             =   1080
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Label Label3 
      Caption         =   "x 100 us"
      Height          =   375
      Left            =   2295
      TabIndex        =   6
      Top             =   1680
      Width           =   690
   End
   Begin VB.Label Label2 
      Caption         =   "Enter time:"
      Height          =   255
      Left            =   540
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Enter number"
      Height          =   375
      Left            =   540
      TabIndex        =   3
      Top             =   2280
      Width           =   960
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6210
      Top             =   480
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
      Begin VB.Menu confirm 
         Caption         =   "confirm"
      End
   End
   Begin VB.Menu test 
      Caption         =   "TEST"
      Begin VB.Menu ok 
         Caption         =   "do it!"
      End
   End
   Begin VB.Menu file 
      Caption         =   "FILE"
      Begin VB.Menu save 
         Caption         =   "save"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim mem$(8000)
Dim myresult(8000)
Public max%



Private Sub Dir1_Change()
    File1.Path = Dir1.Path  ' Set File1 path.
End Sub
Private Sub File1_PathChange()
Beep
    Dir1.Path = File1.Path  ' Set Dir1 path.
End Sub
Private Sub File1_DblClick()
    ' Display the selected filename when double-clicked.
    Label4.Caption = "Your selection: " + UCase(File1.filename)
End Sub


Private Sub Command1_Click()
Static jj%
train_select% = 0

a$ = Chr$(&H70 + train_select%) 'acquire activation
      
  For kk = 1 To 7000: Next kk 'wait loop
  'TIME: N* 100 useconds
      cc% = Val(Text6.Text)
      b$ = Chr$(cc%)
     
  'SAMPLE NUMBER (0 = max  to 127 =min)
    numb% = Val(Text7.Text)
    numb1% = numb% \ 128 'number of blocks of 256 doubled
    ' there are two readings for one result
    numb% = numb1% * 128 ' corrected number
    c$ = Chr$(numb1%)
    Text7.Text = numb% 'corrected number
    For kk = 1 To 7000: Next kk 'waiting loop
         MSComm1.InputLen = 1
 'instruction prepared: channel, spacing and number of samples
         a$ = Chr$(&H70) + b$ + c$
         MSComm1.Output = a$
         
     'waiting for "J"; end of run:
top:
        D$ = MSComm1.Input
           If D$ = "J" Then GoTo over
        GoTo top
over:
        

   'time to read. Instruction is:
    a$ = "D"

       MSComm1.Output = a$
    'waiting for answers
    'reading one after another:
         
     For GG% = 0 To numb% - 2
back:
         mem$(GG%) = MSComm1.Input
        If mem$(GG%) = "" Then GoTo back
     Next GG%
 For jj% = 0 To numb% - 4
         high% = Asc(mem$(jj%))
         jj% = jj% + 1
         low% = Asc(mem$(jj%))
         result% = low% + 256 * high%
         myresult(jj%) = result%
       
       If jj% > 56 Then
         List1.ListIndex = (jj% \ 2) - 1
       End If
         List1.AddItem result%
    Next jj%
         max% = jj% 'number of readings
         MSComm1.InputLen = 0
 Exit Sub

 
        

End Sub


Private Sub confirm_Click()
    End
End Sub

Private Sub Form_Load()
MSComm1.PortOpen = True
End Sub

Private Sub ok_Click()
      a$ = "T"
      MSComm1.Output = a$
      Timer1.Enabled = True
End Sub


Private Sub save_Click()
 CommonDialog1.ShowSave
  filename$ = CommonDialog1.filename$

  Open filename$ For Output As #1
For kk% = 0 To max%
 Print #1, myresult(kk%)

 Next kk%
  Close
End Sub

Private Sub Timer1_Timer()
'timer interval for TEST
       answer$ = Form1.MSComm1.Input 'buffer emptied
       Timer1.Enabled = False
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
     For GG% = 6 To 7
       testres$ = Mid$(answer$, GG%, 1)
       Number% = Asc(testres$)
       rr$ = rr$ + Str$(Number%)
     Next GG%
        
        If rr$ = " 1 2 3 4 5" Then
         MsgBox "SYSTEM OK!"
        End If
        If rr$ <> " 1 2 3 4 5" Then
        MsgBox "FAILURE"
        End If
        
       
down:
        Timer1.Enabled = False
           
End Sub


