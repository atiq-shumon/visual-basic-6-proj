VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Single channel spectroscopy"
   ClientHeight    =   5580
   ClientLeft      =   675
   ClientTop       =   2505
   ClientWidth     =   5865
   Height          =   5985
   Left            =   615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   5865
   Top             =   2160
   Width           =   5985
   Begin VB.CommandButton Command4 
      Caption         =   "TERMINATE"
      Height          =   540
      Left            =   720
      TabIndex        =   8
      Top             =   3255
      Width           =   1215
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2640
      Top             =   1050
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2040
      Top             =   1050
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   240
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   3840
      TabIndex        =   4
      Top             =   210
      Width           =   1590
   End
   Begin VB.CommandButton Command3 
      Caption         =   "START"
      Enabled         =   0   'False
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1365
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   2025
      TabIndex        =   2
      Text            =   "800"
      Top             =   480
      Width           =   705
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE AS..."
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2205
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SELECT HV"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   435
      Left            =   240
      Shape           =   3  'Circle
      Top             =   1365
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      Height          =   495
      Left            =   2835
      TabIndex        =   7
      Top             =   480
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Results:"
      Height          =   225
      Left            =   4320
      TabIndex        =   6
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter high voltage:"
      Height          =   375
      Left            =   2025
      TabIndex        =   5
      Top             =   120
      Width           =   1500
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   3255
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2205
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      CancelError     =   -1  'True
      FilterIndex     =   1
      InitDir         =   "c:\word"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public nn% ' count number of reading



Private Sub Command1_Click()
    hivo% = Val(Text1.Text)
    If hivo% < 300 Then
        MsgBox "Select reasonable voltage"
        Exit Sub
        End If
    If hivo% > 1200 Then
       MsgBox "Are you sure to set it to " + Str$(hivo%) + "?"
    Exit Sub
    End If
 'convert volts into byte:
    req% = Int((hivo% / 1200) * 256)
    MSComm1.Output = "H"
    MSComm1.Output = Chr$(req%)
    label3.Caption = "voltage " + Str$(hivo%) + " applied.  "

 'channel width selection:
     MSComm1.Output = "A"
     MSComm1.Output = Chr$(11) ' 208 mV width set
 'starting conditions fulfilled. Unclock next button;
 'next button can be activated also by ENTER:
    Command3.Enabled = True
    Command3.SetFocus
End Sub


Private Sub Command2_Click()
 CommonDialog1.ShowSave
 filename$ = CommonDialog1.filename$
 Open filename$ For Output As #1
 For jj% = 0 To 31  'array elements into file:
     Print #1, myres&(jj%)
  Next jj%
 Close #1
End Sub


Private Sub Command3_Click()
'at zero level we pick only noise, therefore we start at 1
   nn% = 1
'set level:
   MSComm1.Output = "B" ' call level
   MSComm1.Output = Chr$(nn% * 8) 'level value
  
'clear all, then preset time:
   MSComm1.Output = "C"
'timer inside RIA used. It is more precise than VB
   MSComm1.Output = "E"
   MSComm1.Output = Chr$(1) 'seconds; low
   MSComm1.Output = Chr$(0) 'seconds, high
 'start:
   MSComm1.Output = "S"
   Shape1.BackColor = QBColor(10) 'light from red to green
   MSComm1.Output = "G" ' get status
   Timer2.Enabled = True
  
  
End Sub

Private Sub Command4_Click()
  End
End Sub

Private Sub Form_Load()
 '30 measurements of one second made
 'through full energy range

'Opening Port:
    MSComm1.CommPort = 1
    MSComm1.PortOpen = True
  
 'checking communication
    MSComm1.Output = "T"
    Timer1.Enabled = True

End Sub


Private Sub Timer1_Timer()
  'delay: waiting the TEST answer
 Timer1.Enabled = False
   A$ = MSComm1.Input
 'checking the answer,
   b$ = Mid$(A$, 1, 1) 'string,start,length
   c$ = Mid$(A$, 2, 1)
   d$ = Mid$(A$, 3, 1)
   test% = Asc(b$) + Asc(c$) + Asc(d$)
   If test% = 6 Then MsgBox "SYSTEM OK"
  
End Sub


Private Sub Timer2_Timer()
'checking RUN/STOP
   r$ = MSComm1.Input

     If Left$(r$, 1) = "P" Then   'run finished
        Shape1.BackColor = QBColor(12) 'stop, therefore red
        Timer2.Enabled = False 'counting terminated
        MSComm1.Output = "R"   'asking for result
        timer3.Enabled = True  'delay
        Exit Sub
     End If
     MSComm1.Output = "G" 'next trial
End Sub


Private Sub timer3_Timer()
 'delay:waiting for result after R(ead)command
timer3.Enabled = False 'time to get result
   my$ = MSComm1.Input
 'converting reading into numbers
   hi& = Asc(Mid$(my$, 3, 1)) 'string,start,length
   mi& = Asc(Mid$(my$, 2, 1))
   lo& = Asc(Mid$(my$, 1, 1))
 'calculating result
   myres&(nn%) = 256 ^ 2 * hi& + 256 * mi& + lo&
If nn% > 26 Then List1.Index = nn% 'scrolling at bottom
List1.AddItem Str$(nn%) + "  " + Str$(myres&(nn%))
     
     nn% = nn% + 1
  If nn% < 32 Then
     MSComm1.Output = "C" 'clear before new measur.
     MSComm1.Output = "E" + Chr$(1) + Chr$(0) 'preset 1 sec
     MSComm1.Output = "S" 'start
     Shape1.BackColor = QBColor(10) 'light:red to green
     Timer2.Enabled = True
  End If
End Sub


