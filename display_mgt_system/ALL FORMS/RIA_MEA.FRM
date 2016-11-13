VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFC0&
   Caption         =   "Measurement"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   1455
   ClientWidth     =   8685
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
   Height          =   7455
   Left            =   0
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   8685
   Top             =   1110
   Width           =   8805
   Begin VB.VScrollBar VScroll5 
      Enabled         =   0   'False
      Height          =   3495
      Left            =   3290
      Max             =   0
      Min             =   1200
      TabIndex        =   27
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "HV"
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TEST"
      Height          =   495
      Left            =   240
      TabIndex        =   25
      Top             =   720
      Width           =   1335
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8280
      Top             =   2520
   End
   Begin VB.VScrollBar VScroll4 
      Enabled         =   0   'False
      Height          =   3495
      Left            =   4080
      Max             =   0
      Min             =   60
      TabIndex        =   19
      Top             =   1440
      Width           =   375
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8280
      Top             =   2040
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   5490
      Left            =   6720
      TabIndex        =   18
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   8280
      Top             =   1560
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8280
      Top             =   1080
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   3495
      Left            =   5520
      Max             =   0
      Min             =   255
      TabIndex        =   11
      Top             =   1440
      Width           =   375
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   3495
      Left            =   6120
      Max             =   0
      Min             =   255
      TabIndex        =   10
      Top             =   1440
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3495
      Left            =   4680
      Max             =   0
      Min             =   60
      TabIndex        =   8
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton command6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CHANNEL"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PRESET TIME"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "START"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "READ"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "STOP"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFC0&
      Caption         =   "VOLTS"
      Height          =   255
      Left            =   3120
      TabIndex        =   30
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "P"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   29
      Top             =   2520
      Width           =   255
   End
   Begin VB.Line Line3 
      X1              =   3480
      X2              =   3600
      Y1              =   4920
      Y2              =   5040
   End
   Begin VB.Line Line2 
      X1              =   3480
      X2              =   3360
      Y1              =   4920
      Y2              =   5040
   End
   Begin VB.Line Line1 
      X1              =   3480
      X2              =   3480
      Y1              =   4920
      Y2              =   5160
   End
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   1560
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   28
      Top             =   960
      Width           =   810
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8160
      Top             =   360
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
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "COUNTS READINGS"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6720
      TabIndex        =   24
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "SELECTED"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3480
      TabIndex        =   23
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "TIMER COUNTER"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3840
      TabIndex        =   21
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "time,min"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "width"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "level"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "time,sec"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "COUNTS"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   5640
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()

   
        a$ = "C"
        b$ = "E"
        c$ = "S"
        d$ = "G"
  
'First action: CLEAR all counters =a$
       Form1.MScomm1.Output = a$

'Second action: PRESET time; out: "E" =b$               return nothing
       Form1.MScomm1.Output = b$
       vscroll1.Enabled = False
       vscroll4.Enabled = False

       totaltime% = mymin% * 60 + mysec%
       timehigh% = totaltime% \ 256
       timelow% = totaltime% Mod 256
     If timelow% = 0 Then
       MsgBox "TIME NOT SELECTED!"
       Exit Sub
     End If
       Form1.MScomm1.Output = Chr$(timelow% - 1)
       Form1.MScomm1.Output = Chr$(timehigh%)

       command1.Enabled = False
       Command2.Enabled = True
       Command2.BackColor = 110022
       command3.Enabled = False
       command4.Enabled = False
       command5.Enabled = False
       command6.Enabled = False
       

'Command to uC to start counting is "S" or "s":
      
       Form1.MScomm1.Output = c$  'START
    
'Status checking  command is "G" or "g" =d$
      
       Form1.MScomm1.Output = d$
       Text2.Visible = True
       timer5.Enabled = True  'timer will check status.
End Sub

Private Sub Command2_Click()
'Stop counting; out: "P"(master)  or "p" (slave)
'               return: nothing
       command1.Enabled = True
       Command2.Enabled = True
       command3.Enabled = True
       command4.Enabled = True
       command5.Enabled = True
       command6.Enabled = True
    
          a$ = "P"
   
       
       Form1.MScomm1.Output = a$
       Timer4.Enabled = False
       Text2.Visible = False
End Sub

Private Sub command3_Click()
     Text3.Text = ""

'Clear counter; out: "C"(master)
'               return nothing
   
        a$ = "C"
    
    
    Form1.MScomm1.Output = a$
    localtime% = 0
End Sub

Private Sub Command4_Click()

'Reads the collected counts number; out: "R"
' return: 3 digits for counts,
'CR LF,
'2 digits for result: seconds, and seconds*256
'Initiate reading by "R" :
    
        a$ = "R"
    
    
    
    MScomm1.Output = a$
    Timer3.Enabled = True
 'here we jump to timer 3

End Sub

Private Sub Command5_Click()
'SET TIME ROUTINE
'In: time in minutes, time in seconds
'out: time in sec from 0 to &hFFFF, first low byte, then high byte
'coded as ASCII
'no return
       vscroll1.Enabled = True
       vscroll4.Enabled = True
       vscroll2.Enabled = False
       vscroll3.Enabled = False
   Static aa%
    aa% = aa% Xor 1
   Select Case aa%
    Case Is = 1
       command1.Enabled = False
       Command2.Enabled = False
       command3.Enabled = False
       command4.Enabled = False
       command5.Caption = "ACCEPT"
       command6.Enabled = False
    Case Is = 0
       command1.Enabled = True
       Command2.Enabled = True
       command3.Enabled = True
       command4.Enabled = True
       command5.Caption = "PRESET TIME"
       command6.Enabled = True
     
         a$ = "E"   ' message to  uC : prEset
    
       
       Form1.MScomm1.Output = a$
       vscroll1.Enabled = False
       vscroll4.Enabled = False
       totaltime% = mymin% * 60 + mysec%
       timehigh% = totaltime% \ 256
       timelow% = totaltime% Mod 256
    If timelow% = 0 Then
        MsgBox "Select time."
        Exit Sub
    End If
       Form1.MScomm1.Output = Chr$(timelow% - 1)
       Form1.MScomm1.Output = Chr$(timehigh%)
    End Select

End Sub

Private Sub Command6_Click()
'SET CHANNEL LEVEL, WINDOW
   vscroll1.Enabled = False
   vscroll2.Enabled = True
   vscroll3.Enabled = True
   Static bb%
    bb% = bb% Xor 1
   Select Case bb%
    Case Is = 1
       command1.Enabled = False
       Command2.Enabled = False
       command3.Enabled = False
       command4.Enabled = False
       command5.Enabled = False
       command6.Caption = "ACCEPT"
       
    Case Is = 0
       command1.Enabled = True
       Command2.Enabled = True
       command3.Enabled = True
       command4.Enabled = True
       command5.Enabled = True
       command6.Caption = "CHANNEL"

     
           a$ = "B"    'setting LEVEL
           b$ = "A"    'setting CHWIDTH
    

       vscroll2.Enabled = False
       vscroll3.Enabled = False
       Form1.MScomm1.Output = a$
       Form1.MScomm1.Output = Chr$(level%)
       Form1.MScomm1.Output = b$
       corchwidth% = Int(chwidth% * 1.33)
       If corchwidth% > 255 Then corchwidth% = 255
       Form1.MScomm1.Output = Chr$(chwidth%)
       
    End Select
       

End Sub

Private Sub Command7_Click()
      
           a$ = "T"
      
       Form1.MScomm1.Output = a$
       timer1.Enabled = True

End Sub

Private Sub Command8_Click()
  
     
End Sub


Private Sub Command9_Click()
 Static bb%
  ff = 255 / 1200
  VScroll5.Enabled = True
  Command9.Caption = "ACCEPT HV"
  bb% = bb% Xor 1
  If bb% = 0 Then
    Command9.Caption = "HV"
'activating  high voltage settings
             MScomm1.Output = "H"
'value of the high voltage send out:
            volt% = Int(lhivo% * ff)
            RIA_SET.Text7.Text = lhivo%
            MScomm1.Output = Chr$(volt%)
            VScroll5.Enabled = False
     End If
End Sub

Private Sub Form_Load()
       
       Label2.FontSize = 13.5
       vscroll1.Enabled = False
       vscroll2.Enabled = False
       vscroll3.Enabled = False
      
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Cancel = -1
    Form1.Visible = False
 
End Sub


Private Sub Timer1_Timer()
'timer interval for waiting reply for  TEST
       answer$ = MScomm1.Input 'buffer emptied
       timer1.Enabled = False  'TIMER1 OFF!!!
       rr$ = ""
     For tt% = 1 To 3  'counts
       testres$ = Mid$(answer$, tt%, 1)
       If testres$ = "" Then
         Text3.Text = "FAILURE"
         Exit Sub
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
        Text3.Text = rr$ + "  expexted: 12345"
        'expected string "12345"

End Sub

Private Sub Timer3_Timer()
    'Here we read results:
    Static counts(10)
    instring$ = Form1.MScomm1.Input
    Timer3.Enabled = False  'TIMER1 OFF!!
 'First counts:
     For tt% = 1 To 3
       testres$ = Mid$(instring$, tt%, 1)
       If testres$ = "" Then
          Text3.Text = "FAILURE"
          Exit Sub
       End If

       counts(tt%) = Asc(testres$)
     Next tt%
       totcounts = counts(1) + 256 * counts(2) + (256 ^ 2) * counts(3)
    
 'Second time:
     For tt% = 6 To 7
       testres$ = Mid$(instring$, tt%, 1)
       counts(tt%) = Asc(testres$)
     Next tt%
       tottime = counts(6) + 256 * counts(7)
 'Last  messages:
       Text3.Text = ""

       Text3.Text = Str$(totcounts) + Space$(20) + Str$(tottime)
     
      list1.AddItem totcounts
End Sub

Private Sub Timer4_Timer()
     Static counts(10)
     'looking 3 times per second for status
   
         a$ = "G"
      
    
     MScomm1.Output = a$
     a$ = MScomm1.Input

 If a$ = "S" Then
    'Here we read results:
    
    instring$ = Form1.MScomm1.Input
    Timer3.Enabled = False
 'First counts:
     For tt% = 1 To 3
       testres$ = Mid$(instring$, tt%, 1)
       counts(tt%) = Asc(testres$)
     Next tt%
       totcounts = counts(1) + 256 * counts(2) + (256 ^ 2) * counts(3)
    
 'Second time:
     For tt% = 6 To 7
       testres$ = Mid$(instring$, tt%, 1)
       counts(tt%) = Asc(testres$)
     Next tt%
       tottime = counts(6) + 256 * counts(7)
 'Last  messages:
       Text3.Text = ""
       Text3.Text = Str$(totcounts) + Space$(20) + Str$(tottime)
 End If
End Sub

Private Sub Timer5_Timer()
  
    'timer waits 500ms; and here is an answer:
    instring$ = MScomm1.Input
    'and this is information: Start,stoP,
    'Is the preset time over, too?
    label16.Caption = Left$(instring$, 1)
    If Left$(instring$, 1) = "P" Then
      Beep
      timer5.Enabled = False 'TIMER5 OFF when preset = 0
      Call Command2_Click
      Exit Sub
    End If
    
         MScomm1.Output = "G"
    


End Sub

Private Sub VScroll1_Change()
      Label2.Caption = vscroll1.Value
      mysec% = vscroll1.Value
      Label2.Visible = True
End Sub

Private Sub VScroll1_Scroll()
    'this routine updates, only
       Label2.Caption = vscroll1.Value
End Sub

Private Sub VScroll2_Change()
   label3.Caption = vscroll2.Value
   level% = vscroll2.Value

End Sub

Private Sub VScroll2_Scroll()
   label3.Caption = vscroll2.Value
   level% = vscroll2.Value
End Sub

Private Sub VScroll3_Change()
     label4.Caption = vscroll3.Value
     chwidth% = vscroll3.Value

End Sub

Private Sub VScroll3_Scroll()
     label4.Caption = vscroll3.Value
     chwidth% = vscroll3.Value

End Sub

Private Sub VScroll4_Change()
      Label9.Caption = vscroll4.Value
      mymin% = vscroll4.Value
      Label9.Visible = True

End Sub

Private Sub VScroll4_Scroll()
    'this routine updates, only
       Label9.Caption = vscroll4.Value

End Sub

Private Sub VScroll5_Change()
      lhivo% = VScroll5.Value
      label15 = Str$(lhivo%)
End Sub


Private Sub VScroll5_Scroll()
      lhivo% = VScroll5.Value
      label15 = Str$(lhivo%)
End Sub


