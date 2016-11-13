VERSION 4.00
Begin VB.Form RIA_SPEC 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "SPECTRUM"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   1695
   ClientWidth     =   10815
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
   Height          =   6810
   Left            =   60
   LinkTopic       =   "Form3"
   ScaleHeight     =   6405
   ScaleWidth      =   10815
   Top             =   1350
   Width           =   10935
   Begin VB.CommandButton Command3 
      Caption         =   "STOP"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   26
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "...here"
      Height          =   495
      Left            =   9600
      TabIndex        =   21
      Top             =   5280
      Width           =   975
   End
   Begin VB.HScrollBar Cursor 
      Height          =   255
      Index           =   1
      Left            =   840
      Max             =   0
      Min             =   100
      TabIndex        =   18
      Top             =   5880
      Width           =   8175
   End
   Begin VB.HScrollBar Cursor 
      Height          =   255
      Index           =   0
      Left            =   850
      Max             =   100
      TabIndex        =   17
      Top             =   5400
      Width           =   8175
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4440
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8760
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   8760
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RUN"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "RIA_SPEC.frx":0000
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "RIA_SPEC.frx":0004
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "255"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8700
      TabIndex        =   24
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Define selected channel by pressing...."
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   9600
      TabIndex        =   22
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right Cursor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9120
      TabIndex        =   20
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left Cursor"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   50
      TabIndex        =   19
      Top             =   5400
      Width           =   855
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   8640
      X2              =   8640
      Y1              =   2160
      Y2              =   5880
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1210
      X2              =   1210
      Y1              =   5400
      Y2              =   2160
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "channel number"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Channel counting time (sec)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Last channel counts"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Channel width (0-255 scale)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Level (0 to 255 scale):"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "counts/channel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "CALIBRATION OVER"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "How long to count? "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "How many channels?"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin GraphLib.Graph Graph1 
      Height          =   2655
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   9135
      _Version        =   65536
      _ExtentX        =   16113
      _ExtentY        =   4683
      _StockProps     =   96
      BorderStyle     =   1
      GraphStyle      =   2
      LabelEvery      =   5
      NumPoints       =   30
      RandomData      =   0
      ThickLines      =   0
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontSize        =   4
      FontSize[0]     =   200
      FontSize[1]     =   150
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      GraphData       =   1
      GraphData[]     =   30
      GraphData[0,0]  =   0
      GraphData[0,1]  =   0
      GraphData[0,2]  =   0
      GraphData[0,3]  =   0
      GraphData[0,4]  =   0
      GraphData[0,5]  =   0
      GraphData[0,6]  =   0
      GraphData[0,7]  =   0
      GraphData[0,8]  =   0
      GraphData[0,9]  =   0
      GraphData[0,10] =   0
      GraphData[0,11] =   0
      GraphData[0,12] =   0
      GraphData[0,13] =   0
      GraphData[0,14] =   0
      GraphData[0,15] =   0
      GraphData[0,16] =   0
      GraphData[0,17] =   0
      GraphData[0,18] =   0
      GraphData[0,19] =   0
      GraphData[0,20] =   0
      GraphData[0,21] =   0
      GraphData[0,22] =   0
      GraphData[0,23] =   0
      GraphData[0,24] =   0
      GraphData[0,25] =   0
      GraphData[0,26] =   0
      GraphData[0,27] =   0
      GraphData[0,28] =   0
      GraphData[0,29] =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
End
Attribute VB_Name = "RIA_SPEC"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Static mybase%
  Static repmax%
  Static dec_increment
  command1.Enabled = False

   repmax% = Val(Text1.Text)
   label3.Visible = False
 'Is measurement over?
 If repetition% = repmax% Then
    repetition% = 0
    command1.Enabled = True
    label3.Visible = True
    Exit Sub
 End If
 'This part runs first time only:
 If repetition% > 0 Then GoTo over
     graph1.DrawMode = 3
     graph1.NumPoints = repmax%
     graph1.RandomData = False
     graph1.DataReset = 9
     cnt_time = Val(Text2.Text)
     dec_increment = 255 / Val(Text1.Text)
     mybase% = Int(dec_increment)
     chwidth% = Int(dec_increment)
     
        Form1.MScomm1.Output = "A"
    

     corchwidth% = Int(chwidth% * 1.33)
     If corchwidth% > 255 Then corchwidth% = 255
     Form1.MScomm1.Output = Chr$(corchwidth%)
     
    
over:
     
     level% = Int(mybase% + repetition% * dec_increment)
     repetition% = repetition% + 1
     If level% > 255 Then level% = 255
'Here we go to setting level and channel width
      
       Form1.MScomm1.Output = "B"
      

       Form1.MScomm1.Output = Chr$(level%)
 
'FIRST ACTION:Clear counter; out: "C"
'               return nothing
   
         a$ = "C"
   
    
    Form1.MScomm1.Output = a$
    localtime% = 0

'SECOND ACTION: Preset counter


'Start counting; out : "S"
'                return no
'In: time in minutes, time in seconds
'out: time in sec from 0 to &hFFFF, first low byte, then high byte
'coded as ASCII
'no return
     
         a$ = "E"
         b$ = "S"
         c$ = "G"
       

       Form1.MScomm1.Output = a$
       timehigh% = 0
       timelow% = cnt_time%
       Form1.MScomm1.Output = Chr$(timelow% - 1)
       Form1.MScomm1.Output = Chr$(timehigh%)

'Command to uC to start counting:
       Form1.MScomm1.Output = b$  'START
    
'Status checking  command is "G"or "g":
       Form1.MScomm1.Output = c$
      timer1.Enabled = True
       'timer will check status in 0.3 sec intervals.
      command3.Enabled = True 'stop b utton activated
End Sub

Private Sub Command2_Click()
     low% = Cursor(0).Value
     high% = Cursor(1).Value
     level% = Int(2.55 * low%)
     chwidth% = Int((100 - low% - high%) * 2.55)
    
     Form1.MScomm1.Output = "B"
   

     Form1.MScomm1.Output = Chr$(level%)
    
      Form1.MScomm1.Output = "A"
    

     corchwidth% = Int(chwidth% * 1.33)
     If corchwidth% > 255 Then corchwidth% = 255
     Form1.MScomm1.Output = Chr$(corchwidth%)
   
        RIA_SET.Text5.Text = level%
        RIA_SET.Text6.Text = chwidth%
        
        
    
        
   
    
     RIA_SPEC.Hide
     RIA_SET.Show
     
End Sub


  


Private Sub command3_Click()

    repetition% = 0
    command1.Enabled = True
    'label3.Visible = True
    command3.Enabled = False
    Exit Sub
 
End Sub

Private Sub command3_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

End Sub


Private Sub Cursor_Change(Index As Integer)
    'minwidth% is the minimal channel width
      minwidth% = 90
      If Cursor(Index).Value + Cursor(Index Xor 1).Value > minwidth% Then
         Cursor(Index).Value = minwidth% - Cursor(Index Xor 1).Value
       End If
      Select Case Index
       Case Is = 0
        Line1(Index).X1 = Cursor(Index).Value * 74.5 + 1200
        Line1(Index).X2 = Cursor(Index).Value * 74.5 + 1200
       Case Is = 1
        Line1(Index).X1 = -Cursor(Index).Value * 74.5 + 500 + 8150
        Line1(Index).X2 = -Cursor(Index).Value * 74.5 + 500 + 8150
     End Select

End Sub

Private Sub Cursor_Scroll(Index As Integer)
    'minwidth% is the minimal channel width
      minwidth% = 90
      If Cursor(Index).Value + Cursor(Index Xor 1).Value > minwidth% Then
         Cursor(Index).Value = minwidth% - Cursor(Index Xor 1).Value
       End If
      Select Case Index
       Case Is = 0
        Line1(Index).X1 = Cursor(Index).Value * 74.5 + 1200
        Line1(Index).X2 = Cursor(Index).Value * 74.5 + 1200
       Case Is = 1
        Line1(Index).X1 = -Cursor(Index).Value * 74.5 + 500 + 8150
        Line1(Index).X2 = -Cursor(Index).Value * 74.5 + 500 + 8150
     End Select

End Sub

Private Sub Form_Load()
       Cursor(0).Value = llevel% / 2.55
       Cursor(1).Value = 100 - Int((llevel% + lchwidth%) / 2.55)
       
       Form1.MScomm1.Output = "B"
       Form1.MScomm1.Output = Chr$(llevel%)
       Form1.MScomm1.Output = "A"
       lcorchwidth% = Int(lchwidth% * 1.33)
       If lcorchwidth% > 255 Then lcorchwidth% = 255
       Form1.MScomm1.Output = Chr$(lcorchwidth%)

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text2_DblClick()

End Sub


Private Sub Text2_GotFocus()

End Sub


Private Sub Timer1_Timer()
    
        a$ = "G"
        b$ = "R"
     
   
    'timer waits 300ms; and here is an answer:
    instring$ = Form1.MScomm1.Input
    'Start,stoP: "S" or "P" is an answer
    'Is the preset time over?
    'if not yet, question is repeated:
    If Left$(instring$, 1) = "S" Then
'we check the status again by command "G":
      Form1.MScomm1.Output = a$
      timer1.Enabled = True
'timer will check status in 0.3 sec intervals.
      Exit Sub
    End If
'here measurement is over
    If Left$(instring$, 1) = "P" Then
      timer1.Enabled = False
      Beep
    End If
'Reads the collected counts number; out: "R"
' return: 3 digits for counts,
'CR LF,
'2 digits for result: seconds, and seconds*256
 'Initiate reading by "R":
    
    Form1.MScomm1.Output = b$
    timer2.Enabled = True
 'here we jump to Timer 2, the same FRM

End Sub

Private Sub Timer2_Timer()
    'Here we read results:
    Static counts(10)
    instring$ = Form1.MScomm1.Input
    If instring$ = "" Then
      MsgBox "System failure"
      timer1.Enabled = False
      timer2.Enabled = False
      command1.Enabled = True
      Exit Sub
    End If
 'First the counts:
     For tt% = 1 To 3
       testres$ = Mid$(instring$, tt%, 1)
       counts(tt%) = Asc(testres$)
     Next tt%
       totcounts = counts(1) + 256 * counts(2) + (256 ^ 2) * counts(3)
    
 'Second the :time:
     For tt% = 6 To 7
       testres$ = Mid$(instring$, tt%, 1)
       counts(tt%) = Asc(testres$)
     Next tt%
       tottime = counts(6) + 256 * counts(7)
 'Last  messages:TIME left
       ourtime = Str$(65536 - tottime)
     If ourtime = 65536 Then
          ourtime = 0
          timer2.Enabled = False
     End If
       Text4.Text = chwidth%
       Text5.Text = totcounts
       Text3.Text = level%
       Text6.Text = cnt_time%
       graph1.GraphData = totcounts
      'graph1.ThisPoint = repetition% + 1
       graph1.DrawMode = 3
       timer2.Enabled = False
     Call Command1_Click
End Sub

