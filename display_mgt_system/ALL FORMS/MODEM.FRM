VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "TEST"
   ClientHeight    =   6210
   ClientLeft      =   255
   ClientTop       =   1725
   ClientWidth     =   10800
   Height          =   6900
   Left            =   195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   10800
   Top             =   1095
   Width           =   10920
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
      Height          =   330
      Left            =   8040
      TabIndex        =   10
      Top             =   5565
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   240
      Top             =   1050
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4320
      TabIndex        =   9
      Text            =   "633"
      Top             =   2100
      Width           =   1695
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4320
      Top             =   1050
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   768
      Left            =   1800
      TabIndex        =   7
      Text            =   "CAN YOU HEAR ME?"
      Top             =   2940
      Width           =   2412
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   8040
      TabIndex        =   6
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "HOOK ON"
      Height          =   750
      Left            =   840
      TabIndex        =   5
      Top             =   3885
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   1800
   End
   Begin VB.CommandButton Command5 
      Caption         =   "HOOK OFF"
      Height          =   750
      Left            =   840
      TabIndex        =   4
      Top             =   1995
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SEND MESSAGE"
      Height          =   750
      Left            =   4320
      TabIndex        =   3
      Top             =   2940
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CALL"
      Height          =   750
      Left            =   2640
      TabIndex        =   2
      Top             =   1995
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TEST"
      Height          =   645
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   3600
      Top             =   4800
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   327680
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Type desired number"
      Height          =   330
      Left            =   4320
      TabIndex        =   11
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Message to be transmitted.Type and press SEND"
      Height          =   705
      Left            =   0
      TabIndex        =   8
      Top             =   2940
      Width           =   1815
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
Dim xxx As String
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
  Timer7.Enabled = True
  Do
    DoEvents
  Loop Until mytag% = 1
   mytag% = 0
  Timer7.Enabled = 0
End Sub

Private Sub Command1_Click()
'TEST
    MSComm2.Output = "AT"
    MSComm2.Output = Chr$(13) 'confirm

Call mydelay
    a$ = MSComm2.Input
    
   
End Sub



Private Sub Command2_Click()
  List1.Clear
End Sub

Private Sub Command3_Click()

     
MSComm2.Output = "ATL3" 'high speaker volume
MSComm2.Output = Chr$(13)
  Call mydelay
  x$ = MSComm2.Input
  List1.AddItem x$
 
MSComm2.Output = "ATM1" 'speaker on until connect
MSComm2.Output = Chr$(13)
   Call mydelay
 x$ = MSComm2.Input
 List1.AddItem x$
  numb$ = Text4.Text
  mycall$ = "ATDT" + numb$
  MSComm2.Output = mycall$ '
  MSComm2.Output = Chr$(13) ' carriage return, to confirm
   Call mydelay
List1.AddItem x$
 
End Sub

Private Sub Command4_Click()
'SEND MESSAGE
    mes$ = Text2.Text
    'List1.AddItem mes$
  If mes$ <> "" Then
    MSComm2.Output = mes$
    MSComm2.Output = Chr$(13)
  End If

 End Sub


Private Sub Command5_Click()
'HOOK OFF
 MSComm2.Output = "ATA" ' manual answer, goes off hook
 MSComm2.Output = Chr$(13)
  

End Sub

Private Sub Command6_Click()
'HOOK ON
 ' We have to wait at least 1 second
Call delay_1s
   MSComm2.Output = "+++"
Call delay_1s
   'MSComm2.Output = "ATH0"
   'MSComm2.Output = Chr$(13)
    text1.Text = "Transmition over"
End Sub

Private Sub Command7_Click()
 Timer6.Enabled = True
Label2.Visible = 1

End Sub

Private Sub confirm_Click()
 
 
    MSComm2.Output = "ATH0"   'put down
    MSComm2.Output = Chr$(13) 'confirm
  End
End Sub

Private Sub Form_Load()
   MSComm2.PortOpen = True
   MSComm2.Output = "ATM1" 'speaker on until connect
   MSComm2.Output = Chr$(13)
End Sub


Private Sub Timer1_Timer()
 'reading messages from COMM port.Running permanently
      xx$ = MSComm2.Input
      
   If xx$ <> "" Then
      List1.AddItem xx$
   End If
      
     mypos = InStr(1, x$, "RING", 0) 'Is "RING"?
   If mypos > 0 Then
       text1.Text = "CALL CAME. PRESS HOOK OFF"
       'recognize message "RING"
   End If
   mypos = InStr(1, x$, "CONNECT", 0)
       If mypos > 0 Then text1.Text = "We can talk.Try to send a message"
      'recognize message "connect"

End Sub


Private Sub Timer2_Timer()
'used to terminate communication with +++
  mytag1% = 1

End Sub














Private Sub Timer7_Timer()
  'used for breaks in dialing procedure
   mytag% = 1
End Sub


