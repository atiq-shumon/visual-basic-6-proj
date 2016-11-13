VERSION 4.00
Begin VB.Form RIA_BASE 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "RIA"
   ClientHeight    =   6765
   ClientLeft      =   105
   ClientTop       =   1605
   ClientWidth     =   11325
   FillColor       =   &H00000006&
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
   Left            =   45
   LinkTopic       =   "Form1"
   ScaleHeight     =   28.187
   ScaleMode       =   4  'Character
   ScaleWidth      =   94.375
   Top             =   975
   Width           =   11445
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00E0E0E0&
      Height          =   1155
      Left            =   5280
      Style           =   1  'Simple Combo
      TabIndex        =   44
      Text            =   "Select!"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      Height          =   1545
      Left            =   5280
      Style           =   1  'Simple Combo
      TabIndex        =   43
      Text            =   "Select!"
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Result?"
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
      Left            =   5160
      TabIndex        =   37
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   41
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   40
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   39
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      MultiLine       =   -1  'True
      TabIndex        =   36
      Top             =   4920
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7800
      MultiLine       =   -1  'True
      TabIndex        =   34
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   0
      Top             =   2640
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   1920
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   1320
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   495
      Left            =   7800
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GO!"
      Height          =   495
      Left            =   5160
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   0
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK?"
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   2880
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
      Height          =   420
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox Text5 
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
      Left            =   3720
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text4 
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
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text3 
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
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2400
      Width           =   855
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
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Line Line6 
      Visible         =   0   'False
      X1              =   92
      X2              =   92
      Y1              =   0.5
      Y2              =   25
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Taken at:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9600
      TabIndex        =   42
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "SECOND B"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7800
      TabIndex        =   35
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "CurrentTime:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2520
      TabIndex        =   33
      Top             =   3576
      Width           =   1095
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date:"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   3570
      Width           =   495
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "SELECT SAMPLE"
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
      Left            =   4560
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "SELECT TIME"
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
      Left            =   4680
      TabIndex        =   28
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "   MEASUREMENT OVER"
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
      Height          =   735
      Left            =   1320
      TabIndex        =   27
      Top             =   4080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "File Name"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line5 
      Visible         =   0   'False
      X1              =   79
      X2              =   79
      Y1              =   0.5
      Y2              =   25
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   64
      X2              =   64
      Y1              =   0.5
      Y2              =   25
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   92
      X2              =   42
      Y1              =   25
      Y2              =   25
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   42
      X2              =   42
      Y1              =   25
      Y2              =   0.5
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   42
      X2              =   92
      Y1              =   0.5
      Y2              =   0.5
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "SECOND A"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7800
      TabIndex        =   24
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "FIRST B"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7800
      TabIndex        =   23
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "FIRST A"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7800
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "seconds to END:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7800
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "SELECTED"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "RUN:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "SELECTED"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "COUNT TIME:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Day"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Month"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Born year"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Father's Name"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   240
      TabIndex        =   6
      Top             =   1605
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu New 
         Caption         =   "New"
      End
      Begin VB.Menu view 
         Caption         =   "View"
      End
      Begin VB.Menu saveme 
         Caption         =   "Save"
      End
      Begin VB.Menu edit 
         Caption         =   "Edit"
      End
      Begin VB.Menu report 
         Caption         =   "Print"
      End
      Begin VB.Menu dosshell 
         Caption         =   "DosShell"
      End
   End
   Begin VB.Menu Measure 
      Caption         =   "Measure"
      Begin VB.Menu set 
         Caption         =   "Set..."
      End
      Begin VB.Menu domeas 
         Caption         =   "Do..."
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
   End
   Begin VB.Menu cal 
      Caption         =   "Calibrate"
      Begin VB.Menu spectrum 
         Caption         =   "Spectrum"
      End
      Begin VB.Menu docal 
         Caption         =   "Detail tests..."
      End
   End
   Begin VB.Menu test 
      Caption         =   "Test"
      Begin VB.Menu go_sim 
         Caption         =   "do it!"
         Index           =   0
      End
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
      Begin VB.Menu ex 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu helproot 
      Caption         =   "Help"
      Begin VB.Menu help 
         Caption         =   "?"
      End
   End
End
Attribute VB_Name = "RIA_BASE"
Attribute VB_Creatable = False
Attribute VB_Exposed = False



Private Sub Command1_Click()
   If FileName$ = "" Then
      MsgBox "File name not selected"
      name_sel.Show
      Exit Sub
   End If
   If lchwidth% = 0 Or llevel% = 0 Then
      MsgBox "Set SOURCE or LEVEL and CHANNEL WIDTH. OPEN WINDOW MEASURE"
      
      
      Exit Sub
   End If
   
   command1.Visible = False
   command3.Visible = True
   text10.Visible = True
   text10.Text = FileName$
   Text1.Enabled = False
   Text2.Enabled = False
   Text3.Enabled = False
   Text4.Enabled = False
   Text5.Enabled = False

   label15.Visible = True
   Label7.Visible = True
   Label8.Visible = True
   Label9.Visible = True
   Label10.Visible = True
   Command2.Visible = True
   Combo1.Visible = True
   Combo2.Visible = True
   
   Line1.Visible = True
   Line2.Visible = True
   Line3.Visible = True
   Line4.Visible = True
   Line5.Visible = True
   line6.Visible = True
   Text6.Visible = True
   Text7.Visible = True
   Text8.Visible = True
   Text9.Visible = True
   Text13.Visible = True
   Text14.Visible = True
   Text15.Visible = True
   Text16.Visible = True
   Text17.Visible = True
   Text18.Visible = True
   
   Label11.Visible = True
   Label12.Visible = True
   Label13.Visible = True
   Label14.Visible = True
   label21.Visible = True
   label23.Visible = True
End Sub

Private Sub Command2_Click()
   text10.Text = FileName$
  label16.Visible = False
  Label17.Visible = False
  Label18.Visible = False
  go_stop% = go_stop% Xor 1
  
  Select Case go_stop
 Case Is = 1
   'ria_base.Command2.Caption = "TERMINATE"
   
  'which item selected?
    mm% = RIA_BASE.Combo1.ListIndex
  Select Case mm%
    Case Is = 0
       cnt_time% = 30
    Case Is = 1
       cnt_time% = 60
    Case Is = 2
       cnt_time% = 120
    Case Is = 3
       cnt_time% = 300
    Case Is = 4
       cnt_time% = 600
    Case Else
  End Select
    If mm% = -1 Then
       RIA_BASE.Label17.Visible = True
       go_stop% = go_stop% Xor 1 '&&&&&&
       Exit Sub 'count time not selected
    End If
   
    RIA_BASE.Text6.Enabled = True
   

  'which item selected?
 'Is FIRST/SECOND selected? not selected =-1, first = 0,second = 1
    which_res% = RIA_BASE.Combo2.ListIndex
  Select Case which_res%
    Case Is = -1
       RIA_BASE.Label18.Visible = True
       Exit Sub
    Case Is = 0
       RIA_BASE.Text7.Enabled = True
       RIA_BASE.Text8.Enabled = False
       RIA_BASE.Text9.Enabled = False
       RIA_BASE.Text13.Enabled = False
    Case Is = 1
       RIA_BASE.Text7.Enabled = False
       RIA_BASE.Text8.Enabled = True
       RIA_BASE.Text9.Enabled = False
       RIA_BASE.Text13.Enabled = False

    Case Is = 2
       RIA_BASE.Text7.Enabled = False
       RIA_BASE.Text8.Enabled = False
       RIA_BASE.Text9.Enabled = True
       RIA_BASE.Text13.Enabled = False

    Case Is = 3
       RIA_BASE.Text7.Enabled = False
       RIA_BASE.Text8.Enabled = False
       RIA_BASE.Text9.Enabled = False
       RIA_BASE.Text13.Enabled = True

    Case Else
  End Select

  RIA_BASE.Command2.Caption = "TERMINATE"
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
       
       a$ = "E"  ' message to uC : prEset
       Form1.MScomm1.Output = a$
       timehigh% = 0
       timelow% = cnt_time%
       Form1.MScomm1.Output = Chr$(timelow% - 1)
       Form1.MScomm1.Output = Chr$(timehigh%)

'Command to uC to start counting:
       a$ = "S"
       Form1.MScomm1.Output = a$  'START
    
'Status checking  command is "G":
       a$ = "G"
      Form1.MScomm1.Output = a$
       RIA_BASE.Timer3.Enabled = True
       'timer will check status in 0.5 sec intervals.
  Case Is = 0
    RIA_BASE.Command2.Caption = "GO"
    RIA_BASE.Timer3.Enabled = False
    RIA_BASE.Timer4.Enabled = False
    Exit Sub
  Case Else
End Select

End Sub

Private Sub command3_Click()
     'looking for missing measurements. Trying to replace missing ones
     'if possible
     
     
     ddd = Val(Text7.Text)
     eee = Val(Text8.Text)
     fff = Val(Text9.Text)
     ggg = Val(Text13.Text)

     If nonzero%(0) = 0 And nonzero%(1) = 0 Then
        ' command3.Enabled = False       '#####
         MsgBox "Make measurements!"
         Exit Sub

     End If
     If nonzero%(0) = 0 Then eee = 2 * eee
     If nonzero%(1) = 0 Then ddd = 2 * ddd
     If nonzero%(2) = 0 Then ggg = 2 * ggg
     If nonzero%(3) = 0 Then fff = 2 * fff
     aaa = ddd + eee
     bbb = fff + ggg
     
     finres = Int(1000 * bbb / aaa)
     finres = finres / 10
     Text14.Text = "Uptake is: " + Str$(finres) + "%"
End Sub

Private Sub docal_Click()
     Form1.Show
End Sub

Private Sub domeas_Click()
    ' Is everything selected?

   label16.Visible = False
   Text1.Enabled = False
   Text2.Enabled = False
   Text3.Enabled = False
   Text4.Enabled = False
   Text5.Enabled = False


   Label7.Visible = True
   Label8.Visible = True
   Label9.Visible = True
   Label10.Visible = True
   Command2.Visible = True
   Combo1.Visible = True
   Combo2.Visible = True
   
   Line1.Visible = True
   Line2.Visible = True
   Line3.Visible = True
   Line4.Visible = True
   Line5.Visible = True
  
   Text6.Visible = True
   Text7.Visible = True
   Text8.Visible = True
   Text9.Visible = True
   
   Label11.Visible = True
   Label12.Visible = True
   Label13.Visible = True
   Label14.Visible = True
   
  
End Sub

Private Sub dosshell_Click()
  
  progra$ = "command.com"
  x = Shell(progra$)

End Sub

Private Sub edit_Click()
 Text1.Enabled = True
 Text2.Enabled = True
 Text3.Enabled = True
 Text4.Enabled = True
 Text5.Enabled = True
 text10.Enabled = True
 command1.Visible = True
End Sub

Private Sub ex_Click()
   RIA_BASE.Hide
   End
End Sub

Private Sub Form_Load()
  
  
  'default communication port is 1, it can be changed
    On Error GoTo errhandler
      Form1.MScomm1.CommPort = 1
      Form1.MScomm1.Settings = "9600,N,8,1"
      Form1.MScomm1.InputLen = 0  'reading all characters from buffer
      Form1.MScomm1.PortOpen = True
      
       
    Combo1.AddItem "    30 sec "
    Combo1.AddItem "   1  min  "
    Combo1.AddItem "   2  min  "
    Combo1.AddItem "   5  min  "
    Combo1.AddItem "  10  min  "

    Combo2.AddItem " FIRST A"
    Combo2.AddItem " FIRST B"
    Combo2.AddItem "SECOND A"
    Combo2.AddItem "SECOND B"

    a$ = Date$
   
     mmonth$ = Left$(a$, 2)
     dday$ = Mid$(a$, 4, 2)
     yyear$ = Right$(a$, 4)
     ourdate$ = dday$ + "/" + mmonth$ + "/" + yyear$
     Text11.Text = ourdate$
     Exit Sub
spare_com:
      Form1.MScomm1.CommPort = 2
      Form1.MScomm1.Settings = "9600,N,8,1"
      Form1.MScomm1.InputLen = 0  'reading all characters from buffer
      Form1.MScomm1.PortOpen = True
      Exit Sub
errhandler:
MsgBox "COM1 not free. Trying with COM2"
  GoTo spare_com
     
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Form1.MScomm1.PortOpen = False

'Cancel = -1
    'Form1.Visible = False
 
End Sub


Private Sub go_sim_Click(Index As Integer)
     
        a$ = "T"
        Form1.MScomm1.Output = a$
       
       Text1.Enabled = False
       Text2.Enabled = False
       Text3.Enabled = False
       Text4.Enabled = False
       Text5.Enabled = False
       timer1.Enabled = True ' waiting for reply
 End Sub

Private Sub help_Click()
If missing% = 1 Then GoTo down
   frm_help.Show
down:
   Exit Sub
   
End Sub

Private Sub load_Click()
     LO% = 1
'tryagain:
     
     load_sim.Show
     loadview% = 1
End Sub


Private Sub New_Click()
   name_sel.Show
  FileName$ = ""
  name_sel.Text1.Text = ""
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text4.Enabled = True
  Text5.Enabled = True
  Text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
  Text4.Text = ""
  Text5.Text = ""
  Text6.Text = ""
  Text7.Text = ""
  Text8.Text = ""
  Text9.Text = ""
  text10.Text = ""
  Text11.Text = ourdate$
  Text15.Text = ""
  Text16.Text = ""
  Text17.Text = ""
  Text18.Text = ""
  Text13.Text = ""
  Text14.Text = ""
 'clear all informations:
    
    For nn% = 0 To 19
        info$(nn%) = ""
    Next nn%



End Sub

Private Sub report_Click()
  ' Printer.FontName = sanserif
   Printer.ScaleMode = 4
   Printer.CurrentX = 34
   Printer.CurrentY = 8
   Printer.FontSize = 18
   Printer.Print "UPTAKE RECORD"
   Printer.Print ""
   Printer.FontSize = 13
   Printer.CurrentX = 31
   Printer.Print "YANGON GENERAL HOSPITAL"
   Printer.CurrentX = 31
   Printer.Print "Department of Nuclear Medicine"
   Printer.Print ""
   Printer.Print ""
   Printer.CurrentX = 20
   Printer.Print "File Name:......................."; FileName$
   Printer.CurrentX = 20
   Printer.Print ""
   Printer.CurrentX = 20
   Printer.Print "Name:.............................."; RIA_BASE.Text1.Text
   Printer.CurrentX = 20
   Printer.Print "Father's Name:................"; RIA_BASE.Text2.Text
   Printer.CurrentX = 20
   
   DOB$ = RIA_BASE.Text5.Text + "/" + RIA_BASE.Text4.Text + "/" + RIA_BASE.Text3.Text
   Printer.Print "Date of Birth :.................."; DOB$
   Printer.CurrentX = 20
   Printer.Print ""
   Printer.CurrentX = 20
   Printer.Print "Applied Dose,first:.......... "; RIA_BASE.Text7.Text
   Printer.CurrentX = 20
   Printer.Print "Applied Dose,repeated:.. "; RIA_BASE.Text8.Text
   Printer.CurrentX = 20
   Printer.Print ""
   Printer.CurrentX = 20
   Printer.Print "Uptake,first:......................"; RIA_BASE.Text9.Text
   Printer.CurrentX = 20
   Printer.Print "Uptake,repeated:............."; RIA_BASE.Text13.Text
   Printer.Print ""
   Printer.FontSize = 15
   Printer.CurrentX = 20
   Printer.Print RIA_BASE.Text14.Text
   Printer.CurrentX = 20
   Printer.Print ""
   Printer.CurrentX = 20
   Printer.FontSize = 13
   Printer.Print "Yangon,  "; RIA_BASE.Text11.Text
   Printer.EndDoc

End Sub

Private Sub save_click()
   If FileName$ = "" Then
      MsgBox "File name not selected" ', RIA
      Exit Sub
   End If
     FileName$ = text10.Text
  Open FileName$ For Output As #7
   Print #7, Text1.Text
   Print #7, Text2.Text
   Print #7, Text3.Text
   Print #7, Text4.Text
   Print #7, Text5.Text
   Print #7, cnt_time%
   Print #7, Text7.Text
   Print #7, Text15.Text
   Print #7, Text8.Text
   Print #7, Text16.Text
   Print #7, Text9.Text
   Print #7, Text17.Text
   Print #7, Text13.Text
   Print #7, Text18.Text
   Print #7, Text14.Text
   Print #7, Text11.Text
    '  Print #7, Text18.text
 Close #7

End Sub

Private Sub saveme_Click()
     Call save_click
End Sub

Private Sub set_Click()
  Text1.Enabled = False
   Text2.Enabled = False
  Text3.Enabled = False
   Text4.Enabled = False
  Text5.Enabled = False
  RIA_SET.Show
   
End Sub

Private Sub spectrum_Click()
    RIA_SPEC.Show
End Sub

Private Sub Timer1_Timer()

'timer interval for TEST
       answer$ = Form1.MScomm1.Input 'buffer emptied
       timer1.Enabled = False
       rr$ = ""
     For tt% = 1 To 3  'counts
       testres$ = Mid$(answer$, tt%, 1)
       If testres$ = "" Then
        label6.Caption = "FAILURE!"
        label6.Visible = True
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
        
        If rr$ = " 1 2 3 4 5" Then label6.Caption = "SYSTEM OK!"
        If rr$ <> " 1 2 3 4 5" Then label6.Caption = "FAILURE!"
        label6.Visible = True
       
down:
        timer1.Enabled = False
        timer2.Enabled = True
   
   
  
End Sub

Private Sub Timer2_Timer()
 'Time to show label "OK" or "FAILURE"
  label6.Visible = False
  timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
    'timer waits 500ms; and here is an answer:
    instring$ = Form1.MScomm1.Input
    'Start,stoP: "S or P is an answer
    'Is the preset time over?
    If Left$(instring$, 1) = "P" Then
      Timer3.Enabled = False
Beep: Beep: Beep
     ' Call Command2_Click
      Exit Sub
    End If
'Reads the collected counts number; out: "R"
' return: 3 digits for counts,
'CR LF,
'2 digits for result: seconds, and seconds*256
 'Inintiate reading by "R":
    a$ = "R"
    Form1.MScomm1.Output = a$
    RIA_BASE.Timer4.Enabled = True
 'here we jump to Timer 4

End Sub

Private Sub Timer4_Timer()
    'Here we read results:
    Static counts(10)
    instring$ = Form1.MScomm1.Input
    
 'at first we look for counts:
     For tt% = 1 To 3
       testres$ = Mid$(instring$, tt%, 1)
        If testres$ = "" Then
           MsgBox "Failure. Check operation"
           Command2.Caption = "GO" '#####
           Call Command2_Click
           Exit Sub
        End If

       counts(tt%) = Asc(testres$)
     Next tt%
       totcounts = counts(1) + 256 * counts(2) + (256 ^ 2) * counts(3)
    
 'then for the time:
     For tt% = 6 To 7
       testres$ = Mid$(instring$, tt%, 1)
       counts(tt%) = Asc(testres$)
     Next tt%
       tottime = counts(6) + 256 * counts(7)
 'Last  messages:TIME left
       ourtime = Str$(65536 - tottime)
     If ourtime = 65536 Then
          ourtime = 0
          Timer3.Enabled = False
     End If
       Text6.Text = ourtime
     If ourtime = 0 Then
         label16.Visible = True
         go_stop% = go_stop Xor 1
         RIA_BASE.Command2.Caption = "GO"
Beep: Beep
     End If

       Timer4.Enabled = False
 
 'Last  messages:COUNTS information
     Select Case which_res%
       Case Is = 0
          If totcounts > 0 Then nonzero%(0) = 1
          Text7.Text = ""
          Text7.Text = Str$(totcounts)
          Text15.Text = ""
          Text15.Text = Time$
       Case Is = 1
          If totcounts > 0 Then nonzero%(1) = 1
          Text8.Text = ""
          Text8.Text = Str$(totcounts)
          Text16.Text = ""
          Text16.Text = Time$

       Case Is = 2
          If totcounts > 0 Then nonzero%(2) = 1
          Text9.Text = ""
          Text9.Text = Str$(totcounts)
          Text17.Text = ""
          Text17.Text = Time$

       Case Is = 3
          If totcounts > 0 Then nonzero%(3) = 1
          Text13.Text = ""
          Text18.Text = ""
          Text18.Text = Time$
          Text13.Text = Str$(totcounts)

       Case Else
     End Select

End Sub

Private Sub Timer5_Timer()
   text12.Text = Time$

End Sub

Private Sub view_Click()
   RIA_BASE.Hide
   Label7.Visible = True
   Label8.Visible = True
   Label9.Visible = True
   Label10.Visible = True
   label15.Visible = True
   Command2.Visible = True
   command3.Visible = True
   Combo1.Visible = True
   Combo2.Visible = True
   text10.Enabled = False

   Line1.Visible = True
   Line2.Visible = True
   Line3.Visible = True
   Line4.Visible = True
   Line5.Visible = True
   line6.Visible = True
   
   Text6.Visible = True
   Text7.Visible = True
   Text8.Visible = True
   Text9.Visible = True
   text10.Visible = True
   Text13.Visible = True
   Text14.Visible = True
   Text15.Visible = True
   Text16.Visible = True
   Text17.Visible = True
   Text18.Visible = True
   
   Label11.Visible = True
   Label12.Visible = True
   Label13.Visible = True
   Label14.Visible = True
   label21.Visible = True
   label23.Visible = True
    
    load_sim.Show
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    Text6.Enabled = False
    Text7.Enabled = False
    Text8.Enabled = False
    Text9.Enabled = False
   

End Sub

