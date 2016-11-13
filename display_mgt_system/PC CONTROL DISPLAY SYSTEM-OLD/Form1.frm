VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Display System"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   10515
      TabIndex        =   18
      Top             =   0
      Width           =   10515
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Based Information Display System"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   26.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   600
         TabIndex        =   20
         Top             =   720
         Width           =   11145
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Based Information Display System"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   26.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   585
         TabIndex        =   19
         Top             =   750
         Width           =   11145
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   10515
      TabIndex        =   16
      Top             =   6045
      Width           =   10515
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "Form1.frx":0442
         Stretch         =   -1  'True
         Top             =   600
         Width           =   6870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Developed By:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   420
         Index           =   2
         Left            =   2145
         TabIndex        =   17
         Top             =   0
         Width           =   2550
      End
   End
   Begin MSCommLib.MSComm Mc1 
      Left            =   2400
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808000&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   10455
      TabIndex        =   15
      Top             =   7860
      Width           =   10515
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   6
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   120
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   681
      TabIndex        =   9
      Top             =   1920
      Width           =   10215
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   1440
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   0
         Width           =   8655
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   1440
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   720
         Width           =   8655
      End
      Begin VB.TextBox Text3 
         Height          =   615
         Left            =   1440
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1440
         Width           =   8655
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   1440
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2280
         Width           =   8655
      End
      Begin VB.TextBox Text5 
         Height          =   615
         Left            =   1440
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2880
         Width           =   8655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "MODI - 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "MODI - 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "MODI - 3:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   12
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "MODI - 4:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   11
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "MODI - 5:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   10
         Top             =   3000
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ff = FreeFile
Open App.Path + "\MODIFILE.DAT" For Random As #1 Len = Len(Modfile)
Modfile.modi1 = Text1.Text
Modfile.modi2 = Text2.Text
Modfile.modi3 = Text3.Text
Modfile.modi4 = Text4.Text
Modfile.modi5 = Text5.Text
Put #1, 1, Modfile
Close ff
End Sub

Private Sub Command2_Click()
Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Mc1.PortOpen = True

'Upper Charector
'Text1.Text = UCase(Text1.Text)
'Text2.Text = UCase(Text2.Text)
'Text3.Text = UCase(Text3.Text)
'Text4.Text = UCase(Text4.Text)
'Text5.Text = UCase(Text5.Text)

Text1.Text = Text1.Text
Text2.Text = Text2.Text
Text3.Text = Text3.Text
Text4.Text = Text4.Text
Text5.Text = Text5.Text

Mc1.Settings = "1200,N,8,1"
Mc1.PortOpen = False
Close #1
Open App.Path + "\Output.txt" For Output As #1
Print #1, Chr$(35) + "                        " + Chr$(49) + "          " + Text1.Text
Print #1, Chr$(35) + "                        " + Chr$(50) + "          " + Text2.Text
Print #1, Chr$(35) + "                        " + Chr$(51) + "          " + Text3.Text
Print #1, Chr$(35) + "                        " + Chr$(52) + "          " + Text4.Text
Print #1, Chr$(35) + "                        " + Chr$(53) + "          " + Text5.Text
Close #1
Open "COM1" For Output As #1
Print #1, Chr$(35) + "                        " + Chr$(49) + "          " + Text1.Text
Print #1, Chr$(35) + "                        " + Chr$(50) + "          " + Text2.Text
Print #1, Chr$(35) + "                        " + Chr$(51) + "          " + Text3.Text
Print #1, Chr$(35) + "                        " + Chr$(52) + "          " + Text4.Text
Print #1, Chr$(35) + "                        " + Chr$(53) + "          " + Text5.Text
Close #1
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
ff = FreeFile
Open App.Path + "\MODIFILE.DAT" For Random As #1 Len = Len(Modfile)
Get #1, 1, Modfile
Text1.Text = Trim(Modfile.modi1)
Text2.Text = Trim(Modfile.modi2)
Text3.Text = Trim(Modfile.modi3)
Text4.Text = Trim(Modfile.modi4)
Text5.Text = Trim(Modfile.modi5)
Close ff

'Upper Charector
'Text1.Text = UCase(Text1.Text)
'Text2.Text = UCase(Text2.Text)
'Text3.Text = UCase(Text3.Text)
'Text4.Text = UCase(Text4.Text)
'Text5.Text = UCase(Text5.Text)

Text1.Text = Text1.Text
Text2.Text = Text2.Text
Text3.Text = Text3.Text
Text4.Text = Text4.Text
Text5.Text = Text5.Text

Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
Me.Icon = LoadPicture(App.Path + "\logo.ico")
Me.Left = 0
Me.Top = 0
Me.Width = Screen.Width
Me.Height = Screen.Height
Picture1.Left = (Me.Width - Picture1.Width) / 2
Picture1.Top = (Me.Height - Picture1.Height) / 2
Label7(2).Left = (Me.Width - Label7(2).Width) / 2
Image1.Left = (Me.Width - Image1.Width) / 2
Label6.Left = (Me.Width - Label6.Width) / 2
Label8.Left = (Me.Width - Label8.Width) / 2
End If
Text1.Height = (Picture1.Height / 6) / 15
Text2.Height = (Picture1.Height / 6) / 15
Text3.Height = (Picture1.Height / 6) / 15
Text4.Height = (Picture1.Height / 6) / 15
Text5.Height = (Picture1.Height / 6) / 15

Text2.Top = Text1.Height + 5
Text3.Top = (Text1.Height + Text2.Height) + 10
Text4.Top = (Text1.Height + Text2.Height + Text3.Height) + 15
Text5.Top = (Text1.Height + Text2.Height + Text3.Height + Text4.Height) + 20

Label2.Top = Text1.Height + 5
Label3.Top = (Text1.Height + Text2.Height) + 10
Label4.Top = (Text1.Height + Text2.Height + Text3.Height) + 15
Label5.Top = (Text1.Height + Text2.Height + Text3.Height + Text4.Height) + 20


End Sub


