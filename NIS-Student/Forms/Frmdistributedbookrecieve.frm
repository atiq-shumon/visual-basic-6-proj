VERSION 5.00
Begin VB.Form Frmdistributedbookrecieve 
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Exit"
      Height          =   435
      Left            =   5910
      TabIndex        =   13
      Top             =   5400
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   435
      Left            =   4950
      TabIndex        =   12
      Top             =   5400
      Width           =   945
   End
   Begin VB.Frame Frame2 
      Height          =   3585
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   6855
      Begin VB.ListBox List2 
         Height          =   2985
         Left            =   3570
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   510
         Width           =   3165
      End
      Begin VB.ListBox List1 
         Height          =   2985
         Left            =   90
         TabIndex        =   8
         Top             =   510
         Width           =   3165
      End
      Begin VB.Line Line1 
         X1              =   3390
         X2              =   3390
         Y1              =   120
         Y2              =   3570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student List #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   210
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recieved Book List #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   3600
         TabIndex        =   9
         Top             =   210
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   0
      TabIndex        =   1
      Top             =   810
      Width           =   6855
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Text            =   "Combo2"
         Top             =   540
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Frmdistributedbookrecieve.frx":0000
         Left            =   1080
         List            =   "Frmdistributedbookrecieve.frx":000A
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   180
         Width           =   1995
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Name "
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   795
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Distributed Book Recieval Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   1470
         TabIndex        =   2
         Top             =   180
         Width           =   4335
      End
   End
End
Attribute VB_Name = "Frmdistributedbookrecieve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
