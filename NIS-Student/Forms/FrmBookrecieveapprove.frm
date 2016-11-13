VERSION 5.00
Begin VB.Form FrmBookrecieveapprove 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApproved 
      BackColor       =   &H8000000C&
      Caption         =   "Approved"
      Height          =   435
      Left            =   4650
      TabIndex        =   14
      Top             =   5640
      Width           =   945
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Exit"
      Height          =   435
      Left            =   5580
      TabIndex        =   13
      Top             =   5640
      Width           =   945
   End
   Begin VB.Frame Frame3 
      Height          =   945
      Left            =   0
      TabIndex        =   7
      Top             =   750
      Width           =   6525
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1230
         TabIndex        =   12
         Text            =   "Combo2"
         Top             =   510
         Width           =   2745
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1230
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   150
         Width           =   1725
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Name"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   540
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3945
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   6525
      Begin VB.ListBox List2 
         Height          =   3210
         Left            =   3390
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   540
         Width           =   2985
      End
      Begin VB.ListBox List1 
         Height          =   3180
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   2985
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recieved Book Approval #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   3390
         TabIndex        =   6
         Top             =   180
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student List #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   1125
      End
      Begin VB.Line Line1 
         X1              =   3240
         X2              =   3240
         Y1              =   90
         Y2              =   3900
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   795
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6465
      TabIndex        =   0
      Top             =   0
      Width           =   6525
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   1
         Top             =   930
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recieved Book (By Student) Approval Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   450
         TabIndex        =   10
         Top             =   180
         Width           =   5280
      End
   End
End
Attribute VB_Name = "FrmBookrecieveapprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
