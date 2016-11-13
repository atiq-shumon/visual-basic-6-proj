VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   LinkTopic       =   "Form2"
   ScaleHeight     =   2715
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3120
      Top             =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PC Interface Display System"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1020
      TabIndex        =   1
      Top             =   330
      Width           =   5475
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Developed By:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   420
      Index           =   2
      Left            =   2407
      TabIndex        =   0
      Top             =   1057
      Width           =   2550
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   247
      OLEDropMode     =   1  'Manual
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   1657
      Width           =   6870
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x
Private Sub Form_Load()
x = 0
End Sub

Private Sub Timer1_Timer()
x = x + 100
If x > 1000 Then
Timer1.Enabled = False
Unload Me
End If
End Sub
