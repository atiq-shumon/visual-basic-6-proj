VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2588
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1748
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      MaxLength       =   16
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Registration No.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1185
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sl, cnt

Private Sub Command1_Click()

If cnt = 2 Then
End
End If

cnt = cnt + 1

If Text1.Text <> "EH01713008555" Then
Text1.SetFocus
Exit Sub
Else
Close #1
Open "C:\windows\system\options.dll" For Random As #1 Len = Len(Rgg)
Rgg.Rg = 9.99999999888889E+31
Put #1, 1, Rgg
Close #1
Unload Me
Call main
End If

End Sub

Private Sub Command2_Click()
End
End Sub

