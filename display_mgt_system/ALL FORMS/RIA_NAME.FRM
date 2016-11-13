VERSION 4.00
Begin VB.Form name_sel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Select File Name"
   ClientHeight    =   1935
   ClientLeft      =   1545
   ClientTop       =   1845
   ClientWidth     =   4785
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
   Height          =   2340
   Left            =   1485
   LinkTopic       =   "Form2"
   ScaleHeight     =   1935
   ScaleWidth      =   4785
   Top             =   1500
   Width           =   4905
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK?"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Type file name:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "name_sel"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
   a$ = Text1.Text
   FileName$ = a$ + ".ria"
   RIA_BASE.text10.Text = FileName$
   name_sel.Hide
   RIA_BASE.command1.Visible = True
  
   RIA_BASE.Text6.Text = ""
   RIA_BASE.Text7.Text = ""
   RIA_BASE.Text8.Text = ""
   RIA_BASE.Text9.Text = ""
   RIA_BASE.text10.Visible = True
   RIA_BASE.label15.Visible = True
End Sub

