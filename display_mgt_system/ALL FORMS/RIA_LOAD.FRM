VERSION 4.00
Begin VB.Form load_sim 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Load File..."
   ClientHeight    =   2325
   ClientLeft      =   2445
   ClientTop       =   2610
   ClientWidth     =   4875
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
   Height          =   2730
   Left            =   2385
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4875
   Top             =   2265
   Visible         =   0   'False
   Width           =   4995
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1575
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   3120
      Pattern         =   "*.RIA"
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton OKButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK?"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "load_sim"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
    load_sim.Hide
End Sub

Private Sub Command2_Click()
    load_sim.Hide
End Sub

Private Sub Dir1_Change()
    file1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
 On Error GoTo down
    Dir1.Path = Drive1.Drive
 Exit Sub
down:
    MsgBox "Drive not responding"
    Exit Sub
End Sub

Private Sub File1_DblClick()
    OKButton_Click
End Sub

Private Sub OKButton_Click()
   load_sim.Hide
 'On Error GoTo FileError

    If (Right$(Dir1.Path, 1) = "\") Then
        FileName$ = Dir1.Path + file1.FileName
    Else
        FileName$ = Dir1.Path + "\" + file1.FileName
    End If
    If file1.FileName = "" Then
       MsgBox "File not selected"
         RIA_BASE.Show
       GoTo cellar
    End If
    RIA_BASE.command1.Visible = False
    RIA_BASE.Show
    RIA_BASE.text10.Text = FileName$
    Open FileName$ For Input As #1
    Do
       nn% = nn% + 1
       Input #1, info$(nn%)
    Loop While Not EOF(1)
     Close #1
        
        RIA_BASE.Text1.Text = info$(1)
        RIA_BASE.Text2.Text = info$(2)
        RIA_BASE.Text3.Text = info$(3)
        RIA_BASE.Text4.Text = info$(4)
        RIA_BASE.Text5.Text = info$(5)
        RIA_BASE.Text6.Text = info$(6) 'cnt_time%
        RIA_BASE.Text7.Text = info$(7)  'first A
        RIA_BASE.Text15.Text = info$(8) 'first time
        RIA_BASE.Text8.Text = info$(9)  '
        RIA_BASE.Text16.Text = info$(10) 'date
        RIA_BASE.Text9.Text = info$(11)
        RIA_BASE.Text17.Text = info$(12)  '
        RIA_BASE.Text13.Text = info$(13) '
        RIA_BASE.Text18.Text = info$(14)
        RIA_BASE.Text14.Text = info$(15)  'result
        RIA_BASE.Text11.Text = info$(16) 'date

cellar:
    
    
   End Sub

