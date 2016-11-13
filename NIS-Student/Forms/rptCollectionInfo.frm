VERSION 5.00
Begin VB.Form rptCollection 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report :Collection Information"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Monthly Collection"
      Height          =   1065
      Left            =   0
      TabIndex        =   18
      Top             =   630
      Width           =   5415
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   315
         ItemData        =   "rptCollectionInfo.frx":0000
         Left            =   2640
         List            =   "rptCollectionInfo.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   300
         Width           =   915
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   300
         Width           =   1065
      End
      Begin VB.OptionButton Option3 
         Caption         =   "ALL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   19
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month Year :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   1530
         TabIndex        =   22
         Top             =   330
         Width           =   1110
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   585
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   5325
      TabIndex        =   9
      Top             =   0
      Width           =   5385
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   10
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report : Collection Statement"
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
         Left            =   975
         TabIndex        =   11
         Top             =   120
         Width           =   3360
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   -720
         Picture         =   "rptCollectionInfo.frx":0068
         Stretch         =   -1  'True
         Top             =   -90
         Width           =   8835
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3045
      Left            =   0
      TabIndex        =   6
      Top             =   1710
      Width           =   5415
      Begin VB.OptionButton Option1 
         Caption         =   "By Student "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   0
         Top             =   150
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "By Class"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   150
         Width           =   1275
      End
      Begin VB.Frame Frame3 
         Caption         =   "Select Class"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   120
         TabIndex        =   8
         Top             =   1890
         Visible         =   0   'False
         Width           =   5205
         Begin VB.ComboBox CboClassID 
            Height          =   315
            Left            =   2220
            TabIndex        =   3
            Top             =   330
            Width           =   2475
         End
         Begin VB.TextBox txtClassName 
            Enabled         =   0   'False
            Height          =   345
            Left            =   2220
            TabIndex        =   12
            Top             =   885
            Width           =   2475
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   900
            TabIndex        =   14
            Top             =   960
            Width           =   1035
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Class"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   900
            TabIndex        =   13
            Top             =   390
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Select Student"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   7
         Top             =   570
         Visible         =   0   'False
         Width           =   5205
         Begin VB.ComboBox cboStdID 
            Height          =   315
            Left            =   2220
            TabIndex        =   2
            Top             =   330
            Width           =   2475
         End
         Begin VB.TextBox txtStdName 
            Enabled         =   0   'False
            Height          =   345
            Left            =   2220
            TabIndex        =   15
            Top             =   885
            Width           =   2475
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Student Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   855
            TabIndex        =   17
            Top             =   960
            Width           =   1245
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Student ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   855
            TabIndex        =   16
            Top             =   390
            Width           =   945
         End
      End
   End
   Begin VB.CommandButton cmd 
      Height          =   405
      Index           =   0
      Left            =   4140
      Picture         =   "rptCollectionInfo.frx":CF0D
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Report Veiwer"
      Top             =   4830
      Width           =   555
   End
   Begin VB.CommandButton cmd 
      Height          =   405
      Index           =   1
      Left            =   4755
      Picture         =   "rptCollectionInfo.frx":D7D7
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   4830
      Width           =   585
   End
   Begin VB.Shape Shape1 
      Height          =   525
      Left            =   4080
      Top             =   4770
      Width           =   1305
   End
End
Attribute VB_Name = "rptCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CboClassID_Click()
Dim rs As New ADODB.Recordset

Set rs = getdata("Select ClassName from ClassInfo where ClassID = '" & CboClassID.Text & "' ")

    If Not rs.EOF Then
       txtClassName.Text = rs!ClassName
    Else
       txtClassName = ""
    End If

End Sub
Private Sub generate_yr()
   Dim i As Integer
   Dim j As Integer
    j = 2000
   For i = 0 To 50
      Combo3.AddItem j
      
     j = j + 1
      
  Next i
 Combo3.Text = Format(Date, "YYYY")
 Combo2.Text = Format(Date, "MMM")
End Sub
Private Sub CboClassID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmd(0).SetFocus
End If
End Sub

Private Sub cboStdID_Click()
Dim rs As New ADODB.Recordset

Set rs = getdata("Select StudentName from StudentInfo where StudentID = '" & cboStdID.Text & "' ")

    If Not rs.EOF Then
       txtStdName.Text = rs!StudentName
    Else
       txtStdName = ""
    End If

End Sub

Private Sub cboStdID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmd(0).SetFocus
End If
End Sub

Private Sub cmd_Click(Index As Integer)
Select Case Index
    Case 0
       Param_mode = 1
       rptMode = 19
       Screen.MousePointer = vbHourglass
       frmViewer.Show 1
    Case 1
        Unload Me
End Select
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
generate_yr

Set rs = getdata("Select StudentID from StudentInfo")
If rs.EOF = False Then
    Do Until rs.EOF
        cboStdID.AddItem rs(0)
        rs.MoveNext
    Loop
End If

Set rs1 = getdata("Select ClassID from ClassInfo")
If rs1.EOF = False Then
    Do Until rs1.EOF
        CboClassID.AddItem rs1(0)
        rs1.MoveNext
    Loop
End If

Option1.Value = True
Call Option1_Click


End Sub

Private Sub Option1_Click()
Frame4.Visible = True
Frame4.Top = 1000
Frame4.Left = 120

Frame3.Visible = False

End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cboStdID.SetFocus
End If
End Sub

Private Sub Option2_Click()
Frame3.Visible = True
Frame3.Top = 1000
Frame3.Left = 120

Frame4.Visible = False
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CboClassID.SetFocus
End If

End Sub


