VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Frmsyllabuspreperation 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   435
      Left            =   8910
      TabIndex        =   7
      ToolTipText     =   "Click to Exit"
      Top             =   6540
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   435
      Left            =   7950
      TabIndex        =   6
      ToolTipText     =   "Click to Save"
      Top             =   6540
      Width           =   945
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   435
      Left            =   7020
      TabIndex        =   3
      ToolTipText     =   "Click to insert new information"
      Top             =   6540
      Width           =   945
   End
   Begin VB.Frame Frame4 
      Caption         =   " Syllabus Detail #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4755
      Left            =   3420
      TabIndex        =   14
      Top             =   1770
      Width           =   6435
      Begin VB.CommandButton CmdUnderline 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&U"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2610
         TabIndex        =   18
         Top             =   570
         Width           =   555
      End
      Begin VB.CommandButton CmdBold 
         BackColor       =   &H00FFFFFF&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   17
         Top             =   570
         Width           =   555
      End
      Begin VB.CommandButton cmdItalic 
         BackColor       =   &H00FFFFFF&
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         TabIndex        =   16
         Top             =   570
         Width           =   555
      End
      Begin VB.TextBox txtfields 
         Height          =   315
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Insert Name"
         Top             =   210
         Width           =   4845
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3795
         Left            =   60
         TabIndex        =   5
         ToolTipText     =   "Insert Detail Syllabus of the Subject"
         Top             =   900
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   6694
         _Version        =   393217
         MaxLength       =   4000
         TextRTF         =   $"Frmsyllabuspreperation.frx":0000
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared   By"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Subject List #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4755
      Left            =   0
      TabIndex        =   13
      Top             =   1770
      Width           =   3405
      Begin VB.ListBox List1 
         Height          =   4350
         Left            =   60
         TabIndex        =   2
         ToolTipText     =   "Select Subject "
         Top             =   330
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   0
      TabIndex        =   8
      Top             =   990
      Width           =   9855
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Frmsyllabuspreperation.frx":0082
         Left            =   1170
         List            =   "Frmsyllabuspreperation.frx":0084
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select Class"
         Top             =   270
         Width           =   2985
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select Education Year"
         Top             =   270
         Width           =   2625
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class "
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Education  Year"
         Height          =   195
         Left            =   4800
         TabIndex        =   11
         Top             =   330
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   9825
      TabIndex        =   9
      Top             =   0
      Width           =   9885
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   10
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Syllabus Preperation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   285
         Left            =   3360
         TabIndex        =   19
         Top             =   270
         Width           =   2385
      End
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   -30
         Picture         =   "Frmsyllabuspreperation.frx":0086
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   9885
      End
   End
End
Attribute VB_Name = "Frmsyllabuspreperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBold_Click()
 RichTextBox1.SelBold = IIf(RichTextBox1.SelBold, False, True)
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdItalic_Click()
 RichTextBox1.SelItalic = IIf(RichTextBox1.SelItalic, False, True)
End Sub

Private Sub cmdnew_Click()
txtfields = ""
RichTextBox1 = ""
List1.SetFocus
End Sub

Private Sub cmdSAVE_Click()
If Len(txtfields) = 0 Then
        MsgBox "Please Enter Prepared By.", vbInformation, App.Title
        txtfields.SetFocus
        Exit Sub
End If
If Len(RichTextBox1.Text) = 0 Then
        MsgBox "Please Enter Syllabus.", vbInformation, App.Title
        RichTextBox1.SetFocus
        Exit Sub
End If
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "Syllabuspreperation1"
cmd(1) = Mid((Combo1.Text), 1, 5)
cmd(2) = Trim(Combo2.Text)
cmd(3) = Mid((List1.Text), 1, 5)
cmd(4) = Trim(RichTextBox1.TextRTF)
cmd(5) = Trim(txtfields)
cmd(6) = "DSL"
cmd(7) = Date
cmd.Execute
MsgBox "Save Successfully.", vbInformation, "Student Management System"
cmdnew.SetFocus


Exit Sub

End Sub



Private Sub CmdUnderline_Click()
RichTextBox1.SelUnderline = IIf(RichTextBox1.SelUnderline, False, True)
End Sub

Private Sub Combo1_Click()
List1.Clear


Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("select Sub_code,Sub_title from Subject_Info_sub where class_code ='" & Mid((Combo1.Text), 1, 5) & "'")
If Not rs1.EOF Then
    Do Until rs1.EOF
        List1.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
    If List1.ListCount > 0 Then List1.ListIndex = 0

End If
getsyllabus
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo2.SetFocus
End If

getsyllabus
End Sub

Private Sub Combo2_Click()

getsyllabus
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List1.SetFocus
End If

getsyllabus
End Sub






Private Sub Form_Load()
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
       Combo1.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
'    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0

End If
Dim IL As Integer
For IL = 2000 To 2020
    Combo2.AddItem (IL)
Next IL
'If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
getsyllabus
End Sub
Public Function getdata(SQLString As String) As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = SQLString
 Set rs = cmd.Execute
Set getdata = rs
End Function

Private Sub List1_Click()
getsyllabus
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtfields.SetFocus
End If
End Sub

Private Sub txtfields_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    RichTextBox1.SetFocus
End If
'getsyllabus
End Sub
Public Function getsyllabus()
Dim rs As New ADODB.Recordset
Set rs = getdata("select PreparedBy,Syllabusdetail from Syllabuspreperation where classid='" & Mid((Combo1.Text), 1, 5) & "' and  Eyear='" & Combo2.Text & "'and subjectId='" & Mid((List1.Text), 1, 5) & "'")
If Not rs.EOF Then
    txtfields = rs!PreparedBy
    RichTextBox1 = rs!Syllabusdetail
Else
    RichTextBox1 = ""
    txtfields = ""
End If

End Function
