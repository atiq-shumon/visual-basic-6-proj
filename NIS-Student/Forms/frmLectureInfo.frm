VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmlectureinfo 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Lecture Details"
      ForeColor       =   &H00C00000&
      Height          =   2865
      Left            =   0
      TabIndex        =   18
      Top             =   2010
      Width           =   9225
      Begin RichTextLib.RichTextBox RickLecdetail 
         Height          =   2535
         Left            =   60
         TabIndex        =   19
         ToolTipText     =   "Insert Lecture Description"
         Top             =   240
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   4471
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmLectureInfo.frx":0000
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1485
      Left            =   -30
      TabIndex        =   16
      Top             =   4890
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   2619
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   435
      Left            =   5280
      TabIndex        =   1
      ToolTipText     =   "Click to insert new information"
      Top             =   6420
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   435
      Left            =   6270
      TabIndex        =   5
      ToolTipText     =   "Click to save"
      Top             =   6420
      Width           =   945
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   435
      Left            =   7260
      TabIndex        =   6
      ToolTipText     =   "Click to Delete"
      Top             =   6420
      Width           =   945
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   435
      Left            =   8250
      TabIndex        =   7
      ToolTipText     =   "Click to Exit"
      Top             =   6420
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   765
      Left            =   30
      ScaleHeight     =   705
      ScaleWidth      =   9135
      TabIndex        =   9
      Top             =   0
      Width           =   9195
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lecture Information"
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
         Height          =   345
         Left            =   2850
         TabIndex        =   20
         Top             =   150
         Width           =   2265
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   -60
         Picture         =   "frmLectureInfo.frx":0082
         Stretch         =   -1  'True
         Top             =   -60
         Width           =   9195
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1245
      Left            =   0
      TabIndex        =   8
      Top             =   750
      Width           =   9225
      Begin VB.CheckBox Check1 
         Caption         =   "Is This Form Open for Student ?"
         Height          =   375
         Left            =   6480
         TabIndex        =   17
         ToolTipText     =   "Click if the form is openned for student"
         Top             =   840
         Width           =   2565
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   2
         Left            =   1590
         TabIndex        =   4
         ToolTipText     =   "Insert Name"
         Top             =   870
         Width           =   4635
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   1590
         TabIndex        =   3
         ToolTipText     =   "Insert Lecture Description"
         Top             =   540
         Width           =   7485
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   4500
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select Subject"
         Top             =   180
         Width           =   2115
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select Class"
         Top             =   180
         Width           =   2115
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   7560
         TabIndex        =   10
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lesson Prepared By"
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   900
         Width           =   1425
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lecture Description"
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   570
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lecture ID"
         Height          =   195
         Left            =   6690
         TabIndex        =   13
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject "
         Height          =   195
         Left            =   3780
         TabIndex        =   12
         Top             =   210
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmlectureinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim LectureDetail As String
Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  RickLecdetail.SetFocus
End If
End Sub

Private Sub cmdDelete_Click()
Dim rs As New adodb.Recordset
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
con.Open GConnString

 Set cmd.ActiveConnection = con
    If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical) = vbYes Then

        cmd.CommandType = adCmdText
        cmd.CommandText = "Delete from LectureInfo  where ClassID ='" & Mid(Trim(Combo1.Text), 1, 5) & "' and subjectid='" & Mid(Trim(Combo2.Text), 1, 5) & "' and LectureID= '" & Trim(txtfields(0)) & "' "
        cmd.Execute
        MsgBox "Delete successfully .", vbInformation, App.Title
        txtfields(0) = ""
        txtfields(1) = ""
        txtfields(2) = ""
        RickLecdetail = ""
        
        If Check1.Value = 1 Then
            Check1.Value = False
        End If
        
        Call ShowFlexData
        
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdnew_Click()
If Len(Combo1) = 0 Then Exit Sub
If Len(Combo2) = 0 Then Exit Sub
Dim rs As New adodb.Recordset
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
con.Open GConnString
cmd.ActiveConnection = con
Set rs = getdata("select max (LectureID+1)from LectureInfo where classId='" & Mid(Trim(Combo1.Text), 1, 5) & "' and SubjectID='" & Mid(Trim(Combo2.Text), 1, 5) & "'")
If Not rs.EOF Then
    txtfields(0) = IIf(IsNull(rs(0)) = True, "00001", Format(rs(0), "00000"))
Else
    txtfields(0) = "00001"
End If

For i = 1 To 2
    txtfields(i) = ""
Next
RickLecdetail = ""

txtfields(1).SetFocus
End Sub
Private Sub cmdSAVE_Click()
If Len(txtfields(1)) = 0 Then
    MsgBox "Please Enter Lecture Description.", vbInformation, App.Title
    txtfields(1).SetFocus
    Exit Sub
End If
 If Len(RickLecdetail) = 0 Then
    MsgBox "Please Enter Lecture Details.", vbInformation, App.Title
    txtfields(1).SetFocus
    Exit Sub
End If
If Len(txtfields(2)) = 0 Then
    MsgBox "Please Enter Lecture Prepared BY.", vbInformation, App.Title
    txtfields(1).SetFocus
    Exit Sub
End If
Dim rs As New adodb.Recordset
Dim check As String
If Check1.Value = 1 Then
    check = "Y"
Else
    check = "N"
End If

Dim cmd As New adodb.Command
Dim con As New adodb.Connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LectureInformation"
cmd(1) = Format(Trim(txtfields(0)), "00000")
cmd(2) = Trim(txtfields(1))
cmd(3) = Mid(Trim(Combo1.Text), 1, 5)
cmd(4) = Mid(Trim(Combo2.Text), 1, 5)
cmd(5) = Trim(RickLecdetail.TextRTF)
cmd(6) = Trim(txtfields(2))
cmd(7) = check
cmd(8) = Date
cmd(9) = "DSL"
cmd.Execute
MsgBox "Save successfully.", vbInformation, "Student Management System"
Call ShowFlexData
cmdnew.SetFocus

End Sub
Private Sub Combo1_Click()
'Dim rs2 As New adodb.Recordset
'Set rs2 = GetData("Select Sub_code,Sub_title from subject_info_sub where classId= '" & Mid(Trim(Combo1.Text), 1, 5) & "'")
'If Not rs2.EOF Then
'    Combo2.Clear
'    Do Until rs2.EOF
'        Combo2.AddItem rs2(0) + " - " + rs2(1)
'        rs2.MoveNext
'    Loop
'    Combo2.AddItem (" ")
'End If
'Combo2.Text = " "
'RickLecdetail.Text = ""
'txtfields(1) = ""
'txtfields(2) = ""
'If Check1.Value = 1 Then
'    Check1.Value = 0
'End If
'ShowFlexData
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Dim rs2 As New adodb.Recordset
If KeyAscii = 13 Then
    RickLecdetail.Text = ""
    txtfields(1) = ""
    txtfields(2) = ""
    If Check1.Value = 1 Then
        Check1.Value = 0
    End If
    Set rs2 = getdata("Select subjectID,subjectdsc from subjectinfo where classId= '" & Mid(Trim(Combo1.Text), 1, 5) & "'")
    If Not rs2.EOF Then
        Combo2.Clear
        Do Until rs2.EOF
            Combo2.AddItem rs2(0) + " - " + rs2(1)
            rs2.MoveNext
        Loop
        Combo2.AddItem (" ")
    End If
    Combo2.SetFocus
    ShowFlexData
End If

'ShowFlexData
End Sub
Private Sub Combo2_Click()
RickLecdetail.Text = ""
txtfields(1) = ""
txtfields(2) = ""
If Check1.Value = 1 Then
    Check1.Value = 0
End If
Call ShowFlexData
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    RickLecdetail.Text = ""
    txtfields(1) = ""
    txtfields(2) = ""
    If Check1.Value = 1 Then
        Check1.Value = 0
    End If
    Call ShowFlexData
    cmdnew.SetFocus
End If
End Sub

Private Sub Form_Load()

Dim rs1 As New adodb.Recordset
Dim rs2 As New adodb.Recordset
Set rs1 = getdata("Select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo1.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
    Combo1.AddItem (" ")
'    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End If

'    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0


With MSFlexGrid1
    .Rows = 1
    .Cols = 4
    .Col = 0: .Text = "  Lecture ID #"
    .Col = 1: .Text = " Description"
    .Col = 2: .Text = "     Prepared By "
    .Col = 3: .Text = "     Open For student "
    .ColWidth(0) = 1200
    .ColWidth(1) = 2500
    .ColWidth(2) = 2000
    .ColWidth(3) = 1900
    
    
End With
End Sub


Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New adodb.Recordset

Set rs = getdata("SELECT  LectureID,LectureDsc,LectureDetail,LecLessonPrepareBY,LIOpenForStu From LectureInfo where classid='" & Mid(Trim(Combo1.Text), 1, 5) & "' and SubjectID ='" & Mid(Trim(Combo2.Text), 1, 5) & "'")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = rs!LectureID
                .TextMatrix(i, 1) = rs!LectureDsc
                .TextMatrix(i, 2) = rs!LecLessonPrepareBY
'                LectureDetail = rs!LectureDetail
                If rs!LIOpenForStu = "Y" Then
                    .TextMatrix(i, 3) = "Y"
                Else
                    .TextMatrix(i, 3) = "N"
                End If
                i = i + 1
            rs.MoveNext
        Loop
    End With
Else
    MSFlexGrid1.Rows = 1
 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub
Public Function getdata(SQLString As String) As adodb.Recordset
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
Dim rs As New adodb.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = SQLString

 Set rs = cmd.Execute
Set getdata = rs
End Function


Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
Dim rs As New adodb.Recordset
txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtfields(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = "Y" Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If
RickLecdetail = ""
Set rs = getdata("SELECT  LectureDetail From LectureInfo where classid='" & Mid(Trim(Combo1.Text), 1, 5) & "' and SubjectID ='" & Mid(Trim(Combo2.Text), 1, 5) & "'and LectureID='" & Trim(txtfields(0)) & "'")
If Not rs.EOF Then
    RickLecdetail = rs!LectureDetail
End If
Exit Sub
errdes:
'MsgBox err.Description, vbInformation, App.Title

End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   Select Case Index
     Case 1
         txtfields(2).SetFocus
     Case 2
          Check1.SetFocus
   End Select
End If
End Sub
