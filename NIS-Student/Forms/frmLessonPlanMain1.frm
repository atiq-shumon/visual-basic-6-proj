VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLessonPlanMain1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3915
      Left            =   30
      TabIndex        =   13
      Top             =   4110
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   6906
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
      Top             =   8130
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   435
      Left            =   6270
      TabIndex        =   3
      ToolTipText     =   "Click to save"
      Top             =   8130
      Width           =   945
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   435
      Left            =   7260
      TabIndex        =   4
      ToolTipText     =   "Click to Delete"
      Top             =   8130
      Width           =   945
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   435
      Left            =   8250
      TabIndex        =   5
      ToolTipText     =   "Click to Exit"
      Top             =   8130
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   765
      Left            =   30
      ScaleHeight     =   705
      ScaleWidth      =   9135
      TabIndex        =   7
      Top             =   0
      Width           =   9195
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lesson Plan"
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
         Left            =   3630
         TabIndex        =   8
         Top             =   150
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3315
      Left            =   0
      TabIndex        =   6
      Top             =   750
      Width           =   9225
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3900
         TabIndex        =   30
         Top             =   2130
         Width           =   915
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   645
         Index           =   4
         Left            =   1290
         TabIndex        =   28
         Top             =   2460
         Width           =   7755
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1290
         TabIndex        =   26
         Top             =   2100
         Width           =   1575
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   3900
         TabIndex        =   24
         Text            =   "Combo4"
         Top             =   1740
         Width           =   5145
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "Select Subject"
         Top             =   1680
         Width           =   1545
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   3900
         TabIndex        =   20
         Text            =   "Combo4"
         Top             =   900
         Width           =   5145
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Select Subject"
         Top             =   900
         Width           =   1545
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   3900
         TabIndex        =   15
         Text            =   "Combo4"
         Top             =   1320
         Width           =   5145
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3900
         TabIndex        =   14
         Text            =   "Combo3"
         Top             =   510
         Width           =   5145
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select Subject"
         Top             =   1290
         Width           =   1545
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select Class"
         Top             =   510
         Width           =   1515
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1290
         TabIndex        =   9
         Top             =   135
         Width           =   1485
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week"
         Height          =   195
         Index           =   3
         Left            =   2970
         TabIndex        =   31
         Top             =   2160
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topic Title"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   29
         Top             =   2580
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topic Serial"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   2130
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term Title"
         Height          =   195
         Index           =   4
         Left            =   2940
         TabIndex        =   25
         Top             =   1770
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term ID"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section Title"
         Height          =   195
         Index           =   3
         Left            =   2940
         TabIndex        =   21
         Top             =   930
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section ID"
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   930
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Title"
         Height          =   195
         Index           =   2
         Left            =   2940
         TabIndex        =   17
         Top             =   1350
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Title"
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   16
         Top             =   540
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lesson Serial"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   12
         Top             =   180
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject ID"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class ID"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Top             =   570
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmLessonPlanMain1"
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

Private Sub cmddelete_Click()
Dim rs As New Recordset
Dim cmd As New Command
Dim con As New Connection
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

Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub cmdnew_Click()
If Len(Combo1) = 0 Then Exit Sub
If Len(Combo2) = 0 Then Exit Sub
Dim rs As New Recordset
Dim cmd As New Command
Dim con As New Connection
con.Open GConnString
cmd.ActiveConnection = con
Set rs = GetData("select max (LectureID+1)from LectureInfo where classId='" & Mid(Trim(Combo1.Text), 1, 5) & "' and SubjectID='" & Mid(Trim(Combo2.Text), 1, 5) & "'")
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
Private Sub cmdsave_Click()
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
Dim rs As New Recordset
Dim check As String
If Check1.Value = 1 Then
    check = "Y"
Else
    check = "N"
End If

Dim cmd As New Command
Dim con As New Connection
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
'Dim rs2 As New Recordset
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
Dim rs2 As New Recordset
If KeyAscii = 13 Then
    RickLecdetail.Text = ""
    txtfields(1) = ""
    txtfields(2) = ""
    If Check1.Value = 1 Then
        Check1.Value = 0
    End If
    Set rs2 = GetData("Select subjectID,subjectdsc from subjectinfo where classId= '" & Mid(Trim(Combo1.Text), 1, 5) & "'")
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

Dim rs1 As New Recordset
Dim rs2 As New Recordset
Set rs1 = GetData("Select ClassId,ClassName from ClassInfo")
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
On Error GoTo ErrDes
Dim rs As New Recordset

Set rs = GetData("SELECT  LectureID,LectureDsc,LectureDetail,LecLessonPrepareBY,LIOpenForStu From LectureInfo where classid='" & Mid(Trim(Combo1.Text), 1, 5) & "' and SubjectID ='" & Mid(Trim(Combo2.Text), 1, 5) & "'")
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
ErrDes:
MsgBox err.Description, vbInformation, App.Title
End Sub
Public Function GetData(SQLString As String) As Recordset
Dim cmd As New Command
Dim con As New Connection
Dim rs As New Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = SQLString

 Set rs = cmd.Execute
Set GetData = rs
End Function


Private Sub Label5_Click()

End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo ErrDes
Dim rs As New Recordset
txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtfields(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = "Y" Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If
RickLecdetail = ""
Set rs = GetData("SELECT  LectureDetail From LectureInfo where classid='" & Mid(Trim(Combo1.Text), 1, 5) & "' and SubjectID ='" & Mid(Trim(Combo2.Text), 1, 5) & "'and LectureID='" & Trim(txtfields(0)) & "'")
If Not rs.EOF Then
    RickLecdetail = rs!LectureDetail
End If
Exit Sub
ErrDes:
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
