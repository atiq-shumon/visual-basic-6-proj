VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLessonPlanMain 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Topic"
      Height          =   375
      Left            =   60
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7470
      Width           =   1125
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   345
      Left            =   4620
      TabIndex        =   11
      ToolTipText     =   "Click to insert new information"
      Top             =   7440
      Width           =   885
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Create Lesson Code"
      Height          =   345
      Left            =   5550
      TabIndex        =   10
      ToolTipText     =   "Click to save"
      Top             =   7440
      Width           =   1605
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   345
      Left            =   7200
      TabIndex        =   12
      ToolTipText     =   "Click to Delete"
      Top             =   7440
      Width           =   915
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   345
      Left            =   8160
      TabIndex        =   13
      ToolTipText     =   "Click to Exit"
      Top             =   7440
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   765
      Left            =   30
      ScaleHeight     =   705
      ScaleWidth      =   9135
      TabIndex        =   16
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
         ForeColor       =   &H00CEF0F7&
         Height          =   285
         Left            =   3390
         TabIndex        =   30
         Top             =   150
         Width           =   1425
      End
      Begin VB.Image Image1 
         Height          =   840
         Left            =   -90
         Picture         =   "frmLessonPlanMain.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   9225
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2565
      Left            =   0
      TabIndex        =   15
      Top             =   750
      Width           =   9225
      Begin VB.CommandButton Command2 
         Caption         =   "::"
         Height          =   315
         Left            =   2820
         TabIndex        =   29
         Top             =   150
         Width           =   345
      End
      Begin VB.ComboBox Combo10 
         Height          =   315
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2040
         Width           =   5145
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2040
         Width           =   1545
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select Section"
         Top             =   870
         Width           =   1545
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   5145
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Select Term"
         Top             =   1680
         Width           =   1545
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   900
         Width           =   5145
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1290
         Width           =   5145
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3900
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   510
         Width           =   5145
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   4
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
         Width           =   1545
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1290
         TabIndex        =   14
         Top             =   165
         Width           =   1515
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Title"
         Height          =   195
         Index           =   2
         Left            =   2940
         TabIndex        =   27
         Top             =   2130
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam ID"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   26
         Top             =   2100
         Width           =   600
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
         Index           =   0
         Left            =   150
         TabIndex        =   24
         Top             =   1740
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section Title"
         Height          =   195
         Index           =   3
         Left            =   2940
         TabIndex        =   23
         Top             =   930
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section ID"
         Height          =   195
         Left            =   150
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   540
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lesson Plan #"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   19
         Top             =   180
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject ID"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class ID"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Top             =   570
         Width           =   585
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4035
      Left            =   0
      TabIndex        =   31
      Top             =   3330
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   7117
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   15005934
      BackColorSel    =   -2147483624
      ForeColorSel    =   16711680
      BackColorBkg    =   15724265
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   4560
      Top             =   7380
      Width           =   4605
   End
End
Attribute VB_Name = "frmLessonPlanMain"
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
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString

Set rs = getdata("select Srl_no from ls_plan_topic where Srl_no='" & Trim(txtFields(0)) & "'")

If Not rs.EOF Then
   MsgBox "First Delete all Details under this topic", vbInformation, cmp
   Exit Sub
End If

 Set cmd.ActiveConnection = con
    If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical) = vbYes Then

        cmd.CommandType = adCmdText
        cmd.CommandText = "Delete from LectureInfo  where ClassID ='" & Mid(Trim(Combo1.Text), 1, 5) & "' and subjectid='" & Mid(Trim(Combo2.Text), 1, 5) & "' and LectureID= '" & Trim(txtFields(0)) & "' "
        cmd.Execute
        MsgBox "Delete successfully .", vbInformation, App.Title
        txtFields(0) = ""
'        txtfields(1) = ""
'        txtfields(2) = ""
        RickLecdetail = ""
        
'        If Check1.Value = 1 Then
'            Check1.Value = False
'        End If
'
        Call ShowFlexData
        
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdnew_Click()
txtFields(0).Text = ""
'If Len(Combo1) = 0 Then Exit Sub
'If Len(Combo2) = 0 Then Exit Sub
'Dim rs As New adodb.Recordset
'Dim cmd As New adodb.command
'Dim con As New adodb.connection
'con.Open GConnString
'cmd.ActiveConnection = con
'Set rs = GetData("select max (LectureID+1)from LectureInfo where classId='" & Mid(Trim(Combo1.Text), 1, 5) & "' and SubjectID='" & Mid(Trim(Combo2.Text), 1, 5) & "'")
'If Not rs.EOF Then
'    txtfields(0) = IIf(IsNull(rs(0)) = True, "00001", Format(rs(0), "00000"))
'Else
'    txtfields(0) = "00001"
'End If

'For i = 1 To 2
'    txtfields(i) = ""
'Next
'RickLecdetail = ""
'
'txtfields(1).SetFocus
End Sub
Private Sub cmdSAVE_Click()
If Len(Combo1.Text) = 0 Then
    MsgBox "Class Id Mandatory...Please Verify.", vbInformation, App.Title
    Combo1.SetFocus
    Exit Sub
End If
 If Len(Combo5.Text) = 0 Then
    MsgBox "Section Id Mandatory...Please Verify.", vbInformation, App.Title
    Combo5.SetFocus
    Exit Sub
End If

If Len(Combo2.Text) = 0 Then
    MsgBox "Subject Id Mandatory...Please Verify.", vbInformation, App.Title
    Combo2.SetFocus
    Exit Sub
End If

If Len(Combo7.Text) = 0 Then
    MsgBox "Term Id Mandatory...Please Verify.", vbInformation, App.Title
    Combo7.SetFocus
    Exit Sub
End If

If Len(Combo9.Text) = 0 Then
    MsgBox "Exam Id Mandatory...Please Verify.", vbInformation, App.Title
    Combo1.SetFocus
    Exit Sub
End If

'If Len(txtfields(4)) = 0 Then
'    MsgBox "Topic Title Required.", vbInformation, App.Title
'    txtfields(4).SetFocus
'    Exit Sub
'End If

Dim rs As New ADODB.Recordset


Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LS_PLAN_MASTER_Save"
cmd(1) = "S"
cmd(2) = Val(Trim(txtFields(0)))
cmd(3) = Trim(Combo1.Text)
cmd(4) = Trim(Combo5.Text)
cmd(5) = Trim(Combo7.Text)
cmd(6) = Trim(Combo9.Text)
cmd(7) = Trim(Combo2.Text)
cmd.Execute
MsgBox "Save successfully.", vbInformation, "Student Management System"
Call ShowFlexData
cmdnew.SetFocus

End Sub
Private Sub Combo1_Click()
   load_section
   load_subject
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select ClassName from classinfo where classId= '" & Trim(Combo1.Text) & "'")
   
   If Not rs.EOF Then
      Combo3.Text = Trim(rs(0))
   End If
   ShowFlexData
   
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
'Dim rs2 As New adodb.Recordset
'If KeyAscii = 13 Then
'    Set rs2 = GetData("Select subjectID,subjectdsc from subjectinfo where classId= '" & Mid(Trim(Combo1.Text), 1, 5) & "'")
'    If Not rs2.EOF Then
'        Combo2.Clear
'        Do Until rs2.EOF
'            Combo2.AddItem rs2(0) + " - " + rs2(1)
'            rs2.MoveNext
'        Loop
'        Combo2.AddItem (" ")
'    End If
'    Combo2.SetFocus
'    ShowFlexData
'End If

'ShowFlexData
End Sub

Private Sub Combo2_Click()
Dim rs As New ADODB.Recordset
    Set rs = getdata("Select a.Sub_title from subject_info_sub a ,subjectinfomain b  where a.M_code=b.M_code and a.Class_code='" & Trim(Combo1.Text) & "' and  a.Sub_code='" & Trim(Combo2.Text) & "'")

   If Not rs.EOF Then
      Combo4.Text = Trim(rs(0))
   End If
   ShowFlexData
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'
'    txtfields(1) = ""
'    txtfields(2) = ""
'    If Check1.Value = 1 Then
'        Check1.Value = 0
'    End If
'    Call ShowFlexData
'    cmdnew.SetFocus
'End If
End Sub



Private Sub Combo3_LostFocus()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select classId from classinfo where ClassName= '" & Trim(Combo3.Text) & "'")
   
   If Not rs.EOF Then
      Combo1.Text = Trim(rs(0))
   End If
End Sub

Private Sub Combo4_LostFocus()
 Dim rs As New ADODB.Recordset
   Set rs = getdata("Select a.Sub_code  from subject_info_sub a ,subjectinfomain b  where a.M_code=b.M_code and a.Class_code='" & Trim(Combo1.Text) & "' and a.Sub_title='" & Trim(Combo4.Text) & "'")

   If Not rs.EOF Then
      Combo2.Text = Trim(rs(0))
   End If
End Sub

Private Sub Combo5_Click()
   ShowFlexData
   load_section_title
End Sub

Private Sub Combo6_LostFocus()
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("Select SectionID from SectionInfo where ClassID='" & Trim(Combo1.Text) & "'and Sectiondsc='" & Trim(Combo6.Text) & "'")
If Not rs1.EOF Then
   Combo5.Text = Trim(rs1(0))
End If
Exit Sub
End Sub

Private Sub Combo7_Click()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select ETypeName from ExamTypeInfo where ETypeID= '" & Trim(Combo7.Text) & "'")
   
   If Not rs.EOF Then
      Combo8.Text = rs(0)
   End If
   load_exam_sub
   ShowFlexData
End Sub
Private Sub load_exam_sub()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select Exam_code,Exam_title from Exam_setup where Group_code= '" & Trim(Combo7.Text) & "'")
   
   Combo9.Clear
   Combo10.Clear
   If Not rs.EOF Then
    Do Until rs.EOF
      Combo9.AddItem rs(0)
      Combo10.AddItem rs(1)
      rs.MoveNext
    Loop
   
   End If
   
End Sub
Private Sub Combo8_LostFocus()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select ETypeID from ExamTypeInfo where ETypeName= '" & Trim(Combo8.Text) & "'")

   If Not rs.EOF Then
      Combo7.Text = rs(0)
   End If

End Sub

Private Sub Combo9_Click()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select Exam_title from Exam_setup where Group_code= '" & Trim(Combo7.Text) & "' and Exam_code='" & Trim(Combo9.Text) & "'")
   If Not rs.EOF Then
      Combo10.Text = Trim(rs(0))
   End If
   ShowFlexData
End Sub

Private Sub Command1_Click()

 If Len(txtFields(0).Text) = 0 Then
    MsgBox "Please put a Valid  Lesson Serial First", vbInformation, App.Title
    Exit Sub
 End If
 
 Dim rs As New ADODB.Recordset
 Set rs = getdata("Select Srl_no from Ls_plan_master where Srl_no='" & txtFields(0).Text & "'")
 
 If rs.EOF Then
    MsgBox "No such Serial No exists ...Please Verify.", vbInformation, App.Title
    Exit Sub
 End If

  frmTopic_serial.Show 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys (Chr(9))
    End If
    If KeyAscii = 27 Then
       Unload Me
    End If
End Sub

Private Sub Form_Load()
txtFields(0).Text = 0
 load_class
' load_section
 load_exam

With MSFlexGrid1
    .Rows = 1
    .Cols = 7
    .Col = 0: .Text = " Serial No "
    .Col = 1: .Text = " Class Id"
    .Col = 2: .Text = " Section Id"
    .Col = 3: .Text = " Term Id "
    .Col = 4: .Text = " Exam Id "
    .Col = 5: .Text = " Subject Id "
    .Col = 6: .Text = " Subject Title "
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 3000
    
End With
End Sub
Private Sub load_class()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Combo1.Clear
Combo3.Clear
Set rs1 = getdata("Select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo1.AddItem rs1(0)
        Combo3.AddItem rs1(1)
        rs1.MoveNext
    Loop
    Combo1.AddItem (" ")
'    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End If

'    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0


End Sub

Private Sub load_exam()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Combo7.Clear
Combo8.Clear
Set rs1 = getdata("Select ETypeID,ETypeName from ExamTypeInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo7.AddItem rs1(0)
        Combo8.AddItem rs1(1)
        rs1.MoveNext
    Loop
'    Combo1.AddItem (" ")
End If
End Sub
Private Sub load_subject()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Combo2.Clear
Combo4.Clear
Set rs1 = getdata("Select a.Sub_code,a.Sub_title from subject_info_sub a  where a.Class_code='" & Trim(Combo1.Text) & "'")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo2.AddItem rs1(0)
        Combo4.AddItem rs1(1)
        rs1.MoveNext
    Loop
'    Combo1.AddItem (" ")
End If
End Sub

Private Sub load_section()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Combo5.Clear
Combo6.Clear
Set rs1 = getdata("Select SectionID,Sectiondsc from SectionInfo where ClassID='" & Trim(Combo1.Text) & "'")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo5.AddItem Trim(rs1(0))
        Combo6.AddItem Trim(rs1(1))
        rs1.MoveNext
    Loop
End If
'Combo6.Text = Trim(Combo6.List(0))
'Combo5.Text = Trim(Combo5.List(0))
End Sub
Private Sub load_section_title()
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("Select Sectiondsc from SectionInfo where ClassID='" & Trim(Combo1.Text) & "'and SectionID='" & Trim(Combo5.Text) & "'")
If Not rs1.EOF Then
   Combo6.Text = Trim(rs1(0))
End If
End Sub
Private Sub load_section_id()

End Sub
Private Sub ShowFlexData()
'On Error GoTo ErrDes
Dim rs As New ADODB.Recordset

 Set rs = getdata("Select a.Srl_no,a.Class_id,a.Section_id,a.Term_id,a.Exam_id,a.Sub_id,(select b.Sub_title from subject_info_sub b where b.Class_code=a.Class_id and b.Sub_code=a.Sub_id ) as sub_title from Ls_plan_master a where a.Class_id='" & Trim(Combo1.Text) & " ' and a.Section_id='" & Trim(Combo5.Text) & "' and a.Term_id='" & Trim(Combo7.Text) & "' and  a.Exam_id= '" & Trim(Combo9.Text) & "' and a.Sub_id='" & Trim(Combo2.Text) & "'")

If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = rs!srl_no
                .TextMatrix(i, 1) = rs!Class_id
                .TextMatrix(i, 2) = rs!Section_id
                .TextMatrix(i, 3) = rs!Term_id
                .TextMatrix(i, 4) = rs!Exam_id
                .TextMatrix(i, 5) = rs!Sub_id
                .TextMatrix(i, 6) = rs!Sub_title
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

Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
txtFields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtFields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtFields(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = "Y" Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If
RickLecdetail = ""
Set rs = getdata("SELECT  LectureDetail From LectureInfo where classid='" & Mid(Trim(Combo1.Text), 1, 5) & "' and SubjectID ='" & Mid(Trim(Combo2.Text), 1, 5) & "'and LectureID='" & Trim(txtFields(0)) & "'")
If Not rs.EOF Then
    RickLecdetail = rs!LectureDetail
End If
Exit Sub
errdes:
'MsgBox err.Description, vbInformation, App.Title

End Sub




Private Sub MSFlexGrid1_DblClick()
   
End Sub

Private Sub txtfields_Change(Index As Integer)
      Select Case Index
          Case 0
                  If Len(txtFields(0).Text) > 0 Then
                     cmdSave.Enabled = False
                  Else
                    cmdSave.Enabled = True
                  End If
          End Select
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   Select Case Index
     Case 1
         txtFields(2).SetFocus
     Case 2
          Check1.SetFocus
   End Select
End If
End Sub
