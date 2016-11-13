VERSION 5.00
Begin VB.Form RptMarksDistribution 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   -60
      Picture         =   "rptMarksDistribution.frx":0000
      ScaleHeight     =   690
      ScaleWidth      =   9645
      TabIndex        =   6
      Top             =   0
      Width           =   9705
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   7
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marksdistribution Report"
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
         Left            =   2760
         TabIndex        =   8
         Top             =   150
         Width           =   2835
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   -120
         Picture         =   "rptMarksDistribution.frx":C897
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9795
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H8000000C&
      Caption         =   "View"
      Height          =   375
      Left            =   6930
      TabIndex        =   3
      ToolTipText     =   "Click for Attendance"
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Height          =   1035
      Left            =   0
      TabIndex        =   5
      Top             =   750
      Width           =   10365
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   0
         TabIndex        =   9
         Top             =   180
         Width           =   9015
         Begin VB.ComboBox cboClass 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   330
            Width           =   2865
         End
         Begin VB.ComboBox CboExamType 
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   330
            Width           =   2655
         End
         Begin VB.ComboBox CboExamID 
            Height          =   315
            Left            =   5760
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   330
            Width           =   3165
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   150
            Width           =   375
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exam Term"
            Height          =   195
            Index           =   3
            Left            =   3120
            TabIndex        =   11
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exam"
            Height          =   195
            Index           =   4
            Left            =   5760
            TabIndex        =   10
            Top             =   150
            Width           =   390
         End
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   375
      Left            =   7980
      TabIndex        =   4
      ToolTipText     =   "Click to Exit"
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   6870
      Top             =   1830
      Width           =   2175
   End
End
Attribute VB_Name = "RptMarksDistribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub load_marks()
  Dim rs As New ADODB.Recordset
 Set rs = getdata("select a.fullmarks,a.passmarks from SubjectMarksDistribution  a where a.ClassId = '" & Mid(Trim(cboClass.Text), 1, 5) & "' and a.SubjectId= '" & Mid(Trim(CboSubject.Text), 1, 5) & "' and a.term_code='" & Mid(Trim(CboExamType.Text), 1, 2) & "' and a.Exam_code='" & Mid(Trim(CboExamID.Text), 1, 2) & "'and a.CategoryID='" & Mid(CboCategory, 1, 5) & "'")

 If Not rs.EOF Then
   txtfullMarks = rs(0)
   txtPassMarks = rs(1)
 End If

End Sub

Private Sub CboCategory_Click()
'''get_resultMain
 load_marks
 '''LoadStuID
  '''show_grid
 
End Sub

Private Sub CboExamType_Click()
  load_exam_sub
End Sub
Private Sub cmdDelete_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
Dim classId As String
Dim SectionID As String

If Len(cboClass) = 0 Then
  MsgBox "Please select a class first", vbInformation, cmp
  cboClass.SetFocus
  Exit Sub
End If

If Len(CboSection.Text) = 0 Then
   MsgBox "Please select a Section ", vbInformation, cmp
   CboSection.SetFocus
   Exit Sub
End If

If Len(CboSubject) = 0 Then
   MsgBox "Please select a Subject ", vbInformation, cmp
   CboSubject.SetFocus
   Exit Sub
End If
  
If Len(cmdAcaYear) = 0 Then
   MsgBox "Please select an Academic Year ", vbInformation, cmp
   cmdAcaYear.SetFocus
   Exit Sub
End If

If Len(CboExamType) = 0 Then
   MsgBox "Please select an Exam Type ", vbInformation, cmp
   CboExamType.SetFocus
   Exit Sub
End If

If Len(CboExamID) = 0 Then
   MsgBox "Please select an Exam  ", vbInformation, cmp
   CboExamID.SetFocus
   Exit Sub
End If
  
If Len(Combo2.Text) = 0 Then
   MsgBox "Please select a Shift ", vbInformation, cmp
   Combo2.SetFocus
   Exit Sub
End If


If Len(txtMarks.Text) = 0 Then
    MsgBox "Please put obtained marks ", vbInformation, cmp
    txtMarks.SetFocus
   Exit Sub
  End If



  Set rs = getdata("select StdID from result_sub where M_Slr_no=" & Trim(txtSerial) & " and S_Slr_no =" & Trim(txtSerialSub) & "")
  If rs.EOF Then
    MsgBox "No such Student Exists...please verify.", vbInformation, cmp
    Exit Sub
  End If

  If MsgBox("Are you Sure to Delete ? ", vbInformation + vbYesNo + vbDefaultButton1, cmp) = vbYes Then
        con.Open GConnString
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "Result_Save"
        cmd(1) = "d"
        cmd(2) = Val(txtSerial)
        cmd(3) = Val(txtSerialSub)
        cmd(4) = Mid(cboClass, 1, 5)
        cmd(5) = Mid(CboSection, 1, 5)
        cmd(6) = Mid(Combo2, 1, 1)
        cmd(7) = Mid(CboSubject, 1, 5)
        cmd(8) = Trim(cmdAcaYear)
        cmd(9) = Mid(CboExamType, 1, 2)
        cmd(10) = Mid(CboExamID, 1, 2)
        cmd(11) = Mid(CboCategory, 1, 5)
        cmd(12) = Mid(List1, 1, 10)
        '''cmd(13) = txtfields(3)
        cmd(14) = txtMarks
        cmd(15) = txtPassMarks
        cmd(16) = txtfullMarks
        cmd(17) = soft_user
        cmd(18) = Date
    
        cmd.Execute
        
        '''Call LoadStuID
       
'        For i = 3 To 3
'        txtfields(i) = ""
'        Next

        MsgBox "Deleted Successfully.", vbInformation, "Student Management System"
       
'       get_resultMain
'       Call LoadStuID
'        show_grid
'        txtMarks = 0
'       List1.SetFocus
End If
      
'Else
'        Exit Sub
'End If

End Sub

Private Sub cmdMarksheet_Click()
     rptMode = 8
    Screen.MousePointer = vbHourglass
    frmViewer.Show 1
End Sub

Private Sub cmdSAVE_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
Dim classId As String
Dim SectionID As String

If Len(cboClass) = 0 Then
  MsgBox "Please select a class first", vbInformation, cmp
  cboClass.SetFocus
  Exit Sub
End If

If Len(CboSection.Text) = 0 Then
   MsgBox "Please select a Section ", vbInformation, cmp
   CboSection.SetFocus
   Exit Sub
End If

If Len(CboSubject) = 0 Then
   MsgBox "Please select a Subject ", vbInformation, cmp
   CboSubject.SetFocus
   Exit Sub
End If
  
If Len(cmdAcaYear) = 0 Then
   MsgBox "Please select an Academic Year ", vbInformation, cmp
   cmdAcaYear.SetFocus
   Exit Sub
End If

If Len(CboExamType) = 0 Then
   MsgBox "Please select an Exam Type ", vbInformation, cmp
   CboExamType.SetFocus
   Exit Sub
End If

If Len(CboExamID) = 0 Then
   MsgBox "Please select an Exam  ", vbInformation, cmp
   CboExamID.SetFocus
   Exit Sub
End If
  
If Len(Combo2.Text) = 0 Then
   MsgBox "Please select a Shift ", vbInformation, cmp
   Combo2.SetFocus
   Exit Sub
End If

If Len(List1.Text) = 0 Then
    MsgBox "Please a Student ID from the list ", vbInformation, cmp
    List1.SetFocus
   Exit Sub
  End If

If Len(txtMarks.Text) = 0 Then
    MsgBox "Please put obtained marks ", vbInformation, cmp
    txtMarks.SetFocus
   Exit Sub
  End If

If Len(CboCategory) = 0 Then
   MsgBox "Please Marks category required ", vbInformation, cmp
    CboCategory.SetFocus
   Exit Sub
  End If


'If Option1(0).Value = True Then
'   Set rs = GetData("select ClassId from StudentAttendanceLeaveInfo where ClassID='" & Trim(Mid(List1, 1, 5)) & "' and SectionID ='" & Mid(Combo1, 1, 5) & "'  and attn_date ='" & Format(MaskEdBox3, "dd mmm yyyy") & "'")
'  If Not rs.EOF Then
'    MsgBox "Attendance of all student of " & Trim(Mid(List1, 6, 15)) & " has already been completed on date :" & Format(MaskEdBox3, "dd mmm yyyy") & " ", vbInformation, cmp
'    Exit Sub
'  End If
'End If
'If Option1(1).Value = True Then
'   Set rs = GetData("select ClassId from StudentAttendanceLeaveInfo where ClassID='" & Trim(Mid(List1, 1, 5)) & "' and SectionID ='" & Mid(Combo1, 1, 5) & "'  and attn_date ='" & Format(MaskEdBox3, "dd mmm yyyy") & "' and StudentID='" & Mid(List2, 1, 10) & "'")
'  If Not rs.EOF Then
'     MsgBox "Attendance of Mr." & Mid(List2, 11, 80) & "  already been completed on date: " & MaskEdBox3.Text & " ", vbInformation, cmp
'    Exit Sub
'  End If
'End If
'
'        con.Open GConnString
'        cmd.ActiveConnection = con
'        cmd.CommandType = adCmdStoredProc
'        cmd.CommandText = "Result_Save"
'        cmd(1) = "s"
'        cmd(2) = Val(txtSerial)
'        cmd(3) = Val(txtSerialSub)
'        cmd(4) = Mid(cboClass, 1, 5)
'        cmd(5) = Mid(CboSection, 1, 5)
'        cmd(6) = Mid(Combo2, 1, 1)
'        cmd(7) = Mid(CboSubject, 1, 5)
'        cmd(8) = Trim(cmdAcaYear)
'        cmd(9) = Mid(CboExamType, 1, 2)
'        cmd(10) = Mid(CboExamID, 1, 2)
'        cmd(11) = Mid(CboCategory, 1, 5)
'        cmd(12) = Mid(List1, 1, 10)
'        cmd(13) = txtfields(3)
'        cmd(14) = txtMarks
'        cmd(15) = txtPassMarks
'        cmd(16) = txtfullMarks
'        cmd(17) = soft_user
'        cmd(18) = Date
'
'        cmd.Execute
'
'        Call LoadStuID
'
'        For i = 3 To 3
'        txtfields(i) = ""
'        Next
'
'        MsgBox "Save Successfully.", vbInformation, "Student Management System"
'
'       get_resultMain
'       Call LoadStuID
'        show_grid
'        txtMarks = 0
'       List1.SetFocus
'
''Else
''        Exit Sub
''End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

'
'Private Sub ComStuId_Click()
'Label3.Caption = ""
'
'
'
'txtfields(3) = ""
'
'Dim rs As New ADODB.Recordset
'Set rs = getdata("SELECT  distinct   StudentEvaluation.StudentID, StudentInfo.StudentName,  StudentEvaluation.Shift, " + _
'        "StudentEvaluation.ClassId, ClassInfo.ClassName, StudentEvaluation.SectionId, SectionInfo.Sectiondsc, StudentEvaluation.ClassRoll, " + _
'        "StudentEvaluation.ActiveClass FROM StudentEvaluation INNER JOIN SectionInfo ON StudentEvaluation.SectionId = SectionInfo.SectionID INNER JOIN " + _
'        "ClassInfo ON StudentEvaluation.ClassId = ClassInfo.ClassID INNER JOIN StudentInfo ON StudentEvaluation.StudentID = StudentInfo.StudentID where StudentEvaluation.studentid='" & ComStuId & "'and StudentEvaluation.ActiveClass='Y'")
'
'        If Not rs.EOF Then
'            Label3.Caption = "" & rs!StudentName
'
'
''            txtfields(0) = "" & rs!Shift
'            txtfields(3) = "" & rs!ClassRoll
'
'        End If
'End Sub

Private Sub ComStuId_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    scmdAttend.SetFocus
End If
End Sub


Private Sub ComStuId_LostFocus()
If Len(ComStuId) = 0 Then Exit Sub
Dim rs As New ADODB.Recordset
Set rs = getdata("select StudentId from StudentEvaluation where StudentId ='" & ComStuId.Text & "'")
If rs.EOF Then
    MsgBox "Invalid Id.", vbCritical, "School Management System"
    ComStuId.Text = ""
    Exit Sub
End If
End Sub

Private Sub CmdEdit_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
Dim classId As String
Dim SectionID As String

If Len(cboClass) = 0 Then
  MsgBox "Please select a class first", vbInformation, cmp
  cboClass.SetFocus
  Exit Sub
End If

If Len(CboSection.Text) = 0 Then
   MsgBox "Please select a Section ", vbInformation, cmp
   CboSection.SetFocus
   Exit Sub
End If

If Len(CboSubject) = 0 Then
   MsgBox "Please select a Subject ", vbInformation, cmp
   CboSubject.SetFocus
   Exit Sub
End If
  
If Len(cmdAcaYear) = 0 Then
   MsgBox "Please select an Academic Year ", vbInformation, cmp
   cmdAcaYear.SetFocus
   Exit Sub
End If

If Len(CboExamType) = 0 Then
   MsgBox "Please select an Exam Type ", vbInformation, cmp
   CboExamType.SetFocus
   Exit Sub
End If

If Len(CboExamID) = 0 Then
   MsgBox "Please select an Exam  ", vbInformation, cmp
   CboExamID.SetFocus
   Exit Sub
End If
  
If Len(Combo2.Text) = 0 Then
   MsgBox "Please select a Shift ", vbInformation, cmp
   Combo2.SetFocus
   Exit Sub
End If


If Len(txtMarks.Text) = 0 Then
    MsgBox "Please put obtained marks ", vbInformation, cmp
    txtMarks.SetFocus
   Exit Sub
  End If



  Set rs = getdata("select StdID from result_sub where M_Slr_no=" & Trim(txtSerial) & " and S_Slr_no =" & Trim(txtSerialSub) & "")
  If rs.EOF Then
    MsgBox "No such Student Exists...please verify.", vbInformation, cmp
    Exit Sub
  End If

'If Option1(1).Value = True Then
'   Set rs = GetData("select ClassId from StudentAttendanceLeaveInfo where ClassID='" & Trim(Mid(List1, 1, 5)) & "' and SectionID ='" & Mid(Combo1, 1, 5) & "'  and attn_date ='" & Format(MaskEdBox3, "dd mmm yyyy") & "' and StudentID='" & Mid(List2, 1, 10) & "'")
'  If Not rs.EOF Then
'     MsgBox "Attendance of Mr." & Mid(List2, 11, 80) & "  already been completed on date: " & MaskEdBox3.Text & " ", vbInformation, cmp
'    Exit Sub
'  End If
'End If
'
'        con.Open GConnString
'        cmd.ActiveConnection = con
'        cmd.CommandType = adCmdStoredProc
'        cmd.CommandText = "Result_Save"
'        cmd(1) = "u"
'        cmd(2) = Val(txtSerial)
'        cmd(3) = Val(txtSerialSub)
'        cmd(4) = Mid(cboClass, 1, 5)
'        cmd(5) = Mid(CboSection, 1, 5)
'        cmd(6) = Mid(Combo2, 1, 1)
'        cmd(7) = Mid(CboSubject, 1, 5)
'        cmd(8) = Trim(cmdAcaYear)
'        cmd(9) = Mid(CboExamType, 1, 2)
'        cmd(10) = Mid(CboExamID, 1, 2)
'        cmd(11) = Mid(CboCategory, 1, 5)
'        cmd(12) = Mid(List1, 1, 10)
'        cmd(13) = txtfields(3)
'        cmd(14) = txtMarks
'        cmd(15) = txtPassMarks
'        cmd(16) = txtfullMarks
'        cmd(17) = soft_user
'        cmd(18) = Date
'        cmd.Execute
'
'        Call LoadStuID
'
'        For i = 3 To 3
'        txtfields(i) = ""
'        Next
'
'        MsgBox "Edited Successfully.", vbInformation, "Student Management System"
'
'       get_resultMain
'       Call LoadStuID
'        show_grid
'        txtMarks = 0
'       List1.SetFocus
'
''Else
''        Exit Sub
''End If

End Sub

Private Sub cmdPrint_Click()
    rptMode = 10
    Screen.MousePointer = vbHourglass
    frmViewer.Show 1
End Sub

Private Sub Combo1_Click()
  '' LoadStuID
End Sub


Private Sub Command1_Click()
  
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
load_class
load_Aca_year
load_Scction
load_exam

End Sub
Private Sub load_exam()
Dim rs1 As New ADODB.Recordset
CboExamType.Clear
Set rs1 = getdata("Select ETypeID,ETypeName from ExamTypeInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        CboExamType.AddItem rs1(0) + "-" + rs1(1)
       rs1.MoveNext
    Loop
End If
End Sub
Private Sub load_exam_sub()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select Exam_code,Exam_title from Exam_setup where Group_code= '" & Mid(Trim(CboExamType.Text), 1, 2) & "'")
   
   CboExamID.Clear
   If Not rs.EOF Then
     Do Until rs.EOF
      CboExamID.AddItem rs(0) + "-" + rs(1)
      rs.MoveNext
     Loop
         
   End If
   
End Sub
Private Sub load_Scction()
 Dim rs1 As New ADODB.Recordset
'CboSection.Clear
End Sub
Private Sub load_Aca_year()
'  Dim i As Integer
'For i = 2000 To 2050
'  cmdAcaYear.AddItem i
'Next i
End Sub
 
 Private Sub load_class()
   Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT ClassID, ClassName FROM  classinfo")
cboClass.Clear
If Not rs.EOF Then
    Do Until rs.EOF
        cboClass.AddItem rs(0) + "-" + rs(1)
        rs.MoveNext
    Loop
End If
 End Sub







Private Sub get_roll()
'Dim rs As New ADODB.Recordset
'Set rs = getdata("select classRoll From StudentAdmission where StudentId='" & Mid(List1, 1, 10) & "'" & _
' " and serial_no=(select max(serial_no)  From StudentAdmission  where StudentId='" & Mid(List1, 1, 10) & "')")
'If Not rs.EOF Then
'  txtfields(3).Text = rs(0)
'End If
End Sub

Private Sub Option1_Click(Index As Integer)
'   Select Case Index
'          Case 0
'             List2.Enabled = False
'             txtfields(3).Text = ""
'          Case 1
'             List2.Enabled = True
'  End Select
End Sub

Private Sub List1_Click()
  get_roll
 get_name
 txtSerialSub = ""
End Sub
Private Sub get_name()
  Dim rs As New ADODB.Recordset
Set rs = getdata("select studentname From Studentinfo where StudentId='" & Mid(List1, 1, 10) & "'")
If Not rs.EOF Then
   lblStdname.Text = rs(0)
End If
End Sub

Private Sub MSFlexGrid1_Click()
' If MSFlexGrid1.Row > 0 Then
'  txtid = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
'  txtfields(3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
'  lblStdname = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
'  txtMarks = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
'  txtSerialSub = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
'End If

End Sub

Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub

Private Sub txtMarks_Change()
  If Not IsNumeric(txtMarks) Then
      txtMarks = ""
  ElseIf Val(txtMarks) > Val(txtfullMarks) Then
     txtMarks = 0
  End If
  End Sub
Private Sub txtMarks_GotFocus()
   txtMarks.SelStart = 0
   txtMarks.SelLength = Len(Trim(txtMarks.Text))
End Sub
Private Sub txtMarks_LostFocus()
  If Val(txtMarks) > Val(txtfullMarks) Then
     txtMarks = 0
  End If
End Sub
