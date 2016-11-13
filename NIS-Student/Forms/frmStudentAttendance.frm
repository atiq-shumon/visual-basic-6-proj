VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmStudentAttendance 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10110
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAttendenceReportofYear 
      BackColor       =   &H8000000C&
      Caption         =   "Annually Attendence Report"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      ToolTipText     =   "Click to Close"
      Top             =   6750
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   -60
      Picture         =   "frmStudentAttendance.frx":0000
      ScaleHeight     =   690
      ScaleWidth      =   10275
      TabIndex        =   23
      Top             =   0
      Width           =   10335
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   24
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Attendance Entry"
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
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   150
         Width           =   2940
      End
      Begin VB.Image Image1 
         Height          =   990
         Left            =   -120
         Picture         =   "frmStudentAttendance.frx":C897
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   10305
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H8000000C&
      Caption         =   "Print"
      Height          =   375
      Left            =   7440
      TabIndex        =   22
      ToolTipText     =   "Click to Print"
      Top             =   6750
      Width           =   1035
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000C&
      Caption         =   "Edit"
      Height          =   375
      Left            =   6390
      TabIndex        =   21
      ToolTipText     =   "Click for Editing Attendance"
      Top             =   6750
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   0
      TabIndex        =   10
      Top             =   750
      Width           =   10665
      Begin VB.Frame Frame5 
         Height          =   555
         Index           =   1
         Left            =   7050
         TabIndex        =   27
         Top             =   120
         Width           =   3015
         Begin VB.ComboBox CboExamID 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   30
            ToolTipText     =   "Select Exam Type"
            Top             =   150
            Width           =   915
         End
         Begin VB.ComboBox CboExamType 
            Height          =   315
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   29
            ToolTipText     =   "Select Exam Term"
            Top             =   150
            Width           =   1005
         End
         Begin VB.ComboBox cmdAcaYear 
            Height          =   315
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   28
            ToolTipText     =   "Select Academic Year "
            Top             =   150
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Class"
         Height          =   5055
         Left            =   120
         TabIndex        =   20
         Top             =   660
         Width           =   2415
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4680
            Left            =   90
            TabIndex        =   4
            Top             =   240
            Width           =   2205
         End
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   6510
         TabIndex        =   18
         Top             =   930
         Width           =   3495
      End
      Begin VB.Frame Frame6 
         Caption         =   "Student List"
         Height          =   4485
         Left            =   2550
         TabIndex        =   17
         Top             =   1230
         Width           =   7455
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4050
            Left            =   60
            TabIndex        =   7
            Top             =   180
            Width           =   7305
         End
      End
      Begin VB.Frame Frame5 
         Height          =   555
         Index           =   0
         Left            =   4440
         TabIndex        =   16
         Top             =   120
         Width           =   2595
         Begin VB.OptionButton Option5 
            Caption         =   "Leave"
            Height          =   195
            Left            =   1740
            TabIndex        =   31
            Top             =   210
            Width           =   795
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Absent"
            Height          =   195
            Left            =   900
            TabIndex        =   3
            Top             =   210
            Width           =   795
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Present "
            Height          =   225
            Left            =   30
            TabIndex        =   2
            Top             =   210
            Value           =   -1  'True
            Width           =   885
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmStudentAttendance.frx":1973C
         Left            =   4650
         List            =   "frmStudentAttendance.frx":19746
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   930
         Width           =   1845
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2580
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   930
         Width           =   2055
      End
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   2550
         TabIndex        =   13
         Top             =   120
         Width           =   1875
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   285
            Left            =   450
            TabIndex        =   26
            Top             =   180
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd-mmm-yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   195
            Left            =   30
            TabIndex        =   14
            Top             =   225
            Width           =   345
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Choose Option.."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   2415
         Begin VB.OptionButton Option1 
            Caption         =   "Student Specific"
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
            Height          =   345
            Index           =   1
            Left            =   630
            TabIndex        =   1
            Top             =   180
            Width           =   1755
         End
         Begin VB.OptionButton Option1 
            Caption         =   "All"
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
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   555
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Roll"
         Height          =   345
         Left            =   6480
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section "
         Height          =   195
         Index           =   1
         Left            =   2580
         TabIndex        =   15
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Index           =   0
         Left            =   4650
         TabIndex        =   11
         Top             =   720
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdAttend 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   375
      Left            =   5340
      TabIndex        =   8
      ToolTipText     =   "Click for Attendance"
      Top             =   6750
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   375
      Left            =   8490
      TabIndex        =   9
      ToolTipText     =   "Click to Close"
      Top             =   6750
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      Height          =   435
      Left            =   5310
      Top             =   6720
      Width           =   4245
   End
End
Attribute VB_Name = "frmStudentAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CboExamType_Click()
 load_exam_sub
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

Private Sub cmdAttend_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
Dim classId As String
Dim SectionID As String

If Len(Combo1.Text) = 0 Then
   MsgBox "Please select a Section ", vbInformation, cmp
   Combo1.SetFocus
   Exit Sub
End If

If Len(Combo2.Text) = 0 Then
   MsgBox "Please select a Shift ", vbInformation, cmp
   Combo2.SetFocus
   Exit Sub
End If

If Len(List1.Text) = 0 Then
    MsgBox "Please select a class from the student list ", vbInformation, cmp
    List1.SetFocus
   Exit Sub
  End If
  
If Len(cmdAcaYear.Text) = 0 Then
    MsgBox "Academic Year Required ... ", vbInformation, cmp
    cmdAcaYear.SetFocus
   Exit Sub
  End If
  
  If Len(CboExamType.Text) = 0 Then
    MsgBox "Exam Term Required ... ", vbInformation, cmp
    CboExamType.SetFocus
   Exit Sub
  End If

If Len(CboExamID.Text) = 0 Then
    MsgBox "Exam Type Required ... ", vbInformation, cmp
    CboExamID.SetFocus
   Exit Sub
  End If

If MaskEdBox3.Text = "__/__/__" Then
   MsgBox "Date required.....Please put a date ", vbInformation, cmp
   MaskEdBox3.SetFocus
   Exit Sub
End If

 If Check_ValidDate(MaskEdBox3) = False Then
    MaskEdBox3.SetFocus
    Exit Sub
 End If

If Option1(1).Value = True Then
  If Len(List2.Text) = 0 Then
    MsgBox "Please a student from the student list ", vbInformation, cmp
    List2.SetFocus
   Exit Sub
  End If
End If


If Val(Mid(MaskEdBox3.Text, 7, 8)) <> Val(Mid(cmdAcaYear, 3, 4)) Then
   MsgBox "Academic Year Conflicts....", vbInformation, cmp
    cmdAcaYear.SetFocus
   Exit Sub
End If
   

If Option1(0).Value = True Then
   Set rs = getdata("select ClassId from StudentAttendanceLeaveInfo where ClassID='" & Trim(Mid(List1, 1, 5)) & "' and SectionID ='" & Mid(Combo1, 1, 5) & "'  and attn_date ='" & Format(MaskEdBox3, "dd mmm yyyy") & "'")
  If Not rs.EOF Then
    MsgBox "Attendance of all student of " & Trim(Mid(List1, 6, 15)) & " has already been completed on date :" & Format(MaskEdBox3, "dd mmm yyyy") & " ", vbInformation, cmp
    Exit Sub
  End If
End If
If Option1(1).Value = True Then
   Set rs = getdata("select ClassId from StudentAttendanceLeaveInfo where ClassID='" & Trim(Mid(List1, 1, 5)) & "' and SectionID ='" & Mid(Combo1, 1, 5) & "'  and attn_date ='" & Format(MaskEdBox3, "dd mmm yyyy") & "' and StudentID='" & Mid(List2, 1, 10) & "'")
  If Not rs.EOF Then
     MsgBox "Attendance of Mr." & Mid(List2, 11, 80) & "  already been completed on date: " & MaskEdBox3.Text & " ", vbInformation, cmp
    Exit Sub
  End If
End If





            
        con.Open GConnString
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "StudentAttendance"
        If Option1(0).Value = True Then
           cmd(1) = "a"
        ElseIf Option1(1).Value = True Then
          cmd(1) = "s"
        End If
        cmd(2) = Mid(List2, 1, 10)
        cmd(3) = Mid(Combo2.Text, 1, 1)
        cmd(4) = Mid(List1, 1, 5)
        cmd(5) = Mid(Combo1, 1, 5)
        cmd(6) = Val(txtfields(3))
        If Option3.Value = True Then
           cmd(7) = "P"
           cmd(8) = "N"
           cmd(13) = "N"
        ElseIf Option4.Value = True Then
           cmd(7) = "A"
           cmd(8) = "Y"
           cmd(13) = "N"
           
        ElseIf Option5.Value = True Then
           cmd(7) = "P"
           cmd(8) = "N"
           cmd(13) = "Y"
        End If
        
       cmd(9) = Format(MaskEdBox3, "dd mmm yyyy")
       cmd(10) = Trim(cmdAcaYear)
       cmd(11) = Mid(Trim(CboExamType), 1, 2)
       cmd(12) = Mid(Trim(CboExamID), 1, 2)
        cmd.Execute
        
        Call LoadStuID
       
        For i = 3 To 3
        txtfields(i) = ""
        Next
        
        MsgBox "Save Successfully.", vbInformation, "Student Management System"
'Else
'        Exit Sub
'End If

End Sub

Private Sub cmdAttendenceReportofYear_Click()

If Len(Combo1.Text) = 0 Then
   MsgBox "Please select a Section ", vbInformation, cmp
   Combo1.SetFocus
   Exit Sub
End If

If Len(Combo2.Text) = 0 Then
   MsgBox "Please select a Shift ", vbInformation, cmp
   Combo2.SetFocus
   Exit Sub
End If

If Len(List1.Text) = 0 Then
    MsgBox "Please select a class from the student list ", vbInformation, cmp
    List1.SetFocus
   Exit Sub
  End If
  
If Len(cmdAcaYear.Text) = 0 Then
    MsgBox "Academic Year Required ... ", vbInformation, cmp
    cmdAcaYear.SetFocus
   Exit Sub
  End If
  
  
    
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub


Private Sub ComStuId_Click()
Label3.Caption = ""



txtfields(3) = ""
           
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT  distinct   StudentEvaluation.StudentID, StudentInfo.StudentName,  StudentEvaluation.Shift, " + _
        "StudentEvaluation.ClassId, ClassInfo.ClassName, StudentEvaluation.SectionId, SectionInfo.Sectiondsc, StudentEvaluation.ClassRoll, " + _
        "StudentEvaluation.ActiveClass FROM StudentEvaluation INNER JOIN SectionInfo ON StudentEvaluation.SectionId = SectionInfo.SectionID INNER JOIN " + _
        "ClassInfo ON StudentEvaluation.ClassId = ClassInfo.ClassID INNER JOIN StudentInfo ON StudentEvaluation.StudentID = StudentInfo.StudentID where StudentEvaluation.studentid='" & ComStuId & "'and StudentEvaluation.ActiveClass='Y'")
        
        If Not rs.EOF Then
            Label3.Caption = "" & rs!StudentName
         
           
'            txtfields(0) = "" & rs!Shift
            txtfields(3) = "" & rs!ClassRoll
           
        End If
End Sub

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

Private Sub cmdEdit_Click()
  Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
Dim classId As String
Dim SectionID As String

If Len(Combo1.Text) = 0 Then
   MsgBox "Please select a Section ", vbInformation, cmp
   Combo1.SetFocus
   Exit Sub
End If

If Len(Combo2.Text) = 0 Then
   MsgBox "Please select a Shift ", vbInformation, cmp
   Combo2.SetFocus
   Exit Sub
End If

If Len(List1.Text) = 0 Then
    MsgBox "Please a class from the student list ", vbInformation, cmp
    List1.SetFocus
   Exit Sub
  End If

If MaskEdBox3.Text = "__/__/__" Then
   MsgBox "Date required.....Please put a date ", vbInformation, cmp
   MaskEdBox3.SetFocus
   Exit Sub
End If

 If Check_ValidDate(MaskEdBox3) = False Then
    MaskEdBox3.SetFocus
    Exit Sub
 End If

If Option1(1).Value = True Then
  If Len(List2.Text) = 0 Then
    MsgBox "Please select a student from the student list ", vbInformation, cmp
    List2.SetFocus
   Exit Sub
  End If
End If
If Len(cmdAcaYear.Text) = 0 Then
    MsgBox "Academic Year Required ... ", vbInformation, cmp
    cmdAcaYear.SetFocus
   Exit Sub
  End If
  
  If Len(CboExamType.Text) = 0 Then
    MsgBox "Exam Term Required ... ", vbInformation, cmp
    CboExamType.SetFocus
   Exit Sub
  End If

If Len(CboExamID.Text) = 0 Then
    MsgBox "Exam Type Required ... ", vbInformation, cmp
    CboExamID.SetFocus
   Exit Sub
  End If
  
  If Val(Mid(MaskEdBox3.Text, 7, 8)) <> Val(Mid(cmdAcaYear, 3, 4)) Then
   MsgBox "Academic Year Conflicts....", vbInformation, cmp
    cmdAcaYear.SetFocus
   Exit Sub
End If


If Option1(0).Value = True Then
   If MsgBox("Are you sure to change the attendance status of all student of class  " & Trim(Mid(List1, 6, 15)) & "  on date : " & MaskEdBox3 & " ", vbYesNo + vbInformation, cmp) = vbYes Then
     Set rs = getdata("select classID from StudentAttendanceLeaveInfo where ClassID='" & Trim(Mid(List1, 1, 5)) & "' and SectionID ='" & Mid(Combo1, 1, 5) & "'  and attn_date ='" & Format(MaskEdBox3, "dd mmm yyyy") & "'")
      If Not rs.EOF Then
        con.Open GConnString
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "StudentAttendance"
        cmd(1) = "u"
        cmd(2) = Mid(List2, 1, 10)
        cmd(3) = Mid(Combo2.Text, 1, 1)
        cmd(4) = Mid(List1, 1, 5)
        cmd(5) = Mid(Combo1, 1, 5)
        cmd(6) = Val(txtfields(3))
      If Option3.Value = True Then
           cmd(7) = "P"
           cmd(8) = "N"
           cmd(13) = "N"
        ElseIf Option4.Value = True Then
           cmd(7) = "A"
           cmd(8) = "Y"
           cmd(13) = "N"
           
        ElseIf Option5.Value = True Then
           cmd(7) = "P"
           cmd(8) = "N"
           cmd(13) = "Y"
        End If
        cmd(9) = Format(MaskEdBox3, "dd mmm yyyy")
        cmd(10) = Trim(cmdAcaYear)
       cmd(11) = Mid(Trim(CboExamType), 1, 2)
       cmd(12) = Mid(Trim(CboExamID), 1, 2)
         
        cmd.Execute
        MsgBox "Edited Successfully...", vbInformation, "Student Management System"
  Else
       Exit Sub
 End If
End If
End If

  If Option1(1).Value = True Then
    If MsgBox("Are you sure to change the attendance status of Mr." & Mid(List2, 11, 80) & " of class  " & Trim(Mid(List1, 6, 15)) & "  on date : " & MaskEdBox3 & " ", vbYesNo + vbInformation, cmp) = vbYes Then
     Set rs = getdata("select StudentID from StudentAttendanceLeaveInfo where ClassID='" & Trim(Mid(List1, 1, 5)) & "' and SectionID ='" & Mid(Combo1, 1, 5) & "'  and attn_date ='" & Format(MaskEdBox3, "dd mmm yyyy") & "' and StudentID='" & Mid(List2, 1, 10) & "'")
    If Not rs.EOF Then
     con.Open GConnString
     cmd.ActiveConnection = con
     cmd.CommandType = adCmdStoredProc
     cmd.CommandText = "StudentAttendance"
     cmd(1) = "p"
     cmd(2) = Mid(List2, 1, 10)
     cmd(3) = Mid(Combo2.Text, 1, 1)
     cmd(4) = Mid(List1, 1, 5)
     cmd(5) = Mid(Combo1, 1, 5)
     cmd(6) = Val(txtfields(3))
     If Option3.Value = True Then
           cmd(7) = "P"
           cmd(8) = "N"
           cmd(13) = "N"
        ElseIf Option4.Value = True Then
           cmd(7) = "A"
           cmd(8) = "Y"
           cmd(13) = "N"
           
        ElseIf Option5.Value = True Then
           cmd(7) = "P"
           cmd(8) = "N"
           cmd(13) = "Y"
        End If
         cmd(9) = Format(MaskEdBox3, "dd mmm yyyy")
        cmd(10) = Trim(cmdAcaYear)
       cmd(11) = Mid(Trim(CboExamType), 1, 2)
       cmd(12) = Mid(Trim(CboExamID), 1, 2)
   

        cmd.Execute
       MsgBox "Edited Successfully...", vbInformation, "Student Management System"
    Else
       Exit Sub
   End If
 End If
End If
            
        Call LoadStuID
       
        For i = 3 To 3
        txtfields(i) = ""
        Next
        
        
End Sub

Private Sub cmdPrint_Click()
  If MaskEdBox3.Text = "__/__/__" Then
     MsgBox "Please put a valid Date", vbInformation, cmp
     MaskEdBox3.SetFocus
     Exit Sub
  End If
    rptMode = 5
    Screen.MousePointer = vbHourglass
    frmViewer.Show 1
End Sub

Private Sub Combo1_Click()
   LoadStuID
End Sub

Private Sub Command1_Click()
     
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys (Chr(9))
  End If
End Sub

Private Sub Form_Load()
Call LoadStuID
load_class
load_Aca_year
load_exam
End Sub
Private Sub load_Aca_year()
  Dim i As Integer
For i = 2000 To 2050
  cmdAcaYear.AddItem i
Next i
cmdAcaYear.Text = Format(Date, "YYYY")
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
Private Sub LoadStuID()
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT Distinct a.StudentID,(select StudentName from studentinfo s where s.StudentID=a.StudentID) FROM  Studentadmission a where a.classid='" & Mid(List1, 1, 5) & "' and sectionid='" & Mid(Combo1, 1, 5) & "' and a.approval='Y' and active_std=1 and a.admissionCancel='N'and a.serial_no=(select max(serial_no) from studentadmission where studentid=a.studentid and aca_yr='" & Trim(cmdAcaYear) & "') order by studentid")
List2.Clear
If Not rs.EOF Then
    Do Until rs.EOF
        List2.AddItem rs!studentid + "-" + rs(1)
        rs.MoveNext
    Loop
End If

End Sub

 Private Sub load_class()
   Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT ClassID, ClassName FROM  classinfo")
List1.Clear
If Not rs.EOF Then
    Do Until rs.EOF
        List1.AddItem rs(0) + "-" + rs(1)
        rs.MoveNext
    Loop
End If
 End Sub


Private Sub List1_Click()
    load_section
    LoadStuID
End Sub
Private Sub load_section()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("SELECT SectionID,Sectiondsc from sectioninfo WHERE ClassID='" & Mid(Trim(List1.Text), 1, 5) & "'")
            Combo1.Clear
            If Not rs.EOF Then
               rs.MoveFirst
               Do Until rs.EOF
                 Combo1.AddItem rs(0) + "  -  " + rs(1)
                 rs.MoveNext
              Loop
               
            End If
End Sub

Private Sub List2_Click()
 If Option1(1).Value = True And Len(List2.Text) > 0 Then
     get_roll
 End If
End Sub

Private Sub MaskEdBox3_GotFocus()
  MaskEdBox3.SelStart = 0
  MaskEdBox3.SelLength = Len(MaskEdBox3.Text)
  
End Sub

Private Sub MaskEdBox3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If MaskEdBox3 <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox3) = False Then
                MaskEdBox3.SetFocus
                Exit Sub
            End If
    End If

End If
End Sub

Private Sub get_roll()
Dim rs As New ADODB.Recordset
Set rs = getdata("select classRoll From StudentAdmission where StudentId='" & Mid(List2, 1, 10) & "'" & _
 " and serial_no=(select max(serial_no)  From StudentAdmission  where StudentId='" & Mid(List2, 1, 10) & "')")
txtfields(3).Text = rs(0)
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
          Case 0
             List2.Enabled = False
             txtfields(3).Text = ""
          Case 1
             List2.Enabled = True
  End Select
End Sub

