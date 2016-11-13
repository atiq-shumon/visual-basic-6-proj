VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmstudentadmission 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4305
      Left            =   0
      TabIndex        =   18
      Top             =   2400
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   7594
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   12582912
      BackColorSel    =   12640511
      ForeColorSel    =   16711680
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdAdmitted 
      BackColor       =   &H8000000C&
      Caption         =   "Admitted"
      Height          =   435
      Left            =   7260
      TabIndex        =   5
      ToolTipText     =   "Click To Admitted"
      Top             =   6780
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000C&
      Caption         =   "Cancel"
      Height          =   435
      Left            =   8250
      TabIndex        =   6
      ToolTipText     =   "Click to cancel Admission"
      Top             =   6780
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   435
      Left            =   9240
      TabIndex        =   7
      ToolTipText     =   "Click to Exit"
      Top             =   6780
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Height          =   1605
      Left            =   0
      TabIndex        =   9
      Top             =   780
      Width           =   10275
      Begin VB.ComboBox cmdAcaYear 
         Height          =   315
         Left            =   7740
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   660
         Width           =   1905
      End
      Begin VB.TextBox txtfields 
         Height          =   315
         Left            =   4260
         TabIndex        =   17
         ToolTipText     =   "Insert Roll"
         Top             =   1170
         Width           =   2235
      End
      Begin VB.ComboBox ComboSection 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Select Section"
         Top             =   1170
         Width           =   2595
      End
      Begin VB.ComboBox ComboClass 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select Class"
         Top             =   660
         Width           =   2565
      End
      Begin VB.ComboBox ComboShift 
         Height          =   315
         ItemData        =   "frmstudentadmission.frx":0000
         Left            =   4260
         List            =   "frmstudentadmission.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select Shift"
         Top             =   660
         Width           =   2235
      End
      Begin MSMask.MaskEdBox MaskEdBoxDate 
         Height          =   315
         Left            =   7740
         TabIndex        =   4
         ToolTipText     =   "Insert  Admission Date"
         Top             =   1170
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox ComStuId 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         ToolTipText     =   "Select student"
         Top             =   180
         Width           =   2535
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A.Year"
         Height          =   195
         Left            =   7200
         TabIndex        =   21
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Roll"
         Height          =   195
         Left            =   3810
         TabIndex        =   16
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class "
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   3810
         TabIndex        =   13
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   7200
         TabIndex        =   12
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3480
         TabIndex        =   11
         Top             =   180
         Width           =   6135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   210
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   825
      Left            =   -30
      ScaleHeight     =   765
      ScaleWidth      =   10185
      TabIndex        =   8
      Top             =   -30
      Width           =   10245
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Admission Information"
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
         Left            =   3480
         TabIndex        =   19
         Top             =   150
         Width           =   3600
      End
      Begin VB.Image Image1 
         Height          =   930
         Left            =   0
         Picture         =   "frmstudentadmission.frx":0035
         Stretch         =   -1  'True
         Top             =   -60
         Width           =   10185
      End
   End
   Begin MSMask.MaskEdBox MEEffectiveDatePay 
      Height          =   315
      Left            =   2910
      TabIndex        =   23
      ToolTipText     =   "Insert  Effective Date For Paying Fee"
      Top             =   6810
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd-mmm-yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Effective Date For Paying Fee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   6840
      Width           =   2580
   End
End
Attribute VB_Name = "frmstudentadmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAcaYear_Click()
  load_flex
End Sub

Private Sub cmdAcaYear_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtfields.SetFocus
  End If
End Sub

Private Sub cmdAcaYear_LostFocus()
  get_roll
End Sub

Private Sub cmdAdmitted_Click()
If Len(ComStuId) = 0 Then
  ComStuId.SetFocus
  Exit Sub
End If
If Len(ComStuId) = 0 And Len(ComboClass) = 0 Then Exit Sub
If Len(ComStuId) = 0 Then
    MsgBox "Please Enter Student ID. ", vbInformation, cmp
    ComStuId.SetFocus
    Exit Sub
End If
If Len(ComboShift.Text) = 0 Then
    MsgBox "Select Shift Name.", vbInformation, cmp
    ComboShift.SetFocus
    Exit Sub
End If
If Len(ComboClass.Text) = 0 Then
    MsgBox "Select Class .", vbInformation, cmp
    ComboClass.SetFocus
    Exit Sub
End If
If Len(ComboSection.Text) = 0 Then
    MsgBox "Select Section.", vbInformation, cmp
    ComboSection.SetFocus
    Exit Sub
End If

If Len(cmdAcaYear.Text) = 0 Then
    MsgBox "Academic Year Required...", vbInformation, cmp
    cmdAcaYear.SetFocus
    Exit Sub
End If

If MaskEdBoxDate = "__/__/__" Then
    MsgBox "Enter Date.", vbInformation, cmp
    MaskEdBoxDate.SetFocus
    Exit Sub
End If

If MEEffectiveDatePay = "__/__/__" Then
    MsgBox "Enter Date Effective Date For Paying Fee.", vbInformation, cmp
    MEEffectiveDatePay.SetFocus
    Exit Sub
End If

If Len(ComboSection) <> 0 And Len(ComboShift) <> 0 And Len(ComboClass) <> 0 And Len(cmdAcaYear) <> 0 Then
 Dim rs2 As New ADODB.Recordset
     Set rs2 = getdata("select ClassRoll from StudentAdmission where classid='" & Mid(ComboClass, 1, 5) & "' and shift='" & Mid(ComboShift, 1, 1) & "'and sectionId='" & Mid(ComboSection, 1, 5) & "' and aca_yr='" & Trim(cmdAcaYear) & "' and ClassRoll='" & Trim(txtfields.Text) & "'")
      If Not rs2.EOF Then
        MsgBox "Same Roll No is occupied by another Student...Please Chose a another", vbInformation, cmp
        txtfields.SetFocus
        Exit Sub
   End If
       
 End If
 
 If Len(ComboSection) <> 0 And Len(ComboShift) <> 0 And Len(ComboClass) <> 0 And Len(cmdAcaYear) <> 0 Then
  Dim rs3 As New ADODB.Recordset
     Set rs3 = getdata("select ClassRoll from StudentAdmission where  studentId='" & Trim(ComStuId) & "'")
     If Not rs3.EOF Then
        MsgBox "This Student is already admitted...Please Choose Re-Admission", vbInformation, cmp
        ComStuId.SetFocus
        Exit Sub
  End If
End If

Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set rs = getdata("select StudentId from StudentAdmission where studentId='" & ComStuId & "' and classid='" & Mid(ComboClass, 1, 5) & "' and shift='" & Mid(ComboShift, 1, 1) & "'and sectionId='" & Mid(ComboSection, 1, 5) & "' and aca_yr='" & Trim(cmdAcaYear) & "' and ClassRoll='" & Trim(txtfields.Text) & "'")
If Not rs.EOF Then
    If MsgBox("Information Inserted Previously ,Do you want to update the Admission Information this student? ", vbYesNo + vbInformation) = vbYes Then
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "StuAdmissionEvaluationInformation"
            cmd(1) = 1
            cmd(2) = Trim(ComStuId.Text)
            cmd(3) = Format(MaskEdBoxDate, "dd mmm yyyy")
            cmd(4) = Mid(ComboShift, 1, 1)
            cmd(5) = Mid(ComboClass, 1, 5)
            cmd(6) = Mid(ComboSection, 1, 5)
            cmd(7) = txtfields
            cmd(8) = soft_user
            cmd(9) = Date
            cmd(10) = "Y"
            cmd(11) = "N"
            cmd(12) = "Y"
            cmd(13) = "Y"
            cmd(14) = Trim(cmdAcaYear)
            cmd(15) = "A"
            cmd(16) = Format(MEEffectiveDatePay, "dd mmm yyyy")
            cmd.Execute
            MsgBox "Update Successfully.", vbInformation, "Student Management System"
            ShowFlexData (4)
'            cmdAdmitted.Enabled = False
    Else
            Exit Sub
    End If
Else
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "StuAdmissionEvaluationInformation"
    cmd(1) = 1
    cmd(2) = ComStuId.Text
    cmd(3) = MaskEdBoxDate
    cmd(4) = Mid(ComboShift, 1, 1)
    cmd(5) = Mid(ComboClass, 1, 5)
    cmd(6) = Mid(ComboSection, 1, 5)
    cmd(7) = txtfields
    cmd(8) = soft_user
    cmd(9) = Date
    cmd(10) = "N"
    cmd(11) = "N"
    cmd(12) = "N"
    cmd(13) = "Y"
    cmd(14) = Trim(cmdAcaYear)
    cmd(15) = "A"
    cmd(16) = Format(MEEffectiveDatePay, "dd mmm yyyy")
    cmd.Execute
    MsgBox "Saved Successfully.", vbInformation, "Student Management System"
    ShowFlexData (4)
'    cmdAdmitted.Enabled = False
End If

End Sub

Private Sub cmdCancel_Click()
If Len(ComStuId) = 0 And Len(ComboClass) = 0 Then Exit Sub
If Len(ComStuId) = 0 Then
    MsgBox "Please Enter Student ID. ", vbInformation, cmp
    ComStuId.SetFocus
    Exit Sub
End If
If Len(ComboShift.Text) = 0 Then
    MsgBox "Select Shift Name.", vbInformation, cmp
    ComboShift.SetFocus
    Exit Sub
End If
If Len(ComboClass.Text) = 0 Then
    MsgBox "Select Class .", vbInformation, cmp
    ComboClass.SetFocus
    Exit Sub
End If
If Len(ComboSection.Text) = 0 Then
    MsgBox "Select Section.", vbInformation, cmp
    ComboSection.SetFocus
    Exit Sub
End If

If MaskEdBoxDate = "__/__/__" Then
    MsgBox "Enter Date.", vbInformation, cmp
    MaskEdBoxDate.SetFocus
    Exit Sub
End If

If MsgBox("Are You sure to cancel Admission Information for this Student? ", vbYesNo + vbInformation) = vbYes Then
    Dim con As New ADODB.connection
    con.Open GConnString
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "StudentEvaluation1"
    cmd(1) = ComStuId
    cmd(2) = MaskEdBoxDate
    cmd(3) = "N"
    cmd(4) = "Y"
    cmd(5) = "Y"
    cmd(6) = "DSL"
    cmd(7) = Date
    cmd(8) = Mid(ComboClass.Text, 1, 5)
    cmd.Execute
    MsgBox "Cancellation is completed Successfully.", vbInformation, "Student Management System"
    ShowFlexData (4)
    Label3.Caption = ""
'    ComboClass = ""
'    ComboSection = ""
'    ComboShift = ""
    txtfields = ""
    MaskEdBoxDate = "__/__/__"
 Else
     Exit Sub
 End If
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub ComboClass_Click()
ComboSection.Clear
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("select SectionId,Sectiondsc from SectionInfo where ClassId='" & Mid(ComboClass, 1, 5) & "' and trackid=(select max(trackid) from sectionInfo where ClassId='" & Mid(ComboClass, 1, 5) & "')")
If Not rs1.EOF Then
    Do Until rs1.EOF
        ComboSection.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
End If
load_flex
'ComboSection.SetFocus
End Sub
Private Sub load_flex()
 If Len(cmdAcaYear) <> 0 And Len(ComboClass) = 0 And Len(ComboShift) = 0 And Len(ComboSection) = 0 Then
   ShowFlexData (1)
ElseIf Len(cmdAcaYear) <> 0 And Len(ComboClass) <> 0 And Len(ComboShift) = 0 And Len(ComboSection) = 0 Then
   ShowFlexData (2)
ElseIf Len(cmdAcaYear) <> 0 And Len(ComboClass) <> 0 And Len(ComboShift) = 0 And Len(ComboSection) <> 0 Then
   ShowFlexData (3)
ElseIf Len(cmdAcaYear) <> 0 And Len(ComboClass) <> 0 And Len(ComboShift) <> 0 And Len(ComboSection) <> 0 Then
   ShowFlexData (4)
Else
  ShowFlexData (3)
End If
  
End Sub
Private Sub ComboClass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboSection.SetFocus
End If
End Sub

Private Sub ComboClass_LostFocus()
  get_roll
End Sub

Private Sub ComboSection_Click()
  load_flex
End Sub

Private Sub ComboSection_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    ComboShift.SetFocus
  End If
End Sub

Private Sub ComboSection_LostFocus()
   get_roll
End Sub
Private Sub get_roll()
  Dim rs As New ADODB.Recordset
 If Len(ComboSection) <> 0 And Len(ComboShift) <> 0 And Len(ComboClass) And Len(cmdAcaYear) <> 0 Then
    Set rs = getdata("select max(ClassRoll)+ 1 from Studentadmission where classid='" & Mid(ComboClass, 1, 5) & "' and shift='" & Mid(ComboShift, 1, 1) & "' and sectionId='" & Mid(ComboSection, 1, 5) & "' and aca_yr='" & Trim(cmdAcaYear) & "'")
    If Not rs.EOF Then
        txtfields = IIf(IsNull(rs(0)) = True, "1", rs(0))
    Else
        txtfields = "1"
    End If
 End If

End Sub

Private Sub ComboShift_Click()
 load_flex
End Sub

Private Sub ComboShift_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     cmdAcaYear.SetFocus
  End If
End Sub

Private Sub ComboShift_LostFocus()
   get_roll
End Sub

'Private Sub ComStuId_click()
''ComboShift.SetFocus
''Label3.Caption = ""
'''ComboClass.Text = ""
''''ComboSection = ""
'''ComboShift = ""
''txtfields = ""
''MaskEdBoxDate = "__/__/__"
''cmdAdmitted.Enabled = True
'''load_roll
'End Sub

Private Sub ComStuId_GotFocus()
  load_roll
End Sub

Private Sub ComStuId_Click()

Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT StudentInfo.StudentName, StudentAdmission.AdmissionDate, StudentAdmission.Shift, StudentAdmission.ClassId, StudentAdmission.SectionId," + _
            "StudentAdmission.ClassRoll , ClassInfo.ClassName, SectionInfo.Sectiondsc,StudentAdmission.AdmissionDate " + _
            "FROM StudentAdmission INNER JOIN StudentInfo ON StudentAdmission.StudentId = StudentInfo.StudentID INNER JOIN " + _
            "ClassInfo ON StudentAdmission.ClassId = ClassInfo.ClassID INNER JOIN " + _
            "SectionInfo ON StudentAdmission.SectionId = SectionInfo.SectionID WHERE StudentAdmission.StudentID = '" & ComStuId.Text & "' and AdmissionCancel='Y'")
If Not rs.EOF Then
    Label3.Caption = rs!StudentName
    If rs!Shift = "M" Then
       ComboShift.ListIndex = 0
    ElseIf rs!Shift = "D" Then
       ComboShift.ListIndex = 1
    End If
    
    
'    ComboClass = Trim(rs!classId) & "-" & Trim(rs!ClassName)
'    ComboSection = rs!SectionID + "-" + rs!Sectiondsc
'    txtfields = rs!ClassRoll
    MaskEdBoxDate = Format(rs!AdmissionDate, "dd/mm/yy")
Else
    
         Set rs = getdata("select StudentName from StudentInfo where StudentID = '" & ComStuId.Text & "'")
         If Not rs.EOF Then
             Label3.Caption = rs!StudentName
    
    Else
        If Len(ComboClass.Text) <> 0 And Len(ComboSection.Text) <> 0 And Len(ComboShift) <> 0 Then
            Set rs = getdata("select max(ClassRoll)+ 1 from StudentEvaluation where classid='" & Mid(ComboClass, 1, 5) & "' and shift='" & ComboShift & "'and sectionId='" & Mid(ComboSection, 1, 5) & "'")
            If Not rs.EOF Then
                txtfields = IIf(IsNull(rs(0)) = True, "1", rs(0))
            Else
                txtfields = "1"
            End If
            
        End If
    End If

End If

End Sub

Private Sub Form_Load()
With MSFlexGrid1
    .Rows = 1
    .Cols = 7
    .Col = 0: .Text = " Student ID   #"
    .Col = 1: .Text = "Student Name   "
    .Col = 2: .Text = " Class  "
    .Col = 3: .Text = " Section  "
    .Col = 4: .Text = " Shift  "
    .Col = 5: .Text = " Roll No  "
    .Col = 6: .Text = " Admission Date  "
    .ColWidth(0) = 1100
    .ColWidth(1) = 3000
    .ColWidth(2) = 2000
    .ColWidth(3) = 3000
    .ColWidth(4) = 2000
    .ColWidth(5) = 2000
    .ColWidth(6) = 2000
End With
load_Aca_year
ShowFlexData (1)
load_roll

End Sub
Private Sub load_Aca_year()
  Dim i As Integer
  For i = 2000 To 2050
     cmdAcaYear.AddItem i
  Next i
  cmdAcaYear.Text = Format(Date, "YYYY")
End Sub
Private Sub load_roll()
 Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT StudentID From StudentInfo " + _
"WHERE (StudentID NOT IN(SELECT StudentID FROM StudentAdmission))")
ComStuId.Clear
If Not rs.EOF Then
    Do Until rs.EOF
        ComStuId.AddItem rs(0)
        rs.MoveNext
    Loop
End If
Dim rs1 As New ADODB.Recordset
ComboClass.Clear
Set rs1 = getdata("select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
            ComboClass.AddItem rs1(0) + " - " + rs1(1)
            rs1.MoveNext
    Loop
End If
End Sub
Private Sub MaskEdBoxDate_GotFocus()
  MaskEdBoxDate.SelStart = 0
  MaskEdBoxDate.SelLength = Len(MaskEdBoxDate.Text)
End Sub

Private Sub MaskEdBoxDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MaskEdBoxDate <> "__/__/__" Then
            If Check_ValidDate(MaskEdBoxDate) = False Then
                MaskEdBoxDate.SetFocus
                Exit Sub
            End If
    End If
    MEEffectiveDatePay.SetFocus
    
End If
End Sub
Private Sub ComStuId_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboClass.SetFocus
End If
End Sub
Private Sub ShowFlexData(mode As Integer)
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Dim rsName As New ADODB.Recordset

If mode = 1 Then
'
'Set rs = getdata("select s.Shift,i.StudentName,s.ClassId,c.ClassName,s.SectionId,s.ClassRoll,s.aca_yr from " + _
'   " studentAdmission s,studentInfo i,classInfo c where s.StudentId='" & Trim(ComStuId.Text) & "' and s.ClassId=c.ClassId and s.StudentId=i.StudentId and s.serial_no=(select max(serial_no) from studentAdmission where StudentId='" & Trim(ComStuId.Text) & "')")
'


     Set rs = getdata("select a.studentid,a.shift,a.classid,a.sectionid,a.classroll,a.admissiondate " + _
               " from StudentAdmission a where a.approval='Y' and a.aca_yr='" & cmdAcaYear & "' and Student_status='A'")
ElseIf mode = 2 Then
     Set rs = getdata("select studentid,shift,classid,sectionid,classroll,admissiondate " + _
               " from StudentAdmission where approval='Y' and aca_yr='" & cmdAcaYear & "' and ClassId='" & Mid(ComboClass, 1, 5) & "' and Student_status='A'")
ElseIf mode = 3 Then
     Set rs = getdata("select studentid,shift,classid,sectionid,classroll,admissiondate " + _
               " from StudentAdmission where approval='Y' and aca_yr='" & cmdAcaYear & "' and ClassId='" & Mid(ComboClass, 1, 5) & "' and SectionId='" & Mid(ComboSection, 1, 5) & "' and Student_status='A'")

ElseIf mode = 4 Then
     Set rs = getdata("select studentid,shift,classid,sectionid,classroll,admissiondate " + _
               " from StudentAdmission where approval='Y' and aca_yr='" & cmdAcaYear & "' and ClassId='" & Mid(ComboClass, 1, 5) & "' and SectionId='" & Mid(ComboSection, 1, 5) & "' and shift='" & Mid(ComboShift, 1, 1) & "' and Student_status='A'")
Else
  Set rs = Nothing
End If
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .TextMatrix(i, 0) = "" & rs!studentid
                 Set rsName = getdata("select studentName from studentinfo where studentid='" & Trim(rs!studentid) & "'")
                .TextMatrix(i, 1) = rsName(0)
'                .TextMatrix(i, 2) = "" & rs!classId + "-" + rs!ClassName
'                .TextMatrix(i, 3) = "" & rs!SectionID + "-" + rs!Sectiondsc
'                .TextMatrix(i, 4) = "" & rs!Shift
'                .TextMatrix(i, 5) = "" & rs!ClassRoll
'                .TextMatrix(i, 6) = "" & rs!AdmissionDate
           rs.MoveNext
           i = i + 1
        Loop
    End With
 Else
     MSFlexGrid1.Rows = 1

 End If

Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub MEEffectiveDatePay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdAdmitted.SetFocus
End If
End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
If MSFlexGrid1.Rows > 1 Then
    ComStuId = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    Label3.Caption = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
    ComboClass = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
    ComboSection = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
    ComboShift = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
    txtfields = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
    MaskEdBoxDate = Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6), "dd/mm/yy")
Else
    Exit Sub
End If
Exit Sub
errdes:
'MsgBox err.Description, vbInformation, App.Title

End Sub

Private Sub txtfields_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     MaskEdBoxDate.SetFocus
  End If
  
End Sub
