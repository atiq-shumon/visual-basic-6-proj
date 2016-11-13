VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmExamRoutine 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   0
      TabIndex        =   23
      Top             =   2190
      Width           =   8145
      Begin MSComCtl2.DTPicker DTPickerStartTime 
         Height          =   315
         Left            =   5790
         TabIndex        =   8
         ToolTipText     =   "Insert Exam starting time"
         Top             =   570
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59310082
         CurrentDate     =   38643
      End
      Begin VB.TextBox txtfields 
         Height          =   315
         Index           =   1
         Left            =   1110
         TabIndex        =   9
         ToolTipText     =   "Insert short note"
         Top             =   930
         Width           =   6885
      End
      Begin VB.TextBox txtfields 
         Height          =   315
         Index           =   0
         Left            =   1110
         TabIndex        =   6
         Text            =   "0"
         ToolTipText     =   "Insert Total Marks"
         Top             =   570
         Width           =   1215
      End
      Begin VB.ComboBox ComboExamCatagory 
         Height          =   315
         Left            =   5790
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Select Exam Category"
         Top             =   210
         Width           =   2235
      End
      Begin VB.ComboBox ComboSubject 
         Height          =   315
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Select Subject"
         Top             =   210
         Width           =   3015
      End
      Begin MSMask.MaskEdBox MaskEdBoxExamdate 
         Height          =   315
         Left            =   3420
         TabIndex        =   7
         ToolTipText     =   "Insert Exam adte"
         Top             =   570
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         Height          =   195
         Left            =   90
         TabIndex        =   29
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Marks"
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Time"
         Height          =   195
         Left            =   4860
         TabIndex        =   27
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Date"
         Height          =   195
         Left            =   2430
         TabIndex        =   26
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Catagopry"
         Height          =   195
         Left            =   4380
         TabIndex        =   25
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         Height          =   195
         Left            =   90
         TabIndex        =   24
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   8805
      TabIndex        =   19
      Top             =   0
      Width           =   8865
      Begin VB.Frame Frame2 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   20
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Routine Set Up"
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
         Left            =   2820
         TabIndex        =   30
         Top             =   180
         Width           =   2430
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   -60
         Picture         =   "FrmExamRoutine.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8745
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   0
      TabIndex        =   15
      Top             =   840
      Width           =   8745
      Begin VB.ComboBox ComboYear 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select Year"
         Top             =   180
         Width           =   2295
      End
      Begin VB.ComboBox ComboClass 
         Height          =   315
         Left            =   4770
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select Class"
         Top             =   180
         Width           =   2955
      End
      Begin VB.ComboBox ComboExamName 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Select Exam Name"
         Top             =   540
         Width           =   6615
      End
      Begin VB.CheckBox CheckMarksCategory 
         Caption         =   "Marks Category Applicable"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3930
         TabIndex        =   16
         Top             =   930
         Width           =   2505
      End
      Begin MSMask.MaskEdBox MaskEdStartDate 
         Height          =   315
         Left            =   1140
         TabIndex        =   13
         Top             =   900
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         Height          =   195
         Left            =   3930
         TabIndex        =   22
         Top             =   210
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   210
         Width           =   330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Name"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   990
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   405
      Left            =   4230
      TabIndex        =   0
      ToolTipText     =   "Click to insert new information"
      Top             =   6870
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   405
      Left            =   5220
      TabIndex        =   10
      ToolTipText     =   "Click to Save"
      Top             =   6870
      Width           =   945
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   405
      Left            =   6210
      TabIndex        =   11
      ToolTipText     =   "Click to Delete"
      Top             =   6870
      Width           =   945
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   405
      Left            =   7200
      TabIndex        =   12
      ToolTipText     =   "Click to Exit"
      Top             =   6870
      Width           =   945
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3225
      Left            =   0
      TabIndex        =   14
      Top             =   3570
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   5689
      _Version        =   393216
      FixedCols       =   0
   End
End
Attribute VB_Name = "FrmExamRoutine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdDelete_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
If MsgBox("Are You sure to Delete ?", vbYesNo + vbCritical) = vbYes Then
    cmd.CommandType = adCmdText
    cmd.CommandText = "Delete from ExamRoutine where Examid='" & Mid(ComboExamName, 1, InStr(1, ComboExamName, "-") - 1) & "'"
    cmd.Execute
    MsgBox "Delete successfully Exam Routine Information.", vbInformation, App.Title
    ShowFlexData1
    ComboYear.Text = ""
    ComboClass.Text = ""
    If ComboExamCatagory.Visible = True Then
        ComboExamCatagory.Text = ""
    End If
    ComboExamName.Text = ""
    If txtfields(0).Visible = True Then
        txtfields(0) = ""
    End If
    txtfields(1) = ""
    ComboSubject = ""
    If CheckMarksCategory.Value = 1 Then
        CheckMarksCategory.Value = 0
    End If
    DTPickerStartTime = "00:00:00"
    MaskEdBoxExamdate = "__/__/__"
    MaskEdStartDate = "__/__/__"
Else
    Exit Sub
End If
End Sub

Private Sub cmdnew_Click()
If ComboExamCatagory.Visible = True Then
    txtfields(0).Visible = True
    Label10.Visible = True
Else
    ComboExamCatagory.Visible = True
    Label7.Visible = True
End If
ComboYear.Text = ""
ComboClass.Text = ""
ComboExamCatagory.Text = ""
ComboExamName.Text = ""
txtfields(0) = ""
txtfields(1) = ""
ComboSubject = ""
If CheckMarksCategory.Value = 1 Then
    CheckMarksCategory.Value = 0
End If
DTPickerStartTime = "00:00:00"
MaskEdBoxExamdate = "__/__/__"
MaskEdStartDate = "__/__/__"
ComboYear.SetFocus
End Sub
Private Sub cmdSAVE_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
Dim rs As New ADODB.Recordset
cmd.ActiveConnection = con
    If Len(ComboClass) = 0 And Len(ComboYear) = 0 Then Exit Sub
    If Len(ComboYear) = 0 Then
        MsgBox "Please Enter Exam Year.", vbCritical, App.Title
        ComboYear.SetFocus
        Exit Sub
    End If
   
    If Len(ComboClass) = 0 Then
        MsgBox "Please Enter Class.", vbCritical, App.Title
        ComboClass.SetFocus
        Exit Sub
    End If
    
    If Len(ComboExamName) = 0 Then
        MsgBox "Please Enter Exam Name.", vbCritical, App.Title
        ComboExamName.SetFocus
        Exit Sub
    End If
    If DTPickerStartTime.Hour = 0 Then
        MsgBox "Please Enter Start Time.", vbCritical, App.Title
        DTPickerStartTime.SetFocus
        Exit Sub
    End If
    If MaskEdBoxExamdate = "__/__/__" Then
        MsgBox "Please Enter Exam Date.", vbCritical, App.Title
        MaskEdBoxExamdate.SetFocus
        Exit Sub
    End If
    Dim check As String
    If CheckMarksCategory.Value = 1 Then
        check = "Y"
    Else
        check = "N"
    End If
    If Len(ComboYear) <> 0 And Len(ComboExamName) <> 0 And Len(ComboClass) <> 0 Then
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ExaminationRoutine"
    
    cmd(1) = ComboYear
    cmd(2) = Mid(ComboClass, 1, 5)
    cmd(3) = Mid(ComboExamName, 1, InStr(1, ComboExamName, "-") - 1)
    cmd(4) = Format(MaskEdStartDate, "dd mm yyyy")
    cmd(5) = check
    cmd(6) = Mid(ComboSubject, 1, 5)
    
    If ComboExamCatagory.Visible = True Then
        cmd(7) = Mid(ComboExamCatagory, 1, 5)
    Else
        cmd(7) = ""
    End If
    cmd(8) = Val(txtfields(0).Text)
    
'    If txtfields(0).Visible = True Then
'        cmd(8) = txtfields(0)
'    Else
'        cmd(8) = 0
'    End If
    
    cmd(9) = Format(MaskEdBoxExamdate, "dd mm yyyy")
    
    cmd(10) = Format(DTPickerStartTime.Value, "hh:mm:ss")
      
    cmd(11) = txtfields(1)
    cmd(12) = soft_user
    cmd.Execute
    MsgBox "Saved Successfully.", vbInformation, "Student Management System"
    cmdnew.SetFocus
  
    Call ShowFlexData
    Else
    Exit Sub
    End If

End Sub

Private Sub ComboClass_Click()
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("SELECT     ExamSchedule.ExamID, ExamTypeInfo.ETypeName, ExamSchedule.ExamStartDate " + _
"FROM ExamSchedule INNER JOIN ExamTypeInfo ON ExamSchedule.ExamTypeID = ExamTypeInfo.ETypeID where ExamYear='" & ComboYear & "'and ClassId='" & Mid(ComboClass, 1, 5) & "'")
ComboExamName.Clear
If Not rs1.EOF Then
    Do Until rs1.EOF
        ComboExamName.AddItem rs1!ExamID & " - " & rs1!ETypeName & "-" & CStr(rs1!ExamStartDate)
        rs1.MoveNext
    Loop
End If
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT  Sub_code,Sub_title From Subject_Info_sub WHERE Class_code = '" & Mid(ComboClass, 1, 5) & "'")
ComboSubject.Clear
If Not rs.EOF Then
    Do Until rs.EOF
        ComboSubject.AddItem rs!Sub_code & "-" & rs!Sub_title
        rs.MoveNext
    Loop

End If
ShowFlexData
End Sub

Private Sub ComboClass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim rs As New ADODB.Recordset
    Set rs = getdata("SELECT Sub_title From Subject_Info_sub WHERE Class_code = '" & Mid(ComboClass, 1, 5) & "'")
    If Not rs.EOF Then
        Do Until rs.EOF
            ComboSubject.AddItem rs!Sub_title
            rs.MoveNext
        Loop
    End If
    ShowFlexData
    ComboExamName.SetFocus
End If

End Sub



Private Sub ComboExamCatagory_KeyPress(KeyAscii As Integer)
If ComboExamCatagory.Visible = True Then
    If KeyAscii = 13 Then
        MaskEdBoxExamdate.SetFocus
    End If
End If
End Sub

Private Sub ComboExamName_click()
MaskEdStartDate = "__/__/__"
'ComboSubject = ""
MaskEdBoxExamdate = "__/__/__"
DTPickerStartTime = "00:00:00"
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("SELECT  ExamStartDate,markscataapplied FROM ExamSchedule where ExamId='" & Mid(ComboExamName, 1, (InStr(1, ComboExamName, "-") - 1)) & "'")
If Not rs1.EOF Then
    MaskEdStartDate = Format(rs1!ExamStartDate, "dd/mm/yy")
    If rs1!MarksCataApplied = "Y" Then
        CheckMarksCategory.Value = 1
        ComboExamCatagory.Visible = True
        Label7.Visible = True
        Label10.Visible = False
        txtfields(0).Visible = False
    Else
        CheckMarksCategory.Value = 0
        ComboExamCatagory.Visible = False
        txtfields(0).Visible = True
        Label7.Visible = False
        Label10.Visible = True
    End If
End If
ShowFlexData
End Sub
Private Sub ComboExamName_KeyPress(KeyAscii As Integer)

If Len(ComboExamName) <> 0 Then
    If KeyAscii = 13 Then
        Dim rs1 As New ADODB.Recordset
        Set rs1 = getdata("SELECT ExamStartDate,markscataapplied FROM ExamSchedule where ExamId='" & Mid(ComboExamName, 1, (InStr(1, ComboExamName, "-") - 1)) & "'")
        If Not rs1.EOF Then
            MaskEdStartDate = Format(rs1!ExamStartDate, "dd/mm/yy")
            If rs1!MarksCataApplied = "Y" Then
                CheckMarksCategory.Value = 1
                ComboExamCatagory.Visible = True
                Label7.Visible = True
                txtfields(0).Visible = False
                Label10.Visible = False
            Else
                CheckMarksCategory.Value = 0
                ComboExamCatagory.Visible = False
                Label7.Visible = False
                txtfields(0).Visible = True
                Label10.Visible = True
            End If
        End If
        ComboSubject.SetFocus
    End If
Else
    Exit Sub
End If
End Sub

Private Sub ComboSubject_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ComboExamCatagory.Visible = True Then
            ComboExamCatagory.SetFocus
    Else
            txtfields(0).SetFocus
    End If
End If
End Sub
Private Sub ComboSubject_LostFocus()
If ComboExamCatagory.Visible = True Then
    Dim rs As New ADODB.Recordset
    Set rs = getdata("SELECT     SubjectMarksDistribution.CategoryID, MarksCategory.MCategoryDsc " + _
    "FROM SubjectMarksDistribution INNER JOIN MarksCategory ON SubjectMarksDistribution.CategoryID = MarksCategory.MCategoryID " + _
    "WHERE     (SubjectMarksDistribution.ClassID = '" & Mid(ComboClass, 1, 5) & "') AND (SubjectMarksDistribution.SubjectID = '" & Mid(ComboSubject, 1, 5) & "')")
    
    ComboExamCatagory.Clear
    If Not rs.EOF Then
        Do Until rs.EOF
          ComboExamCatagory.AddItem rs!CategoryID & "-" & rs!mcategoryDsc
          rs.MoveNext
        Loop
    End If
End If
End Sub

Private Sub ComboYear_Click()
Dim rs1 As New ADODB.Recordset

Set rs1 = getdata("SELECT     ExamSchedule.ExamID, ExamTypeInfo.ETypeName, ExamSchedule.ExamStartDate " + _
"FROM ExamSchedule INNER JOIN ExamTypeInfo ON ExamSchedule.ExamTypeID = ExamTypeInfo.ETypeID where ExamYear='" & ComboYear & "'and ClassId='" & Mid(ComboClass, 1, 5) & "'")
If Not rs1.EOF Then
    Do Until rs1.EOF
        ComboExamName.AddItem rs1!ExamID & " - " & rs1!ETypeName & "-" & CStr(rs1!ExamStartDate)
        rs1.MoveNext
    Loop
End If
'ShowFlexData
End Sub
Private Sub ComboYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboClass.SetFocus
    ShowFlexData
End If

End Sub
Private Sub DTPickerStartTime_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    txtfields(1).SetFocus
End If
End Sub
Private Sub Form_Load()

Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("Select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        ComboClass.AddItem rs1!classId + " - " + rs1!ClassName
        rs1.MoveNext
    Loop

End If

Dim IL As Integer
For IL = 2000 To 2020
   ComboYear.AddItem (IL)
Next IL

With MSFlexGrid1
    .Rows = 1
    .Cols = 6
    
    .Col = 0: .Text = "ExamId"
    .Col = 1: .Text = "  Subject Name #"
    .Col = 2: .Text = " Exam Category"
    .Col = 3: .Text = "Total Marks"
    .Col = 4: .Text = "Exam Date"
    .Col = 5: .Text = "  Start Time"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 2500
    .ColWidth(2) = 2200
    .ColWidth(3) = 1000
    .ColWidth(4) = 1200
    .ColWidth(5) = 1000
   
End With


Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MaskEdBoxExamdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MaskEdBoxExamdate <> "__/__/__" Then
            If Check_ValidDate(MaskEdBoxExamdate) = False Then
                MaskEdBoxExamdate.SetFocus
                Exit Sub
            End If
    End If

    DTPickerStartTime.SetFocus
End If
End Sub

Private Sub MaskEdStartDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MaskEdStartDate <> "__/__/__" Then
            If Check_ValidDate(MaskEdStartDate) = False Then
                MaskEdStartDate.SetFocus
                Exit Sub
            End If
    End If

    CheckMarksCategory.SetFocus
End If
End Sub
Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 0
            If txtfields(0).Visible = True Then
              MaskEdBoxExamdate.SetFocus
              
            End If
        Case 1
            cmdSave.SetFocus
    End Select
End If
End Sub

Private Sub ShowFlexData()

'On Error GoTo ErrDes
Dim rs As New ADODB.Recordset

Set rs = getdata("SELECT   ExamRoutine.ExamID, ExamRoutine.TotalMarks, ExamRoutine.ExamDate, ExamRoutine.ExamStartTime " + _
    " from ExamRoutine  where ExamRoutine.examyear='" & ComboYear & "'and ExamRoutine.classid='" & Mid(ComboClass, 1, 5) & "' and examRoutine.ExamID='" & Mid(ComboExamName, 1, 2) & "'")

If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1

                .TextMatrix(i, 0) = "" & rs!ExamID
'                .TextMatrix(i, 1) = "" & rs!SubjectDsc
'                .TextMatrix(i, 2) = IIf(IsNull(rs!mcategoryDsc) = True, "Not Applicable", rs!mcategoryDsc) '"" & rs!mcategoryDsc
                .TextMatrix(i, 3) = "" & rs!totalmarks
                .TextMatrix(i, 4) = "" & rs!ExamDate
                .TextMatrix(i, 5) = "" & rs!ExamStartTime

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
Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
Dim rs As New ADODB.Recordset

If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) <> 0 Then

    Set rs = getdata("SELECT    ExamTypeInfo.ETypeName, ExamRoutine.SubjectID, SubjectInfo.SubjectDsc, ExamRoutine.Startdate, ExamRoutine.Startdate, " + _
    " ExamRoutine.TotalMarks,ExamRoutine.Note FROM  ExamTypeInfo INNER JOIN " + _
    "ExamSchedule ON ExamTypeInfo.ETypeID = ExamSchedule.ExamTypeID INNER JOIN SubjectInfo ON ExamSchedule.ClassId = SubjectInfo.ClassID INNER JOIN " + _
    "ExamRoutine ON ExamSchedule.ExamID = ExamRoutine.ExamID AND SubjectInfo.SubjectID = ExamRoutine.SubjectID AND SubjectInfo.classId = ExamRoutine.classId " + _
    "WHERE     SubjectInfo.SubjectDsc = '" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) & "'  AND ExamRoutine.ExamID = '" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) & "'and ExamRoutine.ExamYear='" & ComboYear & "'")
    
    If Not rs.EOF Then
       
        ComboSubject = rs!SubjectID & "-" & rs!SubjectDsc
        ComboExamName = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) & "-" & rs!ETypeName & "-" & CStr(rs!Startdate)
        txtfields(1) = "" & rs!Note
        Label10.Visible = True
        txtfields(0).Visible = True
        txtfields(0) = rs!totalmarks
        CheckMarksCategory.Value = 0
        ComboExamCatagory.Visible = False
        Label7.Visible = False
        MaskEdStartDate = Format(rs!Startdate, "dd/mm/yy")
       
    End If
    MaskEdBoxExamdate = Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4), "dd/mm/yy")
    DTPickerStartTime.Value = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)

Else
    
    Set rs = getdata("SELECT     ExamTypeInfo.ETypeName, ExamRoutine.Startdate,ExamRoutine.SubjectID,SubjectInfo.SubjectDsc, ExamRoutine.CategoryID, MarksCategory.MCategoryDsc, ExamRoutine.ExamDate,ExamRoutine.Note " + _
    "FROM  ExamRoutine INNER JOIN MarksCategory ON ExamRoutine.CategoryID = MarksCategory.MCategoryID INNER JOIN " + _
    "ExamSchedule ON ExamRoutine.ExamID = ExamSchedule.ExamID AND ExamRoutine.ClassID = ExamSchedule.ClassId INNER JOIN " + _
    "ExamTypeInfo ON ExamSchedule.ExamTypeID = ExamTypeInfo.ETypeID INNER JOIN " + _
    "SubjectInfo ON ExamRoutine.SubjectID = SubjectInfo.SubjectID " + _
    "WHERE  SubjectInfo.SubjectDsc =  '" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) & "' AND ExamRoutine.ExamID = '" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) & "' and MarksCategory.MCategoryDsc ='" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)) & "' and ExamRoutine.ExamYear='" & ComboYear & "'")
    
    If Not rs.EOF Then
       
        ComboSubject = rs!SubjectID & "-" & rs!SubjectDsc
        ComboExamName = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) & "-" & rs!ETypeName & "-" & CStr(rs!Startdate)
        txtfields(1) = "" & rs!Note
        Label10.Visible = False
        txtfields(0).Visible = False
        CheckMarksCategory.Value = 1
        ComboExamCatagory.Visible = True
        Label7.Visible = True
        ComboExamCatagory = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2))
        MaskEdStartDate = Format(rs!Startdate, "dd/mm/yy")
     End If
     MaskEdBoxExamdate = Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4), "dd/mm/yy")
     DTPickerStartTime.Value = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
End If
    
errdes:

End Sub

Private Sub ShowFlexData1()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT     ExamRoutine.ExamID, SubjectInfo.SubjectDsc, ExamRoutine.TotalMarks, ExamRoutine.ExamDate, ExamRoutine.ExamStartTime, " + _
"MarksCategory.MCategoryDsc FROM  ExamRoutine INNER JOIN SubjectInfo ON ExamRoutine.SubjectID = SubjectInfo.SubjectID AND ExamRoutine.ClassID = SubjectInfo.ClassID LEFT OUTER JOIN " + _
"MarksCategory ON ExamRoutine.CategoryID = MarksCategory.MCategoryID where ExamRoutine.examyear='" & ComboYear & "'and ExamRoutine.classid='" & Mid(ComboClass, 1, 5) & "' and ExamRoutine.examId <> '" & Mid(ComboExamName, 1, InStr(1, ComboExamName, "-") - 1) & "'")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .TextMatrix(i, 0) = "" & rs!ExamID
                .TextMatrix(i, 1) = "" & rs!SubjectDsc
                .TextMatrix(i, 2) = "" & rs!mcategoryDsc
                .TextMatrix(i, 3) = "" & rs!totalmarks
                .TextMatrix(i, 4) = "" & rs!ExamDate
                .TextMatrix(i, 5) = "" & rs!ExamStartTime


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
