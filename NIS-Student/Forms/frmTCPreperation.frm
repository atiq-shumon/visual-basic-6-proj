VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmTCPreperation 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Student Information"
      ForeColor       =   &H00C00000&
      Height          =   1965
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Width           =   8595
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   28
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   18
         Top             =   570
         Width           =   7215
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   5100
         TabIndex        =   17
         Top             =   1200
         Width           =   3435
      End
      Begin VB.ComboBox ComStuId 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         ToolTipText     =   "Select Student"
         Top             =   210
         Width           =   2265
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   16
         Top             =   870
         Width           =   7215
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   15
         Top             =   1530
         Width           =   2835
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   5100
         TabIndex        =   14
         Top             =   1530
         Width           =   2055
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Roll"
         Height          =   195
         Left            =   4320
         TabIndex        =   27
         Top             =   1560
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Name"
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   570
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mother's Name"
         Height          =   195
         Left            =   60
         TabIndex        =   25
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID"
         Height          =   195
         Left            =   60
         TabIndex        =   24
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3600
         TabIndex        =   23
         Top             =   210
         Width           =   4935
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   60
         TabIndex        =   22
         Top             =   1590
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class "
         Height          =   195
         Left            =   60
         TabIndex        =   21
         Top             =   1260
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         Height          =   195
         Left            =   4320
         TabIndex        =   20
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4320
         TabIndex        =   19
         Top             =   1710
         Width           =   45
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "TC Information"
      ForeColor       =   &H00C00000&
      Height          =   1785
      Left            =   0
      TabIndex        =   9
      Top             =   2940
      Width           =   8595
      Begin VB.TextBox txtfields 
         Height          =   1035
         Index           =   6
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Insert Short Note"
         Top             =   600
         Width           =   7185
      End
      Begin VB.ComboBox ComboTypeOftc 
         Height          =   315
         ItemData        =   "frmTCPreperation.frx":0000
         Left            =   1320
         List            =   "frmTCPreperation.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select TC Type"
         Top             =   240
         Width           =   4755
      End
      Begin MSMask.MaskEdBox MaskEdDate 
         Height          =   285
         Left            =   7080
         TabIndex        =   2
         ToolTipText     =   "Insert Date of TC Issue"
         Top             =   270
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   6510
         TabIndex        =   12
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type of TC"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   600
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   975
      Left            =   -30
      ScaleHeight     =   915
      ScaleWidth      =   8565
      TabIndex        =   7
      Top             =   -30
      Width           =   8625
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   8
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer Certificate Prepeartion Information"
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
         Left            =   1470
         TabIndex        =   29
         Top             =   300
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   1020
         Left            =   -30
         Picture         =   "frmTCPreperation.frx":0004
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8595
      End
   End
   Begin VB.CommandButton cmdTc 
      BackColor       =   &H8000000C&
      Caption         =   "TC"
      Height          =   435
      Left            =   5700
      TabIndex        =   4
      ToolTipText     =   "Click to Prepare the TC"
      Top             =   4830
      Width           =   945
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   435
      Left            =   6690
      TabIndex        =   5
      ToolTipText     =   "Clickt o Cancel the TC"
      Top             =   4830
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   435
      Left            =   7680
      TabIndex        =   6
      ToolTipText     =   "Click to Exit"
      Top             =   4830
      Width           =   945
   End
End
Attribute VB_Name = "frmTCPreperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
Unload Me
End Sub


Private Sub cmdDelete_Click()
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
Dim rs As New adodb.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
    If MsgBox("Are You sure to Delete ?", vbYesNo + vbCritical) = vbYes Then
            cmd.CommandType = adCmdText
        
            cmd.CommandText = "Delete from TcInformation  where StudentId = '" & ComStuId & "'and Approved='P' "
            cmd.Execute
            MsgBox "Delete Successfully Tc Information for the student.", vbInformation, App.Title
            For i = 0 To 6
             txtfields(i) = ""
            Next
            Label7.Caption = ""
            ComStuId.Text = ""
            ComboTypeOftc.Text = " "
            MaskEdDate = "__/__/__"
    Else
            Exit Sub
    End If



End Sub

Private Sub cmdTc_Click()
Dim cmd As New adodb.Command
Dim con As New adodb.Connection
con.Open GConnString
cmd.ActiveConnection = con

    If Len(ComboTypeOftc) = 0 Then
        MsgBox "Please Enter TC Type Name.", vbCritical, App.Title
        ComboTypeOftc.SetFocus
        Exit Sub
    End If
   If MaskEdDate = "__/__/__" Then
        MsgBox "Please Enter Date", vbCritical, App.Title
        MaskEdDate.SetFocus
        Exit Sub
   End If
 
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "TCInfo"
    cmd(1) = ComStuId
    cmd(2) = Trim(txtfields(4))
    cmd(3) = Trim(Mid(txtfields(2), 1, 5))
    cmd(4) = Trim(Mid(txtfields(3), 1, 5))
    cmd(5) = Trim(txtfields(5))
    cmd(6) = Mid(ComboTypeOftc, 1, 5)
    cmd(7) = Format(MaskEdDate, "dd mm yy")
    cmd(8) = Trim(txtfields(6))
    cmd(9) = "DSL"
    cmd(10) = Date
    cmd(11) = "P"
   
    cmd.Execute
    MsgBox "Save Successfully.", vbInformation, "Student Management System"
   
  
    

End Sub

Private Sub ComboTypeOftc_Click()
Dim con As New adodb.Connection
con.Open GConnString
Dim SQL As String
SQL = "Select Note from TCTypeSetup Where TCID='" & Mid(Trim(ComboTypeOftc.Text), 1, 5) & "'"
Dim rs As adodb.Recordset
Set rs = con.Execute(SQL)
If Not (rs.EOF Or rs.BOF) Then
   txtfields(6).Text = rs!Note
End If
 
End Sub

Private Sub ComboTypeOftc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    MaskEdDate.SetFocus
End If
End Sub

Private Sub ComStuId_Click()
For i = 0 To 6
    txtfields(i) = ""
Next
ComboTypeOftc.Text = " "
MaskEdDate = "__/__/__"

 Dim rs As New adodb.Recordset
 Set rs = getdata("SELECT     TCInformation.StudentID, TCInformation.Shift, TCInformation.ClassID, ClassInfo.ClassName, TCInformation.SectionID, SectionInfo.Sectiondsc, " + _
"TCInformation.ClassRoll, TCInformation.TCTypeID, TCTypeSetUp.TcName, TCInformation.TCDate, TCInformation.TCNote, TCInformation.Approved,StudentInfo.StudentName , StudentInfo.StuFatherName, StudentInfo.StuMotherName " + _
"FROM  TCInformation INNER JOIN ClassInfo ON TCInformation.ClassID = ClassInfo.ClassID INNER JOIN " + _
"SectionInfo ON TCInformation.SectionID = SectionInfo.SectionID INNER JOIN " + _
 "TCTypeSetUp ON TCInformation.TCTypeID = TCTypeSetUp.TCID INNER JOIN " + _
"StudentInfo ON TCInformation.StudentID = StudentInfo.StudentID where TCInformation.studentid='" & ComStuId & "'and TCInformation.approved='P' ")
    If Not rs.EOF Then
                Label7.Caption = "" & rs!StudentName
                txtfields(0) = "" & rs!StuFatherName
                txtfields(1) = "" & rs!StuMotherName
                txtfields(2) = "" & rs!classId + "-" + rs!ClassName
                txtfields(3) = "" & rs!SectionID + "-" + rs!Sectiondsc
                txtfields(4) = "" & rs!Shift
                txtfields(5) = "" & rs!ClassRoll
                ComboTypeOftc = "" & rs!TCTypeID + "-" + rs!TcName
                MaskEdDate = Format(rs!TCDate, "dd/mm/yy")
                txtfields(6) = "" & rs!TCNote
                txtfields(6).SetFocus
                
    Else
        Set rs = getdata("SELECT     StudentEvaluation.StudentID, StudentInfo.StudentName, StudentInfo.StuFatherName, StudentInfo.StuMotherName, StudentEvaluation.Shift, " + _
        "StudentEvaluation.ClassId, ClassInfo.ClassName, StudentEvaluation.SectionId, SectionInfo.Sectiondsc, StudentEvaluation.ClassRoll, " + _
        "StudentEvaluation.ActiveClass FROM StudentEvaluation INNER JOIN SectionInfo ON StudentEvaluation.SectionId = SectionInfo.SectionID INNER JOIN " + _
        "ClassInfo ON StudentEvaluation.ClassId = ClassInfo.ClassID INNER JOIN StudentInfo ON StudentEvaluation.StudentID = StudentInfo.StudentID where StudentEvaluation.studentid='" & ComStuId & "'and StudentEvaluation.ActiveClass='Y' ")
        If Not rs.EOF Then
            Label7.Caption = "" & rs!StudentName
            txtfields(0) = "" & rs!StuFatherName
            txtfields(1) = "" & rs!StuMotherName
            txtfields(2) = "" & rs!classId + "-" + rs!ClassName
            txtfields(3) = "" & rs!SectionID + "-" + rs!Sectiondsc
            txtfields(4) = "" & rs!Shift
            txtfields(5) = "" & rs!ClassRoll
           ComboTypeOftc.SetFocus
        End If
       
    End If
   

End Sub

Private Sub ComStuId_KeyPress(KeyAscii As Integer)


    Dim rs As New adodb.Recordset
    Set rs = getdata("SELECT     TCInformation.StudentID, TCInformation.Shift, TCInformation.ClassID, ClassInfo.ClassName, TCInformation.SectionID, SectionInfo.Sectiondsc, " + _
    "TCInformation.ClassRoll, TCInformation.TCTypeID, TCTypeSetUp.TcName, TCInformation.TCDate, TCInformation.TCNote, TCInformation.Approved,StudentInfo.StudentName , StudentInfo.StuFatherName, StudentInfo.StuMotherName " + _
    "FROM  TCInformation INNER JOIN ClassInfo ON TCInformation.ClassID = ClassInfo.ClassID INNER JOIN " + _
    "SectionInfo ON TCInformation.SectionID = SectionInfo.SectionID INNER JOIN " + _
     "TCTypeSetUp ON TCInformation.TCTypeID = TCTypeSetUp.TCID INNER JOIN " + _
    "StudentInfo ON TCInformation.StudentID = StudentInfo.StudentID where TCInformation.studentid='" & ComStuId & "'and TCInformation.approved='P' ")
    If Not rs.EOF Then
                Label7.Caption = "" & rs!StudentName
                txtfields(0) = "" & rs!StuFatherName
                txtfields(1) = "" & rs!StuMotherName
                txtfields(2) = "" & rs!classId + "-" + rs!ClassName
                txtfields(3) = "" & rs!SectionID + "-" + rs!Sectiondsc
                txtfields(4) = "" & rs!Shift
                txtfields(5) = "" & rs!ClassRoll
                ComboTypeOftc = "" & rs!TCTypeID + "-" + rs!TcName
                MaskEdDate = Format(rs!TCDate, "dd/mm/yy")
                txtfields(6) = "" & rs!TCNote
                txtfields(6).SetFocus
    Else
                Set rs = getdata("SELECT     StudentEvaluation.StudentID, StudentInfo.StudentName, StudentInfo.StuFatherName, StudentInfo.StuMotherName, StudentEvaluation.Shift, " + _
                "StudentEvaluation.ClassId , ClassInfo.ClassName, StudentEvaluation.SectionId, SectionInfo.Sectiondsc, StudentEvaluation.ClassRoll " + _
                "FROM   StudentEvaluation INNER JOIN ClassInfo ON StudentEvaluation.ClassId = ClassInfo.ClassID INNER JOIN " + _
                "SectionInfo ON StudentEvaluation.SectionId = SectionInfo.SectionID INNER JOIN " + _
                "StudentInfo ON StudentEvaluation.StudentID = StudentInfo.StudentID where StudentEvaluation.studentid='" & ComStuId & "'and StudentEvaluation.ActiveClass='Y' ")
                If Not rs.EOF Then
                    Label7.Caption = "" & rs!StudentName
                    txtfields(0) = "" & rs!StuFatherName
                    txtfields(1) = "" & rs!StuMotherName
                    txtfields(2) = "" & rs!classId + "-" + rs!ClassName
                    txtfields(3) = "" & rs!SectionID + "-" + rs!Sectiondsc
                    txtfields(4) = "" & rs!Shift
                    txtfields(5) = "" & rs!ClassRoll
                  ComboTypeOftc.SetFocus
                End If
        
    End If
 

End Sub

Private Sub Form_Load()
Dim rs As New adodb.Recordset
Set rs = getdata("SELECT     Distinct StudentID FROM  StudentEvaluation Where Active='Y'")
If Not rs.EOF Then
    Do Until rs.EOF
        ComStuId.AddItem rs!StudentID
        rs.MoveNext
    Loop

End If
Set rs = getdata("SELECT TCID,TCName From TCTypeSetUp ")
If Not rs.EOF Then
    Do Until rs.EOF
        ComboTypeOftc.AddItem rs!TCID + "-" + rs!TcName
        rs.MoveNext
    Loop
    ComboTypeOftc.AddItem (" ")
End If
End Sub
Private Sub MaskEdDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MaskEdDate <> "__/__/__" Then
            If Check_ValidDate(MaskEdDate) = False Then
                MaskEdDate.SetFocus
                Exit Sub
            End If
    End If
txtfields(6).SetFocus
End If
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 6
            cmdTc.SetFocus
    End Select
End If
End Sub


 
