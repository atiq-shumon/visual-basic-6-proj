VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStudent_performance 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdEdit 
      BackColor       =   &H8000000C&
      Caption         =   "Edit"
      Height          =   375
      Left            =   7860
      TabIndex        =   7
      ToolTipText     =   "Click to save"
      Top             =   8040
      Width           =   945
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5085
      Left            =   1470
      TabIndex        =   14
      Top             =   2790
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8969
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorSel    =   -2147483624
      ForeColorSel    =   12582912
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   375
      Left            =   5940
      TabIndex        =   10
      ToolTipText     =   "Click to insert new information"
      Top             =   8040
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   375
      Left            =   6900
      TabIndex        =   6
      ToolTipText     =   "Click to save"
      Top             =   8040
      Width           =   945
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   375
      Left            =   8820
      TabIndex        =   8
      ToolTipText     =   "Click to Delete"
      Top             =   8040
      Width           =   945
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   375
      Left            =   10740
      TabIndex        =   9
      ToolTipText     =   "Click to Exit"
      Top             =   8040
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   765
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12075
      TabIndex        =   12
      Top             =   0
      Width           =   12135
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Students'  Performance  Status"
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
         Left            =   4410
         TabIndex        =   31
         Top             =   210
         Width           =   3540
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   -60
         Picture         =   "frmStudent_performance.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   11985
      End
   End
   Begin VB.Frame Frame1 
      Height          =   10005
      Left            =   0
      TabIndex        =   11
      Top             =   660
      Width           =   12045
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000C&
         Caption         =   "Print"
         Height          =   375
         Left            =   9780
         TabIndex        =   32
         ToolTipText     =   "Click to save"
         Top             =   7380
         Width           =   945
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3360
         Style           =   1  'Simple Combo
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   300
         Width           =   1695
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00CEF0F7&
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   540
         TabIndex        =   28
         Top             =   7530
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   300
         Width           =   1605
      End
      Begin VB.ListBox List1 
         Height          =   6495
         Left            =   90
         TabIndex        =   26
         Top             =   690
         Width           =   1365
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   585
         Index           =   5
         Left            =   6330
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   870
         Width           =   5385
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   585
         Index           =   4
         Left            =   1470
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   870
         Width           =   4875
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Index           =   1
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   300
         Width           =   5025
      End
      Begin VB.ComboBox CmbDetail 
         Height          =   315
         Left            =   2190
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   1155
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   1470
         TabIndex        =   17
         Top             =   1380
         Width           =   12045
         Begin VB.TextBox txtfields 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Index           =   7
            Left            =   3390
            TabIndex        =   5
            Top             =   180
            Width           =   6765
         End
         Begin VB.TextBox txtfields 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   6
            Left            =   1350
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
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
            Index           =   10
            Left            =   2640
            TabIndex        =   21
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Performance"
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
            Index           =   8
            Left            =   210
            TabIndex        =   20
            Top             =   255
            Width           =   1080
         End
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00CEF0F7&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00CEF0F7&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   60
         TabIndex        =   2
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Academic Year"
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
         Index           =   4
         Left            =   3390
         TabIndex        =   30
         Top             =   90
         Width           =   1290
      End
      Begin VB.Shape Shape1 
         Height          =   435
         Left            =   5910
         Top             =   7350
         Width           =   5805
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lesson(Oral)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   195
         Index           =   2
         Left            =   6330
         TabIndex        =   25
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lesson(Written)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   195
         Index           =   5
         Left            =   1530
         TabIndex        =   24
         Top             =   645
         Width           =   1290
      End
      Begin VB.Label Label4 
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
         Left            =   5040
         TabIndex        =   19
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Roll"
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
         Index           =   7
         Left            =   10020
         TabIndex        =   18
         Top             =   90
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Details #"
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
         Index           =   6
         Left            =   2220
         TabIndex        =   16
         Top             =   90
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topic#"
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
         Left            =   1080
         TabIndex        =   15
         Top             =   90
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lesson #"
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
         Left            =   120
         TabIndex        =   13
         Top             =   90
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmStudent_performance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim font_indicator As Integer
Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  RickLecdetail.SetFocus
End If
End Sub
Private Sub CmbDetail_Click()
   
   Dim rs As New ADODB.Recordset
   Dim rs1 As New ADODB.Recordset
   
   Set rs = getdata("SELECT Oral, Written from Ls_plan_details where Topic_srl= '" & txtFields(3).Text & "' and Details_srl= '" & CmbDetail.Text & "'")
   
   If rs.EOF Or rs.BOF Then
    Exit Sub
   Else
    If font_indicator = 0 Then
        txtFields(4).FontName = "MS Sans Serif"
        txtFields(5).FontName = "MS Sans Serif"
                 
        txtFields(5).Text = rs("Oral")
        txtFields(4).Text = rs("Written")
    Else
        txtFields(4).FontName = "SutonnyMJ"
        txtFields(5).FontName = "SutonnyMJ"
        txtFields(5).Text = rs("Oral")
        txtFields(4).Text = rs("Written")
    End If
    Set rs = Nothing
   End If
   list_data_refresh
   
   Call ShowFlexData
End Sub

Private Sub cmdDelete_Click()

If Len(CmbDetail.Text) = 0 Then
    MsgBox "Detail Serial Mandatory.", vbInformation, cmp
    CmbDetail.SetFocus
    Exit Sub
End If

    If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical) = vbYes Then
         
        
        Dim con As New ADODB.connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        
        con.Open GConnString
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "STD_PERFORMANCE_Save"
        cmd("@mode") = "D"
       cmd("@Student_id") = Trim(txtFields(8).Text)
        cmd("@classid") = Trim(frmLessonPlanMain.Combo1)
        cmd("@sectionid") = Trim(frmLessonPlanMain.Combo5)
        cmd("@Class_roll") = Val(Trim(txtFields(0).Text))
        cmd("@Srl_no") = Val(Trim(txtFields(2).Text))
        cmd("@Topic_srl") = Val(Trim(txtFields(3).Text))
        cmd("@Details_srl") = Val(Trim(CmbDetail.Text))
        cmd("@Prfm") = Trim(txtFields(6).Text)
        cmd("@Remarks") = Trim(txtFields(7).Text)
        cmd("@Entry_by") = soft_user
        cmd("@Entry_date") = Format(Date, "dd mmm yyyy")
        cmd("@Academic_yr") = Trim(Combo1.Text)
        cmd.Execute

        MsgBox "Deleted successfully.", vbInformation, "Student Management System"
        Call ShowFlexData
        Call list_data_refresh
        cmdnew.SetFocus
     End If
End Sub

Private Sub CmdEdit_Click()

If MsgBox("Are you Sure to update?", vbYesNo, "Student Management System") = vbYes Then
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
Set cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "STD_PERFORMANCE_Save"
cmd("@mode") = "U"
cmd("@Student_id") = Trim(txtFields(8).Text)
cmd("@classid") = Trim(frmLessonPlanMain.Combo1)
cmd("@sectionid") = Trim(frmLessonPlanMain.Combo5)
cmd("@Class_roll") = Val(Trim(txtFields(0).Text))
cmd("@Srl_no") = Val(Trim(txtFields(2).Text))
cmd("@Topic_srl") = Val(Trim(txtFields(3).Text))
cmd("@Details_srl") = Val(Trim(CmbDetail.Text))
cmd("@Prfm") = Trim(txtFields(6).Text)
cmd("@Remarks") = Trim(txtFields(7).Text)
cmd("@Entry_by") = soft_user
cmd("@Entry_date") = Format(Date, "dd mmm yyyy")
cmd("@Academic_yr") = Trim(Combo1.Text)
cmd.Execute
MsgBox "Updated successfully.", vbInformation, "Student Management System"
End If
Call ShowFlexData
cmdnew.SetFocus
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdnew_Click()
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

If Len(CmbDetail.Text) = 0 Then
    MsgBox "Detail Serial Mandatory.", vbInformation, cmp
    CmbDetail.SetFocus
    Exit Sub
End If

If Len(txtFields(0).Text) = 0 Then
    MsgBox "Student ID is Mandatory.", vbInformation, cmp
    List1.SetFocus
    Exit Sub
End If

If Len(txtFields(0).Text) = 0 Then
    MsgBox "Class Roll Mandatory.", vbInformation, cmp
    CmbClsRoll.SetFocus
    Exit Sub
End If

If Len(txtFields(6).Text) = 0 Then
    MsgBox "Student Performance is Mandatory.", vbInformation, cmp
    txtFields(6).SetFocus
    Exit Sub
End If

If Len(List1.Text) = 0 Then
    MsgBox "Please Select a student Id from the list", vbInformation, cmp
    List1.SetFocus
    Exit Sub
End If



'If (MaskEdBox1.Text) = "__/__/__" Then
'    MsgBox "Topic Date Required..", vbInformation, cmp
'    MaskEdBox1.SetFocus
'    Exit Sub
'End If

Dim rs As New ADODB.Recordset


Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "STD_PERFORMANCE_Save"
cmd("@mode") = "S"
cmd("@Student_id") = Trim(List1.Text)
cmd("@classid") = Trim(frmLessonPlanMain.Combo1)
cmd("@sectionid") = Trim(frmLessonPlanMain.Combo5)
cmd("@Class_roll") = Trim(txtFields(0))
cmd("@Srl_no") = Val(Trim(txtFields(2)))
cmd("@Topic_srl") = Val(Trim(txtFields(3)))
cmd("@Details_srl") = Val(Trim(CmbDetail.Text))
cmd("@Prfm") = Trim(txtFields(6))
cmd("@Remarks") = Trim(txtFields(7))
cmd("@Entry_by") = soft_user
cmd("@Entry_date") = Format(Date, "dd mmm yyyy")
cmd("@Academic_yr") = Trim(Combo1.Text)
cmd.Execute
MsgBox "Save successfully.", vbInformation, "Student Management System"
Call ShowFlexData
cmdnew.SetFocus
list_data_refresh
End Sub
Private Sub Combo1_Click()
   load_section
   load_subject
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select ClassName from classinfo where classId= '" & Trim(Combo1.Text) & "'")
   
   If Not rs.EOF Then
      Combo3.Text = rs(0)
   End If
   list_data_refresh
   Call ShowFlexData
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
   Set rs = getdata("Select a.Sub_title from subject_info_sub a ,subjectinfomain b  where a.M_code=b.M_code and b.Class_code='" & Trim(Combo1.Text) & "' and Sub_code='" & Combo2.Text & "'")
   
   If Not rs.EOF Then
      Combo4.Text = rs(0)
   End If
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



Private Sub Combo3_Click()
  Dim rs As New ADODB.Recordset
   Set rs = getdata("Select classId from classinfo where ClassName= '" & Trim(Combo3.Text) & "'")
   
   If Not rs.EOF Then
      Combo1.Text = rs(0)
   End If
End Sub

Private Sub Combo4_Click()
Dim rs As New ADODB.Recordset
   Set rs = getdata("Select a.Sub_code  from subject_info_sub a ,subjectinfomain b  where a.M_code=b.M_code and b.Class_code='" & Trim(Combo1.Text) & "' and a.Sub_title='" & Combo4.Text & "'")
   
   If Not rs.EOF Then
      Combo2.Text = rs(0)
   End If
End Sub

Private Sub Combo5_Click()
   'load_section
End Sub

Private Sub Combo7_Click()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select ETypeName from ExamTypeInfo where ETypeID= '" & Trim(Combo7.Text) & "'")
   
   If Not rs.EOF Then
      Combo8.Text = rs(0)
   End If
   load_exam_sub
End Sub
Private Sub load_exam_sub()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select Exam_code,Exam_title from Exam_setup where Group_code= '" & Trim(Combo7.Text) & "'")
   
   Combo9.Clear
   Combo10.Clear
   If Not rs.EOF Then
      Combo9.AddItem rs(0)
      Combo10.AddItem rs(1)
   
   End If
   
End Sub
Private Sub Combo8_Click()
  Dim rs As New ADODB.Recordset
   Set rs = getdata("Select ETypeID from ExamTypeInfo where ETypeName= '" & Trim(Combo8.Text) & "'")
   
   If Not rs.EOF Then
      Combo7.Text = rs(0)
   End If
End Sub

Private Sub Combo9_Click()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select Exam_title from Exam_setup where Group_code= '" & Trim(Combo7.Text) & "' and Exam_code='" & Trim(Combo9.Text) & "'")
   Combo10.Clear
   If Not rs.EOF Then
      Combo10.Text = rs(0)
   End If
End Sub

Private Sub Command1_Click()
  rptMode = 6
  Screen.MousePointer = vbHourglass
  frmViewer.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
SendKeys Chr(9)
End If
End Sub
Private Sub list_data()
  Dim rs As New ADODB.Recordset
  Set rs = getdata("Select a.StudentId from StudentAdmission a where a.approval='Y' and a.admissionCancel='N'and a.ClassId='" & Trim(frmLessonPlanMain.Combo1) & "' and a.SectionId='" & Trim(frmLessonPlanMain.Combo5) & "' and  a.serial_no=(select max(serial_no) from studentadmission where studentid=a.studentid)")
  If Not rs.EOF Then
     While Not rs.EOF
          List1.AddItem Trim(rs(0))
          rs.MoveNext
     Wend
 End If
End Sub
Private Sub list_data_refresh()
  Dim local_rs As New ADODB.Recordset
  
  Set local_rs = getdata("Select a.StudentId,(select  Studentname from studentinfo b where b.studentid=a.studentid) as StudentName from StudentAdmission a where a.approval='Y' and a.admissionCancel='N'and a.ClassId='" & Trim(frmLessonPlanMain.Combo1) & "' and a.SectionId='" & Trim(frmLessonPlanMain.Combo5) & "' and  a.serial_no=(select max(serial_no) from studentadmission where studentid=a.studentid) and a.StudentId not in(select Student_id from std_study_performance where Classid='" & Trim(frmLessonPlanMain.Combo1) & "' and Sectionid='" & Trim(frmLessonPlanMain.Combo5) & "' and Srl_no='" & Trim(txtFields(2).Text) & "' and  Topic_srl='" & Trim(txtFields(3).Text) & "' and  Details_srl='" & Trim(CmbDetail.Text) & "'and Academic_yr='" & Trim(Combo1.Text) & "')")
  
  List1.Clear
  If Not local_rs.EOF Then
  While Not local_rs.EOF
       List1.AddItem Trim(local_rs(0))
       local_rs.MoveNext
  Wend
 End If
End Sub

Private Sub Form_Load()
' Dim IL As Integer
' For IL = 2000 To 2050
'    Combo1.AddItem (IL)
' Next IL
  txtFields(2).Text = frmLessonPlanMain.txtFields(0).Text
  txtFields(3).Text = frmTopic_serial.txtFields(3).Text
  Combo1.Text = Trim(frmTopic_serial.combo_aca)
   list_data

Set rs = getdata("Select Details_srl,font_indicator from Ls_plan_details where Topic_srl= '" & txtFields(3).Text & "'")

    Do Until rs.EOF
        CmbDetail.AddItem rs(0)
        font_indicator = rs!font_indicator
        rs.MoveNext
    Loop
With MSFlexGrid1
    .Rows = 1
    .Cols = 5
    .Col = 0: .Text = " Student ID "
    .Col = 1: .Text = " Class Roll"
    .Col = 2: .Text = " Student Name "
    .Col = 3: .Text = " Performance "
    .Col = 4: .Text = " Remarks "
'    .Col = 5: .Text = " Detail Serial "

        
    
    .ColWidth(0) = 1800
    .ColWidth(1) = 1000
    .ColWidth(2) = 4000
    .ColWidth(3) = 500
    .ColWidth(4) = 3000
'    .ColWidth(5) = 1000
''    .ColWidth(4) = 5500

    
End With
'ShowFlexData
End Sub
Private Sub load_class()
'Dim rs1 As New adodb.Recordset
'Dim rs2 As New adodb.Recordset
'''Combo1.Clear
''Combo3.Clear
'Set rs1 = GetData("Select ClassId,ClassName from ClassInfo")
'If Not rs1.EOF Then
'    Do Until rs1.EOF
'        Combo1.AddItem rs1(0)
'        Combo3.AddItem rs1(1)
'        rs1.MoveNext
'    Loop
'    Combo1.AddItem (" ")
''    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
'End If
'
''    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
'
'
End Sub
Private Sub load_exam()
'Dim rs1 As New adodb.Recordset
'Dim rs2 As New adodb.Recordset
'Combo7.Clear
'Combo8.Clear
'Set rs1 = GetData("Select ETypeID,ETypeName from ExamTypeInfo")
'If Not rs1.EOF Then
'    Do Until rs1.EOF
'        Combo7.AddItem rs1(0)
'        Combo8.AddItem rs1(1)
'        rs1.MoveNext
'    Loop
''    Combo1.AddItem (" ")
'End If
End Sub
Private Sub load_subject()
'Dim rs1 As New adodb.Recordset
'Dim rs2 As New adodb.Recordset
'Combo2.Clear
'Combo4.Clear
'Set rs1 = GetData("Select a.Sub_code,a.Sub_title from subject_info_sub a ,subjectinfomain b  where a.M_code=b.M_code and b.Class_code='" & Trim(Combo1.Text) & "'")
'If Not rs1.EOF Then
'    Do Until rs1.EOF
'        Combo2.AddItem rs1(0)
'        Combo4.AddItem rs1(1)
'        rs1.MoveNext
'    Loop
''    Combo1.AddItem (" ")
'End If
End Sub

Private Sub load_section()
'Dim rs1 As New adodb.Recordset
'Dim rs2 As New adodb.Recordset
'Combo5.Clear
'Combo6.Clear
'Set rs1 = GetData("Select SectionID,Sectiondsc from SectionInfo where ClassID='" & Trim(Combo1.Text) & "'")
'If Not rs1.EOF Then
'    Do Until rs1.EOF
'        Combo5.AddItem rs1(0)
'        Combo6.AddItem rs1(1)
'        rs1.MoveNext
'    Loop
'End If
'Combo6.Text = Combo6.List(0)
End Sub
Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT Std_Study_performance.Student_id,Std_Study_performance.Class_roll,(select StudentName from studentinfo where StudentID=Std_Study_performance.Student_id)as stdname,Std_Study_performance.Prfm,Std_Study_performance.Remarks,Std_Study_performance.Details_srl, Std_Study_performance.Topic_srl, Std_Study_performance.Srl_no, Std_Study_performance.Details_srl FROM Std_Study_performance INNER JOIN StudentInfo ON StudentInfo.StudentID = Std_Study_performance.Student_id where Srl_no='" & txtFields(2).Text & "' and  Topic_srl= '" & txtFields(3).Text & "' and Details_srl= '" & CmbDetail.Text & "' and Classid='" & Trim(frmLessonPlanMain.Combo1) & "' and Sectionid='" & Trim(frmLessonPlanMain.Combo5) & "' and Academic_yr='" & Trim(Combo1.Text) & "'")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
             MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = rs!Student_id
                .TextMatrix(i, 1) = rs!Class_roll
                .TextMatrix(i, 2) = rs!StdName
                .TextMatrix(i, 3) = rs!Prfm
                .TextMatrix(i, 4) = rs!remarks
                
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
Private Sub cmdsearch_Click()
Dim f As New frmFind
Set f.OwnerForm = Me
    f.intInputsel = 0
    f.SQLString = "Select a.StudentId,(select  Studentname from studentinfo b where b.studentid=a.studentid) as StudentName from StudentAdmission a where a.approval='Y' and a.admissionCancel='N'and a.ClassId='" & Trim(frmLessonPlanMain.Combo1) & "' and a.SectionId='" & Trim(frmLessonPlanMain.Combo5) & "' and  a.serial_no=(select max(serial_no) from studentadmission where studentid=a.studentid) "
    f.Show 1
    txtFields(0).SetFocus
    Exit Sub
End Sub

Private Sub get_roll()
Dim rs As New ADODB.Recordset
Set rs = getdata("select classRoll From StudentAdmission where StudentId='" & Trim(List1) & "'" & _
 " and serial_no=(select max(serial_no)  From StudentAdmission  where StudentId='" & Trim(List1) & "')")
txtFields(0).Text = rs(0)
End Sub
Private Sub get_name()
Dim rs As New ADODB.Recordset
Set rs = getdata("select studentname From Studentinfo where StudentId='" & Trim(List1) & "'")
txtFields(1).Text = rs(0)
End Sub

Private Sub List1_Click()
 get_roll
 get_name
End Sub

Private Sub MSFlexGrid1_Click()
'On Error GoTo ErrDes
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

txtFields(8) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtFields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtFields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
txtFields(6) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
txtFields(7) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
'CmbDetail.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))


If Len(txtFields(5)) = 0 Then
    Set rs = getdata("SELECT Topic_srl From Ls_plan_details where Written = '" & txtFields(4) & "'")
    txtFields(3).Text = rs!Topic_srl
Else
    Set rs = getdata("SELECT Topic_srl From Ls_plan_details where Oral = '" & txtFields(5) & "'")
    txtFields(3).Text = rs!Topic_srl
End If


'Set rs1 = GetData("Select Details_srl from Ls_plan_details where Topic_srl= '" & txtfields(3).Text & "'")
'    CmbDetail.Clear
'    Do Until rs1.EOF
'        CmbDetail.AddItem rs1(0)
'        rs1.MoveNext
'    Loop


End Sub

Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub
