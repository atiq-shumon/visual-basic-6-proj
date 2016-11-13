VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTopic_serial 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Student Performance Entry"
      Height          =   375
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6420
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Lesson Details Entry"
      Height          =   375
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6420
      Width           =   1815
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   375
      Left            =   5310
      TabIndex        =   1
      ToolTipText     =   "Click to insert new information"
      Top             =   6390
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   375
      Left            =   6300
      TabIndex        =   0
      ToolTipText     =   "Click to save"
      Top             =   6390
      Width           =   945
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   375
      Left            =   7290
      TabIndex        =   2
      ToolTipText     =   "Click to Delete"
      Top             =   6390
      Width           =   945
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      ToolTipText     =   "Click to Exit"
      Top             =   6390
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   765
      Left            =   30
      ScaleHeight     =   705
      ScaleWidth      =   9225
      TabIndex        =   7
      Top             =   0
      Width           =   9285
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lesson Plan(topic )"
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
         Left            =   3510
         TabIndex        =   14
         Top             =   180
         Width           =   2205
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   -60
         Picture         =   "frmTopic_serial.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   9285
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1845
      Left            =   0
      TabIndex        =   6
      Top             =   750
      Width           =   9315
      Begin VB.TextBox txtfields 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   1290
         TabIndex        =   19
         Top             =   990
         Width           =   7755
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bengali"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   18
         Top             =   1470
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "English"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   17
         Top             =   1230
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.ComboBox combo_aca 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   570
         Width           =   1095
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1290
         TabIndex        =   4
         Top             =   570
         Width           =   1545
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1290
         TabIndex        =   5
         Top             =   150
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Academic Year"
         Height          =   195
         Index           =   3
         Left            =   2910
         TabIndex        =   11
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topic Title"
         Height          =   195
         Index           =   2
         Left            =   30
         TabIndex        =   10
         Top             =   990
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topic Serial"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   570
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lesson Serial"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   150
         Width           =   945
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3735
      Left            =   30
      TabIndex        =   16
      Top             =   2580
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   6588
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
      Height          =   435
      Left            =   5280
      Top             =   6360
      Width           =   3975
   End
End
Attribute VB_Name = "frmTopic_serial"
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
On Error Resume Next
If Len(txtfields(0).Text) = 0 Then
    MsgBox "Lesson Serial Mandatory.", vbInformation, App.Title
    txtfields(0).SetFocus
    Exit Sub
End If

'If Len(txtFields(4)) = 0 Then
'    MsgBox "Topic Title Required.", vbInformation, App.Title
'    txtFields(4).SetFocus
'    Exit Sub
'End If



Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString

Set rs = getdata("select Srl_no from ls_plan_details where Srl_no='" & Trim(txtfields(0)) & "' and Topic_srl='" & Trim(txtfields(3).Text) & "' and Academic_yr='" & Trim(combo_aca.Text) & "'")

If Not rs.EOF Then
   MsgBox "First Delete all Details under this topic", vbInformation, cmp
   Exit Sub
End If

 Set cmd.ActiveConnection = con
    If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical) = vbYes Then

       cmd.ActiveConnection = con
       cmd.CommandType = adCmdStoredProc
       cmd.CommandText = "LS_PLAN_TOPIC_Save"
       cmd(1) = "D"
       cmd(2) = Val(Trim(txtfields(0)))
       cmd(3) = Val(Trim(txtfields(3)))
       If Option1(0).Value = True Then
           txtfields(1).FontName = "MS Sans Serif"
       ElseIf Option1(1).Value = True Then
          txtfields(1).FontName = "SutonnyMJ"
       End If ''Sentance
       cmd(4) = txtfields(1)
       cmd(5) = Trim(combo_aca)
       cmd(6) = soft_user
       cmd(7) = Format(Date, "dd mmm yyyy")
       If Option1(0).Value = True Then
        cmd(8) = 0
       ElseIf Option1(1).Value = True Then
        cmd(8) = 1
       End If
       cmd.Execute
       MsgBox "Deleted successfully.", vbInformation, "Student Management System"
       Call ShowFlexData
       cmdnew.SetFocus
       cmdsave.Enabled = False
     End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdnew_Click()
        txtfields(3).Text = ""
        txtfields(1).Text = ""
        
        cmdsave.Enabled = True
End Sub
Private Sub cmdSAVE_Click()
Dim Sentance As String
If Len(txtfields(0).Text) = 0 Then
    MsgBox "Lesson Serial Mandatory.", vbInformation, App.Title
    txtfields(0).SetFocus
    Exit Sub
End If

If Len(txtfields(1)) = 0 Then
    MsgBox "Topic Title Required.", vbInformation, App.Title
    txtfields(4).SetFocus
    Exit Sub
End If

Dim rs As New ADODB.Recordset


Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LS_PLAN_TOPIC_Save"
cmd(1) = "S"
cmd(2) = Val(Trim(txtfields(0)))
cmd(3) = Val(Trim(txtfields(3)))
If Option1(0).Value = True Then
    txtfields(1).FontName = "MS Sans Serif"
ElseIf Option1(1).Value = True Then
   txtfields(1).FontName = "SutonnyMJ"
End If ''Sentance
cmd(4) = txtfields(1)
cmd(5) = Trim(combo_aca)
cmd(6) = soft_user
cmd(7) = Format(Date, "dd mmm yyyy")

If Option1(0).Value = True Then
 cmd(8) = 0
ElseIf Option1(1).Value = True Then
 cmd(8) = 1
End If
cmd(9) = Trim(combo_aca)

cmd.Execute
MsgBox "Save successfully.", vbInformation, "Student Management System"
Call ShowFlexData
cmdnew.SetFocus
cmdsave.Enabled = False

End Sub
Private Sub Combo1_Click()
   load_section
   load_subject
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select ClassName from classinfo where classId= '" & Trim(Combo1.Text) & "'")
   
   If Not rs.EOF Then
      Combo3.Text = rs(0)
   End If
   
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

Private Sub combo_aca_Click()
  ShowFlexData
  txtfields(3) = ""
End Sub

Private Sub Command1_Click()
  If Len(txtfields(3).Text) = 0 Then
     MsgBox "Please select a topic serial from the list given below ", vbInformation, cmp
     Exit Sub
  End If
  frmTopic_details.Show 1
End Sub

Private Sub Command2_Click()
If Len(txtfields(3).Text) = 0 Then
     MsgBox "Please select a topic serial from the list given below ", vbInformation, cmp
     Exit Sub
  End If
 
 frmStudent_performance.Show 1
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
  txtfields(0).Text = frmLessonPlanMain.txtfields(0).Text
  
' load_class
' load_section
' load_exam

With MSFlexGrid1
    .Rows = 1
    .Cols = 5
    .Col = 0: .Text = " Serial "
    .Col = 1: .Text = " Topic Serial"
    .Col = 2: .Text = " Topic Title "
    .Col = 3: .Text = " Acadmic Year "
    .Col = 4: .Text = " font indicator "
    .ColWidth(0) = 800
    .ColWidth(1) = 800
    .ColWidth(2) = 6000
    .ColWidth(3) = 1500
    .ColWidth(4) = 0
    
End With
ShowFlexData
Dim i As Integer
For i = 2000 To 2050
  combo_aca.AddItem i
Next i
combo_aca.Text = Format(Date, "YYYY")
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
'On Error GoTo errdes
Dim rs As New ADODB.Recordset

Set rs = getdata("SELECT  Srl_no,Topic_srl,Topic_title,AcademicYr,font_indicator From Ls_plan_topic where Srl_no='" & Trim(txtfields(0).Text) & " ' and AcademicYr='" & Trim(combo_aca.Text) & "'")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = rs!srl_no
                .TextMatrix(i, 1) = rs!Topic_srl
                If rs!font_indicator = 0 Then
                  .Col = 2
                  .Row = i
                  .CellFontName = "MS Sans Serif"
                  .TextMatrix(i, 2) = rs!Topic_title
                ElseIf rs!font_indicator = 1 Then
                  .Col = 2
                  .Row = i
                  .CellFontName = "SutonnyMJ"
                  .TextMatrix(i, 2) = rs!Topic_title
                End If
             
                .TextMatrix(i, 3) = rs!AcademicYr
                .TextMatrix(i, 4) = rs!font_indicator
                
                i = i + 1
            rs.MoveNext
        Loop
'       .FontName = "MS Sans Serif"
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
txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfields(3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4) = 0 Then
  Option1(0).Value = True
  txtfields(1).FontName = "MS Sans Serif"
  txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
Else
  Option1(1).Value = True
  txtfields(1).FontName = "SutonnyMJ"
  txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
End If
  
txtfields(5) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
'If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = "Y" Then
'    Check1.Value = 1
'Else
'    Check1.Value = 0
'End If
'RickLecdetail = ""
'Set rs = GetData("SELECT  LectureDetail From LectureInfo where classid='" & Mid(Trim(Combo1.Text), 1, 5) & "' and SubjectID ='" & Mid(Trim(Combo2.Text), 1, 5) & "'and LectureID='" & Trim(txtfields(0)) & "'")
'If Not rs.EOF Then
'    RickLecdetail = rs!LectureDetail
'End If
Exit Sub
errdes:
'MsgBox err.Description, vbInformation, App.Title
End Sub
Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub

Private Sub Option1_Click(Index As Integer)
      Select Case Index
             Case 0
                
                 With txtfields(1)
                   .FontName = "MS Sans Serif"
                 End With
             Case 1
                 With txtfields(1)
                   .FontName = "SutonnyMJ"
                 End With
     End Select
End Sub

Private Sub txtfields_Change(Index As Integer)
        Select Case Index
               Case 3
                    If Len(txtfields(3)) > 0 Then
                       cmdsave.Enabled = False
                    Else
                       cmdsave.Enabled = True
                    End If
        End Select
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   Select Case Index
     Case 1
'         txtfields(2).SetFocus
     Case 2
          Check1.SetFocus
   End Select
End If
End Sub
