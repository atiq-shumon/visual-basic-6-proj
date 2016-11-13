VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTopic_details 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdEdit 
      BackColor       =   &H8000000C&
      Caption         =   "Edit"
      Height          =   375
      Left            =   8940
      TabIndex        =   24
      ToolTipText     =   "Click to save"
      Top             =   8160
      Width           =   945
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   375
      Left            =   7020
      TabIndex        =   7
      ToolTipText     =   "Click to insert new information"
      Top             =   8160
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   375
      Left            =   7980
      TabIndex        =   6
      ToolTipText     =   "Click to save"
      Top             =   8160
      Width           =   945
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   375
      Left            =   9900
      TabIndex        =   9
      ToolTipText     =   "Click to Delete"
      Top             =   8160
      Width           =   945
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   375
      Left            =   10860
      TabIndex        =   8
      ToolTipText     =   "Click to Exit"
      Top             =   8160
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   765
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12075
      TabIndex        =   13
      Top             =   0
      Width           =   12135
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lesson Plan(Details )"
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
         Left            =   3990
         TabIndex        =   25
         Top             =   150
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   -120
         Picture         =   "frmTopic_details.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   12015
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5085
      Left            =   90
      TabIndex        =   28
      Top             =   2970
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8969
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
   Begin VB.Frame Frame1 
      Height          =   3345
      Left            =   -30
      TabIndex        =   12
      Top             =   720
      Width           =   12045
      Begin VB.TextBox txtfields 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   5070
         TabIndex        =   26
         Top             =   300
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   0
         TabIndex        =   20
         Top             =   780
         Width           =   12045
         Begin VB.Frame Frame2 
            Height          =   495
            Left            =   0
            TabIndex        =   23
            Top             =   -90
            Width           =   12165
            Begin VB.OptionButton Option1 
               Caption         =   "Home Work"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   0
               Left            =   180
               TabIndex        =   2
               Top             =   150
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Class  Work"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   1
               Left            =   1410
               TabIndex        =   3
               Top             =   150
               Width           =   1155
            End
         End
         Begin VB.TextBox txtfields 
            Height          =   465
            Index           =   4
            Left            =   6150
            TabIndex        =   5
            Top             =   660
            Width           =   5475
         End
         Begin VB.TextBox txtfields 
            Height          =   465
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   660
            Width           =   5475
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
            Left            =   6150
            TabIndex        =   22
            Top             =   450
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
            Left            =   150
            TabIndex        =   21
            Top             =   435
            Width           =   1290
         End
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   8370
         TabIndex        =   18
         Top             =   300
         Width           =   945
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   10260
         TabIndex        =   1
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   6540
         TabIndex        =   0
         Top             =   300
         Width           =   795
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3180
         TabIndex        =   10
         Top             =   330
         Width           =   855
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   11
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Academic Yr"
         Height          =   195
         Index           =   7
         Left            =   4080
         TabIndex        =   27
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Details Serial "
         Height          =   195
         Index           =   6
         Left            =   7380
         TabIndex        =   19
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topic Date"
         Height          =   195
         Index           =   4
         Left            =   9360
         TabIndex        =   17
         Top             =   330
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week"
         Height          =   195
         Index           =   3
         Left            =   6030
         TabIndex        =   16
         Top             =   330
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topic Serial"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   15
         Top             =   330
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lesson Serial #"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   330
         Width           =   1095
      End
   End
   Begin VB.Shape Shape1 
      Height          =   435
      Left            =   6990
      Top             =   8130
      Width           =   4845
   End
End
Attribute VB_Name = "frmTopic_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim font_indicator As Integer
'Dim LectureDetail As String
Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  RickLecdetail.SetFocus
End If
End Sub

Private Sub cmdDelete_Click()

If Len(txtFields(0).Text) = 0 Then
    MsgBox "Lesson Serial Mandatory.", vbInformation, App.Title
    txtFields(0).SetFocus
    Exit Sub
End If

    If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical) = vbYes Then
         
        
        Dim con As New ADODB.connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
        con.Open GConnString
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "LS_PLAN_DETAILS_Save"
        cmd(1) = "D"
        cmd(2) = Val(Trim(txtFields(0)))
        cmd(3) = Val(Trim(txtFields(3)))
        cmd(4) = Val(Trim(txtFields(2)))
        cmd(5) = Trim(txtFields(6))
        cmd(6) = Format(MaskEdBox1.Text, "dd mmm yyyy")
        cmd(7) = Trim(txtFields(5))
        If Option1(0).Value = True Then
           cmd(8) = "HW"
        Else
           cmd(8) = "CW"
        End If
        cmd(9) = Trim(txtFields(4))
        cmd(10) = Trim(txtFields(1))
        cmd(11) = soft_user
        cmd(12) = Format(Date, "dd mmm yyyy")
        cmd(13) = Val(font_indicator)
        cmd.Execute
        MsgBox "Deleted successfully.", vbInformation, "Student Management System"
        Call ShowFlexData
        cmdnew.SetFocus
     End If
End Sub

Private Sub CmdEdit_Click()
 If Len(txtFields(0).Text) = 0 Then
    MsgBox "Lesson Serial Mandatory.", vbInformation, cmp
    txtFields(0).SetFocus
    Exit Sub
End If

If (MaskEdBox1.Text) = "__/__/__" Then
    MsgBox "Topic Date Required..", vbInformation, cmp
    MaskEdBox1.SetFocus
    Exit Sub
End If

If Len(txtFields(2).Text) = 0 Then
    MsgBox "Lesson Details Serial Mandatory.Please select Details from the list below", vbInformation, cmp
'    txtfields(2).SetFocus
    Exit Sub
End If

Dim rs As New ADODB.Recordset


Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LS_PLAN_DETAILS_Save"
cmd(1) = "U"
cmd(2) = Val(Trim(txtFields(0)))
cmd(3) = Val(Trim(txtFields(3)))
cmd(4) = Val(Trim(txtFields(2)))
cmd(5) = Trim(txtFields(6))
cmd(6) = Format(MaskEdBox1.Text, "dd mmm yyyy")
cmd(7) = Trim(txtFields(5))
If Option1(0).Value = True Then
   cmd(8) = "HW"
Else
   cmd(8) = "CW"
End If
cmd(9) = Trim(txtFields(4))
cmd(10) = Trim(txtFields(1))
cmd(11) = soft_user
cmd(12) = Format(Date, "dd mmm yyyy")
cmd(13) = Val(font_indicator)
cmd.Execute
MsgBox "Edited successfully.", vbInformation, "Student Management System"
Call ShowFlexData
cmdnew.SetFocus


End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdnew_Click()
 MaskEdBox1.SetFocus
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

If Len(txtFields(0).Text) = 0 Then
    MsgBox "Lesson Serial Mandatory.", vbInformation, cmp
    txtFields(0).SetFocus
    Exit Sub
End If

If (MaskEdBox1.Text) = "__/__/__" Then
    MsgBox "Topic Date Required..", vbInformation, cmp
    MaskEdBox1.SetFocus
    Exit Sub
End If

Dim rs As New ADODB.Recordset


Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "LS_PLAN_DETAILS_Save"
cmd(1) = "S"
cmd(2) = Val(Trim(txtFields(0)))
cmd(3) = Val(Trim(txtFields(3)))
cmd(4) = Val(Trim(txtFields(2)))
cmd(5) = Trim(txtFields(6))
cmd(6) = Format(MaskEdBox1.Text, "dd mmm yyyy")
cmd(7) = Trim(txtFields(5))
If Option1(0).Value = True Then
   cmd(8) = "HW"
Else
   cmd(8) = "CW"
End If
cmd(9) = Trim(txtFields(4))
cmd(10) = Trim(txtFields(1))
cmd(11) = soft_user
cmd(12) = Format(Date, "dd mmm yyyy")
cmd(13) = Val(font_indicator)
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

Private Sub Command1_Click()
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
If frmTopic_serial.Option1(0).Value = True Then
   font_indicator = 0
   txtFields(1).FontName = "MS Sans Serif"
   txtFields(4).FontName = "MS Sans Serif"
Else
   font_indicator = 1
   txtFields(1).FontName = "SutonnyMJ"
   txtFields(4).FontName = "SutonnyMJ"
End If
  
  txtFields(0).Text = frmLessonPlanMain.txtFields(0).Text
  txtFields(3).Text = frmTopic_serial.txtFields(3).Text
  txtFields(6).Text = frmTopic_serial.combo_aca.Text
'  txtfields(5).Text = frmTopic_serial.txtfields(5).Text
  
' load_class
' load_section
' load_exam

With MSFlexGrid1
    .Rows = 1
    .Cols = 7
    .Col = 0: .Text = " # "
    .Col = 1: .Text = " HW/CW "
    .Col = 2: .Text = " Date "
    .Col = 3: .Text = " Written "
    .Col = 4: .Text = " Oral "
    .Col = 5: .Text = "font Indicator"
    
    .ColWidth(0) = 500
    .ColWidth(1) = 500
    .ColWidth(2) = 1000
    .ColWidth(3) = 5500
    .ColWidth(4) = 5500
    .ColWidth(5) = 0
    .ColWidth(6) = 0
 
End With
ShowFlexData
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

Set rs = getdata("SELECT  Details_srl,Ls_date,HW_CW,Oral,Written,font_indicator,LS_Week From Ls_plan_details where Srl_no='" & Trim(txtFields(0).Text) & " ' and Topic_srl='" & Trim(txtFields(3).Text) & "' ")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .Rows = i + 1
                .TextMatrix(i, 0) = rs!Details_srl
                .TextMatrix(i, 1) = rs!HW_CW
                .TextMatrix(i, 2) = rs!Ls_date
                
                If rs!font_indicator = 0 Then
                  .Col = 3
                  .Row = i
                  .CellFontName = "MS Sans Serif"
                  .TextMatrix(i, 3) = rs!Written
                  .Col = 4
                  .Row = i
                  .CellFontName = "MS Sans Serif"
                  .TextMatrix(i, 4) = rs!Oral
                ElseIf rs!font_indicator = 1 Then
                  .Col = 3
                  .Row = i
                  .CellFontName = "SutonnyMJ"
                  .TextMatrix(i, 3) = rs!Written
                  .Col = 4
                  .Row = i
                  .CellFontName = "SutonnyMJ"
                  .TextMatrix(i, 4) = rs!Oral
                                   
                 
                End If
                .Col = 6
                .Row = i
                .TextMatrix(i, 6) = rs!LS_Week
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

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If MaskEdDate <> "__/__/__" Then
            If Check_ValidDate(MaskEdBox1) = False Then
                MaskEdBox1.SetFocus
                Exit Sub
            End If
    End If

End If
End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
txtFields(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)

If (MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)) = "HW" Then
    Option1(0).Value = True
 Else
   Option1(1).Value = True
End If
MaskEdBox1 = Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2), "dd/mm/yy")
txtFields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
txtFields(4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
txtFields(5) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
Exit Sub
errdes:
  MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MSFlexGrid1_SelChange()
      MSFlexGrid1_Click
End Sub
