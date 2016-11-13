VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmsubmarksdis 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8775
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000C&
      Caption         =   "Edit"
      Height          =   405
      Left            =   5730
      TabIndex        =   21
      ToolTipText     =   "Click to Save"
      Top             =   5730
      Width           =   945
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   405
      Left            =   7710
      TabIndex        =   10
      ToolTipText     =   "Click to Exit"
      Top             =   5730
      Width           =   945
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   405
      Left            =   6720
      TabIndex        =   9
      ToolTipText     =   "Click to Delete"
      Top             =   5730
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   405
      Left            =   4740
      TabIndex        =   7
      ToolTipText     =   "Click to Save"
      Top             =   5730
      Width           =   945
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   405
      Left            =   3750
      TabIndex        =   8
      ToolTipText     =   "Click to insert new infornmation"
      Top             =   5730
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8805
      TabIndex        =   12
      Top             =   0
      Width           =   8865
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Marks Distribution"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEF0F7&
         Height          =   270
         Left            =   2760
         TabIndex        =   20
         Top             =   150
         Width           =   2820
      End
      Begin VB.Image Image1 
         Height          =   930
         Left            =   -30
         Picture         =   "frmsubmarksdis.frx":0000
         Stretch         =   -1  'True
         Top             =   -60
         Width           =   8775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   2175
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   8775
      Begin VB.TextBox txtfieldsPass 
         Height          =   285
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   5
         ToolTipText     =   "Select Marks"
         Top             =   1470
         Width           =   975
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   870
         Width           =   2955
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   870
         Width           =   2895
      End
      Begin VB.TextBox txtfieldsFull 
         Height          =   285
         Left            =   7440
         MaxLength       =   3
         TabIndex        =   6
         ToolTipText     =   "Select Marks"
         Top             =   1485
         Width           =   975
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Select Marks Category"
         Top             =   1470
         Width           =   2895
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select Subject"
         Top             =   330
         Width           =   2955
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select Class"
         Top             =   330
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Code"
         Height          =   195
         Index           =   2
         Left            =   4530
         TabIndex        =   19
         Top             =   900
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Term Code"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   18
         Top             =   930
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pass Marks"
         Height          =   195
         Left            =   4530
         TabIndex        =   17
         Top             =   1500
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Marks"
         Height          =   195
         Left            =   6510
         TabIndex        =   16
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mark's Category"
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   1500
         Width           =   1140
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject ID"
         Height          =   195
         Left            =   4530
         TabIndex        =   14
         Top             =   390
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class ID"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   390
         Width           =   585
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2865
      Left            =   0
      TabIndex        =   22
      Top             =   2760
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5054
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
      Height          =   495
      Left            =   3690
      Top             =   5670
      Width           =   4995
   End
End
Attribute VB_Name = "frmsubmarksdis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDelete_Click()
'If Len(Combo1.Text) <> 0 And Len(Combo2.Text) <> 0 And Len(Combo3.Text) <> 0 And Len(txtFields) <> 0 Then
        If (MsgBox("Are You sure to delete ?", vbYesNo + vbInformation) = vbYes) Then
                Dim rs As New ADODB.Recordset
                Dim cmd As New ADODB.Command
                Dim con As New ADODB.connection
                con.Open GConnString
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "Subjectmarksdistribution1"
                cmd(1) = "D"
                cmd(2) = Mid(Trim(Combo1.Text), 1, 5)
                cmd(3) = Mid(Trim(Combo2.Text), 1, 5)
                cmd(4) = Mid(Trim(Combo4.Text), 1, 2)
                cmd(5) = Mid(Trim(Combo5.Text), 1, 2)
                cmd(6) = Mid(Trim(Combo3.Text), 1, 5)
                cmd(7) = Abs(Val(Trim(txtfieldsPass.Text)))
                cmd(8) = Abs(Val(Trim(txtfieldsFull.Text)))
                
                cmd.Execute
                  MsgBox "Delete successfully .", vbInformation, App.Title
                txtFields = ""
'                Combo1.Text = ""
'                Combo2.Text = ""
'                Combo3.Text = ""
'                Label6.Caption = ""
         Else
                Exit Sub
         End If
' Else
'        MsgBox "Data doesn't exist to delete.", vbCritical, "School Management System"
'        Exit Sub
 
Call ShowFlexData
End Sub

Private Sub CmdEdit_Click()
  If Len(Trim(Combo1.Text)) = 0 Then
    MsgBox "Class ID required..", vbInformation, App.Title
    Combo1.SetFocus
    Exit Sub
End If

If Len(Trim(Combo2.Text)) = 0 Then
    MsgBox "Subject ID required..", vbInformation, App.Title
    Combo2.SetFocus
    Exit Sub
End If

If Len(Trim(Combo4.Text)) = 0 Then
    MsgBox "Term code required..", vbInformation, App.Title
    Combo4.SetFocus
    Exit Sub
End If

If Len(Trim(Combo5.Text)) = 0 Then
    MsgBox "Exam code required..", vbInformation, App.Title
    Combo5.SetFocus
    Exit Sub
End If

If Len(Trim(Combo3.Text)) = 0 Then
    MsgBox "Marks Category required..", vbInformation, App.Title
    Combo3.SetFocus
    Exit Sub
End If



If Len(Trim(txtfieldsPass)) = 0 Then
    MsgBox "Enter Pass marks for the category.", vbInformation, App.Title
    txtfieldsPass.SetFocus
    Exit Sub
End If
If Len(Trim(txtfieldsFull)) = 0 Then
    MsgBox "Enter Full marks for the category.", vbInformation, App.Title
    txtfieldsFull.SetFocus
    Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "Subjectmarksdistribution1"
cmd(1) = "u"
cmd(2) = Mid(Trim(Combo1.Text), 1, 5)
cmd(3) = Mid(Trim(Combo2.Text), 1, 5)
cmd(4) = Mid(Trim(Combo4.Text), 1, 2)
cmd(5) = Mid(Trim(Combo5.Text), 1, 2)
cmd(6) = Mid(Trim(Combo3.Text), 1, 5)
cmd(7) = Abs(Val(Trim(txtfieldsPass.Text)))
cmd(8) = Abs(Val(Trim(txtfieldsFull.Text)))

cmd.Execute
MsgBox "Edited successfully.", vbInformation, "Student Management System"
cmdnew.SetFocus
Call ShowFlexData

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
cmdSave.Enabled = True
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
txtFields = ""
Combo1.SetFocus
Call ShowFlexData
End Sub

Private Sub cmdSAVE_Click()
If Len(Trim(Combo1.Text)) = 0 Then
    MsgBox "Class ID required..", vbInformation, App.Title
    Combo1.SetFocus
    Exit Sub
End If

If Len(Trim(Combo2.Text)) = 0 Then
    MsgBox "Subject ID required..", vbInformation, App.Title
    Combo2.SetFocus
    Exit Sub
End If

If Len(Trim(Combo4.Text)) = 0 Then
    MsgBox "Term code required..", vbInformation, App.Title
    Combo4.SetFocus
    Exit Sub
End If

If Len(Trim(Combo5.Text)) = 0 Then
    MsgBox "Exam code required..", vbInformation, App.Title
    Combo5.SetFocus
    Exit Sub
End If

If Len(Trim(Combo3.Text)) = 0 Then
    MsgBox "Marks Category required..", vbInformation, App.Title
    Combo3.SetFocus
    Exit Sub
End If



If Len(Trim(txtfieldsPass)) = 0 Then
    MsgBox "Enter Pass marks for the category.", vbInformation, App.Title
    txtfieldsPass.SetFocus
    Exit Sub
End If
If Len(Trim(txtfieldsFull)) = 0 Then
    MsgBox "Enter Full marks for the category.", vbInformation, App.Title
    txtfieldsFull.SetFocus
    Exit Sub
End If
If Val(txtfieldsPass) > Val(txtfieldsFull) Then
    MsgBox "Pass marks can't be greater than full marks...Please verify", vbInformation, App.Title
    txtfieldsFull.SetFocus
    Exit Sub
End If
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "Subjectmarksdistribution1"
cmd(1) = "S"
cmd(2) = Mid(Trim(Combo1.Text), 1, 5)
cmd(3) = Mid(Trim(Combo2.Text), 1, 5)
cmd(4) = Mid(Trim(Combo4.Text), 1, 2)
cmd(5) = Mid(Trim(Combo5.Text), 1, 2)
cmd(6) = Mid(Trim(Combo3.Text), 1, 5)
cmd(7) = Abs(Val(Trim(txtfieldsPass.Text)))
cmd(8) = Abs(Val(Trim(txtfieldsFull.Text)))

cmd.Execute
MsgBox "Save successfully.", vbInformation, "Student Management System"
cmdnew.SetFocus
Call ShowFlexData
End Sub

Private Sub Combo1_Click()
    load_subject
    ShowFlexData
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    Dim rs2 As New adodb.Recordset
'    Set rs2 = GetData("Select subjectID,subjectdsc from subjectinfo where classId= '" & Mid(Trim(Combo1.Text), 1, 5) & "'")
'        If Not rs2.EOF Then
'            Combo2.Clear
'            Do Until rs2.EOF
'                Combo2.AddItem rs2(0) + " - " + rs2(1)
'                rs2.MoveNext
'            Loop
''            If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
''
'        Else
'            Combo2.Clear
'
'        End If
'        Combo2.SetFocus
'        txtfields = ""
'End If

End Sub


Private Sub Combo1_LostFocus()
'Dim rss As New adodb.Recordset
'Set rss = GetData("select totalmarks from SubjectInfo where classID = '" & Mid(Trim(Combo1), 1, 5) & "' and SubjectID='" & Mid(Trim(Combo2), 1, 5) & "' ")
'If Not (rss.EOF Or rss.BOF) Then
'    Label6.Caption = rss!totalmarks
'End If

End Sub

Private Sub Combo2_Click()
'Call ShowFlexData
'txtfields = ""
ShowFlexData
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    Combo3.SetFocus
'    Call ShowFlexData
'End If
txtFields = ""
End Sub

Private Sub Combo2_LostFocus()
'Dim rss As New adodb.Recordset
'Set rss = GetData("select totalmarks from SubjectInfo where classID = '" & Mid(Trim(Combo1), 1, 5) & "' and SubjectID='" & Mid(Trim(Combo2), 1, 5) & "' ")
'If Not (rss.EOF Or rss.BOF) Then
'    Label6.Caption = rss!totalmarks
'End If
End Sub

Private Sub Combo3_Click()
'    ShowFlexData
End Sub
Private Sub Combo4_Click()
    Dim rs As New ADODB.Recordset
   Set rs = getdata("Select Exam_code,Exam_title from exam_setup where Group_code='" & Mid(Trim(Combo4.Text), 1, 2) & "'")
  
  If Not rs.EOF Then
    Combo5.Clear
    Do Until rs.EOF
        Combo5.AddItem rs(0) + " - " + rs(1)
        rs.MoveNext
    Loop

End If
ShowFlexData
End Sub

Private Sub Combo5_Click()
  ShowFlexData
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        SendKeys (Chr(9))
     End If
End Sub

Private Sub Form_Load()

Dim rs1 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Set rs1 = getdata("Select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo1.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
'    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End If
Set rs3 = getdata("Select McategoryID,McategoryDsc from Markscategory ")
If Not rs3.EOF Then
    Do Until rs3.EOF
        Combo3.AddItem rs3(0) + " - " + rs3(1)
        rs3.MoveNext
    Loop
'    If Combo3.ListCount > 0 Then Combo3.ListIndex = 0
End If
With MSFlexGrid1
    .Rows = 1
    .Cols = 3
    .Col = 0: .Text = " Marks Category "
    .Col = 1: .Text = " Pass Marks "
    .Col = 2: .Text = " Full Marks "
    
    .ColWidth(0) = 4400
    .ColWidth(1) = 2000
    .ColWidth(2) = 2300
    
    
End With

load_term

End Sub
Private Sub load_subject()
Dim rs2 As New ADODB.Recordset
Set rs2 = getdata("Select a.Sub_code,a.Sub_title from subject_info_sub a where  a.Class_code= '" & Mid(Trim(Combo1.Text), 1, 5) & "'")
    If Not rs2.EOF Then
        Combo2.Clear
        Do Until rs2.EOF
            Combo2.AddItem rs2(0) + " - " + rs2(1)
            rs2.MoveNext
        Loop
'        If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
        
    Else
        Combo2.Clear
    End If
    txtFields = ""
    End Sub
Private Sub load_term()
  Dim rs As New ADODB.Recordset
  Set rs = getdata("Select ETypeID,ETypeName from examtypeinfo")
  
  If Not rs.EOF Then
    Combo4.Clear
    Do Until rs.EOF
        Combo4.AddItem rs(0) + " - " + rs(1)
        rs.MoveNext
    Loop

End If

End Sub

Private Sub Label6_Click()

End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
cmdSave.Enabled = True
Combo3.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtfieldsPass = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
txtfieldsFull = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)

Exit Sub
errdes:

End Sub

Private Sub txtfields_KeyPress(KeyAscii As Integer)

'If KeyAscii = 13 Then
'    If Len(Trim(txtfields)) <> 0 Then
'        cmdsave.SetFocus
'    Else
'        cmdnew.SetFocus
'    End If
'End If
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
Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New ADODB.Recordset

Set rs = getdata("select a.CategoryID,(select b.MCategoryDsc  from markscategory b where b.MCategoryID=a.CategoryID),a.passmarks,a.fullmarks from SubjectMarksDistribution  a where a.ClassId = '" & Mid(Trim(Combo1.Text), 1, 5) & "' and a.SubjectId= '" & Mid(Trim(Combo2.Text), 1, 5) & "' and a.term_code='" & Mid(Trim(Combo4.Text), 1, 2) & "' and a.Exam_code='" & Mid(Trim(Combo5.Text), 1, 2) & "'")
                    
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            .Rows = i + 1
            .TextMatrix(i, 0) = rs(0) + " - " + rs(1)
            .TextMatrix(i, 1) = rs(2)
            .TextMatrix(i, 2) = rs(3)
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

Private Sub txtfields_LostFocus()
'If Len(Trim(txtfields)) = 0 Then Exit Sub
'Dim rs As New adodb.Recordset
'Dim rss As New adodb.Recordset
'Dim SubMarks As Double
'
'Set rs = GetData("select submarks from SubjectMarksDistribution where categoryID = '" & Mid(Trim(Combo3), 1, 5) & "'")
'If Not rs.EOF Then
'        Set rs = GetData("select sum(submarks) from SubjectMarksDistribution where classID = '" & Mid(Trim(Combo1), 1, 5) & "' and SubjectID='" & Mid(Trim(Combo2), 1, 5) & "'and categoryId <> '" & Mid(Trim(Combo2), 1, 5) & "' ")
'        If Not rs.EOF Then
'            SubMarks = IIf(IsNull(rs(0)) = True, 0, rs(0))
'        Else
'            SubMarks = 0
'        End If
'
'        If SubMarks + Abs(Val(Trim(txtfields))) > Val(Trim(Label6.Caption)) Then
'            MsgBox "Marks should be in range of total marks.", vbInformation, App.Title
'            txtfields = ""
'            txtfields.SetFocus
'            Exit Sub
'        End If
'Else
'        Set rs = GetData("select sum(submarks) from SubjectMarksDistribution where classID = '" & Mid(Trim(Combo1), 1, 5) & "' and SubjectID='" & Mid(Trim(Combo2), 1, 5) & "' ")
'        If Not rs.EOF Then
'            SubMarks = IIf(IsNull(rs(0)) = True, 0, rs(0))
'        Else
'            SubMarks = 0
'        End If
'
'        If SubMarks + Abs(Val(Trim(txtfields))) > Val(Trim(Label6.Caption)) Then
'            MsgBox "Marks should be in range of total marks.", vbInformation, App.Title
'            txtfields = ""
'            txtfields.SetFocus
'            Exit Sub
'        End If
'End If
'
'If Len(txtfields) = 0 Then
'    cmdsave.Enabled = False
'
'Else
'    If IsNumeric(txtfields) = False Then
'        MsgBox "Please Enter Numeric Value.", vbInformation, App.Title
'        txtfields = ""
'        txtfields.SetFocus
'        Exit Sub
'    End If
'    cmdsave.Enabled = True
'End If
End Sub

Private Sub MSFlexGrid1_SelChange()
   MSFlexGrid1_Click
End Sub

Private Sub txtfieldsFull_Change()
  If Not IsNumeric(txtfieldsFull) Then
         txtfieldsFull = ""
  End If
End Sub

Private Sub txtfieldsFull_GotFocus()
  txtfieldsFull.SelStart = 0
  txtfieldsFull.SelLength = Len(txtfieldsFull)
End Sub

Private Sub txtfieldsPass_Change()
  If Not IsNumeric(txtfieldsPass) Then
     txtfieldsPass = ""
  End If
End Sub

Private Sub txtfieldsPass_GotFocus()
  txtfieldsPass.SelStart = 0
  txtfieldsPass.SelLength = Len(txtfieldsPass)
End Sub
