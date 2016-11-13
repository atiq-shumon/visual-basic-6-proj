VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmsubjectinfo_sub_teacher 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   5400
      Width           =   9765
      Begin VB.TextBox txttrackid 
         Height          =   285
         Left            =   3780
         TabIndex        =   23
         Top             =   210
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H8000000C&
         Caption         =   "Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6510
         TabIndex        =   4
         ToolTipText     =   "Click to Update information"
         Top             =   270
         Width           =   975
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H8000000C&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8490
         TabIndex        =   6
         ToolTipText     =   "Click to Exit"
         Top             =   270
         Width           =   975
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H8000000C&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7500
         TabIndex        =   5
         ToolTipText     =   "Click to Delete"
         Top             =   270
         Width           =   975
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H8000000C&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5490
         TabIndex        =   2
         ToolTipText     =   "Click to Save"
         Top             =   270
         Width           =   975
      End
      Begin VB.CommandButton cmdnew 
         BackColor       =   &H8000000C&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4500
         TabIndex        =   3
         ToolTipText     =   "Click to insert new information"
         Top             =   270
         Width           =   975
      End
      Begin VB.Shape Shape1 
         Height          =   465
         Left            =   4470
         Top             =   240
         Width           =   5025
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   795
      Left            =   30
      ScaleHeight     =   735
      ScaleWidth      =   9645
      TabIndex        =   10
      Top             =   0
      Width           =   9705
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Teacher Information"
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
         Left            =   3000
         TabIndex        =   21
         Top             =   150
         Width           =   3225
      End
      Begin VB.Image Image1 
         Height          =   1020
         Left            =   -30
         Picture         =   "frmsubjectinfo_sub_teacher.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   9705
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   1845
      Left            =   0
      TabIndex        =   9
      Top             =   810
      Width           =   9765
      Begin VB.TextBox txtfields 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   6
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   690
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00CEF0F7&
         Height          =   285
         Index           =   5
         Left            =   3990
         MaxLength       =   100
         TabIndex        =   24
         ToolTipText     =   "Select Subject"
         Top             =   1470
         Width           =   3315
      End
      Begin VB.ComboBox cmdTeacher 
         Height          =   315
         ItemData        =   "frmsubjectinfo_sub_teacher.frx":CEA5
         Left            =   1620
         List            =   "frmsubjectinfo_sub_teacher.frx":CEA7
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select Class"
         Top             =   1470
         Width           =   1095
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00CEF0F7&
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
         Height          =   315
         Index           =   3
         Left            =   3990
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   300
         Width           =   5475
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00CEF0F7&
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
         Height          =   315
         Index           =   2
         Left            =   3990
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   690
         Width           =   5475
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00CEF0F7&
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
         Height          =   285
         Index           =   1
         Left            =   3990
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   8
         ToolTipText     =   "Select Subject"
         Top             =   1065
         Width           =   5475
      End
      Begin VB.TextBox txtfields 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1065
         Width           =   1065
      End
      Begin MSMask.MaskEdBox MaskEdBoxDate 
         Height          =   315
         Index           =   1
         Left            =   8400
         TabIndex        =   1
         ToolTipText     =   "Insert  To  Date"
         Top             =   1470
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         Height          =   195
         Index           =   4
         Left            =   7350
         TabIndex        =   27
         Top             =   1530
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teacher  Name"
         Height          =   195
         Index           =   3
         Left            =   2850
         TabIndex        =   26
         Top             =   1500
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teacher  Id"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   25
         Top             =   1500
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Title"
         Height          =   195
         Index           =   2
         Left            =   2850
         TabIndex        =   19
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Id"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Name"
         Height          =   195
         Index           =   1
         Left            =   2850
         TabIndex        =   16
         Top             =   750
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Code(Main)"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Name"
         Height          =   195
         Index           =   0
         Left            =   2850
         TabIndex        =   12
         Top             =   1110
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Code(Sub)"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   1110
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2775
      Left            =   0
      TabIndex        =   22
      Top             =   2640
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   4895
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Information (Sub)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   2910
   End
End
Attribute VB_Name = "frmsubjectinfo_sub_teacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbSubjectType_Change()
 
      cmbSubjectType.Text = "Compulsory"
End Sub
Private Sub cmbSubjectType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Combo2.SetFocus
   End If
End Sub
Private Sub cmdAssignTeacher_Click()
  frmsubjectinfo_sub_teacher.Show 1
End Sub
Private Sub cmdDelete_Click()

If Len(txttrackid) = 0 Then
    MsgBox "Please Select a row from the grid below.", vbInformation, App.Title
    MSFlexGrid1.SetFocus
    Exit Sub
End If

If MsgBox("Are you sure to Delete ?", vbInformation + vbYesNo + vbDefaultButton1, cmp) = vbYes Then
        Dim rs As New ADODB.Recordset
        Dim cmd As New ADODB.Command
        Dim con As New ADODB.connection
        con.Open GConnString
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SubjectInformation_SUB_Teacher"
        cmd(1) = "D"
        cmd(2) = Trim(txtfields(6))
        cmd(3) = Trim(txtfields(0))
        cmd(4) = Trim(txtfields(4))
        cmd(5) = Trim(cmdTeacher.Text)
        cmd(6) = Format(Date, "dd mmm yyyy")
        cmd(7) = Trim(soft_user)
        cmd(8) = Format(Date, "dd mmm yyyy")
        cmd(9) = Val(txttrackid)
        cmd.Execute

        MsgBox "Deleted successfully.", vbInformation, cmp
        cmdnew.SetFocus
        Call ShowFlexData
        txttrackid = ""
End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdnew_Click()

MaskEdBoxDate(1) = "__/__/__"
txttrackid = ""

cmdTeacher.SetFocus
cmdSave.Enabled = True
End Sub

Private Sub cmdSAVE_Click()
If Len(cmdTeacher.Text) = 0 Then
    MsgBox "Please Select a Teacher Id", vbInformation, App.Title
    cmdTeacher.SetFocus
    Exit Sub
End If
If MaskEdBoxDate(1).Text = "__/__/__" Then
   MsgBox "Please Enter an Effective Date fo assignment", vbInformation, cmp
   MaskEdBoxDate(1).SetFocus
   Exit Sub
End If




Dim rs1 As New ADODB.Recordset

Set rs1 = getdata("SELECT Sub_code From Subject_Info_sub_history where  Sub_code='" & Trim(txtfields(0).Text) & "' and class_code='" & Trim(txtfields(4).Text) & "' and M_code='" & Trim(txtfields(6).Text) & "' and Teacher_id ='" & Trim(cmdTeacher.Text) & "' and Effective_Date='" & MaskEdBoxDate(1).Text & "'")
If Not rs1.EOF Then
   MsgBox "Same Teacher already exists at same date,Please try another", vbInformation, cmp
   cmdnew.SetFocus
   Exit Sub
End If

Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "SubjectInformation_SUB_Teacher"
cmd(1) = "S"
cmd(2) = Trim(txtfields(6))
cmd(3) = Trim(txtfields(0))
cmd(4) = Trim(txtfields(4))
cmd(5) = Trim(cmdTeacher.Text)
cmd(6) = Format(MaskEdBoxDate(1).Text, "dd mmm yyyy")
cmd(7) = Trim(soft_user)
cmd(8) = Format(Date, "dd mmm yyyy")
cmd(9) = Val(txttrackid)
cmd.Execute
MsgBox "Saved successfully.", vbInformation, cmp
cmdnew.SetFocus
txttrackid = ""
Call ShowFlexData
cmdSave.Enabled = False
End Sub

Private Sub cmdteacher_Click()
     load_teacher_title
End Sub
Private Sub load_teacher_title()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(cmdTeacher) & "'")
   If Not rs.EOF Then
     txtfields(5).Text = "" & rs!name
   End If
   
End Sub
Private Sub cmdUpdate_Click()
If Len(Combo1.Text) = 0 Then
    MsgBox "Please Select a Main code of Subject", vbInformation, App.Title
    Combo1.SetFocus
    Exit Sub
End If
If Len(txtfields(0)) = 0 Then
    MsgBox "Please Enter a subject Id.", vbInformation, App.Title
    cmdnew.SetFocus
    Exit Sub
End If

If Len(txtfields(1)) = 0 Then
    
    MsgBox "Please Enter subject Name.", vbInformation, App.Title
    txtfields(1).SetFocus
    Exit Sub
End If

If Len(cmdTeacher.Text) = 0 Then
    MsgBox "Please select a teacher Id.", vbInformation, App.Title
    cmdTeacher.SetFocus
    Exit Sub
End If

Dim rs1 As New ADODB.Recordset

Set rs1 = getdata("SELECT Sub_code From Subject_Info_sub where M_code='" & Trim(Combo1.Text) & "' and Sub_code='" & Trim(txtfields(0).Text) & "' and class_code='" & Trim(Combo2.Text) & "'")
If rs1.EOF Then
   MsgBox "No Such Subject exists..Please verify.", vbInformation, cmp
   Exit Sub
End If

Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "SubjectInformation_SUB"
cmd(1) = "U"
cmd(2) = Trim(Combo1.Text)
cmd(3) = Trim(txtfields(0))
cmd(4) = Trim(Combo2.Text)
cmd(5) = Trim(txtfields(1))
cmd(6) = Trim(cmdTeacher.Text)
cmd(7) = Trim(soft_user)
cmd(8) = Format(Date, "dd mmm yyyy")
cmd(9) = Val(txttrackid)
cmd.Execute
MsgBox "Updated successfully.", vbInformation, cmp
txttrackid = ""
cmdnew.SetFocus
Call ShowFlexData
End Sub
Private Sub Combo1_Click()
  If Len(Combo1) <> 0 Then
    Call sub_name
    Call ShowFlexData
  End If
End Sub
Private Sub sub_name()
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT M_title From SubjectInfoMain where M_code='" & Trim(Combo1.Text) & "'")
If Not rs.EOF Then
  txtfields(2).Text = "" & rs!M_title
End If
End Sub
Private Sub Combo1_LostFocus()
    Call cmdnew_Click
End Sub

Private Sub Combo2_Click()
  load_class_title
  Call ShowFlexData
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

With MSFlexGrid1
    .Rows = 1
    .Cols = 8
    .Col = 0: .Text = " Class ID#"
    .Col = 1: .Text = " Main Code #"
    .Col = 2: .Text = " Course"
    .Col = 3: .Text = " Title"
    .Col = 4: .Text = " Teacher ID "
    .Col = 5: .Text = " Teacher Name"
    .Col = 6: .Text = " Effective Date "
    .Col = 7: .Text = " Trackid "
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 700
    .ColWidth(3) = 2000
    .ColWidth(4) = 1000
    .ColWidth(5) = 4500
    .ColWidth(6) = 1500
    .ColWidth(7) = 0
    
End With
frmsubjectinfo_sub_teacher.txtfields(4) = frmsubjectinfo_sub.Combo2.Text
load_class_title
frmsubjectinfo_sub_teacher.txtfields(6) = frmsubjectinfo_sub.Combo1.Text
load_main_sub_title
frmsubjectinfo_sub_teacher.txtfields(0) = frmsubjectinfo_sub.txtfields(0).Text
load_main_sub_subject_title
load_teacher
ShowFlexData
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title

End Sub
Private Sub load_teacher()
  cmdTeacher.Clear
  Dim rs As New ADODB.Recordset
  Set rs = getdata("SELECT Emp_id  FROM  Emp_Per_Info")
  If Not rs.EOF Then
     Do Until rs.EOF
       cmdTeacher.AddItem Trim(rs(0))
       rs.MoveNext
     Loop
   End If
     
End Sub

Private Sub load_class_title()
  Dim rs As New ADODB.Recordset
 Set rs = getdata("Select ClassName from ClassInfo where classid='" & Trim(txtfields(4).Text) & "'")
 If Not rs.EOF Then
    txtfields(3).Text = rs(0)
  End If
End Sub
Private Sub load_main_sub_title()
  Dim rs As New ADODB.Recordset
 Set rs = getdata("Select M_title from SubjectInfoMain where M_code='" & Trim(txtfields(6).Text) & "'")
 If Not rs.EOF Then
    txtfields(2).Text = rs(0)
  End If
End Sub
Private Sub load_main_sub_subject_title()
  Dim rs As New ADODB.Recordset
 Set rs = getdata("Select Sub_title from Subject_Info_sub where M_code='" & Trim(txtfields(6).Text) & "' and Class_code='" & Trim(txtfields(4).Text) & "' and Sub_code='" & Trim(txtfields(0).Text) & "'")
 If Not rs.EOF Then
    txtfields(1).Text = rs(0)
  End If
End Sub

Private Sub MaskEdBoxDate_GotFocus(Index As Integer)
    Select Case Index
             Case Index
             MaskEdBoxDate(Index).SelStart = 0
             MaskEdBoxDate(Index).SelLength = Len(MaskEdBoxDate(Index))
      End Select
End Sub

Private Sub MSFlexGrid1_Click()
On Error Resume Next
If MSFlexGrid1.Row > 0 Then
   'Combo2.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
   'Combo1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
   txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
   txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
   cmdTeacher.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4))
   If Len(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)) = 0 Then
      MaskEdBoxDate(1).Text = "__/__/__"
   Else
     MaskEdBoxDate(1).Text = Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6), "dd/mm/yy")
   End If
   
   txttrackid.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7))
   
End If


Exit Sub
errdes:
   MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
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
Private Sub txtfields_LostFocus(Index As Integer)
'txtfields(0) = Format(txtfields(0), "00000")
'Dim rs As New adodb.Recordset
'
'Select Case Index
'    Case 0
'        If Len(Trim(txtfields(0))) = 0 Then Exit Sub
'
'            txtfields(0) = Format(txtfields(0), "00000")
'
'            Set rs = GetData("SELECT * from SubjectInfo WHERE (SubjectID = '" & txtfields(0) & "') and ClassID= '" & Combo1.Text & "'")
'                 If Not rs.EOF Then
'                        Combo1.Text = rs!classId
'                        txtfields(1) = rs!SubjectDsc
'                        txtfields(2) = rs!totalmarks
'                        Combo2.Text = rs!SubjectUnit
'
'                End If
'
'    Case 2
'        Dim SubMarks As Double
'        If Len(Trim(txtfields(2))) = 0 Then Exit Sub
'        If IsNumeric(txtfields(2)) = False Then
'            MsgBox "Please Enter Numeric Value.", vbInformation, App.Title
'            txtfields(2) = ""
'            txtfields(2).SetFocus
'            Exit Sub
'        End If
'End Select
End Sub
Private Sub ShowFlexData()
Dim rs1 As New ADODB.Recordset
Dim sub_title_rs As New ADODB.Recordset
Dim teacher_rs As New ADODB.Recordset
Set rs1 = getdata("SELECT class_code,M_code,Sub_code,Teacher_id,trackid,Effective_Date From Subject_Info_sub_history where Class_code='" & Trim(txtfields(4).Text) & "' and m_code='" & Trim(txtfields(6).Text) & "' and Sub_code='" & Trim(txtfields(0).Text) & "' order by trackid desc")
If Not rs1.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs1.EOF
            MSFlexGrid1.Rows = i + 1
                .TextMatrix(i, 0) = rs1!class_code
                .TextMatrix(i, 1) = rs1!M_code
                .TextMatrix(i, 2) = rs1!Sub_code
                 Set sub_title_rs = getdata("select Sub_title from subject_info_sub where Class_code='" & Trim(txtfields(4).Text) & "' and m_code='" & Trim(txtfields(6).Text) & "' and Sub_code='" & Trim(txtfields(0).Text) & "'")
                .TextMatrix(i, 3) = sub_title_rs!Sub_title
               .TextMatrix(i, 4) = Trim(rs1!Teacher_id)
                Set teacher_rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(rs1!Teacher_id) & "'")
                .TextMatrix(i, 5) = teacher_rs(0)
                .ColAlignment(6) = 0
               .TextMatrix(i, 6) = "" & rs1!effective_date
               .TextMatrix(i, 7) = Trim(rs1!trackid)
                i = i + 1
            rs1.MoveNext
        Loop
        
    End With
 Else
        MSFlexGrid1.Rows = 1
        
 End If
Set teacher_rs = Nothing
Set rs = Nothing
Set rs1 = Nothing
Set sub_title_rs = Nothing
End Sub
