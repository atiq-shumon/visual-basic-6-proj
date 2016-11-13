VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   5400
      Width           =   9765
      Begin VB.CommandButton cmdAssignTeacher 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Caption         =   "Assign Teacher"
         Height          =   345
         Left            =   30
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1305
      End
      Begin VB.TextBox txttrackid 
         Height          =   285
         Left            =   3780
         TabIndex        =   23
         Text            =   "0"
         Top             =   210
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H8000000C&
         Caption         =   "Update"
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
         TabIndex        =   14
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   4
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
         TabIndex        =   10
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
      TabIndex        =   6
      Top             =   0
      Width           =   9705
      Begin VB.Label Label4 
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
         ForeColor       =   &H00CEF0F7&
         Height          =   285
         Left            =   3000
         TabIndex        =   21
         Top             =   150
         Width           =   2910
      End
      Begin VB.Image Image1 
         Height          =   1020
         Left            =   -30
         Picture         =   "frmsubjectinfo_sub_teacher1.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   9705
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   1485
      Left            =   0
      TabIndex        =   5
      Top             =   810
      Width           =   9765
      Begin VB.TextBox txtfields 
         BackColor       =   &H00CEF0F7&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   4140
         TabIndex        =   18
         Top             =   270
         Width           =   5475
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmsubjectinfo_sub_teacher1.frx":CEA5
         Left            =   1650
         List            =   "frmsubjectinfo_sub_teacher1.frx":CEA7
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select Class"
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00CEF0F7&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   4140
         TabIndex        =   15
         Top             =   690
         Width           =   5475
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmsubjectinfo_sub_teacher1.frx":CEA9
         Left            =   1650
         List            =   "frmsubjectinfo_sub_teacher1.frx":CEAB
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select Class"
         Top             =   690
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   1
         Left            =   4140
         MaxLength       =   100
         TabIndex        =   3
         ToolTipText     =   "Select Subject"
         Top             =   1125
         Width           =   5475
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   0
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1125
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Title"
         Height          =   195
         Index           =   2
         Left            =   2970
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
         Left            =   2970
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
         Left            =   180
         TabIndex        =   13
         Top             =   750
         Width           =   1395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Name"
         Height          =   195
         Index           =   0
         Left            =   2970
         TabIndex        =   8
         Top             =   1170
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Code(Sub)"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   1170
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3135
      Left            =   0
      TabIndex        =   22
      Top             =   2280
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   5530
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

If Len(cmdteacher.Text) = 0 Then
    MsgBox "Please select a teacher Id.", vbInformation, App.Title
    cmdteacher.SetFocus
    Exit Sub
End If

Dim rs1 As New ADODB.Recordset

Set rs1 = getdata("SELECT Sub_code From Subject_Info_sub where M_code='" & Trim(Combo1.Text) & "' and Sub_code='" & Trim(txtfields(0).Text) & "' and class_code='" & Trim(Combo2.Text) & "'")
If rs1.EOF Then
   MsgBox "No Such Subject exists..Please verify.", vbInformation, cmp
   Exit Sub
End If


If MsgBox("Are you sure to Delete ?", vbInformation + vbYesNo + vbDefaultButton1, cmp) = vbYes Then
        Dim rs As New ADODB.Recordset
        Dim cmd As New ADODB.Command
        Dim con As New ADODB.connection
        con.Open GConnString
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SubjectInformation_SUB"
        cmd(1) = "D"
        cmd(2) = Trim(Combo1.Text)
        cmd(3) = Trim(txtfields(0))
        cmd(4) = Trim(Combo2.Text)
        cmd(5) = Trim(txtfields(1))
        cmd(6) = Trim(cmdteacher.Text)
        cmd(7) = Trim(soft_user)
        cmd(8) = Format(Date, "dd mmm yyyy")
        cmd(9) = Val(txttrackid)
        cmd.Execute
        MsgBox "Deleted successfully.", vbInformation, cmp
        cmdnew.SetFocus
        Call ShowFlexData
End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdnew_Click()
If Len(Combo1.Text) = 0 Then Exit Sub
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
'con.Open connectionstring.GConnString
con.Open GConnString
cmd.ActiveConnection = con
Set rs = getdata("select max(cast(Sub_code as int))+1 from Subject_Info_sub where Class_code ='" & Mid(Trim(Combo2.Text), 1, 5) & "'")
If Not rs.EOF Then
    txtfields(0) = IIf(IsNull(rs(0)) = True, "00001", Format(rs(0), "00000"))
Else
    txtfields(0) = "00001"
End If

For i = 1 To 1
    txtfields(i) = ""
Next

'cmbSubjectType.Text = " "
cmdAssignTeacher.Enabled = False
txtfields(1).SetFocus
cmdsave.Enabled = True
End Sub

Private Sub cmdSAVE_Click()
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



If Len(cmdteacher.Text) = 0 Then
    MsgBox "Please select a teacher Id.", vbInformation, App.Title
    cmdteacher.SetFocus
    Exit Sub
End If



Dim rs1 As New ADODB.Recordset

Set rs1 = getdata("SELECT Sub_code From Subject_Info_sub where  Sub_code='" & Trim(txtfields(0).Text) & "' and class_code='" & Trim(Combo2.Text) & "'")
If Not rs1.EOF Then
   MsgBox "Same Subject already exists,Please try another", vbInformation, cmp
   cmdnew.SetFocus
   Exit Sub
End If

Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "SubjectInformation_SUB"
cmd(1) = "S"
cmd(2) = Trim(Combo1.Text)
cmd(3) = Trim(txtfields(0))
cmd(4) = Trim(Combo2.Text)
cmd(5) = Trim(txtfields(1))
cmd(6) = Trim(cmdteacher.Text)
cmd(7) = Trim(soft_user)
cmd(8) = Format(Date, "dd mmm yyyy")
cmd(9) = Val(txttrackid)
cmd.Execute
MsgBox "Saved successfully.", vbInformation, cmp
cmdnew.SetFocus
Call ShowFlexData
cmdsave.Enabled = False
End Sub

Private Sub cmdteacher_Click()
     load_teacher_title
End Sub
Private Sub load_teacher_title()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(cmdteacher) & "'")
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

If Len(cmdteacher.Text) = 0 Then
    MsgBox "Please select a teacher Id.", vbInformation, App.Title
    cmdteacher.SetFocus
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
cmd(6) = Trim(cmdteacher.Text)
cmd(7) = Trim(soft_user)
cmd(8) = Format(Date, "dd mmm yyyy")
cmd(9) = Val(txttrackid)
cmd.Execute
MsgBox "Updated successfully.", vbInformation, cmp
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
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Set rs1 = getdata("Select M_code from SubjectInfoMain")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo1.AddItem rs1(0)
        rs1.MoveNext
    Loop
    Combo1.AddItem (" ")
'    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End If
With MSFlexGrid1
    .Rows = 1
    .Cols = 6
    .Col = 0: .Text = " Class ID#"
    .Col = 1: .Text = " Main Code #"
    .Col = 2: .Text = " Sub Course"
    .Col = 3: .Text = " Title"
    .Col = 4: .Text = " Teacher ID "
    .Col = 5: .Text = " Trackid "
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 2000
    .ColWidth(3) = 5000
    .ColWidth(4) = 1000
    .ColWidth(5) = 0
    
End With
laod_class
load_teacher
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title

End Sub
Private Sub load_teacher()
  cmdteacher.Clear
  Dim rs As New ADODB.Recordset
  Set rs = getdata("SELECT Emp_id  FROM  Emp_Per_Info")
  If Not rs.EOF Then
     Do Until rs.EOF
       cmdteacher.AddItem Trim(rs(0))
       rs.MoveNext
     Loop
   End If
     
End Sub
Private Sub laod_class()
 Dim rs As New ADODB.Recordset
 Set rs = getdata("Select ClassId from ClassInfo")
 If Not rs.EOF Then
    Do Until rs.EOF
        Combo2.AddItem rs(0)
        rs.MoveNext
    Loop
  End If
End Sub
Private Sub load_class_title()
  Dim rs As New ADODB.Recordset
 Set rs = getdata("Select ClassName from ClassInfo where classid='" & Trim(Combo2.Text) & "'")
 If Not rs.EOF Then
    txtfields(3).Text = rs(0)
  End If
End Sub


Private Sub MSFlexGrid1_Click()
On Error Resume Next
If MSFlexGrid1.Row > 0 Then
   'Combo2.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
   'Combo1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
   txtfields(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
   txtfields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
   cmdteacher.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4))
   txttrackid.Text = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
   cmdAssignTeacher.Enabled = True
End If


Exit Sub
errdes:
   MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub

Private Sub txtfields_Change(Index As Integer)
            Select Case Index
                  
                   Case 4
                        If Not IsNumeric(txtfields(4).Text) Or Len(txtfields(4).Text) > 3 Then
                               txtfields(4) = ""
                         End If
                 End Select
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
Set rs1 = getdata("SELECT class_code,M_code,Sub_code,Sub_title,Teacher_id,trackid From Subject_Info_sub_history where Class_code='" & Trim(Combo2.Text) & "' and m_code='" & Trim(Combo1.Text) & "' order by trackid desc")
If Not rs1.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs1.EOF
            MSFlexGrid1.Rows = i + 1
                .TextMatrix(i, 0) = rs1!class_code
                .TextMatrix(i, 1) = rs1!M_code
                .TextMatrix(i, 2) = rs1!Sub_code
                .TextMatrix(i, 3) = rs1!Sub_title
               .TextMatrix(i, 4) = Trim(rs1!Teacher_id)
               .TextMatrix(i, 5) = Trim(rs1!trackid)
                i = i + 1
            rs1.MoveNext
        Loop
        
    End With
 Else
        MSFlexGrid1.Rows = 1
        
 End If

End Sub
