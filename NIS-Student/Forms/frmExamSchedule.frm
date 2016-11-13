VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExamSchedule 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000C&
      Caption         =   "Print"
      Height          =   375
      Left            =   6660
      TabIndex        =   27
      ToolTipText     =   "Click to insert new information"
      Top             =   6990
      Width           =   945
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000C&
      Caption         =   "Edit"
      Height          =   375
      Left            =   4740
      TabIndex        =   26
      ToolTipText     =   "Click to Save"
      Top             =   6990
      Width           =   945
   End
   Begin VB.Frame Frame4 
      Height          =   1065
      Left            =   0
      TabIndex        =   21
      Top             =   1980
      Width           =   8625
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   5490
         TabIndex        =   7
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   609
         _Version        =   393216
         Format          =   68222978
         CurrentDate     =   38838
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   1320
         TabIndex        =   6
         Top             =   630
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   609
         _Version        =   393216
         Format          =   68222978
         CurrentDate     =   38838
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5490
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   2775
      End
      Begin MSMask.MaskEdBox MaskEdStartDate 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         ToolTipText     =   "Insert Exam Starting "
         Top             =   210
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Time "
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   25
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Time"
         Height          =   195
         Index           =   1
         Left            =   4260
         TabIndex        =   24
         Top             =   690
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Date"
         Height          =   195
         Left            =   90
         TabIndex        =   23
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Title"
         Height          =   195
         Index           =   0
         Left            =   4230
         TabIndex        =   22
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   30
      TabIndex        =   20
      Text            =   "0"
      Top             =   7020
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   375
      Left            =   7650
      TabIndex        =   11
      ToolTipText     =   "Click to Exit"
      Top             =   6990
      Width           =   945
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   375
      Left            =   5700
      TabIndex        =   10
      ToolTipText     =   "Click to Delete"
      Top             =   6990
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   375
      Left            =   3750
      TabIndex        =   8
      ToolTipText     =   "Click to Save"
      Top             =   6990
      Width           =   945
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   375
      Left            =   2790
      TabIndex        =   9
      ToolTipText     =   "Click to insert new information"
      Top             =   6990
      Width           =   945
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   1380
      Width           =   8625
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5490
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   210
         Width           =   2745
      End
      Begin VB.ComboBox ComboExamName 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select Exam Name"
         Top             =   180
         Width           =   2475
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Name"
         Height          =   195
         Index           =   1
         Left            =   4230
         TabIndex        =   19
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Term Name"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   18
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   8595
      TabIndex        =   13
      Top             =   0
      Width           =   8655
      Begin VB.Frame Frame2 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   14
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Routine Entry"
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
         Left            =   2730
         TabIndex        =   28
         Top             =   180
         Width           =   2280
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   -30
         Picture         =   "frmExamSchedule.frx":0000
         Stretch         =   -1  'True
         Top             =   -60
         Width           =   8895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      TabIndex        =   12
      Top             =   750
      Width           =   8625
      Begin VB.ComboBox ComboYear 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select Year"
         Top             =   180
         Width           =   2475
      End
      Begin VB.ComboBox ComboClass 
         Height          =   315
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select Calss"
         Top             =   180
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         Height          =   195
         Left            =   4230
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   330
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3885
      Left            =   0
      TabIndex        =   29
      Top             =   3030
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   6853
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
      Left            =   2730
      Top             =   6960
      Width           =   5895
   End
End
Attribute VB_Name = "frmExamSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExamIDA As Integer
Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdDelete_Click()
'On Error GoTo ErrDes
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
Dim rs As New ADODB.Recordset
cmd.ActiveConnection = con

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
  
 
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("select Sub_id from examschedule where ClassId='" & Mid(ComboClass, 1, 5) & "' and  ExamYear='" & Trim(ComboYear) & "' and  ExamId ='" & Mid(ComboExamName, 1, 2) & "' and  ExamTypeID='" & Mid(Combo1, 1, 2) & "' and Sub_id='" & Trim(Mid(Combo2, 1, 5)) & "'")
If rs1.EOF Then
  MsgBox "No such subject  exists ,Please verify...", vbInformation, cmp
  Combo2.SetFocus
  Exit Sub
End If

    Dim check As String
If MsgBox("Are you sure to Delete?", vbDefaultButton1 + vbInformation + vbYesNo, cmp) = vbYes Then
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ExaminatioSchedule"
    cmd(1) = "d"
    cmd(2) = Trim(Text1)
    cmd(3) = Mid(ComboClass, 1, 5)
    cmd(4) = Trim(ComboYear)
    cmd(5) = Mid(ComboExamName, 1, 2)
    cmd(6) = Mid(Combo1, 1, 2)
    cmd(7) = Format(Date, "dd mm yy")
    cmd(8) = 1
    cmd(9) = Trim(Mid(Combo2, 1, 5))
    cmd(10) = DTPicker1.Value
    cmd(11) = DTPicker1.Value
    cmd(12) = soft_user
    cmd.Execute
    MsgBox "Deleted Successfully.", vbInformation, cmp
    cmdnew.SetFocus
  
    Call ShowFlexData
End If
Exit Sub
errdes:
    MsgBox Err.Description, vbInformation, "Student Management System"
End Sub

Private Sub CmdEdit_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
Dim rs As New ADODB.Recordset
cmd.ActiveConnection = con

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
  
 
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("select Sub_id from examschedule where ClassId='" & Mid(ComboClass, 1, 5) & "' and  ExamYear='" & Trim(ComboYear) & "' and  ExamId ='" & Mid(ComboExamName, 1, 2) & "' and  ExamTypeID='" & Mid(Combo1, 1, 2) & "' and Sub_id='" & Trim(Mid(Combo2, 1, 5)) & "'")
If rs1.EOF Then
  MsgBox "No such subject  exists ,Please verify...", vbInformation, cmp
  Combo2.SetFocus
  Exit Sub
End If

    Dim check As String
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ExaminatioSchedule"
    cmd(1) = "u"
    cmd(2) = Trim(Text1)
    cmd(3) = Mid(ComboClass, 1, 5)
    cmd(4) = Trim(ComboYear)
    cmd(5) = Mid(ComboExamName, 1, 2)
    cmd(6) = Mid(Combo1, 1, 2)
    cmd(7) = Format(Date, "dd mm yy")
    cmd(8) = 1
    cmd(9) = Trim(Mid(Combo2, 1, 5))
    cmd(10) = Format(DTPicker1.Value, "hh:mm:ss AM/PM")
    cmd(11) = Format(DTPicker2.Value, "hh:mm:ss AM/PM")
    cmd(12) = soft_user
    cmd.Execute
    MsgBox "Updated Successfully.", vbInformation, cmp
    cmdnew.SetFocus
  
    Call ShowFlexData
Exit Sub
errdes:
    MsgBox Err.Description, vbInformation, "Student Management System"

End Sub

Private Sub cmdnew_Click()
MaskEdStartDate.Text = "__/__/__"
txtfields = ""
ExamIDA = 0

ComboClass.SetFocus
End Sub
Private Sub cmdSAVE_Click()
'On Error GoTo ErrDes
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
Dim rs As New ADODB.Recordset
cmd.ActiveConnection = con

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
    If Len(Combo2) = 0 Then
        MsgBox "Please Enter Subject Name.", vbCritical, App.Title
        Combo2.SetFocus
        Exit Sub
    End If
  
  If MaskEdStartDate = "__/__/__" Then
     MsgBox " Please put a valid exam date", vbiformation, cmp
     MaskEdStartDate.SetFocus
     Exit Sub
 End If

Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("select Sub_id from examschedule where ClassId='" & Mid(ComboClass, 1, 5) & "' and  ExamYear='" & Trim(ComboYear) & "' and  ExamId ='" & Mid(ComboExamName, 1, 2) & "' and  ExamTypeID='" & Mid(Combo1, 1, 2) & "' and Sub_id='" & Trim(Mid(Combo2, 1, 5)) & "'")
If Not rs1.EOF Then
  MsgBox "Same subject already exists ,Please verify...", vbInformation, cmp
  Combo2.SetFocus
  Exit Sub
End If

    Dim check As String
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ExaminatioSchedule"
    cmd(1) = "s"
    cmd(2) = Trim(Text1)
    cmd(3) = Mid(ComboClass, 1, 5)
    cmd(4) = Trim(ComboYear)
    cmd(5) = Mid(ComboExamName, 1, 2)
    cmd(6) = Mid(Combo1, 1, 2)
    cmd(7) = Format(MaskEdStartDate, "dd mm yy")
    cmd(8) = 1
    cmd(9) = Trim(Mid(Combo2, 1, 5))
    cmd(10) = Format(DTPicker1.Value, "hh:mm:ss")
    cmd(11) = Format(DTPicker2.Value, "hh:mm:ss")
    cmd(12) = soft_user
    cmd.Execute
    MsgBox "Saved Successfully.", vbInformation, "Student Management System"
    cmdnew.SetFocus
  
    Call ShowFlexData
Exit Sub
errdes:
    MsgBox Err.Description, vbInformation, "Student Management System"
End Sub

Private Sub Combo1_Click()
   ShowFlexData
   MSFlexGrid1_Click
End Sub

Private Sub ComboClass_Click()
 Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT  Sub_code,Sub_title From Subject_Info_sub WHERE Class_code = '" & Mid(ComboClass, 1, 5) & "'")
Combo2.Clear
If Not rs.EOF Then
    Do Until rs.EOF
        Combo2.AddItem rs!Sub_code & "-" & rs!Sub_title
        rs.MoveNext
    Loop

End If
   ShowFlexData
End Sub

Private Sub ComboClass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ShowFlexData
End If
End Sub

Private Sub ComboClass_LostFocus()
ShowFlexData
End Sub



Private Sub ComboExamName_click()
   Combo1.SetFocus
   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select Exam_code,Exam_title from exam_setup where Group_code='" & Mid(Trim(ComboExamName.Text), 1, 2) & "'")
  
  If Not rs.EOF Then
    Combo1.Clear
    Do Until rs.EOF
        Combo1.AddItem rs(0) + " - " + rs(1)
        rs.MoveNext
    Loop
 End If
ShowFlexData
End Sub

Private Sub ComboYear_Click()
  ShowFlexData
End Sub

Private Sub ComboYear_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    ShowFlexData
End If
End Sub

Private Sub Command1_Click()
       rptMode = 4
       Screen.MousePointer = vbHourglass
       frmViewer.Show 1
  
End Sub





Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    DTPicker2.SetFocus
  End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     DTPicker2.SetFocus
  End If
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     cmdSave.SetFocus
  End If
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
ExamIDA = 0
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("Select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        ComboClass.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop

End If
Set rs1 = getdata("Select ETypeId,ETypeName from ExamTypeInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        ComboExamName.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop

End If
Dim IL As Integer
For IL = 2000 To 2020
   ComboYear.AddItem (IL)
Next IL
ComboYear.Text = Format(Date, "YYYY")
With MSFlexGrid1
    .Rows = 1
    .Cols = 6
    .Col = 0: .Text = "Serial#"
    .Col = 1: .Text = "Date"
    .Col = 2: .Text = "Subject Id"
    .Col = 3: .Text = "Title"
    .Col = 4: .Text = "Start"
    .Col = 5: .Text = "End"
    .ColWidth(0) = 500
    .ColWidth(1) = 900
    .ColWidth(2) = 1200
    .ColWidth(3) = 3500
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    

   
End With

Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title

End Sub
Private Sub ShowFlexData()
'On Error GoTo ErrDes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT   a.serial_no, " + _
" a.ExamDate,a.Sub_id,(select Sub_title from subject_info_sub where Class_code=a.ClassId and Sub_code=a.Sub_id ) as sub_title,a.ExamStartTime,a.ExamEndTime  FROM    ExamSchedule a where  a.ExamYear='" & Trim(ComboYear.Text) & "' and a.ClassId='" & Mid(ComboClass, 1, 5) & "' and examid='" & Mid(ComboExamName, 1, 2) & "' and ExamTypeID='" & Mid(Combo1, 1, 2) & "' order by a.serial_no")
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1

                .TextMatrix(i, 0) = "" & rs!serial_no
                .TextMatrix(i, 1) = "" & rs!ExamDate
                .TextMatrix(i, 2) = "" & rs!Sub_id
                .TextMatrix(i, 3) = "" & rs!Sub_title
                .TextMatrix(i, 4) = "" & rs!ExamStartTime
                .TextMatrix(i, 5) = "" & rs!ExamEndTime



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

Private Sub MaskEdStartDate_GotFocus()
  MaskEdStartDate.SelStart = 0
  MaskEdStartDate.SelLength = Len(MaskEdStartDate.Text)
End Sub

Private Sub MaskEdStartDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MaskEdStartDate <> "__/__/__" Then
            If Check_ValidDate(MaskEdStartDate) = False Then
                MaskEdStartDate.SetFocus
                Exit Sub
            End If
    End If

    
End If

End Sub
Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
If MSFlexGrid1.Rows > 1 Then
    Text1 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    MaskEdStartDate = Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1), "DD/MM/YY")
    Combo2 = Trim(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)) + "-" + Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)))
    DTPicker1.Value = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
    DTPicker2.Value = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
'    Combo2 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
End If

Exit Sub
errdes:
  MsgBox Err.Description, vbInformation, App.Title

End Sub


Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub
