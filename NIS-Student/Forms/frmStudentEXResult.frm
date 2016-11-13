VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStudentExResult 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtid 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   540
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   7350
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   375
      Left            =   6390
      TabIndex        =   38
      ToolTipText     =   "Click to Delete"
      Top             =   7320
      Width           =   1035
   End
   Begin VB.TextBox txtSerialSub 
      Height          =   315
      Left            =   210
      TabIndex        =   37
      Text            =   "0"
      Top             =   7350
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txtSerial 
      Height          =   315
      Left            =   930
      TabIndex        =   36
      Text            =   "0"
      Top             =   7350
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   -60
      Picture         =   "frmStudentEXResult.frx":0000
      ScaleHeight     =   690
      ScaleWidth      =   9645
      TabIndex        =   21
      Top             =   0
      Width           =   9705
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   22
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ex. B P. Marks Distribution Entry"
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
         Left            =   2460
         TabIndex        =   23
         Top             =   150
         Width           =   3705
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   -120
         Picture         =   "frmStudentEXResult.frx":C897
         Stretch         =   -1  'True
         Top             =   0
         Width           =   9795
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H000000C0&
      Caption         =   "Print"
      Height          =   375
      Left            =   7440
      TabIndex        =   20
      ToolTipText     =   "Click to Print"
      Top             =   7320
      Width           =   1035
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H8000000C&
      Caption         =   "Edit"
      Height          =   375
      Left            =   5340
      TabIndex        =   19
      ToolTipText     =   "Click to Edit"
      Top             =   7320
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Height          =   6465
      Left            =   0
      TabIndex        =   15
      Top             =   750
      Width           =   9675
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3825
         Left            =   1800
         TabIndex        =   40
         Top             =   2550
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   6747
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         BackColor       =   15005934
         BackColorSel    =   -2147483624
         ForeColorSel    =   16711680
         BackColorBkg    =   15724265
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3300
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   4020
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.ComboBox CboCategory 
         Height          =   315
         Left            =   3390
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   3930
         Visible         =   0   'False
         Width           =   2685
      End
      Begin VB.Frame Frame4 
         Height          =   1065
         Left            =   1800
         TabIndex        =   30
         Top             =   1440
         Width           =   7755
         Begin VB.TextBox txtManners 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   6930
            MaxLength       =   5
            TabIndex        =   12
            Text            =   "0.00"
            Top             =   660
            Width           =   705
         End
         Begin VB.TextBox txtCleanness 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   5430
            MaxLength       =   5
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   660
            Width           =   705
         End
         Begin VB.TextBox txtAttentiveness 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3900
            MaxLength       =   5
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   660
            Width           =   705
         End
         Begin VB.TextBox txtHWMarks 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   2070
            MaxLength       =   5
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   660
            Width           =   705
         End
         Begin VB.TextBox txtfields 
            BackColor       =   &H00CEF0F7&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   6780
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtCWMarks 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00;(0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   810
            MaxLength       =   5
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   660
            Width           =   705
         End
         Begin VB.TextBox lblStdname 
            BackColor       =   &H00CEF0F7&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H009F5620&
            Height          =   285
            Left            =   810
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   4995
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Manners"
            Height          =   195
            Index           =   10
            Left            =   6240
            TabIndex        =   44
            Top             =   690
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cleanness"
            Height          =   195
            Index           =   9
            Left            =   4650
            TabIndex        =   43
            Top             =   690
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Attentiveness"
            Height          =   195
            Index           =   12
            Left            =   2880
            TabIndex        =   42
            Top             =   690
            Width           =   960
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "H.W. "
            Height          =   195
            Index           =   11
            Left            =   1620
            TabIndex        =   41
            Top             =   690
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Roll :"
            Height          =   195
            Left            =   6330
            TabIndex        =   35
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C.W. "
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   34
            Top             =   690
            Width           =   405
         End
         Begin VB.Label Label2 
            Caption         =   "Name :"
            Height          =   225
            Left            =   90
            TabIndex        =   33
            Top             =   270
            Width           =   585
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Exam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   795
         Left            =   1830
         TabIndex        =   26
         Top             =   660
         Width           =   7725
         Begin VB.ComboBox CboExamType 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   390
            Width           =   2295
         End
         Begin VB.ComboBox CboExamID 
            Height          =   315
            Left            =   4950
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   390
            Width           =   2685
         End
         Begin VB.ComboBox cmdAcaYear 
            Height          =   315
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   390
            Width           =   2535
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exam Term"
            Height          =   195
            Index           =   3
            Left            =   2640
            TabIndex        =   29
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exam"
            Height          =   195
            Index           =   4
            Left            =   4980
            TabIndex        =   28
            Top             =   150
            Width           =   510
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Academic Year"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   27
            Top             =   180
            Width           =   1080
         End
      End
      Begin VB.ComboBox CboSubject 
         Height          =   315
         Left            =   5430
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   300
         Width           =   4005
      End
      Begin VB.ComboBox cboClass 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   1755
      End
      Begin VB.Frame Frame7 
         Caption         =   "Student"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   5775
         Left            =   120
         TabIndex        =   18
         Top             =   660
         Width           =   1665
         Begin VB.ListBox List1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5520
            Left            =   90
            TabIndex        =   7
            Top             =   180
            Width           =   1485
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmStudentEXResult.frx":1973C
         Left            =   3960
         List            =   "frmStudentEXResult.frx":19746
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   1485
      End
      Begin VB.ComboBox CboSection 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject "
         Height          =   195
         Index           =   6
         Left            =   5430
         TabIndex        =   25
         Top             =   120
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   24
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section "
         Height          =   195
         Index           =   1
         Left            =   1950
         TabIndex        =   17
         Top             =   90
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Index           =   0
         Left            =   3960
         TabIndex        =   16
         Top             =   120
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   375
      Left            =   4290
      TabIndex        =   13
      ToolTipText     =   "Click to Save"
      Top             =   7320
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   375
      Left            =   8490
      TabIndex        =   14
      ToolTipText     =   "Click to Close"
      Top             =   7320
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   4230
      Top             =   7260
      Width           =   5325
   End
End
Attribute VB_Name = "frmStudentExResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub load_marks()
On Error GoTo err_des

Dim rs As New ADODB.Recordset
 Set rs = getdata("select a.fullmarks,a.passmarks from SubjectMarksDistribution  a where a.ClassId = '" & Mid(Trim(cboClass.Text), 1, 5) & "' and a.SubjectId= '" & Mid(Trim(CboSubject.Text), 1, 5) & "' and a.term_code='" & Mid(Trim(CboExamType.Text), 1, 2) & "' and a.Exam_code='" & Mid(Trim(CboExamID.Text), 1, 2) & "'and a.CategoryID='" & Mid(CboCategory, 1, 5) & "'")

 If Not rs.EOF Then
   txtfullMarks = rs(0)
   txtPassMarks = rs(1)
 End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub CboCategory_Click()
On Error GoTo err_des

get_resultMain
load_marks
LoadStuID
show_grid

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cboClass_Click()
On Error GoTo err_des

 load_Scction
 load_subject
 LoadStuID
 get_resultMain
 LoadStuID
show_grid
load_category

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub load_subject()
On Error GoTo err_des

Dim rs As New ADODB.Recordset
CboSubject.Clear
Set rs = getdata("SELECT  Sub_code,Sub_title From Subject_Info_sub WHERE Class_code = '" & Mid(cboClass, 1, 5) & "'")
If Not rs.EOF Then
    Do Until rs.EOF
        CboSubject.AddItem rs!Sub_code & "-" & rs!Sub_title
        rs.MoveNext
    Loop

End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub CboExamID_Click()
On Error GoTo err_des

  get_resultMain
  LoadStuID
  show_grid
  load_category
  
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub CboExamType_Click()
On Error GoTo err_des
  load_exam_sub
  get_resultMain
  LoadStuID
  show_grid
  load_category
  
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub CboSection_Click()
On Error GoTo err_des

 LoadStuID
 get_resultMain

 show_grid
 
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub CboSubject_Click()
On Error GoTo err_des
 get_resultMain
 LoadStuID
 show_grid
load_category

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdAcaYear_Click()
On Error GoTo err_des

  get_resultMain
  LoadStuID
  show_grid
  
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdDelete_Click()
On Error GoTo err_des
  Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
Dim classId As String
Dim SectionID As String

If Len(cboClass) = 0 Then
  MsgBox "Please select a class first", vbInformation, cmp
  cboClass.SetFocus
  Exit Sub
End If

If Len(CboSection.Text) = 0 Then
   MsgBox "Please select a Section ", vbInformation, cmp
   CboSection.SetFocus
   Exit Sub
End If

If Len(CboSubject) = 0 Then
   MsgBox "Please select a Subject ", vbInformation, cmp
   CboSubject.SetFocus
   Exit Sub
End If
  
If Len(cmdAcaYear) = 0 Then
   MsgBox "Please select an Academic Year ", vbInformation, cmp
   cmdAcaYear.SetFocus
   Exit Sub
End If

If Len(CboExamType) = 0 Then
   MsgBox "Please select an Exam Type ", vbInformation, cmp
   CboExamType.SetFocus
   Exit Sub
End If

If Len(CboExamID) = 0 Then
   MsgBox "Please select an Exam  ", vbInformation, cmp
   CboExamID.SetFocus
   Exit Sub
End If
  
If Len(Combo2.Text) = 0 Then
   MsgBox "Please select a Shift ", vbInformation, cmp
   Combo2.SetFocus
   Exit Sub
End If

  Set rs = getdata("select StdID from EXB_P_Result_Sub where M_Slr_no=" & Trim(txtSerial) & " and S_Slr_no =" & Trim(txtSerialSub) & "")
  If rs.EOF Then
    MsgBox "No such Student Exists...please verify.", vbInformation, cmp
    Exit Sub
  End If

  If MsgBox("Are you Sure to Delete ? ", vbInformation + vbYesNo + vbDefaultButton1, cmp) = vbYes Then
        con.Open GConnString
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "EX_B_P_Result_Save"
        cmd(1) = "d"
        cmd(2) = Val(txtSerial)
        cmd(3) = Val(txtSerialSub)
        cmd(4) = Mid(cboClass, 1, 5)
        cmd(5) = Mid(CboSection, 1, 5)
        cmd(6) = Mid(Combo2, 1, 1)
        cmd(7) = Mid(CboSubject, 1, 5)
        cmd(8) = Trim(cmdAcaYear)
        cmd(9) = Mid(CboExamType, 1, 2)
        cmd(10) = Mid(CboExamID, 1, 2)
        cmd(11) = Mid(CboCategory, 1, 5)
        cmd(12) = Mid(List1, 1, 10)
        cmd(13) = txtfields(3)
        cmd(14) = Val(txtCWMarks)
        cmd(15) = Val(txtHWMarks)
        cmd(16) = Val(txtAttentiveness)
        cmd(17) = Val(txtCleanness)
        cmd(18) = Val(txtManners)
        cmd(19) = soft_user
        cmd(20) = Date
    
        cmd.Execute
        
        Call LoadStuID
       
        For i = 3 To 3
        txtfields(i) = ""
        Next

        MsgBox "Deleted Successfully.", vbInformation, "Student Management System"
       
       get_resultMain
       Call LoadStuID
       show_grid
       
       txtCWMarks.Text = "0.00"
       txtHWMarks.Text = "0.00"
       txtAttentiveness.Text = "0.00"
       txtCleanness.Text = "0.00"
       txtManners.Text = "0.00"
       
       List1.SetFocus
End If
      
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdMarksheet_Click()
On Error GoTo err_des

Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
Dim classId As String
Dim SectionID As String

If Len(cboClass) = 0 Then
  MsgBox "Please select a class first", vbInformation, cmp
  cboClass.SetFocus
  Exit Sub
End If

If Len(CboSection.Text) = 0 Then
   MsgBox "Please select a Section ", vbInformation, cmp
   CboSection.SetFocus
   Exit Sub
End If

If Len(CboSubject) = 0 Then
   MsgBox "Please select a Subject ", vbInformation, cmp
   CboSubject.SetFocus
   Exit Sub
End If
  
If Len(cmdAcaYear) = 0 Then
   MsgBox "Please select an Academic Year ", vbInformation, cmp
   cmdAcaYear.SetFocus
   Exit Sub
End If

If Len(CboExamType) = 0 Then
   MsgBox "Please select an Exam Type ", vbInformation, cmp
   CboExamType.SetFocus
   Exit Sub
End If

If Len(CboExamID) = 0 Then
   MsgBox "Please select an Exam  ", vbInformation, cmp
   CboExamID.SetFocus
   Exit Sub
End If
  
If Len(Combo2.Text) = 0 Then
   MsgBox "Please select a Shift ", vbInformation, cmp
   Combo2.SetFocus
   Exit Sub
End If

If Len(List1.Text) = 0 Then
    MsgBox "Please a Student ID from the list ", vbInformation, cmp
    List1.SetFocus
   Exit Sub
  End If

If Len(txtMarks.Text) = 0 Then
    MsgBox "Please put obtained marks ", vbInformation, cmp
    txtMarks.SetFocus
   Exit Sub
  End If

If Len(CboCategory) = 0 Then
   MsgBox "Please Marks category required ", vbInformation, cmp
    CboCategory.SetFocus
   Exit Sub
  End If


        con.Open GConnString
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "Result_Save"
        cmd(1) = "s"
        cmd(2) = Val(txtSerial)
        cmd(3) = Val(txtSerialSub)
        cmd(4) = Mid(cboClass, 1, 5)
        cmd(5) = Mid(CboSection, 1, 5)
        cmd(6) = Mid(Combo2, 1, 1)
        cmd(7) = Mid(CboSubject, 1, 5)
        cmd(8) = Trim(cmdAcaYear)
        cmd(9) = Mid(CboExamType, 1, 2)
        cmd(10) = Mid(CboExamID, 1, 2)
        cmd(11) = Mid(CboCategory, 1, 5)
        cmd(12) = Mid(List1, 1, 10)
        cmd(13) = txtfields(3)
        cmd(14) = txtMarks
        cmd(15) = txtPassMarks
        cmd(16) = txtfullMarks
        cmd(17) = soft_user
        cmd(18) = Date
    
        cmd.Execute
        
        Call LoadStuID
       
        For i = 3 To 3
        txtfields(i) = ""
        Next

        MsgBox "Save Successfully.", vbInformation, "Student Management System"
       
       get_resultMain
       Call LoadStuID
        show_grid
        txtMarks = 0
       List1.SetFocus
       
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdSAVE_Click()
On Error GoTo err_des

Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
Dim classId As String
Dim SectionID As String

If Len(cboClass) = 0 Then
  MsgBox "Please select a class first", vbInformation, cmp
  cboClass.SetFocus
  Exit Sub
End If

If Len(CboSection.Text) = 0 Then
   MsgBox "Please select a Section ", vbInformation, cmp
   CboSection.SetFocus
   Exit Sub
End If

If Len(CboSubject) = 0 Then
   MsgBox "Please select a Subject ", vbInformation, cmp
   CboSubject.SetFocus
   Exit Sub
End If
  
If Len(cmdAcaYear) = 0 Then
   MsgBox "Please select an Academic Year ", vbInformation, cmp
   cmdAcaYear.SetFocus
   Exit Sub
End If

If Len(CboExamType) = 0 Then
   MsgBox "Please select an Exam Type ", vbInformation, cmp
   CboExamType.SetFocus
   Exit Sub
End If

If Len(CboExamID) = 0 Then
   MsgBox "Please select an Exam  ", vbInformation, cmp
   CboExamID.SetFocus
   Exit Sub
End If
  
If Len(Combo2.Text) = 0 Then
   MsgBox "Please select a Shift ", vbInformation, cmp
   Combo2.SetFocus
   Exit Sub
End If

If Len(List1.Text) = 0 Then
    MsgBox "Please a Student ID from the list ", vbInformation, cmp
    List1.SetFocus
   Exit Sub
  End If

'If Len(txtMarks.Text) = 0 Then
'    MsgBox "Please put obtained marks ", vbInformation, cmp
'    txtMarks.SetFocus
'   Exit Sub
'  End If

'If Len(CboCategory) = 0 Then
'   MsgBox "Please Marks category required ", vbInformation, cmp
'    CboCategory.SetFocus
'   Exit Sub
'  End If


'If Option1(0).Value = True Then
'   Set rs = GetData("select ClassId from StudentAttendanceLeaveInfo where ClassID='" & Trim(Mid(List1, 1, 5)) & "' and SectionID ='" & Mid(Combo1, 1, 5) & "'  and attn_date ='" & Format(MaskEdBox3, "dd mmm yyyy") & "'")
'  If Not rs.EOF Then
'    MsgBox "Attendance of all student of " & Trim(Mid(List1, 6, 15)) & " has already been completed on date :" & Format(MaskEdBox3, "dd mmm yyyy") & " ", vbInformation, cmp
'    Exit Sub
'  End If
'End If
'If Option1(1).Value = True Then
'   Set rs = GetData("select ClassId from StudentAttendanceLeaveInfo where ClassID='" & Trim(Mid(List1, 1, 5)) & "' and SectionID ='" & Mid(Combo1, 1, 5) & "'  and attn_date ='" & Format(MaskEdBox3, "dd mmm yyyy") & "' and StudentID='" & Mid(List2, 1, 10) & "'")
'  If Not rs.EOF Then
'     MsgBox "Attendance of Mr." & Mid(List2, 11, 80) & "  already been completed on date: " & MaskEdBox3.Text & " ", vbInformation, cmp
'    Exit Sub
'  End If
'End If
'
        con.Open GConnString
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "EX_B_P_Result_Save"
        cmd(1) = "s"
        cmd(2) = Val(txtSerial)
        cmd(3) = Val(txtSerialSub)
        cmd(4) = Mid(cboClass, 1, 5)
        cmd(5) = Mid(CboSection, 1, 5)
        cmd(6) = Mid(Combo2, 1, 1)
        cmd(7) = Mid(CboSubject, 1, 5)
        cmd(8) = Trim(cmdAcaYear)
        cmd(9) = Mid(CboExamType, 1, 2)
        cmd(10) = Mid(CboExamID, 1, 2)
        cmd(11) = Mid(CboCategory, 1, 5)
        cmd(12) = Mid(List1, 1, 10)
        cmd(13) = txtfields(3)
        cmd(14) = Val(txtCWMarks)
        cmd(15) = Val(txtHWMarks)
        cmd(16) = Val(txtAttentiveness)
        cmd(17) = Val(txtCleanness)
        cmd(18) = Val(txtManners)
        cmd(19) = soft_user
        cmd(20) = Date
    
        cmd.Execute
        
        generate_position
        Call LoadStuID
       
        For i = 3 To 3
        txtfields(i) = ""
        Next

        MsgBox "Save Successfully.", vbInformation, "Student Management System"
       
       get_resultMain
       Call LoadStuID
        show_grid
        txtMarks = 0
       List1.SetFocus
      
'Else
'        Exit Sub
'End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub generate_position()
On Error GoTo err_des
       If con.State = 0 Then
           con.Open GConnString
       End If
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "generate_position"
        cmd(1) = "a"
        cmd(2) = Mid(cboClass, 1, 5)
        cmd(3) = Mid(CboSection, 1, 5)
        cmd(4) = Mid(Combo2, 1, 1)
        cmd(5) = Mid(CboExamType, 1, 2)
        cmd(6) = Mid(CboExamID, 1, 2)
        cmd(7) = Trim(cmdAcaYear)
        cmd.Execute
        
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub cmdClose_Click()
On Error GoTo err_des

    Unload Me

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub


Private Sub ComStuId_Click()
On Error GoTo err_des

Label3.Caption = ""



txtfields(3) = ""
           
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT  distinct   StudentEvaluation.StudentID, StudentInfo.StudentName,  StudentEvaluation.Shift, " + _
        "StudentEvaluation.ClassId, ClassInfo.ClassName, StudentEvaluation.SectionId, SectionInfo.Sectiondsc, StudentEvaluation.ClassRoll, " + _
        "StudentEvaluation.ActiveClass FROM StudentEvaluation INNER JOIN SectionInfo ON StudentEvaluation.SectionId = SectionInfo.SectionID INNER JOIN " + _
        "ClassInfo ON StudentEvaluation.ClassId = ClassInfo.ClassID INNER JOIN StudentInfo ON StudentEvaluation.StudentID = StudentInfo.StudentID where StudentEvaluation.studentid='" & ComStuId & "'and StudentEvaluation.ActiveClass='Y'")
        
        If Not rs.EOF Then
            Label3.Caption = "" & rs!StudentName
         
           
'            txtfields(0) = "" & rs!Shift
            txtfields(3) = "" & rs!ClassRoll
           
        End If
        
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub ComStuId_KeyPress(KeyAscii As Integer)
On Error GoTo err_des

If KeyAscii = 13 Then
    scmdAttend.SetFocus
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub


Private Sub ComStuId_LostFocus()
On Error GoTo err_des

If Len(ComStuId) = 0 Then Exit Sub
Dim rs As New ADODB.Recordset
Set rs = getdata("select StudentId from StudentEvaluation where StudentId ='" & ComStuId.Text & "'")
If rs.EOF Then
    MsgBox "Invalid Id.", vbCritical, "School Management System"
    ComStuId.Text = ""
    Exit Sub
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub CmdEdit_Click()
On Error GoTo err_des

Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
Dim classId As String
Dim SectionID As String

If Len(cboClass) = 0 Then
  MsgBox "Please select a class first", vbInformation, cmp
  cboClass.SetFocus
  Exit Sub
End If

If Len(CboSection.Text) = 0 Then
   MsgBox "Please select a Section ", vbInformation, cmp
   CboSection.SetFocus
   Exit Sub
End If

If Len(CboSubject) = 0 Then
   MsgBox "Please select a Subject ", vbInformation, cmp
   CboSubject.SetFocus
   Exit Sub
End If
  
If Len(cmdAcaYear) = 0 Then
   MsgBox "Please select an Academic Year ", vbInformation, cmp
   cmdAcaYear.SetFocus
   Exit Sub
End If

If Len(CboExamType) = 0 Then
   MsgBox "Please select an Exam Type ", vbInformation, cmp
   CboExamType.SetFocus
   Exit Sub
End If

If Len(CboExamID) = 0 Then
   MsgBox "Please select an Exam  ", vbInformation, cmp
   CboExamID.SetFocus
   Exit Sub
End If
  
If Len(Combo2.Text) = 0 Then
   MsgBox "Please select a Shift ", vbInformation, cmp
   Combo2.SetFocus
   Exit Sub
End If





  Set rs = getdata("select StdID from EXB_P_Result_Sub where M_Slr_no=" & Trim(txtSerial) & " and S_Slr_no =" & Trim(txtSerialSub) & "")
  If rs.EOF Then
    MsgBox "No such Student Exists...please verify.", vbInformation, cmp
    Exit Sub
  End If

        con.Open GConnString
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "EX_B_P_Result_Save"
        cmd(1) = "u"
        cmd(2) = Val(txtSerial)
        cmd(3) = Val(txtSerialSub)
        cmd(4) = Mid(cboClass, 1, 5)
        cmd(5) = Mid(CboSection, 1, 5)
        cmd(6) = Mid(Combo2, 1, 1)
        cmd(7) = Mid(CboSubject, 1, 5)
        cmd(8) = Trim(cmdAcaYear)
        cmd(9) = Mid(CboExamType, 1, 2)
        cmd(10) = Mid(CboExamID, 1, 2)
        cmd(11) = Mid(CboCategory, 1, 5)
        cmd(12) = Mid(List1, 1, 10)
        cmd(13) = txtfields(3)
        cmd(14) = Val(txtCWMarks)
        cmd(15) = Val(txtHWMarks)
        cmd(16) = Val(txtAttentiveness)
        cmd(17) = Val(txtCleanness)
        cmd(18) = Val(txtManners)
        cmd(19) = soft_user
        cmd(20) = Date
        
        cmd.Execute
        
        Call LoadStuID
       
        For i = 3 To 3
        txtfields(i) = ""
        Next

        MsgBox "Edited Successfully.", vbInformation, "Student Management System"
       
       get_resultMain
       Call LoadStuID
        show_grid
        txtMarks = 0
       List1.SetFocus
       
       txtCWMarks.Text = "0.00"
       txtHWMarks.Text = "0.00"
       txtAttentiveness.Text = "0.00"
       txtCleanness.Text = "0.00"
       txtManners.Text = "0.00"

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdPrint_Click()
On Error GoTo err_des

If Len(cboClass) = 0 Then
  MsgBox "Please select a class first", vbInformation, cmp
  cboClass.SetFocus
  Exit Sub
End If

If Len(CboSection.Text) = 0 Then
   MsgBox "Please select a Section ", vbInformation, cmp
   CboSection.SetFocus
   Exit Sub
End If

If Len(CboSubject) = 0 Then
   MsgBox "Please select a Subject ", vbInformation, cmp
   CboSubject.SetFocus
   Exit Sub
End If
  
If Len(cmdAcaYear) = 0 Then
   MsgBox "Please select an Academic Year ", vbInformation, cmp
   cmdAcaYear.SetFocus
   Exit Sub
End If

If Len(CboExamType) = 0 Then
   MsgBox "Please select an Exam Type ", vbInformation, cmp
   CboExamType.SetFocus
   Exit Sub
End If

If Len(CboExamID) = 0 Then
   MsgBox "Please select an Exam  ", vbInformation, cmp
   CboExamID.SetFocus
   Exit Sub
End If
  
If Len(Combo2.Text) = 0 Then
   MsgBox "Please select a Shift ", vbInformation, cmp
   Combo2.SetFocus
   Exit Sub
End If


rptMode = 13
Screen.MousePointer = vbHourglass
frmViewer.Show 1

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub Combo1_Click()
On Error GoTo err_des
   LoadStuID

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub Combo2_Click()
On Error GoTo err_des

  get_resultMain
  LoadStuID
  show_grid
  
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub load_category()
On Error GoTo err_des

 Dim rs As New ADODB.Recordset
 Set rs = getdata("select a.CategoryID,(select b.MCategoryDsc  from markscategory b where b.MCategoryID=a.CategoryID) from SubjectMarksDistribution  a where a.ClassId = '" & Mid(Trim(cboClass.Text), 1, 5) & "' and a.SubjectId= '" & Mid(Trim(CboSubject.Text), 1, 5) & "' and a.term_code='" & Mid(Trim(CboExamType.Text), 1, 2) & "' and a.Exam_code='" & Mid(Trim(CboExamID.Text), 1, 2) & "' order by a.CategoryID")
 CboCategory.Clear

 If Not rs.EOF Then
  rs.MoveFirst
    Do Until rs.EOF
    CboCategory.AddItem rs(0) + "-" + rs(1)
    rs.MoveNext
    Loop

End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo err_des
  If KeyAscii = 13 Then
    SendKeys (Chr(9))
  End If
  If KeyAscii = 27 Then
    Unload Me
  End If
  
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub show_grid()
On Error GoTo err_des
  
Dim rs As New ADODB.Recordset
'Set rs = getdata("select a.StdID,(select studentname from studentinfo where studentid=a.stdid) as stdname,a.Roll,a.ObtainedMarks,a.S_Slr_no from result_sub a where a.M_Slr_no='" & Trim(txtSerial) & "' order by a.StdID")
  
  Set rs = getdata("SELECT a.StdID,(SELECT studentname FROM studentinfo WHERE studentid = a.stdid) AS stdname, " _
        & " a.Roll, a.CWObtainedMarks, a.HWObtainedMarks, a.AttentivenessObtainedMarks," _
        & " a.CleannessObtainedMarks, a.MannersObtainedMarks, a.S_Slr_no, a.M_Slr_no " _
        & " FROM EXB_P_Result_Sub a INNER JOIN " _
        & " EXB_P_Result_Main ON a.M_Slr_no = EXB_P_Result_Main.M_Slr_no " _
        & " where EXB_P_Result_Main.ClassID='" & Mid(cboClass, 1, 5) & "'" _
        & " and EXB_P_Result_Main.SectionID='" & Mid(CboSection, 1, 5) & "'" _
        & " and EXB_P_Result_Main.Shift='" & Mid(Combo2, 1, 1) & "'" _
        & " and EXB_P_Result_Main.SubID='" & Mid(CboSubject, 1, 5) & "'" _
        & " and EXB_P_Result_Main.AcaYr='" & Trim(cmdAcaYear) & "'" _
        & " and  EXB_P_Result_Main.ExamType='" & Mid(CboExamType, 1, 2) & "'" _
        & " and EXB_P_Result_Main.ExamID='" & Mid(CboExamID, 1, 2) & "' ORDER BY a.StdID")
        
        
  If Not rs.EOF Then
   i = 1
    With MSFlexGrid1
        Do Until rs.EOF
                .Rows = i + 1
                .TextMatrix(i, 0) = rs!stdid
                .TextMatrix(i, 1) = rs!roll
                .TextMatrix(i, 2) = "" & rs!stdname
                .TextMatrix(i, 3) = rs!CWObtainedMarks
                .TextMatrix(i, 4) = rs!HWObtainedMarks
                .TextMatrix(i, 5) = rs!AttentivenessObtainedMarks
                .TextMatrix(i, 6) = rs!CleannessObtainedMarks
                .TextMatrix(i, 7) = rs!MannersObtainedMarks
                .TextMatrix(i, 8) = "100.00"
                .TextMatrix(i, 9) = Val(rs!CWObtainedMarks) + Val(rs!HWObtainedMarks) + Val(rs!AttentivenessObtainedMarks) + Val(rs!CleannessObtainedMarks) + Val(rs!MannersObtainedMarks)
                .TextMatrix(i, 10) = rs!S_Slr_no
                .TextMatrix(i, 11) = rs!M_Slr_no
                
                i = i + 1
                
            rs.MoveNext
        Loop
    End With
Else
    MSFlexGrid1.Rows = 1
'    MSFlexGrid1.Clear
 End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub get_resultMain()
On Error GoTo err_des

  Dim rs As New ADODB.Recordset
  Set rs = getdata("select M_Slr_no from EXB_P_Result_Main where ClassID='" & Mid(cboClass, 1, 5) & "' and SectionID='" & Mid(CboSection, 1, 5) & "' and Shift='" & Mid(Combo2, 1, 1) & "' and  SubID='" & Mid(CboSubject, 1, 5) & "' and  AcaYr='" & Trim(cmdAcaYear) & "' and  ExamType='" & Mid(CboExamType, 1, 2) & "'and  ExamID='" & Mid(CboExamID, 1, 2) & "' and categoryid='" & Mid(CboCategory, 1, 5) & "'")

  If Not rs.EOF Then
     txtSerial = rs(0)
  Else
    txtSerial = 0
 End If


'        cmd(11) = Mid(List1, 1, 10)
'        cmd(12) = txtfields(3)
'        cmd(13) = txtMarks
'        cmd(14) = soft_user
'        cmd(15) = Date
'

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub Form_Load()
On Error GoTo err_des

Call LoadStuID
load_class
load_Aca_year
load_Scction
load_exam

With MSFlexGrid1
     .Rows = 1
     .Cols = 12
     .Col = 0: .Text = "ID #"
     .Col = 1: .Text = "Roll"
     .Col = 2: .Text = "Name"
     .Col = 3: .Text = "C.W. Marks"
     .Col = 4: .Text = "H.W. Marks"
     .Col = 5: .Text = "Attentiveness Marks"
     .Col = 6: .Text = "Cleanness Marks"
     .Col = 7: .Text = "Manners Marks"
     .Col = 8: .Text = "Total Marks"
     .Col = 9: .Text = "Obtain Marks"
     .Col = 10: .Text = ""
     .Col = 11: .Text = ""
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 800
    .ColWidth(2) = 4500
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1600
    .ColWidth(6) = 1500
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 10
    .ColWidth(11) = 10
    
    .ColAlignment(0) = 1
    .ColAlignment(1) = 2
    .ColAlignment(2) = 1
    .ColAlignment(3) = 6
    .ColAlignment(4) = 6
    .ColAlignment(5) = 6
    .ColAlignment(6) = 6
    .ColAlignment(7) = 6
    .ColAlignment(8) = 6
    .ColAlignment(9) = 6
    
    
End With

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub load_exam()
On Error GoTo err_des

Dim rs1 As New ADODB.Recordset
CboExamType.Clear
Set rs1 = getdata("Select ETypeID,ETypeName from ExamTypeInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        CboExamType.AddItem rs1(0) + "-" + rs1(1)
       rs1.MoveNext
    Loop
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub load_exam_sub()
On Error GoTo err_des

   Dim rs As New ADODB.Recordset
   Set rs = getdata("Select Exam_code,Exam_title from Exam_setup where Group_code= '" & Mid(Trim(CboExamType.Text), 1, 2) & "'")
   
   CboExamID.Clear
   If Not rs.EOF Then
     Do Until rs.EOF
      CboExamID.AddItem rs(0) + "-" + rs(1)
      rs.MoveNext
     Loop
         
   End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub load_Scction()
On Error GoTo err_des

 Dim rs1 As New ADODB.Recordset
CboSection.Clear
Set rs1 = getdata("Select SectionID,Sectiondsc from SectionInfo where ClassID='" & Mid(Trim(cboClass.Text), 1, 5) & "'")
If Not rs1.EOF Then
    Do Until rs1.EOF
        CboSection.AddItem Trim(rs1(0)) + "-" + Trim(rs1(1))
        rs1.MoveNext
    Loop
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub load_Aca_year()
On Error GoTo err_des

  Dim i As Integer
For i = 2000 To 2050
  cmdAcaYear.AddItem i
Next i
cmdAcaYear.Text = Format(Date, "YYYY")

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub LoadStuID()
On Error GoTo err_des

 List1.Clear
 lblStdname.Text = ""
 txtfields(3).Text = ""
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT Distinct a.StudentID FROM  Studentadmission a where a.classid='" _
    & Mid(cboClass, 1, 5) & "' and a.sectionid='" & Mid(CboSection, 1, 5) & "' and a.approval='Y' and  " _
    & "active_std=1 and a.admissionCancel='N'and a.Shift='" & Mid(Combo2, 1, 1) & "' and a.aca_yr='" _
    & Trim(cmdAcaYear) & "' and  a.serial_no=(select max(serial_no) from studentadmission where studentid=a.studentid)" _
    & " and a.StudentID not in(select StdID from EXB_P_Result_Sub where M_Slr_no='" & _
    Trim(txtSerial) & "') order by a.studentid")


If Not rs.EOF Then
    Do Until rs.EOF
        List1.AddItem rs!studentid
        rs.MoveNext
    Loop
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub load_class()
On Error GoTo err_des
   Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT ClassID, ClassName FROM  classinfo")
cboClass.Clear
If Not rs.EOF Then
    Do Until rs.EOF
        cboClass.AddItem rs(0) + "-" + rs(1)
        rs.MoveNext
    Loop
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub get_roll()
On Error GoTo err_des

Dim rs As New ADODB.Recordset
Set rs = getdata("select classRoll From StudentAdmission where StudentId='" & Mid(List1, 1, 10) & "'" & _
 " and serial_no=(select max(serial_no)  From StudentAdmission  where StudentId='" & Mid(List1, 1, 10) & "')")
If Not rs.EOF Then
  txtfields(3).Text = rs(0)
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub Option1_Click(Index As Integer)
On Error GoTo err_des

   Select Case Index
          Case 0
             List2.Enabled = False
             txtfields(3).Text = ""
          Case 1
             List2.Enabled = True
  End Select
  
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub List1_Click()
On Error GoTo err_des

 get_roll
 get_name
 txtSerialSub = ""
    
    txtCWMarks.Text = "0.00"
    txtHWMarks.Text = "0.00"
    txtAttentiveness.Text = "0.00"
    txtCleanness.Text = "0.00"
    txtManners.Text = "0.00"

 
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub get_name()
On Error GoTo err_des

Dim rs As New ADODB.Recordset
Set rs = getdata("select studentname From Studentinfo where StudentId='" & Mid(List1, 1, 10) & "'")
If Not rs.EOF Then
   lblStdname.Text = rs(0)
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo err_des

Dim rs As New ADODB.Recordset


'Set rs = getdata("SELECT a.StdID,(SELECT studentname FROM studentinfo WHERE studentid = a.stdid) AS stdname, " _
'        & " a.Roll, a.CWObtainedMarks, a.HWObtainedMarks, a.AttentivenessObtainedMarks," _
'        & " a.CleannessObtainedMarks, a.MannersObtainedMarks, a.S_Slr_no, a.M_Slr_no, EXB_P_Result_Main.ClassID," _
'        & " EXB_P_Result_Main.SectionID,EXB_P_Result_Main.Shift, EXB_P_Result_Main.SubID, " _
'        & " EXB_P_Result_Main.AcaYr, EXB_P_Result_Main.ExamType, EXB_P_Result_Main.ExamID," _
'        & " EXB_P_Result_Main.categoryid FROM         EXB_P_Result_Sub a INNER JOIN " _
'        & " EXB_P_Result_Main ON a.M_Slr_no = EXB_P_Result_Main.M_Slr_no ORDER BY a.StdID")
'
'        For intcmbItem = 0 To CboSection.ListCount
'            If Trim(Mid(Trim(CboSection.List(intcmbItem)), 1, InStr(Trim(CboSection.List(intcmbItem)), "-") - 1)) = rs!SectionID Then
'                CboSection.ListIndex = intcmbItem
'                Exit For
'            End If
'        Next
'
'
'        If Trim("M") = rs!Shift Then
'            Combo2.ListIndex = 0
'        ElseIf Trim("D") = rs!Shift Then
'            Combo2.ListIndex = 1
'        End If
'
'
'        For intcmbItem = 0 To CboSubject.ListCount
'            If Trim(Mid(Trim(CboSubject.List(intcmbItem)), 1, InStr(Trim(CboSubject.List(intcmbItem)), "-") - 1)) = rs!SubID Then
'                CboSubject.ListIndex = intcmbItem
'                Exit For
'            End If
'        Next
'
'        For intcmbItem = 0 To cmdAcaYear.ListCount
'            If cmdAcaYear.List(intcmbItem) = rs!AcaYr Then
'                cmdAcaYear.ListIndex = intcmbItem
'                Exit For
'            End If
'        Next
'
'        For intcmbItem = 0 To CboExamType.ListCount
'            If Trim(Mid(Trim(CboExamType.List(intcmbItem)), 1, InStr(Trim(CboExamType.List(intcmbItem)), "-") - 1)) = rs!ExamType Then
'                CboExamType.ListIndex = intcmbItem
'                Exit For
'            End If
'        Next
'
'        For intcmbItem = 0 To CboExamID.ListCount
'            If Trim(Mid(Trim(CboExamID.List(intcmbItem)), 1, InStr(Trim(CboExamID.List(intcmbItem)), "-") - 1)) = rs!ExamID Then
'                CboExamID.ListIndex = intcmbItem
'                Exit For
'            End If
'        Next

 If MSFlexGrid1.Row > 0 Then
  
  
  txtid = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
  txtfields(3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
  lblStdname = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
  txtCWMarks = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
  txtHWMarks = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
  txtAttentiveness = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
  txtCleanness = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
  txtManners = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
  txtSerialSub = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)
  txtSerial = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11)
  
End If

 'LoadStuID
 
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub MSFlexGrid1_SelChange()
On Error GoTo err_des

  MSFlexGrid1_Click
  
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub txtMarks_Change()
On Error GoTo err_des

  If Not IsNumeric(txtMarks) Then
      txtMarks = ""
  ElseIf Val(txtMarks) > Val(txtfullMarks) Then
     txtMarks = 0
  End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub txtMarks_GotFocus()
On Error GoTo err_des

   txtMarks.SelStart = 0
   txtMarks.SelLength = Len(Trim(txtMarks.Text))
   
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
Private Sub txtMarks_LostFocus()
On Error GoTo err_des

  If Val(txtMarks) > Val(txtfullMarks) Then
     txtMarks = 0
  End If
  
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub txtAttentiveness_GotFocus()
On Error GoTo err_des

txtAttentiveness.SelStart = 0
txtAttentiveness.SelLength = Len(Trim(txtAttentiveness.Text))

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

'Private Sub txtAttentiveness_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    txtCleanness.SetFocus
'End If
'End Sub

Private Sub txtAttentiveness_LostFocus()
On Error GoTo err_des

If Val(Trim(txtAttentiveness.Text)) > 20 Then
    txtAttentiveness.Text = 0
    txtAttentiveness.SetFocus
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub txtCleanness_GotFocus()
On Error GoTo err_des

txtCleanness.SelStart = 0
txtCleanness.SelLength = Len(Trim(txtCleanness.Text))

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

'Private Sub txtCleanness_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    txtManners.SetFocus
'End If
'End Sub

Private Sub txtCleanness_LostFocus()
On Error GoTo err_des

If Val(Trim(txtCleanness.Text)) > 20 Then
    txtCleanness.Text = 0
    txtCleanness.SetFocus
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub txtCWMarks_GotFocus()
On Error GoTo err_des

txtCWMarks.SelStart = 0
txtCWMarks.SelLength = Len(Trim(txtCWMarks.Text))

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

'Private Sub txtCWMarks_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    txtHWMarks.SetFocus
'End If
'End Sub

Private Sub txtCWMarks_LostFocus()
On Error GoTo err_des

If Val(Trim(txtCWMarks.Text)) > 20 Then
    txtCWMarks.Text = 0
    txtCWMarks.SetFocus
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub txtHWMarks_GotFocus()
On Error GoTo err_des

txtHWMarks.SelStart = 0
txtHWMarks.SelLength = Len(Trim(txtHWMarks.Text))

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

'Private Sub txtHWMarks_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    txtAttentiveness.SetFocus
'End If
'End Sub

Private Sub txtHWMarks_LostFocus()
On Error GoTo err_des

If Val(Trim(txtHWMarks.Text)) > 20 Then
    txtHWMarks.Text = 0
    txtHWMarks.SetFocus
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub txtManners_GotFocus()
On Error GoTo err_des

txtManners.SelStart = 0
txtManners.SelLength = Len(Trim(txtManners.Text))

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

'Private Sub txtManners_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    cmdSave.SetFocus
'End If
'End Sub

Private Sub txtManners_LostFocus()
On Error GoTo err_des

If Val(Trim(txtManners.Text)) > 20 Then
    txtManners.Text = 0
    txtManners.SetFocus
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
