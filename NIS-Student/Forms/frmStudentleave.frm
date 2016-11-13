VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmStudentleave 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSerial 
      Height          =   285
      Left            =   750
      TabIndex        =   27
      Top             =   8280
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Caption         =   "Leave Specification"
      Height          =   1335
      Left            =   -90
      TabIndex        =   23
      Top             =   2640
      Width           =   10005
      Begin VB.TextBox txtCause 
         Height          =   585
         Left            =   1140
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   660
         Width           =   8715
      End
      Begin MSMask.MaskEdBox MaskEdBoxDate 
         Height          =   315
         Index           =   0
         Left            =   1140
         TabIndex        =   2
         ToolTipText     =   "Insert  form  Date"
         Top             =   240
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBoxDate 
         Height          =   315
         Index           =   1
         Left            =   4260
         TabIndex        =   3
         ToolTipText     =   "Insert  To  Date"
         Top             =   240
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002B3AFD&
         Height          =   375
         Left            =   6900
         TabIndex        =   28
         Top             =   270
         Width           =   1665
      End
      Begin VB.Label Label4 
         Caption         =   "Cause "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   270
         TabIndex        =   26
         Top             =   690
         Width           =   885
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3690
         TabIndex        =   25
         Top             =   270
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   24
         Top             =   300
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdAdmitted 
      BackColor       =   &H8000000C&
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      TabIndex        =   5
      ToolTipText     =   "Click To Accept"
      Top             =   7980
      Width           =   1275
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3945
      Left            =   0
      TabIndex        =   12
      Top             =   3960
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   6959
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorSel    =   12648447
      ForeColorSel    =   192
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000C&
      Caption         =   "Reject"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4950
      TabIndex        =   6
      ToolTipText     =   "Click to cancel Admission"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8640
      TabIndex        =   7
      ToolTipText     =   "Click to Close"
      Top             =   7980
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1845
      Left            =   0
      TabIndex        =   9
      Top             =   810
      Width           =   9945
      Begin VB.ComboBox cboAcaYr 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   1605
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H80000018&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   4
         Left            =   8160
         TabIndex        =   20
         Top             =   1380
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H80000018&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   3
         Left            =   5490
         TabIndex        =   18
         Top             =   1380
         Width           =   1845
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H80000018&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   2
         Left            =   2670
         TabIndex        =   16
         Top             =   1380
         Width           =   1965
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H80000018&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   1050
         TabIndex        =   14
         Top             =   1380
         Width           =   795
      End
      Begin VB.ComboBox ComStuId 
         Height          =   315
         Left            =   1050
         TabIndex        =   1
         ToolTipText     =   "Select student"
         Top             =   870
         Width           =   1605
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aca.Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   330
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   4
         Left            =   7380
         TabIndex        =   21
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   3
         Left            =   4950
         TabIndex        =   19
         Top             =   1410
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   2
         Left            =   1950
         TabIndex        =   17
         Top             =   1410
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Roll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   15
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2640
         TabIndex        =   11
         Top             =   870
         Width           =   7065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   10
         Top             =   900
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   9825
      TabIndex        =   8
      Top             =   0
      Width           =   9885
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Leave Entry"
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
         Left            =   3240
         TabIndex        =   13
         Top             =   150
         Width           =   2325
      End
      Begin VB.Image Image1 
         Height          =   930
         Left            =   0
         Picture         =   "frmStudentleave.frx":0000
         Stretch         =   -1  'True
         Top             =   -60
         Width           =   9915
      End
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   7290
      Top             =   7950
      Width           =   2535
   End
   Begin VB.Menu mnuDel 
      Caption         =   "Delete"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu gfsdgfsd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRef 
         Caption         =   "Refresh"
      End
      Begin VB.Menu fdsafdsa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmStudentleave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAcaYear_LostFocus()
     get_roll
     ShowFlexData
End Sub
Private Sub cmdAdmitted_Click()
  If Len(cboAcaYr.Text) = 0 Then
      MsgBox "Plsease select an academic year.", vbInformation, cmp
      cboAcaYr.SetFocus
      Exit Sub
  End If

  If MaskEdBoxDate(0) = "__/__/__" Then
      MsgBox "From Date required..", vbInformation, cmp
      MaskEdBoxDate(0).SetFocus
      Exit Sub
  End If
  
  If MaskEdBoxDate(1) = "__/__/__" Then
      MsgBox "To Date required..", vbInformation, cmp
      MaskEdBoxDate(1).SetFocus
      Exit Sub
  End If
    
  If Len(ComStuId) = 0 Then
      MsgBox "Student Id required.", vbInformation, cmp
      ComStuId.SetFocus
      Exit Sub
  End If
 
 If Format(MaskEdBoxDate(0), "yyyy") <> cboAcaYr.Text Then
    MsgBox "Year Mismatch...Please Verify", vbInformation, cmp
    MaskEdBoxDate(0).SetFocus
    Exit Sub
 End If
 
 If Format(MaskEdBoxDate(1), "yyyy") <> cboAcaYr.Text Then
    MsgBox "Year Mismatch...Please Verify", vbInformation, cmp
    MaskEdBoxDate(1).SetFocus
    Exit Sub
 End If
 
 If CDate(MaskEdBoxDate(0)) > CDate(MaskEdBoxDate(1)) Then
    MsgBox "Invalid Date Range Please Verify", vbInformation, cmp
    MaskEdBoxDate(0).SetFocus
    Exit Sub
 End If
 
 If Len(txtCause) = 0 Then
    MsgBox "Please Specify the cause of leave..", vbInformation, cmp
    txtCause.SetFocus
    Exit Sub
 End If
 
 
  Set rs = getdata("select StudentId from StudentAdmission where aca_yr='" & Trim(cboAcaYr) & "' and studentId='" & Trim(ComStuId) & "'")
  
  If rs.EOF Then
     MsgBox "Invalid Student Id...Please Verify", vbInformation, cmp
     Exit Sub
 End If
  
Set rs = getdata("select L_st_dt from Student_leave_info where AcademicYr='" & Trim(cboAcaYr) & "' and StudentId='" & Trim(ComStuId) & "' and ('" & Format(MaskEdBoxDate(0).Text, "dd mmm yyyy") & "' BETWEEN  L_st_dt and L_ed_dt)")
  
  If Not rs.EOF Then
     MsgBox "Date already exists between this days...Please Verify", vbInformation, cmp
     Exit Sub
 End If
  
 Set rs = getdata("select L_st_dt from Student_leave_info where AcademicYr='" & Trim(cboAcaYr) & "' and StudentId='" & Trim(ComStuId) & "' and ('" & Format(MaskEdBoxDate(1).Text, "dd mmm yyyy") & "' BETWEEN  L_st_dt and L_ed_dt)")
  
  If Not rs.EOF Then
     MsgBox "Date already exists between this days...Please Verify", vbInformation, cmp
     Exit Sub
 End If
 
  
  
    s_u_d_leave_info (1)
    MsgBox "Saved Successfully.", vbInformation, "Student Management System"
    ShowFlexData


ComStuId.SetFocus
Exit Sub
End Sub
Private Sub s_u_d_leave_info(mode As Integer)
   Dim cmd As New ADODB.Command
   Dim con As New ADODB.connection

   con.Open GConnString
   cmd.ActiveConnection = con
   cmd.CommandType = adCmdStoredProc
   cmd.CommandText = "S_U_D_leave_info"
   cmd(1) = mode
   cmd(2) = Trim(ComStuId.Text)
   cmd(3) = Trim(cboAcaYr.Text)
   cmd(4) = MaskEdBoxDate(0)
   cmd(5) = MaskEdBoxDate(1)
   cmd(6) = Trim(txtCause)
   If mode = 1 Then
      cmd(7) = 0
   Else
     cmd(7) = Trim(txtSerial)
   End If
   cmd.Execute
End Sub
Private Sub cmdCancel_Click()
If Len(ComStuId) = 0 And Len(ComboClass) = 0 Then Exit Sub
If Len(ComStuId) = 0 Then
    MsgBox "Please Enter Student ID. ", vbInformation, cmp
    ComStuId.SetFocus
    Exit Sub
End If

If Len(txtfields(0)) = 0 Then
    MsgBox "Invalid Student Id..Please Verify", vbInformation, cmp
    ComStuId.SetFocus
    Exit Sub
End If

If Len(ComboShift.Text) = 0 Then
    MsgBox "Select Shift Name.", vbInformation, cmp
    ComboShift.SetFocus
    Exit Sub
End If
If Len(ComboClass.Text) = 0 Then
    MsgBox "Select Class .", vbInformation, cmp
    ComboClass.SetFocus
    Exit Sub
End If
If Len(ComboSection.Text) = 0 Then
    MsgBox "Select Section.", vbInformation, cmp
    ComboSection.SetFocus
    Exit Sub
End If
If Len(cmdAcaYear.Text) = 0 Then
    MsgBox "Plsease select an academic year from the list.", vbInformation, cmp
    cmdAcaYear.SetFocus
    Exit Sub
End If
If MaskEdBoxDate(1) = "__/__/__" Then
    MsgBox "Enter Date.", vbInformation, cmp
    MaskEdBoxDate(1).SetFocus
    Exit Sub
End If
If MaskEdBoxDate(0) = "__/__/__" Then
    MsgBox "Enter Date.", vbInformation, cmp
    MaskEdBoxDate(0).SetFocus
    Exit Sub
End If
  Dim rs3 As New ADODB.Recordset
  Set rs3 = getdata("select count(ClassRoll) from StudentAdmission where studentId='" & Trim(ComStuId) & "'")
    If Not rs3.EOF Then
      If rs3(0) < 2 Then
        MsgBox "You can only cancel Re-Admission ..not admission by this operation", vbInformation, cmp
         ComStuId.SetFocus
      Exit Sub
  End If
End If

If Len(ComboSection) <> 0 And Len(ComboShift) <> 0 And Len(ComboClass) <> 0 And Len(cmdAcaYear) <> 0 Then
 
     Set rs3 = getdata("select StudentId from StudentAttendanceLeaveInfo where classid='" & Mid(ComboClass, 1, 5) & "' and sectionId='" & Mid(ComboSection, 1, 5) & "' and shift='" & Mid(ComboShift, 1, 1) & "' and aca_yr='" & Trim(cmdAcaYear) & "' and studentId='" & Trim(ComStuId) & "'")
      If Not rs3.EOF Then
        MsgBox "You can't delete a regular Student..Please Chose a another", vbInformation, cmp
        ComStuId.SetFocus
        Exit Sub
      End If
End If



If MsgBox("Are You sure to cancel Admission Information for this Student? ", vbYesNo + vbInformation) = vbYes Then
    Dim con As New ADODB.connection
    con.Open GConnString
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "StuAdmissionEvaluationInformation"
            cmd(1) = 2
            cmd(2) = Trim(ComStuId.Text)
            cmd(3) = Format(MaskEdBoxDate, "dd mmm yyyy")
            cmd(4) = Mid(ComboShift, 1, 1)
            cmd(5) = Mid(ComboClass, 1, 5)
            cmd(6) = Mid(ComboSection, 1, 5)
            cmd(7) = txtfields(1)
            cmd(8) = soft_user
            cmd(9) = Date
            cmd(10) = "Y"
            cmd(11) = "N"
            cmd(12) = "Y"
            cmd(13) = "Y"
            cmd(14) = Trim(cmdAcaYear)
         cmd.Execute
    MsgBox "Cancelled Successfully.", vbInformation, "Student Management System"
      ShowFlexData
  Else
     Exit Sub
 End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub ComboClass_Click()
ComboSection.Clear
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("select SectionId,Sectiondsc from SectionInfo where ClassId='" & Mid(ComboClass, 1, 5) & "'")
If Not rs1.EOF Then
    Do Until rs1.EOF
        ComboSection.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
End If
'ComboSection.SetFocus
End Sub

Private Sub ComboClass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboSection.SetFocus
End If
End Sub

Private Sub ComboClass_LostFocus()
    get_roll
End Sub


Private Sub ComboSection_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboShift.SetFocus
End If
End Sub

Private Sub ComboSection_LostFocus()
    get_roll
End Sub


Private Sub ComboShift_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdAcaYear.SetFocus
End If
End Sub
Private Sub get_roll()
  If Len(ComboSection) <> 0 And Len(ComboShift) <> 0 And Len(ComboClass) <> 0 And Len(cmdAcaYear) <> 0 Then
  Dim rs As New ADODB.Recordset
     Set rs = getdata("select max(ClassRoll)+ 1 from StudentAdmission where classid='" & Mid(ComboClass, 1, 5) & "' and shift='" & Mid(ComboShift, 1, 1) & "'and sectionId='" & Mid(ComboSection, 1, 5) & "' and aca_yr='" & Trim(cmdAcaYear) & "'")
      If Not rs.EOF Then
        txtfields(1) = IIf(IsNull(rs(0)) = True, "1", rs(0))
    Else
        txtfields(1) = "1"
    End If
 End If
End Sub
'Private Sub ComStuId_click()
''ComboShift.SetFocus
''Label3.Caption = ""
'''ComboClass.Text = ""
''''ComboSection = ""
'''ComboShift = ""
''txtfields = ""
''MaskEdBoxDate = "__/__/__"
''cmdAdmitted.Enabled = True
'''load_roll
'End Sub

Private Sub ComStuId_GotFocus()
    load_roll
End Sub
Private Sub ComStuId_Click()
   Dim rs As New ADODB.Recordset
   Dim sec_rs As New ADODB.Recordset
   Set rs = getdata("select s.Shift,i.StudentName,s.ClassId,c.ClassName,s.SectionId,s.ClassRoll,s.aca_yr from " + _
   " studentAdmission s,studentInfo i,classInfo c where s.StudentId='" & Trim(ComStuId.Text) & "' and s.ClassId=c.ClassId and s.StudentId=i.StudentId and s.serial_no=(select max(serial_no) from studentAdmission where StudentId='" & Trim(ComStuId.Text) & "' AND aca_yr='" & Trim(cboAcaYr.Text) & "') ")
   If Not rs.EOF Then
     txtfields(0) = rs!ClassRoll
     Label3.Caption = rs!StudentName
     Set sec_rs = getdata("select Sectiondsc from sectionInfo where sectionId='" & Trim(rs!SectionID) & "' and classId='" & Trim(rs!classId) & "'")
     If Not sec_rs.EOF Then
        txtfields(2).Text = sec_rs(0)
     End If
     txtfields(3).Text = IIf(rs!Shift = "M", "Morning", "Day")
     txtfields(4).Text = rs!ClassName
     ShowFlexData
  Else
     ShowFlexData
     txtfields(0) = ""
     Label3.Caption = ""
     txtfields(2).Text = ""
     txtfields(3).Text = ""
     txtfields(4).Text = ""
'     txtFields(5).Text = ""
  End If
  txtSerial = ""
  Set rs = Nothing
  Set sec_rs = Nothing
End Sub
Private Sub Form_Load()
    format_grid
    load_roll
    load_Aca_year
    get_class
End Sub
Private Sub format_grid()

With MSFlexGrid1
        .Rows = 1
        .Cols = 6
        .Col = 0: .Text = " Srl #"
        .Col = 1: .Text = "            Start Date"
        .Col = 2: .Text = "            End date"
        .Col = 3: .Text = "Duration(in Days)"
        .Col = 4: .Text = "Remarks"
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 5000
        .ColWidth(5) = 0
        .Rows = 50
   End With
End Sub
Private Sub load_Aca_year()
  Dim i As Integer
  For i = 2000 To 2050
     cboAcaYr.AddItem i
  Next i
  cboAcaYr.Text = Format(Date, "YYYY")
End Sub
Private Sub load_roll()
 Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT distinct StudentID From StudentInfo " + _
"WHERE (StudentID  IN(SELECT StudentID FROM StudentAdmission where Approval='Y'))")
ComStuId.Clear
If Not rs.EOF Then
    Do Until rs.EOF
        ComStuId.AddItem rs(0)
        rs.MoveNext
    Loop
End If

End Sub
Private Sub get_class()
'Dim rs1 As New ADODB.Recordset
'ComboClass.Clear
'Set rs1 = getdata("select ClassId,ClassName from ClassInfo")
'If Not rs1.EOF Then
'    Do Until rs1.EOF
'            ComboClass.AddItem rs1(0) + " - " + rs1(1)
'            rs1.MoveNext
'    Loop
'End If
End Sub



Private Sub MaskEdBoxDate_GotFocus(Index As Integer)
      Select Case Index
             Case Index
             MaskEdBoxDate(Index).SelStart = 0
             MaskEdBoxDate(Index).SelLength = Len(MaskEdBoxDate(Index))
      End Select
End Sub

Private Sub ComStuId_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(ComStuId) <> 0 Then
              If Mid(ComStuId, 1, 3) <> "STI" Then
                   ComStuId.Text = Format(ComStuId, "000000")
                   ComStuId.Text = "STI-" + ComStuId.Text
                End If
              
    End If
   ComStuId_Click
   MaskEdBoxDate(0).SetFocus
    
End If
End Sub
Private Sub ShowFlexData()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Dim total As Integer
total = 0
'format_grid
 Set rs = getdata("select L_st_dt,L_ed_dt,Remarks,Serial from Student_leave_Info where StudentId='" & ComStuId & "' and AcademicYr='" & cboAcaYr & "' order by Serial desc ")
 If Not rs.EOF Then
    format_grid
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = "" & rs!L_st_dt
                .TextMatrix(i, 2) = "" & rs!L_ed_dt
                .ColAlignment(3) = 0
                .TextMatrix(i, 3) = (CDate(rs!L_ed_dt) - CDate(rs!L_st_dt)) + 1
                .ColAlignment(4) = 0
                .TextMatrix(i, 4) = "" & rs!remarks
                .TextMatrix(i, 5) = "" & rs!serial
                total = total + .TextMatrix(i, 3)
            rs.MoveNext
           i = i + 1
        Loop
        .Row = i + 1
        .Col = 2
        .CellFontBold = True
        .CellForeColor = vbBlue
        
        .Text = "        Total ="
        .Row = i + 1
        .Col = 3
        .CellFontBold = True
        .CellForeColor = vbBlue
        If Val(total) = 1 Then
          .Text = Val(total) & "  Day"
        Else
          .Text = Val(total) & "  Days"
        End If
        
    End With
 Else
     MSFlexGrid1.Clear
     format_grid
 End If

Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub MaskEdBoxDate_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
       Case Index
          If KeyAscii = 13 Then
           If MaskEdBoxDate(Index) <> "__/__/__" Then
                   If Check_ValidDate(MaskEdBoxDate(Index)) = False Then
                       MaskEdBoxDate(Index).SetFocus
                       Exit Sub
                   Else
                      If Format(MaskEdBoxDate(Index), "yyyy") <> cboAcaYr.Text Then
                         MsgBox "Year Mismatch...Please Verify", vbInformation, cmp
                         Exit Sub
                       End If
                   End If
           End If
           If Index = 0 Then
             MaskEdBoxDate(1).SetFocus
           Else
             Dim d As Integer
             d = CDate(MaskEdBoxDate(1).Text) - CDate(MaskEdBoxDate(0))
             Label5.Caption = "Total : " & d + 1 & " days"
             txtCause.SetFocus
           End If
End If
End Select
End Sub

Private Sub mnuClose_Click()
  Unload Me
End Sub

Private Sub mnuDelete_Click()
  If Len(txtSerial) = 0 Then
    MsgBox "Please select a Date range from grid below...", vbInformation, cmp
    Exit Sub
  End If
  If MsgBox("Are you Sure to Delete ?", vbInformation + vbYesNo, cmp) = vbYes Then
      s_u_d_leave_info (3)
      MsgBox "Deleted Successfully.", vbInformation, "Student Management System"
    ShowFlexData
 Else
    Exit Sub
    
 End If
  
End Sub

Private Sub mnuRef_Click()
  txtSerial = ""
End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo errdes

If MSFlexGrid1.Rows > 1 Then
  If Len(Trim(Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1), "dd/mm/yy"))) = 0 Then
     MaskEdBoxDate(0).Text = "__/__/__"
     MaskEdBoxDate(1).Text = "__/__/__"
     txtSerial = ""
     txtCause = ""
     Exit Sub
  End If
    MaskEdBoxDate(0).Text = Trim(Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1), "dd/mm/yy"))
    MaskEdBoxDate(1).Text = Trim(Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2), "dd/mm/yy"))
    txtCause = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
    txtSerial = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
    
Else
    Exit Sub
End If
Exit Sub
errdes:
  MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       PopupMenu mnuDel, 2
    End If
End Sub

Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub

Private Sub txtCause_LostFocus()
    cmdAdmitted.SetFocus
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
       Case 1
            If KeyAscii = 13 Then
                   MaskEdBoxDate(0).SetFocus
             End If
   End Select
End Sub
