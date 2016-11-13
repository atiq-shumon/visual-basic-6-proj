VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmstudentREAdmission 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdmitted 
      BackColor       =   &H8000000C&
      Caption         =   "Admit"
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
      Left            =   6360
      TabIndex        =   30
      ToolTipText     =   "Click To Admitted"
      Top             =   8190
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Admission To"
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
      Height          =   1425
      Left            =   0
      TabIndex        =   9
      Top             =   2250
      Width           =   9885
      Begin VB.ComboBox cmdAcaYear 
         Height          =   315
         Left            =   8130
         Style           =   2  'Dropdown List
         TabIndex        =   28
         ToolTipText     =   "Select Section"
         Top             =   390
         Width           =   1425
      End
      Begin VB.ComboBox ComboShift 
         Height          =   315
         ItemData        =   "frmstudentaREAdmission.frx":0000
         Left            =   4530
         List            =   "frmstudentaREAdmission.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Select Shift"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox ComboClass 
         Height          =   315
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Select Class"
         Top             =   360
         Width           =   2025
      End
      Begin VB.ComboBox ComboSection 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Select Section"
         Top             =   900
         Width           =   2085
      End
      Begin VB.TextBox txtFields 
         Height          =   315
         Index           =   1
         Left            =   4560
         TabIndex        =   17
         ToolTipText     =   "Insert Roll"
         Top             =   930
         Width           =   1515
      End
      Begin MSMask.MaskEdBox MaskEdBoxDate 
         Height          =   315
         Left            =   8130
         TabIndex        =   19
         ToolTipText     =   "Insert  Admission Date"
         Top             =   930
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
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
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   6690
         TabIndex        =   29
         Top             =   420
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   390
         Width           =   525
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   930
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   6660
         TabIndex        =   12
         Top             =   990
         Width           =   1320
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   3840
         TabIndex        =   11
         Top             =   390
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Roll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   3810
         TabIndex        =   10
         Top             =   960
         Width           =   405
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4455
      Left            =   0
      TabIndex        =   7
      Top             =   3660
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   7858
      _Version        =   393216
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
      Caption         =   "Cancel"
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
      Left            =   7740
      TabIndex        =   1
      ToolTipText     =   "Click to cancel Admission"
      Top             =   8190
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   405
      Left            =   8880
      TabIndex        =   2
      ToolTipText     =   "Click to Exit"
      Top             =   8190
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Information"
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
      Height          =   1425
      Left            =   0
      TabIndex        =   4
      Top             =   810
      Width           =   9945
      Begin VB.TextBox txtFields 
         BackColor       =   &H80000018&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   5
         Left            =   8130
         TabIndex        =   32
         Top             =   360
         Width           =   1605
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H80000018&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   4
         Left            =   8160
         TabIndex        =   26
         Top             =   930
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H80000018&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   3
         Left            =   5490
         TabIndex        =   24
         Top             =   930
         Width           =   1845
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H80000018&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   2
         Left            =   2670
         TabIndex        =   22
         Top             =   930
         Width           =   1965
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H80000018&
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   1050
         TabIndex        =   20
         Top             =   930
         Width           =   795
      End
      Begin VB.ComboBox ComStuId 
         Height          =   315
         Left            =   1050
         TabIndex        =   0
         ToolTipText     =   "Select student"
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A.Year"
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
         Left            =   7380
         TabIndex        =   31
         Top             =   420
         Width           =   585
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
         TabIndex        =   27
         Top             =   990
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
         TabIndex        =   25
         Top             =   960
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
         TabIndex        =   23
         Top             =   990
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
         TabIndex        =   21
         Top             =   990
         Width           =   345
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   4665
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
         TabIndex        =   5
         Top             =   390
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   9825
      TabIndex        =   3
      Top             =   0
      Width           =   9885
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Re-Admission Entry"
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
         TabIndex        =   8
         Top             =   150
         Width           =   3240
      End
      Begin VB.Image Image1 
         Height          =   930
         Left            =   0
         Picture         =   "frmstudentaREAdmission.frx":0035
         Stretch         =   -1  'True
         Top             =   -60
         Width           =   9915
      End
   End
End
Attribute VB_Name = "frmstudentREAdmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAcaYear_KeyPress(KeyAscii As Integer)
       If KeyAscii = 13 Then
                      If MaskEdBox2 <> "__/__/__" Then
                        If Check_ValidDate(MaskEdBoxDate.Text) = False Then
                              MaskEdBoxDate.SetFocus
                              Exit Sub
                           End If
                      End If
             txtfields(1).SetFocus
       End If
  
End Sub

Private Sub cmdAcaYear_LostFocus()
     get_roll
     ShowFlexData
End Sub

Private Sub cmdAdmitted_Click()
  If Len(ComStuId) = 0 And Len(ComboClass) = 0 Then Exit Sub
  If Len(ComStuId) = 0 Then
      MsgBox "Please Enter Student ID. ", vbInformation, cmp
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

  If MaskEdBoxDate = "__/__/__" Then
      MsgBox "Enter Date.", vbInformation, cmp
      MaskEdBoxDate.SetFocus
      Exit Sub
  End If
  If Len(txtfields(0)) = 0 Then
      MsgBox "Invalid Student Id..Please Verify", vbInformation, cmp
      ComStuId.SetFocus
      Exit Sub
  End If
'If Len(ComboSection) <> 0 And Len(ComboShift) <> 0 And Len(ComboClass) <> 0 And Len(cmdAcaYear) <> 0 Then
'  Dim rs1 As New ADODB.Recordset
'     Set rs1 = getdata("select ClassRoll from StudentAdmission where classid='" & Mid(ComboClass, 1, 5) & "' and shift='" & Mid(ComboShift, 1, 1) & "'and sectionId='" & Mid(ComboSection, 1, 5) & "' and aca_yr='" & Trim(cmdAcaYear) & "'")
'      If Not rs1.EOF Then
'        txtfields(1) = IIf(IsNull(rs1(0)) = True, "1", rs1(0))
'    Else
'        txtfields(1) = "1"
'    End If
' End If
 
 If Len(ComboSection) <> 0 And Len(ComboShift) <> 0 And Len(ComboClass) <> 0 And Len(cmdAcaYear) <> 0 Then
  Dim rs2 As New ADODB.Recordset
     Set rs2 = getdata("select ClassRoll from StudentAdmission where classid='" & Mid(ComboClass, 1, 5) & "' and shift='" & Mid(ComboShift, 1, 1) & "'and sectionId='" & Mid(ComboSection, 1, 5) & "' and aca_yr='" & Trim(cmdAcaYear) & "' and ClassRoll='" & Trim(txtfields(1).Text) & "'")
      If Not rs2.EOF Then
        MsgBox "Same Roll No is occupied by another Student...Please Chose a another", vbInformation, cmp
        txtfields(1).SetFocus
        Exit Sub
      End If
        
 End If
 
  If Len(ComboSection) <> 0 And Len(ComboShift) <> 0 And Len(ComboClass) <> 0 And Len(cmdAcaYear) <> 0 Then
  Dim rs3 As New ADODB.Recordset
     Set rs3 = getdata("select ClassRoll from StudentAdmission where aca_yr='" & Trim(cmdAcaYear) & "' and studentId='" & Trim(ComStuId) & "'")
      If Not rs3.EOF Then
        MsgBox "This Student is already admitted...Please Chose a another", vbInformation, cmp
        ComStuId.SetFocus
        Exit Sub
      End If
        
 End If



Dim cmd As New ADODB.Command
Dim con As New ADODB.connection

con.Open GConnString
Set rs = getdata("select StudentId from StudentAdmission where studentId='" & ComStuId & "' and classid='" & Trim(Mid(ComboClass, 1, 5)) & "' and aca_yr='" & Trim(cmdAcaYear) & "'")
If Not rs.EOF Then
    If MsgBox("Information Inserted Previously ,Do you want to update the Admission Information this student? ", vbYesNo + vbInformation) = vbYes Then
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "StuAdmissionEvaluationInformation"
            cmd(1) = 1
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
            cmd(15) = "R"
            cmd.Execute
            MsgBox "Updated Successfully.", vbInformation, "Student Management System"
            ShowFlexData
'            cmdAdmitted.Enabled = False
    Else
            Exit Sub
    End If
Else
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "StuAdmissionEvaluationInformation"
            cmd(1) = 1
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
            cmd(15) = "R"
         cmd.Execute
    MsgBox "Saved Successfully.", vbInformation, "Student Management System"
    ShowFlexData

End If
ComStuId.SetFocus
Exit Sub
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
If MaskEdBoxDate = "__/__/__" Then
    MsgBox "Enter Date.", vbInformation, cmp
    MaskEdBoxDate.SetFocus
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
            cmd(15) = "R"
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
   " studentAdmission s,studentInfo i,classInfo c where s.StudentId='" & Trim(ComStuId.Text) & "' and s.ClassId=c.ClassId and s.StudentId=i.StudentId and s.serial_no=(select max(serial_no) from studentAdmission where StudentId='" & Trim(ComStuId.Text) & "')")
   If Not rs.EOF Then
     txtfields(0) = rs!ClassRoll
     Label3.Caption = rs!StudentName
     Set sec_rs = getdata("select Sectiondsc from sectionInfo where sectionId='" & Trim(rs!SectionID) & "' and classId='" & Trim(rs!classId) & "'")
     If Not sec_rs.EOF Then
        txtfields(2).Text = sec_rs(0)
     End If
     txtfields(3).Text = IIf(rs!Shift = "M", "Morning", "Day")
     txtfields(4).Text = rs!ClassName
     txtfields(5).Text = rs!Aca_yr
  Else
     txtfields(0) = ""
     Label3.Caption = ""
     txtfields(2).Text = ""
     txtfields(3).Text = ""
     txtfields(4).Text = ""
     txtfields(5).Text = ""
  End If
End Sub
Private Sub Form_Load()
   With MSFlexGrid1
        .Rows = 1
        .Cols = 3
        .Col = 0: .Text = " Student ID   #"
        .Col = 1: .Text = "Student Name   "
        .Col = 2: .Text = " Roll No  "
        .ColWidth(0) = 3000
        .ColWidth(1) = 5000
        .ColWidth(2) = 1540
        .ColAlignment(2) = 0
   End With

ShowFlexData
load_roll
load_Aca_year
get_class
End Sub
Private Sub load_Aca_year()
  Dim i As Integer
  For i = 2000 To 2050
     cmdAcaYear.AddItem i
  Next i
  cmdAcaYear.Text = Format(Date, "YYYY")
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
Dim rs1 As New ADODB.Recordset
ComboClass.Clear
Set rs1 = getdata("select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
            ComboClass.AddItem rs1(0) + " - " + rs1(1)
            rs1.MoveNext
    Loop
End If
End Sub
Private Sub MaskEdBoxDate_GotFocus()
        MaskEdBoxDate.SelStart = 0
        MaskEdBoxDate.SelLength = Len(MaskEdBoxDate)
End Sub

Private Sub MaskEdBoxDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MaskEdBoxDate <> "__/__/__" Then
            If Check_ValidDate(MaskEdBoxDate) = False Then
                MaskEdBoxDate.SetFocus
                Exit Sub
            End If
    End If
    cmdAdmitted.SetFocus
End If
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
   ComboClass.SetFocus
    
End If
End Sub
Private Sub ShowFlexData()

On Error GoTo errdes
Dim rs As New ADODB.Recordset
MSFlexGrid1.Clear
 Set rs = getdata("select s.studentID, i.StudentName ,s.ClassRoll from " + _
   " studentAdmission s ,studentInfo i where s.studentID=i.studentID and s.classid='" & Mid(ComboClass, 1, 5) & "' " + _
   "and s.shift='" & Mid(ComboShift, 1, 1) & "'and s.sectionId='" & Mid(ComboSection, 1, 5) & "' " + _
   "and s.aca_yr='" & Trim(cmdAcaYear) & "'and s.serial_no=(select max(serial_no) from studentAdmission where StudentId=s.StudentId) order by s.ClassRoll")
  
If Not rs.EOF Then
    i = 1
    With MSFlexGrid1
        Do Until rs.EOF
            MSFlexGrid1.Rows = i + 1
                .TextMatrix(i, 0) = "" & rs!studentid
                .TextMatrix(i, 1) = "" & rs!StudentName
                .TextMatrix(i, 2) = "" & rs!ClassRoll
            rs.MoveNext
           i = i + 1
        Loop
    End With
 Else
     MSFlexGrid1.Rows = 50

 End If

Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo errdes
If MSFlexGrid1.Rows > 1 Then
    ComStuId = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    ComStuId_Click
'    Label3.Caption = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
'    ComboClass = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
'    ComboSection = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
'    ComboShift = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
'    txtFields(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
'    MaskEdBoxDate = Format(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6), "dd/mm/yy")
Else
    Exit Sub
End If
Exit Sub
errdes:
'MsgBox err.Description, vbInformation, App.Title

End Sub

Private Sub MSFlexGrid1_SelChange()
  MSFlexGrid1_Click
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
       Case 1
            If KeyAscii = 13 Then
                   MaskEdBoxDate.SetFocus
             End If
   End Select
End Sub
