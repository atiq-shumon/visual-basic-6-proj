VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClassRoutine 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   Icon            =   "frmClassRoutine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   765
      Left            =   0
      TabIndex        =   28
      Top             =   7350
      Width           =   11925
      Begin VB.TextBox txtTrackid 
         Height          =   315
         Left            =   2340
         TabIndex        =   38
         Top             =   300
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000C&
         Caption         =   "Delete Subject"
         Height          =   405
         Left            =   9630
         TabIndex        =   36
         ToolTipText     =   "Click to Delete"
         Top             =   270
         Width           =   1245
      End
      Begin VB.CommandButton CmdPrint 
         BackColor       =   &H008080FF&
         Caption         =   "View Report"
         Height          =   405
         Left            =   180
         MaskColor       =   &H008080FF&
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Click to Print"
         Top             =   240
         Width           =   1515
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H8000000C&
         Caption         =   "Add Subject"
         Height          =   405
         Left            =   8340
         TabIndex        =   10
         ToolTipText     =   "Click to Save"
         Top             =   270
         Width           =   1245
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H8000000C&
         Caption         =   "Close"
         Height          =   405
         Left            =   10890
         TabIndex        =   29
         ToolTipText     =   "Click to Close"
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   4875
      Left            =   690
      TabIndex        =   27
      Top             =   2550
      Width           =   11205
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4875
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   8599
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColor       =   15005934
         BackColorSel    =   -2147483624
         ForeColorSel    =   16711680
         BackColorBkg    =   15724265
         WordWrap        =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5685
      Left            =   -30
      TabIndex        =   26
      Top             =   2460
      Width           =   735
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   4260
         Left            =   60
         TabIndex        =   32
         Top             =   390
         Width           =   555
      End
      Begin VB.Label Label11 
         Caption         =   "Srl#"
         Height          =   225
         Left            =   120
         TabIndex        =   31
         Top             =   150
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1155
      Left            =   0
      TabIndex        =   17
      Top             =   1410
      Width           =   11895
      Begin VB.CommandButton CmdDeleteSerial 
         BackColor       =   &H8000000C&
         Caption         =   "Delete Serial "
         Height          =   405
         Left            =   10290
         TabIndex        =   35
         ToolTipText     =   "Click to Delete Serial "
         Top             =   630
         Width           =   1395
      End
      Begin VB.CommandButton cmdGenerate 
         BackColor       =   &H8000000C&
         Caption         =   "Generate New Serial "
         Height          =   405
         Left            =   8640
         TabIndex        =   34
         ToolTipText     =   "Click to Generate New Serial "
         Top             =   630
         Width           =   1635
      End
      Begin VB.ComboBox CboSubject 
         Height          =   315
         Left            =   690
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   210
         Width           =   2565
      End
      Begin VB.ComboBox CboDays 
         Height          =   315
         ItemData        =   "frmClassRoutine.frx":0442
         Left            =   4020
         List            =   "frmClassRoutine.frx":045B
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   210
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   690
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   2565
      End
      Begin MSComCtl2.DTPicker DtpicStart 
         Height          =   315
         Left            =   6630
         TabIndex        =   6
         ToolTipText     =   "Select Start time"
         Top             =   210
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49414146
         CurrentDate     =   38623
      End
      Begin MSComCtl2.DTPicker dtpickend 
         Height          =   315
         Left            =   9630
         TabIndex        =   7
         ToolTipText     =   "Select End time"
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   49414146
         CurrentDate     =   38623
      End
      Begin MSMask.MaskEdBox MaskEdBoxDate 
         Height          =   315
         Left            =   6630
         TabIndex        =   9
         ToolTipText     =   "Insert  Effective Date"
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E.Date"
         Height          =   195
         Left            =   6120
         TabIndex        =   37
         Top             =   780
         Width           =   495
      End
      Begin VB.Shape Shape1 
         Height          =   465
         Left            =   8610
         Top             =   600
         Width           =   3105
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   285
         Left            =   3810
         TabIndex        =   25
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
         Height          =   195
         Left            =   3420
         TabIndex        =   24
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub"
         Height          =   195
         Left            =   90
         TabIndex        =   23
         Top             =   210
         Width           =   285
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teacher"
         Height          =   195
         Left            =   30
         TabIndex        =   20
         Top             =   750
         Width           =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Time"
         Height          =   195
         Left            =   8640
         TabIndex        =   19
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Time"
         Height          =   195
         Left            =   5610
         TabIndex        =   18
         Top             =   270
         Width           =   930
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   11895
      TabIndex        =   14
      Top             =   0
      Width           =   11955
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   15
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Routine"
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
         Left            =   5310
         TabIndex        =   22
         Top             =   210
         Width           =   1605
      End
      Begin VB.Image Image1 
         Height          =   990
         Left            =   -60
         Picture         =   "frmClassRoutine.frx":049F
         Stretch         =   -1  'True
         Top             =   -90
         Width           =   11985
      End
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   0
      TabIndex        =   11
      Top             =   750
      Width           =   11925
      Begin VB.ComboBox Combo5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1485
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   4050
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select Section"
         Top             =   240
         Width           =   1545
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmClassRoutine.frx":D344
         Left            =   690
         List            =   "frmClassRoutine.frx":D346
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select Class"
         Top             =   240
         Width           =   2565
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmClassRoutine.frx":D348
         Left            =   6630
         List            =   "frmClassRoutine.frx":D358
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select Shift"
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Academic Yr "
         Height          =   195
         Index           =   1
         Left            =   8580
         TabIndex        =   21
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         Height          =   195
         Index           =   0
         Left            =   3450
         TabIndex        =   16
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class "
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   6150
         TabIndex        =   12
         Top             =   270
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmClassRoutine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ClickFromList As Boolean
Dim selectedText As String

Private Sub AddAll_Click()
List3.Clear
List3.TopIndex = 0
For i = 0 To List2.ListCount - 1
    List3.AddItem List2.List(i)
Next
End Sub

Private Sub AddOne_Click()
List3.AddItem List2.Text
End Sub

Private Sub CmdDeleteSerial_Click()
  If Len(Combo1.Text) = 0 Then
    MsgBox "Please Select Class ", vbInformation, "School Management System"
    Combo1.SetFocus
    Exit Sub
End If
If Len(Combo2.Text) = 0 Then
    MsgBox "Please Select Shift ", vbInformation, "School Management System"
    Combo2.SetFocus
    Exit Sub
End If

If Len(List1.Text) = 0 Then
      MsgBox "Please Select serial ", vbInformation, "School Management System"
      List1.SetFocus
    Exit Sub
End If

If Len(Combo5.Text) = 0 Then
    MsgBox "Please Select an Academic Year ", vbInformation, "School Management System"
    Combo5.SetFocus
    Exit Sub
End If
If Len(Combo3.Text) = 0 Then
    MsgBox "Please Select Section ", vbInformation, "School Management System"
    Combo3.SetFocus
    Exit Sub
End If

If MsgBox("Are the sure to delete the whole class routine?", vbCritical + vbYesNo, cmp) = vbYes Then
    Dim cmd As New ADODB.Command
    Dim con As New ADODB.connection
    con.Open GConnString
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ClassRoutine1"
    cmd(1) = 3
    cmd(2) = Get_Code(Trim(Combo1.Text))
    cmd(3) = Mid(Trim(Combo2.Text), 1, 1)
    cmd(4) = Get_Code(Trim(Combo3.Text))
    cmd(5) = Trim(CboDays.Text)
    cmd(6) = Get_Code(Trim(CboSubject))
    cmd(7) = Format(DtpicStart.Value, "hh:mm:ss AM/PM")
    cmd(8) = Format(dtpickend.Value, "hh:mm:ss AM/PM")
    cmd(9) = Trim(Combo4.Text)
    cmd(10) = Trim(soft_user)
    cmd(11) = Format(Date, "dd mmm yyyy")
    cmd(12) = Trim(Combo5.Text)
    cmd(13) = Trim(List1)
    cmd.Execute
    MsgBox "Deleted Successfully.", vbInformation, "Student Management System"
    Combo1.SetFocus
    load_serial
    List1_Click
  End If
Exit Sub

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
  If Len(Combo1.Text) = 0 Then
    MsgBox "Please Select Class ", vbInformation, "School Management System"
    Combo1.SetFocus
    Exit Sub
End If
If Len(Combo2.Text) = 0 Then
    MsgBox "Please Select Shift ", vbInformation, "School Management System"
    Combo2.SetFocus
    Exit Sub
End If
If Len(Combo5.Text) = 0 Then
    MsgBox "Please Select an Academic Year ", vbInformation, "School Management System"
    Combo5.SetFocus
    Exit Sub
End If
If Len(Combo3.Text) = 0 Then
    MsgBox "Please Select Section ", vbInformation, "School Management System"
    Combo3.SetFocus
    Exit Sub
End If

Screen.MousePointer = vbHourglass
rptMode = 12
frmViewer.Show 1

End Sub

Private Sub cmdSAVE_Click()
If Len(Combo1.Text) = 0 Then
    MsgBox "Please Select Class ", vbInformation, "School Management System"
    Combo1.SetFocus
    Exit Sub
End If
If Len(Combo2.Text) = 0 Then
    MsgBox "Please Select Shift ", vbInformation, "School Management System"
    Combo2.SetFocus
    Exit Sub
End If
If Len(Combo5.Text) = 0 Then
    MsgBox "Please Select an Academic Year ", vbInformation, "School Management System"
    Combo5.SetFocus
    Exit Sub
End If
If Len(Combo3.Text) = 0 Then
    MsgBox "Please Select Section ", vbInformation, "School Management System"
    Combo3.SetFocus
    Exit Sub
End If
If Len(Combo4.Text) = 0 Then
    MsgBox "Please Select Teacher ", vbInformation, "School Management System"
    Combo4.SetFocus
    Exit Sub
End If

If Len(CboDays) = 0 Then
   MsgBox "Please Select a day ", vbInformation, "School Management System"
   CboDays.SetFocus
   Exit Sub
End If
If Len(List1.Text) = 0 Then
   MsgBox "Please Select a serial from the list", vbInformation, "School Management System"
   List1.SetFocus
   Exit Sub
End If

If CDate(Format(DtpicStart.Value, "hh:mm:ss AM/PM")) >= CDate(Format(dtpickend.Value, "hh:mm:ss AM/PM")) Then
   MsgBox "Invalid Time...Please varify", vbInformation, cmp
   DtpicStart.SetFocus
   Exit Sub
End If

Dim rs As New ADODB.Recordset
Set rs = getdata("select subjectid from classRoutine where classId='" & Get_Code(Combo1) & "' and Shift='" & Mid(Combo2, 1, 1) & "' and SectionId='" & Get_Code(Combo3) & "' and listofday='" & Trim(CboDays.Text) & "' and academic_yr='" & Trim(Combo5) & "' and SerialNo='" & Trim(List1.Text) & "' and  CONVERT(varchar, Starttime, 8) between '" & Format(DtpicStart, "hh:mm:ss") & "' and '" & Format(dtpickend, "hh:mm:ss") & "' ")
If Not rs.EOF Then
    MsgBox "This Start time is assigned for another Subject for the day.", vbCritical, "School Management System"
    DtpicStart = "00:00:00"
    DtpicStart.SetFocus
    Exit Sub
End If

Set rs = getdata("select subjectid from classRoutine where classId='" & Get_Code(Combo1) & "' and Shift='" & Mid(Combo2, 1, 1) & "' and SectionId='" & Get_Code(Combo3) & "' and listofday='" & Trim(CboDays.Text) & "' and academic_yr='" & Trim(Combo5) & "' and SerialNo='" & Trim(List1.Text) & "' and  CONVERT(varchar, EndTime, 8) between '" & Format(DtpicStart, "hh:mm:ss") & "' and '" & Format(dtpickend, "hh:mm:ss") & "' ")
If Not rs.EOF Then
    MsgBox "This End time is assigned for another Subject for the day.Please verify the time ", vbCritical, "School Management System"
    dtpickend = "00:00:00"
    dtpickend.SetFocus
    Exit Sub
End If


Set rs = getdata("select TeacherId from classRoutine where TeacherId='" & Trim(Combo4.Text) & "' and listofday='" & Trim(CboDays.Text) & "' and academic_yr='" & Trim(Combo5) & "' and  CONVERT(varchar, EndTime, 8) between '" & Format(DtpicStart, "hh:mm:ss") & "' and '" & Format(dtpickend, "hh:mm:ss") & "' and SerialNo='" & Trim(List1.Text) & "'")
If Not rs.EOF Then
    MsgBox "This teacher is already assigned for another Subject for the day of this time.Please verify the time ", vbCritical, "School Management System"
    dtpickend = "00:00:00"
    dtpickend.SetFocus
    Exit Sub
End If

Set rs = getdata("select TeacherId from classRoutine where TeacherId='" & Trim(Combo4.Text) & "' and listofday='" & Trim(CboDays.Text) & "' and academic_yr='" & Trim(Combo5) & "' and  CONVERT(varchar, Starttime, 8) between '" & Format(DtpicStart, "hh:mm:ss") & "' and '" & Format(dtpickend, "hh:mm:ss") & "' and SerialNo='" & Trim(List1.Text) & "'")
If Not rs.EOF Then
    MsgBox "This teacher is already assigned for another Subject for the day of this time.Please verify the time ", vbCritical, "School Management System"
    DtpicStart = "00:00:00"
    DtpicStart.SetFocus
    Exit Sub
End If

    Dim cmd As New ADODB.Command
    Dim con As New ADODB.connection
    con.Open GConnString
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ClassRoutine1"
    cmd(1) = 2
    cmd(2) = Get_Code(Trim(Combo1.Text))
    cmd(3) = Mid(Trim(Combo2.Text), 1, 1)
    cmd(4) = Get_Code(Trim(Combo3.Text))
    cmd(5) = Trim(CboDays.Text)
    cmd(6) = Get_Code(Trim(CboSubject))
    cmd(7) = Format(DtpicStart.Value, "hh:mm:ss AM/PM")
    cmd(8) = Format(dtpickend.Value, "hh:mm:ss AM/PM")
    cmd(9) = Trim(Combo4.Text)
    cmd(10) = Trim(soft_user)
    cmd(11) = Format(MaskEdBoxDate, "dd mmm yyyy")
    cmd(12) = Trim(Combo5.Text)
    cmd(13) = Trim(List1.Text)
    cmd(14) = 1
    cmd.Execute
    MsgBox "Save Successfully.", vbInformation, "Student Management System"
    Combo1.SetFocus
    load_serial
    List1_Click
Exit Sub
End Sub

Private Sub Combo1_Click()

Combo3.Clear
load_subject
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("select SectionId,Sectiondsc from SectionInfo where ClassId='" & Get_Code((Combo1.Text)) & "'")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo3.AddItem rs1(1) + " ~ " + rs1(0)
        rs1.MoveNext
    Loop
    Combo3.AddItem (" ")

End If
End Sub
Private Sub load_subject()
 Dim rs As New ADODB.Recordset
 Set rs = getdata("SELECT  Sub_code,Sub_title From Subject_Info_sub WHERE Class_code = '" & Get_Code(Combo1) & "'")
 CboSubject.Clear
 If Not rs.EOF Then
    Do Until rs.EOF
        CboSubject.AddItem rs!Sub_title & "~" & rs!Sub_code
        rs.MoveNext
    Loop

 End If
End Sub
Private Sub Combo2_Click()

'Dim rs1 As New ADODB.Recordset
'Set rs1 = getdata("select SubjectId from Classroutine  where classid='" & Mid((Combo1.Text), 1, 5) & "'and shift='" & (Combo2.Text) & "'and listofday='" & Trim(List1.Text) & "'and sectionId='" & Mid((Combo3.Text), 1, 5) & "' ")
'If Not rs1.EOF Then
'    Do Until rs1.EOF
'        Set rs2 = getdata("select Subjectdsc from SubjectInfo  where classid='" & Mid((Combo1.Text), 1, 5) & "' and  SubjectId='" & rs1!SubjectID & "' ")
'        List3.AddItem rs1!SubjectID + "-" + rs2!SubjectDsc
'        rs1.MoveNext
'    Loop
'Else
'    List3.Clear
'End If
End Sub

Private Sub Combo4_Click()
   load_teacher_title
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'    cmdSave.SetFocus
End If
End Sub

Private Sub Combo5_Click()
   load_serial
End Sub
Private Sub load_serial()
  List1.Clear
  Dim rs As New ADODB.Recordset
  Set rs = getdata("SELECT distinct SerialNo  FROM ClassRoutine where classid='" & Get_Code(Combo1) & "' and  SectionId='" & Get_Code(Combo3) & "'and shift='" & Mid(Trim(Combo2), 1, 1) & "' and academic_yr='" & Trim(Combo5) & "'order by SerialNo desc")
  If Not rs.EOF Then
    cmdSave.Enabled = True
     Do Until rs.EOF
       List1.AddItem rs(0)
       rs.MoveNext
     Loop
  Else
    cmdSave.Enabled = False
  End If

End Sub
Private Sub cmdGenerate_Click()
If Len(Combo1.Text) = 0 Then
    MsgBox "Please Select Class ", vbInformation, "School Management System"
    Combo1.SetFocus
    Exit Sub
End If
If Len(Combo2.Text) = 0 Then
    MsgBox "Please Select Shift ", vbInformation, "School Management System"
    Combo2.SetFocus
    Exit Sub
End If
If Len(Combo5.Text) = 0 Then
    MsgBox "Please Select an Academic Year ", vbInformation, "School Management System"
    Combo5.SetFocus
    Exit Sub
End If
If Len(Combo3.Text) = 0 Then
    MsgBox "Please Select Section ", vbInformation, "School Management System"
    Combo3.SetFocus
    Exit Sub
End If
If Len(Combo4.Text) = 0 Then
    MsgBox "Please Select Teacher ", vbInformation, "School Management System"
    Combo4.SetFocus
    Exit Sub
End If

If Len(CboDays) = 0 Then
   MsgBox "Please Select a day ", vbInformation, "School Management System"
   CboDays.SetFocus
   Exit Sub
End If

If DtpicStart.Value >= dtpickend.Value Then
   MsgBox "Invalid Time...Please varify", vbInformation, cmp
   DtpicStart.SetFocus
   Exit Sub
End If

If MaskEdBoxDate = "__/__/__" Then
   MsgBox "Effective Date Required..Please Verify ", vbInformation, "School Management System"
   MaskEdBoxDate.SetFocus
   Exit Sub
End If


If Val(Mid(MaskEdBoxDate.Text, 7, 8)) <> Val(Mid(Combo5, 3, 4)) Then
   MsgBox "Academic Year Conflicts....", vbInformation, cmp
    Combo5.SetFocus
   Exit Sub
End If

    Dim cmd As New ADODB.Command
    Dim con As New ADODB.connection
    con.Open GConnString
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ClassRoutine1"
    cmd(1) = 1
    cmd(2) = Get_Code(Trim(Combo1.Text))
    cmd(3) = Mid(Trim(Combo2.Text), 1, 1)
    cmd(4) = Get_Code(Trim(Combo3.Text))
    cmd(5) = Trim(CboDays.Text)
    cmd(6) = Get_Code(Trim(CboSubject))
    cmd(7) = Format(DtpicStart.Value, "hh:mm:ss AM/PM")
    cmd(8) = Format(dtpickend.Value, "hh:mm:ss AM/PM")
    cmd(9) = Trim(Combo4.Text)
    cmd(10) = Trim(soft_user)
    cmd(11) = Format(MaskEdBoxDate, "dd mmm yyyy")
    cmd(12) = Trim(Combo5.Text)
    cmd(13) = 1
    cmd(14) = 1
    cmd.Execute
    MsgBox "Save Successfully.", vbInformation, "Student Management System"
    Combo1.SetFocus
    load_serial
    List1_Click
Exit Sub

End Sub
Private Sub Command1_Click()
 If Len(Combo1.Text) = 0 Then
    MsgBox "Please Select Class ", vbInformation, "School Management System"
    Combo1.SetFocus
    Exit Sub
End If
If Len(Combo2.Text) = 0 Then
    MsgBox "Please Select Shift ", vbInformation, "School Management System"
    Combo2.SetFocus
    Exit Sub
End If
If Len(Combo5.Text) = 0 Then
    MsgBox "Please Select an Academic Year ", vbInformation, "School Management System"
    Combo5.SetFocus
    Exit Sub
End If
If Len(Combo3.Text) = 0 Then
    MsgBox "Please Select Section ", vbInformation, "School Management System"
    Combo3.SetFocus
    Exit Sub
End If

If Len(List1.Text) = 0 Then
   MsgBox "Please Select a serial from the list", vbInformation, "School Management System"
   List1.SetFocus
   Exit Sub
End If

If Len(txtTrackid.Text) = 0 Then
   MsgBox "Please clik on the grid to select a row", vbInformation, "School Management System"
   MSFlexGrid1.SetFocus
   Exit Sub
End If



    Dim cmd As New ADODB.Command
    Dim con As New ADODB.connection
    con.Open GConnString
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ClassRoutine1"
    cmd(1) = 3
    cmd(2) = Get_Code(Trim(Combo1.Text))
    cmd(3) = Mid(Trim(Combo2.Text), 1, 1)
    cmd(4) = Get_Code(Trim(Combo3.Text))
    cmd(5) = Trim(CboDays.Text)
    cmd(6) = Get_Code(Trim(CboSubject))
    cmd(7) = Format(DtpicStart.Value, "hh:mm:ss AM/PM")
    cmd(8) = Format(dtpickend.Value, "hh:mm:ss AM/PM")
    cmd(9) = Trim(Combo4.Text)
    cmd(10) = Trim(soft_user)
    cmd(11) = Format(Date, "dd mmm yyyy")
    cmd(12) = Trim(Combo5.Text)
    cmd(13) = Trim(List1.Text)
    cmd(14) = Trim(txtTrackid.Text)
    cmd.Execute
    MsgBox "Deleted Successfully.", vbInformation, "Student Management System"
    Combo1.SetFocus
    load_serial
    List1_Click
Exit Sub

End Sub

Private Sub dtpickend_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Combo4.SetFocus
End If
End Sub

Private Sub DtpicStart_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtpickend.SetFocus
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys (Chr(9))
  End If
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 2000 To 2050
    Combo5.AddItem i
Next i
Combo5.Text = Format(Date, "YYYY")

Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("select TeacherId,TeacherName from TeacherInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Combo4.AddItem rs1(0) + "-" + rs1(1)
        rs1.MoveNext
    Loop
    Combo4.AddItem ""
  ' If Combo4.ListCount > 0 Then Combo4.ListIndex = 0
End If
Set rs1 = getdata("select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
       Combo1.AddItem rs1(1) + "~" + rs1(0)
        rs1.MoveNext
    Loop

End If
load_teacher
With MSFlexGrid1
    .Rows = 1
    .Cols = 9
    .ColWidth(0) = 0
    .Col = 0: .Text = ""
    .ColWidth(1) = 2000
    .Col = 1: .Text = ""
    .ColWidth(2) = 2000
    .Col = 2: .Text = "Saturday"
    .ColWidth(3) = 2000
    .Col = 3: .Text = "Sunday"
    .ColWidth(4) = 2000
    .Col = 4: .Text = "Monday"
    .ColWidth(5) = 2000
    .Col = 5: .Text = "Tuesday"
    .ColWidth(6) = 2000
    .Col = 6: .Text = "Wednesday"
    .ColWidth(7) = 2000
    .Col = 7: .Text = "Thrusday"
    .ColWidth(8) = 2000
    .Col = 8: .Text = "Friday"
    
End With
  
End Sub
Private Sub load_teacher_title()
   Dim rs As New ADODB.Recordset
   Set rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(Combo4) & "'")
   If Not rs.EOF Then
     Label7.Caption = "" & rs!name
   End If
   
End Sub
Private Sub load_teacher()
 Combo4.Clear
  Dim rs As New ADODB.Recordset
  Set rs = getdata("SELECT Emp_id  FROM  Emp_Per_Info")
  If Not rs.EOF Then
     Do Until rs.EOF
       Combo4.AddItem rs(0)
       rs.MoveNext
     Loop
   End If
     
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

Public Function getsubforteacher()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

For i = 0 To List2.ListCount - 1
Set rs1 = getdata("select SubjectId from Classroutine  where classid='" & Mid((Combo1.Text), 1, 5) & "'and shift='" & (Combo2.Text) & "' and  TeacherId='" & Mid((Combo4.Text), 1, 5) & "' and listofday='" & List1.Text & "'")
If Not rs1.EOF Then
    Do Until rs1.EOF
        If Mid((List3.List(i)), 1, 5) = rs1!SubjectID Then
           List3.Selected(i) = True
        End If
       rs1.MoveNext
    Loop
End If
Next
End Function

Public Function getteacherforsub()
'Combo4.Text = ""
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Set rs1 = getdata("select * from Classroutine  where classid='" & Mid((Combo1.Text), 1, 5) & "'and shift='" & (Combo2.Text) & "' and  SubjectId='" & Mid((List3.Text), 1, 5) & "'and listofday='" & List1.Text & "' ")
If Not rs1.EOF Then
    Set rs2 = getdata("select * from TeacherInfo  where teacherid='" & rs1!teacherId & "'")
    If Not rs2.EOF Then
       Combo4.Text = rs1!teacherId + "-" + rs2!teacherName
       DtpicStart = rs1!Starttime
       dtpickend = rs1!Endtime
    Else
        Combo4.Text = ""
        DtpicStart.Value = "00:00:00"
        dtpickend.Value = "00:00:00"
     End If
Else
'    Combo4.Text = ""
    DtpicStart.Value = "00:00:00"
    dtpickend.Value = "00:00:00"
End If
End Function

Private Sub List1_Click()
Dim rs As New ADODB.Recordset
Dim local_rs As New ADODB.Recordset
Dim saturday_rs As New ADODB.Recordset
Dim sunday_rs As New ADODB.Recordset
Dim monday_rs As New ADODB.Recordset
Dim tuesday_rs As New ADODB.Recordset
Dim wednesday_rs As New ADODB.Recordset
Dim thursday_rs As New ADODB.Recordset
Dim friday_rs As New ADODB.Recordset
Dim teacher_name_rs As New ADODB.Recordset
MSFlexGrid1.Clear
With MSFlexGrid1
    .Rows = 1
    .Cols = 10
    .ColWidth(0) = 0
    .Col = 0: .Text = ""
    .ColWidth(1) = 2000
    .Col = 1: .Text = ""
    .ColWidth(2) = 2000
    .Col = 2: .Text = "Saturday"
    .ColWidth(3) = 2000
    .Col = 3: .Text = "Sunday"
    .ColWidth(4) = 2000
    .Col = 4: .Text = "Monday"
    .ColWidth(5) = 2000
    .Col = 5: .Text = "Tuesday"
    .ColWidth(6) = 2000
    .Col = 6: .Text = "Wednesday"
    .ColWidth(7) = 2000
    .Col = 7: .Text = "Thrusday"
    .ColWidth(8) = 2000
    .Col = 8: .Text = "Friday"
    .ColWidth(9) = 0
    .Col = 9: .Text = "Track id"
End With
Set rs = getdata("SELECT  distinct subjectid,Entrydate,trackid from ClassRoutine  where serialno='" & List1.Text & "' and academic_yr='" & Combo5 & "' and Classid='" & Get_Code(Combo1.Text) & "' and  Shift='" & Mid(Combo2, 1, 1) & "' and SectionId='" & Get_Code(Combo3.Text) & "'")

If Not rs.EOF Then
    MaskEdBoxDate.Text = Format(rs!Entrydate, "dd/mm/yy")
    i = 1
    
    With MSFlexGrid1
        Do Until rs.EOF
            .Rows = i + 1
            .RowHeight(i) = 600
            .ColAlignment(i) = 0
            .TextMatrix(i, 0) = "" & rs!SubjectID
            Set local_rs = getdata("SELECT  Sub_title from subject_info_sub  where Sub_code='" & rs!SubjectID & "'")
                  .TextMatrix(i, 1) = "" & local_rs!Sub_title
                  Set saturday_rs = getdata("SELECT  TeacherId,starttime,Endtime,ListOfday from ClassRoutine  where Subjectid='" & rs!SubjectID & "' and serialno='" & List1.Text & "'and academic_yr='" & Combo5 & "' and Classid='" & Get_Code(Combo1.Text) & "' and  Shift='" & Mid(Combo2, 1, 1) & "' and ListOfday = 'Saturday'")
                  If Not saturday_rs.EOF Then
                     Set teacher_name_rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(saturday_rs!teacherId) & "'")
                      .TextMatrix(i, 2) = saturday_rs!Starttime & "-" & saturday_rs!Endtime & " " & teacher_name_rs!name
                  End If
                  Set sunday_rs = getdata("SELECT  TeacherId,starttime,Endtime,ListOfday from ClassRoutine  where Subjectid='" & rs!SubjectID & "' and serialno='" & List1.Text & "'and academic_yr='" & Combo5 & "' and Classid='" & Get_Code(Combo1.Text) & "' and  Shift='" & Mid(Combo2, 1, 1) & "' and ListOfday = 'Sunday'")
                 
                  If Not sunday_rs.EOF Then
                       Set teacher_name_rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(sunday_rs!teacherId) & "'")
                      .TextMatrix(i, 3) = sunday_rs!Starttime & "-" & sunday_rs!Endtime & "   " & teacher_name_rs!name
                  End If
                  
                  Set monday_rs = getdata("SELECT  TeacherId,starttime,Endtime,ListOfday from ClassRoutine  where Subjectid='" & rs!SubjectID & "' and serialno='" & List1.Text & "'and academic_yr='" & Combo5 & "' and Classid='" & Get_Code(Combo1.Text) & "' and  Shift='" & Mid(Combo2, 1, 1) & "' and ListOfday = 'Monday'")

                  If Not monday_rs.EOF Then
                      Set teacher_name_rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(monday_rs!teacherId) & "'")
                      .TextMatrix(i, 4) = monday_rs!Starttime & "-" & monday_rs!Endtime & "   " & teacher_name_rs!name

                  End If
                  
                    Set tuesday_rs = getdata("SELECT  TeacherId,starttime,Endtime,ListOfday from ClassRoutine  where Subjectid='" & rs!SubjectID & "' and serialno='" & List1.Text & "'and academic_yr='" & Combo5 & "' and Classid='" & Get_Code(Combo1.Text) & "' and  Shift='" & Mid(Combo2, 1, 1) & "' and ListOfday = 'Tuesday'")

                   If Not tuesday_rs.EOF Then
                     Set teacher_name_rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(tuesday_rs!teacherId) & "'")
                     .TextMatrix(i, 5) = tuesday_rs!Starttime & "-" & tuesday_rs!Endtime & "   " & teacher_name_rs!name
                  End If
                  
                    Set wednesday_rs = getdata("SELECT  TeacherId,starttime,Endtime,ListOfday from ClassRoutine  where Subjectid='" & rs!SubjectID & "' and serialno='" & List1.Text & "'and academic_yr='" & Combo5 & "' and Classid='" & Get_Code(Combo1.Text) & "' and  Shift='" & Mid(Combo2, 1, 1) & "' and ListOfday = 'Wednesday'")

                  If Not wednesday_rs.EOF Then
                   Set teacher_name_rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(wednesday_rs!teacherId) & "'")
                   .TextMatrix(i, 6) = wednesday_rs!Starttime & "-" & wednesday_rs!Endtime & "   " & teacher_name_rs!name

                  End If
                Set thursday_rs = getdata("SELECT  TeacherId,starttime,Endtime,ListOfday from ClassRoutine  where Subjectid='" & rs!SubjectID & "' and serialno='" & List1.Text & "'and academic_yr='" & Combo5 & "' and Classid='" & Get_Code(Combo1.Text) & "' and  Shift='" & Mid(Combo2, 1, 1) & "' and ListOfday = 'Thursday'")
                 If Not thursday_rs.EOF Then
                     Set teacher_name_rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(thursday_rs!teacherId) & "'")
                     .TextMatrix(i, 7) = thursday_rs!Starttime & "-" & thursday_rs!Endtime & "   " & teacher_name_rs!name

                  End If
                  
                  Set friday_rs = getdata("SELECT  TeacherId,starttime,Endtime,ListOfday from ClassRoutine  where Subjectid='" & rs!SubjectID & "' and serialno='" & List1.Text & "'and academic_yr='" & Combo5 & "' and Classid='" & Get_Code(Combo1.Text) & "' and  Shift='" & Mid(Combo2, 1, 1) & "' and ListOfday = 'Friday'")

                  If Not friday_rs.EOF Then
                      Set teacher_name_rs = getdata("SELECT Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Name From Emp_Per_Info where Emp_Per_Info.Emp_id='" & Trim(friday_rs!teacherId) & "'")
                      .TextMatrix(i, 8) = friday_rs!Starttime & "-" & friday_rs!Endtime & "   " & teacher_name_rs!name

                  End If
                .TextMatrix(i, 9) = rs!trackid
            rs.MoveNext
           i = i + 1
        Loop
    End With
 Else
     MSFlexGrid1.Rows = 1
 End If
 cmdGenerate.Enabled = False
 MaskEdBoxDate.Enabled = False
 If Len(Trim(List1.Text)) = 0 Then
    cmdSave.Enabled = False
 Else
    cmdSave.Enabled = True
 End If

End Sub
Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List2.SetFocus
End If
End Sub
Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
List3.Clear
DtpicStart = "00:0:00"
dtpickend = "00:00:00"
Combo4.Text = ""
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Set rs1 = getdata("select SubjectId from Classroutine  where classid='" & Mid((Combo1.Text), 1, 5) & "'and shift='" & (Combo2.Text) & "'and listofday='" & Trim(List1.Text) & "'and sectionId='" & Mid((Combo3.Text), 1, 5) & "' ")
If Not rs1.EOF Then
    Do Until rs1.EOF
        Set rs2 = getdata("select Subjectdsc from SubjectInfo  where classid='" & Mid((Combo1.Text), 1, 5) & "' and  SubjectId='" & rs1!SubjectID & "' ")
        List3.AddItem rs1!SubjectID + "-" + rs2!SubjectDsc
        rs1.MoveNext
    Loop
Else
    List3.Clear
End If
End Sub
Private Sub List2_Click()
DtpicStart = "00:0:00"
dtpickend = "00:00:00"
'Combo4.Text = ""
End Sub
Private Sub List2_DblClick()
List3.AddItem List2.Text
getteacherforsub
End Sub
Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
DtpicStart = "00:0:00"
dtpickend = "00:00:00"
Combo4.Text = ""
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command2.SetFocus
End If
End Sub
Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
DtpicStart = "00:0:00"
dtpickend = "00:00:00"
Combo4.Text = ""
End Sub
Private Sub List3_Click()
  selectedText = List3.Text
  getteacherforsub
End Sub

Private Sub List3_DblClick()
    List3.RemoveItem List3.ListIndex
    Combo4.Clear
End Sub
Private Sub List3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DtpicStart.SetFocus
    getteacherforsub
End If
End Sub
Private Sub RemoveAll_Click()
If MsgBox("Do you want to delete the all the subjects for the day?", vbYesNo) = vbYes Then
    Dim cmd As New ADODB.Command
    Dim con As New ADODB.connection
    Dim rs As New ADODB.Recordset
    con.Open GConnString
    Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = "Delete from ClassRoutine  where (ClassID = '" & Mid((Combo1.Text), 1, 5) & "')and sectionId= '" & Mid((Combo3.Text), 1, 5) & "'and shift='" & (Combo2.Text) & "'and listofday='" & List1.Text & "'"
    cmd.Execute
    MsgBox "Delete successfully Subject Information for the day.", vbInformation, App.Title
    List3.Clear
    Combo4.Text = ""
Else
    List3.Clear
End If

End Sub
Private Sub removeone_Click()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
If Combo4.Text = "" Then
    List3.RemoveItem List3.ListIndex
End If
If MsgBox("Are you sure to delete  the subject of the day?", vbYesNo) = vbYes Then
    Dim cmd As New ADODB.Command
    Dim con As New ADODB.connection
    con.Open GConnString
    Set cmd.ActiveConnection = con
    cmd.CommandType = adCmdText
    cmd.CommandText = "Delete from ClassRoutine  where (ClassID = '" & Mid((Combo1.Text), 1, 5) & "')and sectionId= '" & Mid((Combo3.Text), 1, 5) & "'and shift='" & (Combo2.Text) & "'and listofday='" & List1.Text & "'and subjectid ='" & Mid((List3.Text), 1, 5) & "' and teacherid='" & Mid((Combo4.Text), 1, 5) & "'"
    cmd.Execute
    MsgBox "Delete successfully Information for the Subject of the day.", vbInformation, App.Title
    List3.RemoveItem List3.ListIndex
    Combo4.Text = ""
End If
Exit Sub
errdes:
End Sub

Private Sub MSFlexGrid1_Click()
  If MSFlexGrid1.Row > 0 Then
     txtTrackid.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9)
  Else
    txtTrackid = ""
    
  End If
End Sub
