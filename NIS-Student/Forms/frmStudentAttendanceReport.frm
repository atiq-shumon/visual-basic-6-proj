VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmStudentAttendanceReport 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   915
      Left            =   0
      TabIndex        =   23
      Top             =   2370
      Width           =   6555
      Begin VB.OptionButton optSpecificShift 
         Caption         =   "Specific Shift"
         Height          =   225
         Left            =   150
         TabIndex        =   25
         Top             =   390
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.ComboBox cmbSpecificShift 
         Height          =   315
         ItemData        =   "frmStudentAttendanceReport.frx":0000
         Left            =   1680
         List            =   "frmStudentAttendanceReport.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   360
         Width           =   4665
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   375
      Left            =   5430
      TabIndex        =   22
      ToolTipText     =   "Click to Close"
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H8000000C&
      Caption         =   "Print"
      Height          =   375
      Left            =   4380
      TabIndex        =   21
      ToolTipText     =   "Click to Print"
      Top             =   5760
      Width           =   1035
   End
   Begin VB.Frame Frame3 
      Height          =   1425
      Left            =   0
      TabIndex        =   12
      Top             =   4230
      Width           =   6555
      Begin VB.OptionButton optToDay 
         Caption         =   "ToDay"
         Height          =   375
         Left            =   180
         TabIndex        =   26
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optSpecificDate 
         Caption         =   "Specific Date"
         Height          =   375
         Left            =   150
         TabIndex        =   18
         Top             =   600
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optyearly 
         Caption         =   "Yearly"
         Height          =   405
         Left            =   150
         TabIndex        =   17
         Top             =   240
         Width           =   1185
      End
      Begin MSMask.MaskEdBox MaskFromDate 
         Height          =   285
         Left            =   2850
         TabIndex        =   14
         Top             =   690
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskToDate 
         Height          =   285
         Left            =   4950
         TabIndex        =   15
         Top             =   690
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "From Date"
         Height          =   255
         Left            =   1860
         TabIndex        =   16
         Top             =   690
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "To Date"
         Height          =   255
         Left            =   4260
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   0
      TabIndex        =   8
      Top             =   750
      Width           =   6555
      Begin VB.ComboBox cmdAcademicYear 
         Height          =   315
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   300
         Width           =   1665
      End
      Begin VB.ComboBox cmbClass 
         Height          =   315
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Academic Year"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3330
         TabIndex        =   19
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   9
         Top             =   330
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   6555
      Begin VB.ComboBox cmdSection 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   390
         Width           =   4665
      End
      Begin VB.OptionButton optSpecificSection 
         Caption         =   "Specific Section"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   330
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   -60
      Picture         =   "frmStudentAttendanceReport.frx":0004
      ScaleHeight     =   690
      ScaleWidth      =   10275
      TabIndex        =   1
      Top             =   0
      Width           =   10335
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Attendance Report"
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
         Left            =   1950
         TabIndex        =   2
         Top             =   150
         Width           =   3105
      End
      Begin VB.Image Image1 
         Height          =   990
         Left            =   -120
         Picture         =   "frmStudentAttendanceReport.frx":C89B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   10305
      End
   End
   Begin VB.Frame famStudentInfo 
      Height          =   1125
      Left            =   0
      TabIndex        =   3
      Top             =   3180
      Width           =   6555
      Begin VB.ComboBox cmdStudentID 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   540
         Width           =   4665
      End
      Begin VB.OptionButton optSpecifickStudent 
         Caption         =   "Specific Student"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optAllStudent 
         Caption         =   "All Student "
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmStudentAttendanceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FromDate As Date
Public toDate As Date

Private Sub cmbClass_Change()
On Error GoTo err_des
    Call load_section
    Call LoadStuID

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmbClass_Click()
On Error GoTo err_des
    Call load_section
    Call LoadStuID

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub


Private Sub cmbClass_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_des

If KeyCode = vbKeyReturn Then
    cmdAcademicYear.SetFocus
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmbSpecificShift_Change()
On Error GoTo err_des

    Call LoadStuID

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmbSpecificShift_Click()
On Error GoTo err_des

    Call LoadStuID

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub cmbSpecificShift_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_des

If KeyCode = vbKeyReturn Then
    MaskFromDate.SetFocus
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdAcademicYear_Change()
On Error GoTo err_des
    Call LoadStuID

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdAcademicYear_Click()
On Error GoTo err_des
    Call LoadStuID

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdAcademicYear_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_des

    If KeyCode = vbKeyReturn Then
        cmdSection.SetFocus
    End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdClose_Click()
On Error GoTo err_des
Unload Me
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdPrint_Click()
On Error GoTo err_des

    If optSpecificDate.Value = True Then
        If MaskFromDate.Text = "__/__/__" Then
           MsgBox "Please Input a valid Date", vbInformation, cmp
           MaskFromDate.SetFocus
           Exit Sub
        End If
        If MaskToDate.Text = "__/__/__" Then
           MsgBox "Please Input a valid Date", vbInformation, cmp
           MaskToDate.SetFocus
           Exit Sub
        End If
    End If
    
    If optSpecificDate.Value = True Then
        FromDate = MaskFromDate
        toDate = MaskToDate
    ElseIf optyearly.Value = True Then
        FromDate = "01/01/" & cmdAcademicYear & ""
        toDate = "31/12/" & cmdAcademicYear & ""
    ElseIf optToDay.Value = True Then
        FromDate = Date
    End If
    
    
    If (optSpecificSection.Value = True) And (optSpecificShift.Value = True) _
    And (optAllStudent.Value = True) And (optyearly.Value = True) Then
    
        rptMode = 14
        Screen.MousePointer = vbHourglass
        frmViewer.Show 1
        
    End If
    
    
    If (optSpecificSection.Value = True) And (optSpecificShift.Value = True) _
    And (optAllStudent.Value = True) And (optyearly.Value = False) And (optToDay.Value = True) Then
        
        rptMode = 17
        Screen.MousePointer = vbHourglass
        frmViewer.Show 1
    End If
    
    If optSpecificDate.Value = True Then
        If (optSpecificSection.Value = True) And (optSpecificShift.Value = True) _
            And (optAllStudent.Value = True) And ((CDate(MaskFromDate) = CDate(MaskToDate))) Then
            
            rptMode = 17
            Screen.MousePointer = vbHourglass
            frmViewer.Show 1
        End If
    End If
     
    If optSpecificDate.Value = True Then
        If (optSpecificSection.Value = True) And (optSpecificShift.Value = True) _
        And (optAllStudent.Value = True) And (optSpecificDate.Value = True) And (CDate(MaskFromDate) <> CDate(MaskToDate)) Then
            rptMode = 16
            Screen.MousePointer = vbHourglass
            frmViewer.Show 1
        End If
    End If
    
    If (optSpecificDate.Value = True) Or (optyearly.Value = True) Then
        If (optSpecificSection.Value = True) And (optSpecificShift.Value = True) _
        And (optSpecifickStudent.Value = True) And ((optSpecificDate.Value = True) Or ((optyearly.Value = True))) Then
            
            rptMode = 18
            Screen.MousePointer = vbHourglass
            frmViewer.Show 1
            
        End If
    End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdSection_Change()
On Error GoTo err_des
    Call LoadStuID
Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
    
End Sub

Private Sub cmdSection_Click()
On Error GoTo err_des
    Call LoadStuID

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdSection_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_des
If KeyCode = vbKeyReturn Then
    cmbSpecificShift.SetFocus
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub cmdStudentID_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_des
If KeyCode = vbKeyReturn Then
    MaskFromDate.SetFocus
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub


Private Sub Form_Load()
On Error GoTo err_des

load_Aca_year
load_class
load_section
LoadStuID

cmbSpecificShift.AddItem "Morning~Shift"
cmbSpecificShift.AddItem "Day~Shift"

cmbSpecificShift.ListIndex = 0

optToDay.Visible = False

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub


Private Sub load_class()
On Error GoTo err_des
    Dim rs As New ADODB.Recordset
    Set rs = getdata("SELECT ClassID, ClassName FROM  classinfo")
    cmbClass.Clear
    If Not rs.EOF Then
        Do Until rs.EOF
            cmbClass.AddItem rs(1) + "~" + rs(0)
            rs.MoveNext
        Loop
    End If
    
    cmbClass.ListIndex = 0

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub load_Aca_year()
On Error GoTo err_des
    Dim i As Integer
        For i = 2000 To 2050
            cmdAcademicYear.AddItem i
        Next i
    cmdAcademicYear.Text = Format(Date, "YYYY")
    
    cmdAcademicYear.ListIndex = 7

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub load_section()
On Error GoTo err_des
Dim rs As New ADODB.Recordset

    Set rs = getdata("SELECT SectionID,Sectiondsc from sectioninfo WHERE ClassID='" & Get_Code(Trim(cmbClass)) & "'")
    
    cmdSection.Clear

    If Not rs.EOF Then
        rs.MoveFirst
            Do Until rs.EOF
                cmdSection.AddItem rs(1) + "~" + rs(0)
                rs.MoveNext
            Loop
    End If
    
    If cmdSection.ListCount > 0 Then
        cmdSection.ListIndex = 0
    End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub LoadStuID()
On Error GoTo err_des
Dim rs As New ADODB.Recordset
    Set rs = getdata("SELECT Distinct a.StudentID,(select StudentName from studentinfo s where s.StudentID=a.StudentID)" _
        & " FROM  Studentadmission a where a.classid='" & Get_Code(cmbClass) & "' and sectionid='" & Get_Code(cmdSection) _
        & "' and a.approval='Y' and active_std=1 and a.admissionCancel='N'and a.serial_no=(select max(serial_no) from " _
        & " studentadmission where a.Shift='" & Mid(frmStudentAttendanceReport.cmbSpecificShift, 1, 1) & "' and studentid=a.studentid and aca_yr='" & Trim(cmdAcademicYear) & "') order by studentid")
    
    cmdStudentID.Clear
    If Not rs.EOF Then
        Do Until rs.EOF
            cmdStudentID.AddItem rs(1) + "~" + rs(0)
            rs.MoveNext
        Loop
    End If
        
    If cmdStudentID.ListCount > 0 Then
        cmdStudentID.ListIndex = 0
    End If
    

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub MaskFromDate_GotFocus()
On Error GoTo err_des
  MaskFromDate.SelStart = 0
  MaskFromDate.SelLength = Len(MaskFromDate)

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub MaskFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_des
If KeyCode = vbKeyReturn Then
    MaskToDate.SetFocus
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub MaskToDate_GotFocus()
On Error GoTo err_des
MaskToDate.SelStart = 0
MaskToDate.SelLength = Len(MaskToDate)

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub MaskToDate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err_des
If KeyCode = vbKeyReturn Then
    cmdPrint.SetFocus
End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub optAllSection_Click()
On Error GoTo err_des
    If optAllSection.Value = True Then
        cmdSection.Visible = False
    Else
        cmdSection.Visible = True
    End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub optAllStudent_Click()
On Error GoTo err_des
    If optAllStudent.Value = True Then
        cmdStudentID.Visible = False
        optToDay.Visible = True
    Else
        cmdStudentID.Visible = True
        optToDay.Visible = False
    End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub optSpecificDate_Click()
On Error GoTo err_des
    If optSpecificDate.Value = True Then
        Label3.Visible = True
        Label4.Visible = True
        MaskFromDate.Visible = True
        MaskToDate.Visible = True
    Else
        Label3.Visible = False
        Label4.Visible = False
        MaskFromDate.Visible = False
        MaskToDate.Visible = False
    End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub optSpecifickStudent_Click()
On Error GoTo err_des
    If optSpecifickStudent.Value = True Then
        cmdStudentID.Visible = True
        optToDay.Visible = False
    Else
        cmdStudentID.Visible = False
        optToDay.Visible = True
    End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub optSpecificSection_Click()
On Error GoTo err_des
    If optSpecificSection.Value = True Then
        cmdSection.Visible = True
    Else
        cmdSection.Visible = False
    End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub


Private Sub optToDay_Click()
On Error GoTo err_des
    If optToDay.Value = True Then
        Label3.Visible = False
        Label4.Visible = False
        MaskFromDate.Visible = False
        MaskToDate.Visible = False
    Else
        Label3.Visible = True
        Label4.Visible = True
        MaskFromDate.Visible = True
        MaskToDate.Visible = True
    End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub optyearly_Click()
On Error GoTo err_des
    If optyearly.Value = True Then
        Label3.Visible = False
        Label4.Visible = False
        MaskFromDate.Visible = False
        MaskToDate.Visible = False
    Else
        Label3.Visible = True
        Label4.Visible = True
        MaskFromDate.Visible = True
        MaskToDate.Visible = True
    End If

Exit Sub
err_des: MsgBox Err.Description, vbInformation, App.Title
End Sub
