VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAdmissionApprove 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   660
      Width           =   7935
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   900
         TabIndex        =   9
         Top             =   540
         Width           =   2295
      End
      Begin VB.ComboBox ComStuId 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Select Student"
         Top             =   180
         Width           =   2295
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3930
         TabIndex        =   7
         Top             =   510
         Width           =   3885
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   900
         TabIndex        =   6
         Top             =   870
         Width           =   2295
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3930
         TabIndex        =   5
         Top             =   840
         Width           =   1245
      End
      Begin MSMask.MaskEdBox MaskEdBoxDate 
         Height          =   285
         Left            =   5970
         TabIndex        =   10
         Top             =   840
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Roll"
         Height          =   195
         Left            =   3330
         TabIndex        =   17
         Top             =   900
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   900
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class "
         Height          =   195
         Left            =   3360
         TabIndex        =   15
         Top             =   570
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   540
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   5400
         TabIndex        =   13
         Top             =   900
         Width           =   555
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   285
         Left            =   3180
         TabIndex        =   12
         Top             =   180
         Width           =   4635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   210
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   465
      Left            =   7080
      TabIndex        =   3
      ToolTipText     =   "Click to Exit"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdAdmitApprove 
      BackColor       =   &H8000000C&
      Caption         =   "Admission Approve"
      Height          =   465
      Left            =   5460
      TabIndex        =   2
      ToolTipText     =   "Click to Approva the Admission"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7905
      TabIndex        =   0
      Top             =   0
      Width           =   7965
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   1
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Admission Approval"
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
         Left            =   2190
         TabIndex        =   18
         Top             =   120
         Width           =   2310
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   -240
         Picture         =   "frmAdmissionApprove.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   8145
      End
   End
End
Attribute VB_Name = "frmAdmissionApprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdmitApprove_Click()

If Len(ComStuId) = 0 Then
 MsgBox "Please enter a student Id ...", vbInformation, cmp
 ComStuId.SetFocus
 Exit Sub
End If

If MaskEdBoxDate = "__/__/__" Then
    MsgBox "Enter Date.", vbInformation, App.Title
    MaskEdBoxDate.SetFocus
    Exit Sub
End If

Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "Studentadmission1"

cmd(1) = ComStuId.Text
cmd(2) = "Y"
cmd(3) = Format(MaskEdBoxDate, "dd mmm yyyy")
cmd(4) = "Y"
cmd(5) = soft_user
cmd.Execute

MsgBox "Admission is Approved Successfully.", vbInformation, "Student Management System"
loadstudent
Dim i As Integer
For i = 0 To 3
  txtfields(i) = ""
Next i
Label3.Caption = ""

Exit Sub

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub ComStuId_DropDown()
  loadstudent
End Sub

Private Sub ComStuId_LostFocus()

If Len(ComStuId) <> 0 Then
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT StudentID from StudentInfo where StudentID='" & ComStuId & "' ")
    If rs.EOF Then
        MsgBox "Invalid Student ID.", vbCritical, "Shool Management System"
        ComStuId.Text = ""
        Exit Sub
    End If
End If

End Sub

Private Sub Form_Load()
   loadstudent
End Sub
Private Sub loadstudent()
  Dim rs As New ADODB.Recordset
  ComStuId.Clear
Set rs = getdata("SELECT     StudentID From StudentInfo " + _
"WHERE (StudentID  IN(SELECT StudentID FROM StudentAdmission where Approval='N'and Admissioncancel<>'Y'))")
If Not rs.EOF Then
    Do Until rs.EOF
        ComStuId.AddItem rs(0)
        rs.MoveNext
    Loop
End If
End Sub
Private Sub ComStuId_Click()

MaskEdBoxDate = Format(MaskEdBoxDate, "__/__/__")
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT  StudentInfo.StudentName, StudentAdmission.AdmissionDate, StudentAdmission.Shift, StudentAdmission.ClassId, StudentAdmission.SectionId," + _
            "StudentAdmission.ClassRoll , ClassInfo.ClassName, SectionInfo.Sectiondsc,StudentAdmission.AdmissionDate " + _
            "FROM StudentAdmission INNER JOIN StudentInfo ON StudentAdmission.StudentId = StudentInfo.StudentID INNER JOIN " + _
            "ClassInfo ON StudentAdmission.ClassId = ClassInfo.ClassID INNER JOIN " + _
            "SectionInfo ON StudentAdmission.SectionId = SectionInfo.SectionID WHERE StudentAdmission.StudentID = '" & ComStuId.Text & "'")
If Not rs.EOF Then
    Label3.Caption = rs!StudentName
    
    If rs!Shift = "M" Then
        txtfields(0) = "Morning Shift"
    ElseIf rs!Shift = "D" Then
        txtfields(0) = "Day Shift"
    End If
    
    
    txtfields(1) = rs!classId + "-" + rs!ClassName
    txtfields(2) = rs!SectionID + "-" + rs!Sectiondsc
    txtfields(3) = rs!ClassRoll
    
End If
MaskEdBoxDate.SetFocus

End Sub

Private Sub MaskEdBoxDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    If MaskEdBoxDate <> "__/__/__" Then
            If Check_ValidDate(MaskEdBoxDate) = False Then
                MaskEdBoxDate.SetFocus
                Exit Sub
            End If
    End If
    
    cmdAdmitApprove.SetFocus
End If
End Sub
