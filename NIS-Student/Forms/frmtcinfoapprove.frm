VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmtcinfoapprove 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   375
      Left            =   7560
      TabIndex        =   27
      ToolTipText     =   "Click to Exit"
      Top             =   4980
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000C&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6570
      TabIndex        =   26
      ToolTipText     =   "Click to Cancel the TC"
      Top             =   4980
      Width           =   945
   End
   Begin VB.CommandButton cmdTcApprove 
      BackColor       =   &H8000000C&
      Caption         =   "TC Approve"
      Height          =   375
      Left            =   5340
      TabIndex        =   25
      ToolTipText     =   "Click to Approve the TC"
      Top             =   4980
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   975
      Left            =   -30
      ScaleHeight     =   915
      ScaleWidth      =   8565
      TabIndex        =   23
      Top             =   0
      Width           =   8625
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   24
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer Certificate Prepeartion Approval"
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
         Left            =   1440
         TabIndex        =   29
         Top             =   270
         Width           =   4665
      End
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   0
         Picture         =   "frmtcinfoapprove.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   8595
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "TC Information"
      ForeColor       =   &H00C00000&
      Height          =   1935
      Left            =   0
      TabIndex        =   17
      Top             =   2970
      Width           =   8595
      Begin VB.TextBox txtfields 
         Height          =   1335
         Index           =   7
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   510
         Width           =   7185
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   18
         Top             =   210
         Width           =   4815
      End
      Begin MSMask.MaskEdBox MaskEdDate 
         Height          =   285
         Left            =   7080
         TabIndex        =   19
         Top             =   210
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd-mmm-yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   510
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type of TC"
         Height          =   195
         Left            =   60
         TabIndex        =   21
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   6510
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Student Information"
      ForeColor       =   &H00C00000&
      Height          =   1965
      Left            =   0
      TabIndex        =   0
      Top             =   990
      Width           =   8595
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   5100
         TabIndex        =   7
         Top             =   1530
         Width           =   2055
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   6
         Top             =   1530
         Width           =   2835
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   5
         Top             =   870
         Width           =   7215
      End
      Begin VB.ComboBox ComStuId 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         ToolTipText     =   "Select Student"
         Top             =   210
         Width           =   2265
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   5100
         TabIndex        =   3
         Top             =   1200
         Width           =   3435
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Top             =   570
         Width           =   7215
      End
      Begin VB.TextBox txtfields 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   1
         Top             =   1200
         Width           =   2835
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4320
         TabIndex        =   16
         Top             =   1710
         Width           =   45
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         Height          =   195
         Left            =   4320
         TabIndex        =   15
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class "
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   1260
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   1590
         Width           =   315
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3600
         TabIndex        =   12
         Top             =   210
         Width           =   4935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mother's Name"
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Father's Name"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   570
         Width           =   1020
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Roll"
         Height          =   195
         Left            =   4320
         TabIndex        =   8
         Top             =   1560
         Width           =   270
      End
   End
   Begin VB.Shape Shape1 
      Height          =   435
      Left            =   5310
      Top             =   4950
      Width           =   3225
   End
End
Attribute VB_Name = "frmtcinfoapprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
If Len(ComStuId.Text) <> 0 And Len(txtfields(2)) <> 0 And Len(txtfields(6)) <> 0 Then
        Dim cmd As New ADODB.Command
        Dim con As New ADODB.connection
        Dim rs As New ADODB.Recordset
        con.Open GConnString
        Set cmd.ActiveConnection = con
        If MsgBox("Are You Sure To Delete ?", vbYesNo + vbCritical) = vbYes Then
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "TCInfoApprove"
                cmd(1) = ComStuId
                cmd(2) = txtfields(7)
                cmd(3) = "N"
                cmd(4) = soft_user
                cmd(5) = Format(MaskEdDate, "dd mm yyyy")
                cmd(6) = "Y"
                
                cmd.Execute
                MsgBox "Cancel Successfully Tc Information for the student.", vbInformation, App.Title
                For i = 0 To 7
                 txtfields(i) = ""
                Next
                Label7.Caption = ""
                ComStuId.Text = ""
                
                MaskEdDate = "__/__/__"
        Else
                 Exit Sub
        End If
Else
    Exit Sub
End If

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdTcApprove_Click()
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con

    
   If MaskEdDate = "__/__/__" Then
        MsgBox "Please Enter Date", vbCritical, App.Title
        MaskEdDate.SetFocus
        Exit Sub
   End If
    
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "TCInfoApprove"
    cmd(1) = ComStuId
    cmd(2) = txtfields(7)
    cmd(3) = "Y"
    cmd(4) = soft_user
    cmd(5) = MaskEdDate
    cmd(6) = "N"
   
    cmd.Execute
    
    
    MsgBox "Save Successfully.", vbInformation, "Student Management System"
   
  
    
End Sub

Private Sub ComStuId_Click()

 Dim rs As New ADODB.Recordset
 Set rs = getdata("SELECT     TCInformation.StudentID, TCInformation.Shift, TCInformation.ClassID, ClassInfo.ClassName, TCInformation.SectionID, SectionInfo.Sectiondsc, " + _
"TCInformation.ClassRoll, TCInformation.TCTypeID, TCTypeSetUp.TcName,  TCInformation.TCNote, TCInformation.Approved,StudentInfo.StudentName , StudentInfo.StuFatherName, StudentInfo.StuMotherName " + _
"FROM  TCInformation INNER JOIN ClassInfo ON TCInformation.ClassID = ClassInfo.ClassID INNER JOIN " + _
"SectionInfo ON TCInformation.SectionID = SectionInfo.SectionID INNER JOIN " + _
 "TCTypeSetUp ON TCInformation.TCTypeID = TCTypeSetUp.TCID INNER JOIN " + _
"StudentInfo ON TCInformation.StudentID = StudentInfo.StudentID where TCInformation.studentid='" & ComStuId & "'and TCInformation.approved='P' ")
    If Not rs.EOF Then
                Label7.Caption = "" & rs!StudentName
                txtfields(0) = "" & rs!StuFatherName
                txtfields(1) = "" & rs!StuMotherName
                txtfields(2) = "" & rs!classId + "-" + rs!ClassName
                txtfields(3) = "" & rs!SectionID + "-" + rs!Sectiondsc
                txtfields(4) = "" & rs!Shift
                txtfields(5) = "" & rs!ClassRoll
                txtfields(6) = "" & rs!TCTypeID + "-" + rs!TcName
                txtfields(7) = "" & rs!TCNote
            
                
    
    End If
    MaskEdDate.SetFocus

End Sub

Private Sub ComStuId_GotFocus()

For i = 0 To 7
    txtfields(i) = ""
Next

MaskEdDate = "__/__/__"
End Sub

Private Sub ComStuId_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

     Dim rs As New ADODB.Recordset
     Set rs = getdata("SELECT     TCInformation.StudentID, TCInformation.Shift, TCInformation.ClassID, ClassInfo.ClassName, TCInformation.SectionID, SectionInfo.Sectiondsc, " + _
    "TCInformation.ClassRoll, TCInformation.TCTypeID, TCTypeSetUp.TcName, TCInformation.TCNote, TCInformation.Approved,StudentInfo.StudentName , StudentInfo.StuFatherName, StudentInfo.StuMotherName " + _
    "FROM  TCInformation INNER JOIN ClassInfo ON TCInformation.ClassID = ClassInfo.ClassID INNER JOIN " + _
    "SectionInfo ON TCInformation.SectionID = SectionInfo.SectionID INNER JOIN " + _
     "TCTypeSetUp ON TCInformation.TCTypeID = TCTypeSetUp.TCID INNER JOIN " + _
    "StudentInfo ON TCInformation.StudentID = StudentInfo.StudentID where TCInformation.studentid='" & ComStuId & "'and TCInformation.approved='P' ")
        If Not rs.EOF Then
                    Label7.Caption = "" & rs!StudentName
                    txtfields(0) = "" & rs!StuFatherName
                    txtfields(1) = "" & rs!StuMotherName
                    txtfields(2) = "" & rs!classId + "-" + rs!ClassName
                    txtfields(3) = "" & rs!SectionID + "-" + rs!Sectiondsc
                    txtfields(4) = "" & rs!Shift
                    txtfields(5) = "" & rs!ClassRoll
                    txtfields(6) = "" & rs!TCTypeID + "-" + rs!TcName
                    txtfields(7) = "" & rs!TCNote
                    
        
            MaskEdDate.SetFocus
        End If
     
End If
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT StudentID From TcInformation where approved='P'")
If Not rs.EOF Then
    Do Until rs.EOF
        ComStuId.AddItem rs!StudentID
        rs.MoveNext
    Loop

End If

End Sub





Private Sub MaskEdDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MaskEdDate <> "__/__/__" Then
            If Check_ValidDate(MaskEdDate) = False Then
                MaskEdDate.SetFocus
                Exit Sub
            End If
    End If
txtfields(7).SetFocus
End If
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 7
            cmdTcApprove.SetFocus
    End Select
End If
End Sub


 


