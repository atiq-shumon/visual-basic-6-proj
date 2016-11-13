VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmExamSeatPlan 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H8000000C&
      Caption         =   "New"
      Height          =   435
      Left            =   4860
      TabIndex        =   7
      ToolTipText     =   "Click to insert new information"
      Top             =   5490
      Width           =   945
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H8000000C&
      Caption         =   "Save"
      Height          =   435
      Left            =   5850
      TabIndex        =   6
      ToolTipText     =   "Click to Save"
      Top             =   5490
      Width           =   945
   End
   Begin VB.CommandButton cmddelete 
      BackColor       =   &H8000000C&
      Caption         =   "Delete"
      Height          =   435
      Left            =   6840
      TabIndex        =   5
      ToolTipText     =   "Click to Delete"
      Top             =   5490
      Width           =   945
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H8000000C&
      Caption         =   "Close"
      Height          =   435
      Left            =   7830
      TabIndex        =   4
      ToolTipText     =   "Click to Exit"
      Top             =   5490
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   8685
      TabIndex        =   1
      Top             =   0
      Width           =   8745
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   30
         Left            =   0
         TabIndex        =   2
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exam Seat Plan Preperation"
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
         Height          =   405
         Left            =   2370
         TabIndex        =   39
         Top             =   180
         Width           =   3180
      End
      Begin VB.Image Image1 
         Height          =   1020
         Left            =   -30
         Picture         =   "frmExamSeatPlan.frx":0000
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   8745
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4605
      Left            =   0
      TabIndex        =   0
      Top             =   870
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   8123
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "Exam Seat Plan"
      TabPicture(0)   =   "frmExamSeatPlan.frx":CEA5
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "MSFlexGridSeat"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Exam Teacher Plan"
      TabPicture(1)   =   "frmExamSeatPlan.frx":CEC1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "MSFlexGridTeacher"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridTeacher 
         Height          =   2295
         Left            =   -74940
         TabIndex        =   36
         Top             =   2250
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   4048
         _Version        =   393216
         FixedCols       =   0
      End
      Begin VB.Frame Frame5 
         Height          =   1005
         Left            =   -74910
         TabIndex        =   31
         Top             =   1200
         Width           =   8625
         Begin VB.ComboBox ComboResponsibility 
            Height          =   315
            ItemData        =   "frmExamSeatPlan.frx":CEDD
            Left            =   1350
            List            =   "frmExamSeatPlan.frx":CEE7
            TabIndex        =   35
            ToolTipText     =   "Select Responsibility"
            Top             =   540
            Width           =   3495
         End
         Begin VB.ComboBox ComboTeacherName 
            Height          =   315
            Left            =   1350
            TabIndex        =   33
            ToolTipText     =   "Select Teacher Name"
            Top             =   180
            Width           =   5985
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsibility"
            Height          =   195
            Left            =   60
            TabIndex        =   34
            Top             =   600
            Width           =   960
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teacher Name"
            Height          =   195
            Left            =   60
            TabIndex        =   32
            Top             =   240
            Width           =   1065
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridSeat 
         Height          =   2025
         Left            =   30
         TabIndex        =   25
         Top             =   2550
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   3572
         _Version        =   393216
         FixedCols       =   0
      End
      Begin VB.Frame Frame4 
         Height          =   1425
         Left            =   60
         TabIndex        =   14
         Top             =   1080
         Width           =   8655
         Begin VB.ComboBox ComboEndRoll 
            Height          =   315
            Left            =   4470
            TabIndex        =   24
            ToolTipText     =   "Select End Roll"
            Top             =   960
            Width           =   2745
         End
         Begin VB.ComboBox ComboStratRoll 
            Height          =   315
            Left            =   1050
            TabIndex        =   23
            ToolTipText     =   "Select start roll"
            Top             =   990
            Width           =   2535
         End
         Begin VB.ComboBox ComboSection 
            Height          =   315
            Left            =   1050
            TabIndex        =   22
            ToolTipText     =   "Select section"
            Top             =   630
            Width           =   2535
         End
         Begin VB.ComboBox ComboClass 
            Height          =   315
            Left            =   4470
            TabIndex        =   21
            ToolTipText     =   "Select Class"
            Top             =   300
            Width           =   4005
         End
         Begin VB.ComboBox ComboShift 
            Height          =   315
            ItemData        =   "frmExamSeatPlan.frx":CF08
            Left            =   1050
            List            =   "frmExamSeatPlan.frx":CF12
            TabIndex        =   20
            ToolTipText     =   "Select Shift"
            Top             =   270
            Width           =   2535
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Roll"
            Height          =   195
            Left            =   3810
            TabIndex        =   19
            Top             =   1020
            Width           =   600
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Roll"
            Height          =   195
            Left            =   90
            TabIndex        =   18
            Top             =   1050
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Section"
            Height          =   195
            Left            =   90
            TabIndex        =   17
            Top             =   690
            Width           =   540
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class"
            Height          =   195
            Left            =   3810
            TabIndex        =   16
            Top             =   330
            Width           =   375
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shift"
            Height          =   195
            Left            =   90
            TabIndex        =   15
            Top             =   330
            Width           =   315
         End
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00C00000&
         Height          =   885
         Left            =   -74910
         TabIndex        =   8
         Top             =   330
         Width           =   8625
         Begin MSMask.MaskEdBox DTPickerExamdateForTeacherPlan 
            Height          =   315
            Left            =   1080
            TabIndex        =   37
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd-mmm-yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox ComboRoomForTeacher 
            Height          =   315
            Left            =   3780
            TabIndex        =   28
            Top             =   300
            Width           =   2145
         End
         Begin MSComCtl2.DTPicker DTPickerStartTimeForTeacherPlan 
            Height          =   345
            Left            =   7050
            TabIndex        =   30
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   609
            _Version        =   393216
            Format          =   45875202
            CurrentDate     =   38633
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Strat Time"
            Height          =   195
            Left            =   6180
            TabIndex        =   29
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Room No"
            Height          =   195
            Left            =   2910
            TabIndex        =   27
            Top             =   330
            Width           =   675
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exam Date"
            Height          =   195
            Left            =   90
            TabIndex        =   26
            Top             =   300
            Width           =   780
         End
      End
      Begin VB.Frame Frame2 
         ForeColor       =   &H00C00000&
         Height          =   765
         Left            =   60
         TabIndex        =   3
         Top             =   330
         Width           =   8655
         Begin MSMask.MaskEdBox DTPickExamdateForSeatPlan 
            Height          =   285
            Left            =   1020
            TabIndex        =   38
            ToolTipText     =   "Insert Exam date"
            Top             =   270
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd-mmm-yy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox ComboRoom 
            Height          =   315
            Left            =   3810
            TabIndex        =   13
            ToolTipText     =   "Select Room No"
            Top             =   270
            Width           =   2145
         End
         Begin MSComCtl2.DTPicker DTPickerStratTimeforseatPlan 
            Height          =   285
            Left            =   6960
            TabIndex        =   12
            ToolTipText     =   "Insert Start time"
            Top             =   270
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   503
            _Version        =   393216
            Format          =   45875202
            CurrentDate     =   38633
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Strat Time"
            Height          =   195
            Left            =   6090
            TabIndex        =   11
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Room No"
            Height          =   195
            Left            =   2940
            TabIndex        =   10
            Top             =   300
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exam Date"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   300
            Width           =   780
         End
      End
   End
End
Attribute VB_Name = "frmExamSeatPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()

Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
Dim rs As New ADODB.Recordset
con.Open GConnString
Set cmd.ActiveConnection = con
If SSTab1.Tab = 0 Then
    If Not DTPickExamdateForSeatPlan.Mask <> "__/__/__" And Len(ComboRoom) <> 0 And DTPickerStratTimeforseatPlan.Hour = 0 Then
            Set rs = getdata("select * from ExamGuardPlan where ExamDate= '" & Format(DTPickExamdateForSeatPlan, "dd mmm yyyy") & "' and RoomNo='" & ComboRoom.Text & "' and startTime='" & DTPickerStratTimeforseatPlan & "'")
            If rs.EOF Then
                If MsgBox("Are You sure to Delete ?", vbYesNo + vbCritical) = vbYes Then
                    cmd.CommandType = adCmdText
                    cmd.CommandText = "Delete from ExamSitPlan where ExamDate= '" & Format(DTPickExamdateForSeatPlan, "dd mmm yyyy") & "' and RoomNo='" & ComboRoom.Text & "' and startTime='" & DTPickerStratTimeforseatPlan & "'"
                    cmd.Execute
                    MsgBox "Delete successfully Exam Seat Plan Information.", vbInformation, App.Title
                    
                    ComboClass.Text = ""
                    ComboEndRoll = ""
                    ComboRoom.Text = ""
                    ComboSection.Text = ""
                    ComboShift.Text = ""
                    ComboStratRoll.Text = ""
                    DTPickerStratTimeforseatPlan = "00:00:00"
                    Call ShowFlexData
                    DTPickExamdateForSeatPlan.Mask = "__/__/__"
                Else
                    Exit Sub
                End If
            Else
                    MsgBox "Please Delete Teacher Imformation for the Date 1st.", vbCritical, App.Title
                    SSTab1.Tab = 1
                    ComboRoomForTeacher.SetFocus
                    Exit Sub
            End If
    Else
        Exit Sub
    End If

Else
        If Not DTPickerExamdateForTeacherPlan.Mask <> "__/__/__" And Len(ComboRoomForTeacher) <> 0 And DTPickerStartTimeForTeacherPlan.Hour = 0 Then
            If MsgBox("Are You Sure To Delete?", vbYesNo + vbCritical) = vbYes Then
                cmd.CommandType = adCmdText
                cmd.CommandText = "Delete from ExamGuardPlan where (ExamDate= '" & Format(DTPickExamdateForSeatPlan, "dd mmm yyyy") & "') and RoomNo='" & ComboRoom.Text & "'and startTime='" & DTPickerStratTimeforseatPlan & "'"
                cmd.Execute
                MsgBox "Delete successfully Exam Teacher Plan Information.", vbInformation, App.Title
                Call ShowFlexData2
                ComboRoomForTeacher.Text = ""
                ComboResponsibility = ""
                ComboTeacherName = ""
                DTPickerStartTimeForTeacherPlan = "00:00:00"
                DTPickerExamdateForTeacherPlan.Mask = "__/__/__"
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
End If
End Sub

Private Sub cmdnew_Click()
If SSTab1.Tab = 0 Then
     ComboClass.Text = ""
     ComboEndRoll = ""
     ComboRoom.Text = ""
     ComboRoomForTeacher = ""
     ComboSection.Text = ""
     ComboShift.Text = ""
     ComboStratRoll.Text = ""
     ComboResponsibility = ""
     ComboTeacherName = ""
     DTPickerExamdateForTeacherPlan.Mask = "__/__/__"
     DTPickerStartTimeForTeacherPlan = "00:00:00"
     DTPickExamdateForSeatPlan.Mask = "__/__/__"
     DTPickerStratTimeforseatPlan = "00:00:00"
    DTPickExamdateForSeatPlan.SetFocus
Else
    ComboRoomForTeacher.Text = ""
     ComboResponsibility = ""
     ComboTeacherName = ""
     DTPickerExamdateForTeacherPlan = Format(DTPickerExamdateForTeacherPlan, "__/__/__")
     DTPickerStartTimeForTeacherPlan = "00:00:00"
     
    ComboTeacherName.SetFocus
End If

End Sub

Private Sub cmdSAVE_Click()
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim con As New ADODB.connection
con.Open GConnString
cmd.ActiveConnection = con
If SSTab1.Tab = 0 Then
    If Len(ComboClass) = 0 Then

        MsgBox "Please Enter Class .", vbCritical, App.Title
        ComboClass.SetFocus
        Exit Sub
     End If
     If Len(ComboRoom) = 0 Then

        MsgBox "Please Enter Room No .", vbCritical, App.Title
        ComboRoom.SetFocus
        Exit Sub
     End If
      If Len(ComboSection) = 0 Then

        MsgBox "Please Enter Section .", vbCritical, App.Title
        ComboSection.SetFocus
        Exit Sub
     End If
     If Len(ComboShift) = 0 Then

        MsgBox "Please Enter Shift .", vbCritical, App.Title
        ComboShift.SetFocus
        Exit Sub
     End If
      If Len(ComboStratRoll) = 0 Then
        MsgBox "Please Enter Start Roll .", vbCritical, App.Title
        ComboStratRoll.SetFocus
        Exit Sub
     End If
     If Len(ComboEndRoll) = 0 Then

        MsgBox "Please Enter End Roll .", vbCritical, App.Title
        ComboEndRoll.SetFocus
        Exit Sub
     End If
     If DTPickerStratTimeforseatPlan.Hour = "0" Then

        MsgBox "Please Enter Start Time .", vbCritical, App.Title
        DTPickerStratTimeforseatPlan.SetFocus
        Exit Sub
     End If
     If DTPickExamdateForSeatPlan = Format(DTPickExamdateForSeatPlan, "__/__/__") Then

        MsgBox "Please Enter Exam Date .", vbCritical, App.Title
        DTPickExamdateForSeatPlan.SetFocus
        Exit Sub
     End If
     
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ExamSeatPlan"
    cmd(1) = Format(DTPickExamdateForSeatPlan, "dd mmm yyyy")
    cmd(2) = ComboRoom
    cmd(3) = Format(DTPickerStratTimeforseatPlan.Value, "hh:mm:ss")
    cmd(4) = ComboShift
    cmd(5) = Mid(ComboClass, 1, 5)
    cmd(6) = Mid(ComboSection, 1, 5)
    cmd(7) = ComboStratRoll
    cmd(8) = ComboEndRoll
    cmd(9) = "DSL"
    cmd(10) = Date
    cmd.Execute
    MsgBox "Information Save successfully.", vbInformation, "Student Management System"
    Call ShowFlexData
End If
If SSTab1.Tab = 1 Then
    If Len(ComboTeacherName) = 0 Then

        MsgBox "Please Enter Teacher Name .", vbCritical, App.Title
        ComboTeacherName.SetFocus
        Exit Sub
     End If
     If Len(ComboResponsibility) = 0 Then

        MsgBox "Please Enter responsibility .", vbCritical, App.Title
        ComboResponsibility.SetFocus
        Exit Sub
     End If
     If DTPickerStartTimeForTeacherPlan.Hour = "0" Then

        MsgBox "Please Enter Start Time .", vbCritical, App.Title
        DTPickerStartTimeForTeacherPlan.SetFocus
        Exit Sub
     End If
     If DTPickerExamdateForTeacherPlan = Format(DTPickerExamdateForTeacherPlan, "__/__/__") Then

        MsgBox "Please Enter Exam Date .", vbCritical, App.Title
        DTPickerExamdateForTeacherPlan.SetFocus
        Exit Sub
     End If

    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "ExamguardPlan1"
    cmd(1) = Format(DTPickerExamdateForTeacherPlan, "dd mm yyyy")
    cmd(2) = ComboRoomForTeacher
    cmd(3) = Mid(ComboTeacherName, 1, 5)
    cmd(4) = ComboResponsibility
    cmd(5) = Format(DTPickerStartTimeForTeacherPlan.Value, "hh:mm:ss")
    cmd(6) = "DSL"
    cmd(7) = Date
    cmd.Execute
    Call ShowFlexData2
    MsgBox "Information Save successfully.", vbInformation, "Student Management System"

End If


cmdnew.SetFocus
End Sub

Private Sub ComboClass_Click()
ComboSection.Clear
Set rs1 = getdata("select SectionId,Sectiondsc from SectionInfo where ClassId='" & Mid((ComboClass.Text), 1, 5) & "'")
If Not rs1.EOF Then
    Do Until rs1.EOF
       ComboSection.AddItem rs1(0) + " - " + rs1(1)
       rs1.MoveNext
    Loop

End If
End Sub

Private Sub ComboClass_KeyPress(KeyAscii As Integer)
Set rs1 = getdata("select SectionId,Sectiondsc from SectionInfo where ClassId='" & Mid((ComboClass.Text), 1, 5) & "'")
If Not rs1.EOF Then
    Do Until rs1.EOF
       ComboSection.AddItem rs1(0) + " - " + rs1(1)
       rs1.MoveNext
    Loop

End If
ComboSection.SetFocus
End Sub

Private Sub ComboEndRoll_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSAVE.SetFocus
End If
End Sub

Private Sub ComboResponsibility_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSAVE.SetFocus
End If
End Sub

Private Sub ComboRoom_Click()
If DTPickExamdateForSeatPlan = "__/__/__" Then Exit Sub
End Sub

Private Sub ComboRoom_GotFocus()
ComboRoom.Clear
Dim rs As New ADODB.Recordset
Set rs = getdata("select distinct RoomNo from ExamSitPlan")
If Not rs.EOF Then
    Do Until rs.EOF
    ComboRoom.AddItem rs!RoomNo
    rs.MoveNext
    Loop
End If
DTPickerStratTimeforseatPlan = "00:00:00"
End Sub

Private Sub ComboRoom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    DTPickerStratTimeforseatPlan.SetFocus
End If
End Sub

Private Sub ComboRoom_LostFocus()
If DTPickExamdateForSeatPlan = "__/__/__" Then Exit Sub
ShowFlexData
End Sub

Private Sub ComboRoomForTeacher_GotFocus()
Dim rs As New ADODB.Recordset

Set rs = getdata("select RoomNo from ExamSitPlan")
If Not rs.EOF Then
    ComboRoomForTeacher = rs!RoomNo
End If
ShowFlexData2
End Sub

Private Sub ComboRoomForTeacher_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    DTPickerStartTimeForTeacherPlan.SetFocus
End If

End Sub

Private Sub ComboRoomForTeacher_LostFocus()
ShowFlexData2
End Sub

Private Sub ComboSection_Click()
Set rs1 = getdata("select ClassRoll from StudentEvaluation where ClassId='" & Mid((ComboClass.Text), 1, 5) & "'and SectionId='" & Mid(ComboSection, 1, 5) & "'and Active='Y'")
If Not rs1.EOF Then
    Do Until rs1.EOF
        ComboStratRoll.AddItem rs1(0)
        ComboEndRoll.AddItem rs1(0)
        rs1.MoveNext
    Loop

End If
End Sub

Private Sub ComboSection_KeyPress(KeyAscii As Integer)
Set rs1 = getdata("select ClassRoll from StudentEvaluation where ClassId='" & Mid((ComboClass.Text), 1, 5) & "'and SectionId='" & Mid(ComboSection, 1, 5) & "'and Active='Y'")
If Not rs1.EOF Then
    Do Until rs1.EOF
        ComboStratRoll.AddItem rs1(0)
        ComboEndRoll.AddItem rs1(0)
        rs1.MoveNext
    Loop

End If
ComboStratRoll.SetFocus
End Sub

Private Sub ComboShift_Click()
'ComboClass.SetFocus
End Sub

Private Sub ComboShift_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboClass.SetFocus
End If

End Sub

Private Sub ComboStratRoll_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboEndRoll.SetFocus
End If
End Sub



Private Sub ComboTeacherName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ComboResponsibility.SetFocus
End If
End Sub



Private Sub DTPickerExamdateForTeacherPlan_GotFocus()
If DTPickerExamdateForTeacherPlan = "__/__/__" Then Exit Sub
ShowFlexData2
End Sub


Private Sub DTPickerExamdateForTeacherPlan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If DTPickExamdateForSeatPlan <> "__/__/__" Then
            If Check_ValidDate(DTPickExamdateForSeatPlan) = False Then
                DTPickExamdateForSeatPlan.SetFocus
                Exit Sub
            End If
    End If
    ComboRoomForTeacher.SetFocus
End If
End Sub


Private Sub DTPickerExamdateForTeacherPlan_LostFocus()
If DTPickerExamdateForTeacherPlan = "__/__/__" Then Exit Sub
End Sub

Private Sub DTPickerStartTimeForTeacherPlan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    ComboTeacherName.SetFocus
End If
End Sub

Private Sub DTPickerStratTimeforseatPlan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    ComboShift.SetFocus
End If
End Sub

Private Sub DTPickExamdateForSeatPlan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If DTPickExamdateForSeatPlan <> "__/__/__" Then
            If Check_ValidDate(DTPickExamdateForSeatPlan) = False Then
                DTPickExamdateForSeatPlan.SetFocus
                Exit Sub
            End If
    End If
    ComboRoom.SetFocus
End If
End Sub



Private Sub Form_Load()
SSTab1.Tab = 0
Dim rs1 As New ADODB.Recordset
Set rs1 = getdata("select TeacherId,TeacherName from TeacherInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
       ComboTeacherName.AddItem rs1(0) + "-" + rs1(1)
        rs1.MoveNext
    Loop
 
End If

Set rs1 = getdata("select ClassId,ClassName from ClassInfo")
If Not rs1.EOF Then
    Do Until rs1.EOF
       ComboClass.AddItem rs1(0) + " - " + rs1(1)
        rs1.MoveNext
    Loop
End If
With MSFlexGridSeat
    .Rows = 1
    .Cols = 6
    .Col = 0: .Text = " Shift   #"
    .Col = 1: .Text = " Class  "
    .Col = 2: .Text = " Section  "
    .Col = 3: .Text = " Start Roll  "
    .Col = 4: .Text = " End Roll  "
    .Col = 5: .Text = " Start Time  "
    .ColWidth(0) = 3000
    .ColWidth(1) = 5000
    .ColWidth(2) = 4000
    .ColWidth(3) = 3000
    .ColWidth(4) = 3000
    .ColWidth(5) = 2000
    
End With
With MSFlexGridTeacher
    .Rows = 1
    .Cols = 4
    .Col = 0: .Text = " ID   #"
    .Col = 1: .Text = " Name  "
    .Col = 2: .Text = " Reaponsibility  "
     .Col = 3: .Text = " Start Time  "
     .ColWidth(0) = 3000
    .ColWidth(1) = 5000
    .ColWidth(2) = 5000
     .ColWidth(2) = 3000
    
End With
End Sub

Private Sub MSFlexGridSeat_Click()
On Error GoTo errdes
ComboShift = MSFlexGridSeat.TextMatrix(MSFlexGridSeat.Row, 0)
ComboClass = MSFlexGridSeat.TextMatrix(MSFlexGridSeat.Row, 1)
ComboSection = MSFlexGridSeat.TextMatrix(MSFlexGridSeat.Row, 2)
ComboStratRoll = MSFlexGridSeat.TextMatrix(MSFlexGridSeat.Row, 3)
ComboEndRoll = MSFlexGridSeat.TextMatrix(MSFlexGridSeat.Row, 4)
DTPickerStratTimeforseatPlan = MSFlexGridSeat.TextMatrix(MSFlexGridSeat.Row, 5)

Exit Sub
errdes:
'MsgBox err.Descripti on, vbInformation, App.Title

End Sub

Private Sub MSFlexGridTeacher_Click()
On Error GoTo errdes
ComboTeacherName = (MSFlexGridTeacher.TextMatrix(MSFlexGridTeacher.Row, 0) + "-" + MSFlexGridTeacher.TextMatrix(MSFlexGridTeacher.Row, 1))
'ComboClass = MSFlexGridSeat.TextMatrix(MSFlexGridSeat.Row, 1)
ComboResponsibility = MSFlexGridTeacher.TextMatrix(MSFlexGridTeacher.Row, 2)
DTPickerStartTimeForTeacherPlan = MSFlexGridTeacher.TextMatrix(MSFlexGridTeacher.Row, 3)


Exit Sub
errdes:
'MsgBox err.Description, vbInformation, App.Title

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
If DTPickExamdateForSeatPlan = "__/__/__" Then Exit Sub
If Len(ComboRoom.Text) = 0 Then Exit Sub
    If Not DTPickExamdateForSeatPlan = "__/__/__" Then
        DTPickerExamdateForTeacherPlan.Mask = DTPickExamdateForSeatPlan
    
    End If
    If Len(ComboRoom.Text) <> 0 Then
        ComboRoomForTeacher.Text = ComboRoom.Text
   
    End If
    If Not DTPickerStratTimeforseatPlan.Value = "" Then
        DTPickerStartTimeForTeacherPlan.Value = DTPickerStratTimeforseatPlan
   
    End If
  DTPickerExamdateForTeacherPlan.SetFocus
End If
  
End Sub

Private Sub ShowFlexData()
'On Error GoTo ErrDes
Dim rs As New ADODB.Recordset

Set rs = getdata("SELECT     ExamSitPlan.Shift, ExamSitPlan.ClassId, ClassInfo.ClassName, ExamSitPlan.SectionId, SectionInfo.Sectiondsc, ExamSitPlan.StartRoll,ExamSitPlan.EndRoll , ExamSitPlan.StartTime " + _
"FROM  ExamSitPlan INNER JOIN ClassInfo ON ExamSitPlan.ClassId = ClassInfo.ClassID INNER JOIN " + _
"SectionInfo ON ExamSitPlan.SectionId = SectionInfo.SectionID AND ClassInfo.ClassID = SectionInfo.ClassID WHERE ExamSitPlan.ExamDate = '" & Format(DTPickExamdateForSeatPlan, "dd mmm yyyy") & "'AND ExamSitPlan.RoomNo = '" & ComboRoom.Text & "'")
    
    If Not rs.EOF Then
        i = 1
        With MSFlexGridSeat
        Do Until rs.EOF
            MSFlexGridSeat.Rows = i + 1

                
                .TextMatrix(i, 0) = "" & rs!Shift
                .TextMatrix(i, 1) = "" & rs!classId + "-" + rs!ClassName
                .TextMatrix(i, 2) = "" & rs!SectionID + "-" + rs!Sectiondsc
                .TextMatrix(i, 3) = "" & rs!StartRoll
                .TextMatrix(i, 4) = "" & rs!EndRoll
                .TextMatrix(i, 5) = "" & rs!Starttime
                

            rs.MoveNext
           i = i + 1
        Loop
    End With
Else
    MSFlexGridSeat.Rows = 1

 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub

Private Sub ShowFlexData1()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT     ExamGuardPlan.TeacherID, TeacherInfo.TeacherName, ExamGuardPlan.Responsibility,ExamGuardPlan.StartTime FROM ExamGuardPlan INNER JOIN " + _
"TeacherInfo ON ExamGuardPlan.TeacherID = TeacherInfo.TeacherId  ")
If Not rs.EOF Then
        i = 1
        With MSFlexGridTeacher
        Do Until rs.EOF
            MSFlexGridTeacher.Rows = i + 1

                .TextMatrix(i, 0) = "" & rs!teacherId
                .TextMatrix(i, 1) = "" & rs!teacherName
                .TextMatrix(i, 2) = "" & rs!Responsibility
                 .TextMatrix(i, 3) = "" & rs!Starttime

            rs.MoveNext
           i = i + 1
        Loop
    End With
Else
    MSFlexGridTeacher.Rows = 1

 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub


Private Sub ShowFlexData2()
On Error GoTo errdes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT     ExamGuardPlan.TeacherID, TeacherInfo.TeacherName, ExamGuardPlan.Responsibility,ExamGuardPlan.StartTime FROM ExamGuardPlan INNER JOIN " + _
"TeacherInfo ON ExamGuardPlan.TeacherID = TeacherInfo.TeacherId where ExamGuardPlan.examdate ='" & Format(DTPickerExamdateForTeacherPlan, "dd mmm yyyy") & "' and ExamGuardPlan.RoomNo='" & ComboRoomForTeacher.Text & "' and ExamGuardPlan.StartTime='" & DTPickerStratTimeforseatPlan & "'")
If Not rs.EOF Then
        i = 1
        With MSFlexGridTeacher
        Do Until rs.EOF
            MSFlexGridTeacher.Rows = i + 1

                
                .TextMatrix(i, 0) = "" & rs!teacherId
                .TextMatrix(i, 1) = "" & rs!teacherName
                .TextMatrix(i, 2) = "" & rs!Responsibility
                .TextMatrix(i, 3) = "" & rs!Starttime

            rs.MoveNext
           i = i + 1
        Loop
    End With
Else
    MSFlexGridTeacher.Rows = 1


 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub


Private Sub ShowFlexData3()
'On Error GoTo ErrDes
Dim rs As New ADODB.Recordset
Set rs = getdata("SELECT     ExamSitPlan.Shift, ExamSitPlan.ClassId, ClassInfo.ClassName, ExamSitPlan.SectionId, SectionInfo.Sectiondsc, ExamSitPlan.StartRoll, " + _
    "ExamSitPlan.EndRoll , ExamSitPlan.StartTime FROM         ExamSitPlan INNER JOIN " + _
    "ClassInfo ON ExamSitPlan.ClassId = ClassInfo.ClassID INNER JOIN " + _
    "SectionInfo ON ExamSitPlan.SectionId = SectionInfo.SectionID ")
    If Not rs.EOF Then
        i = 1
        With MSFlexGridSeat
        Do Until rs.EOF
            MSFlexGridSeat.Rows = i + 1

                
                .TextMatrix(i, 0) = "" & rs!Shift
                .TextMatrix(i, 1) = "" & rs!classId + "-" + rs!ClassName
                .TextMatrix(i, 2) = "" & rs!SectionID + "-" + rs!Sectiondsc
                .TextMatrix(i, 3) = "" & rs!StartRoll
                .TextMatrix(i, 4) = "" & rs!EndRoll
                .TextMatrix(i, 5) = "" & rs!Starttime
                

            rs.MoveNext
           i = i + 1
        Loop
        
    End With
 Else
     MSFlexGridSeat.Rows = 1

 End If
Exit Sub
errdes:
MsgBox Err.Description, vbInformation, App.Title
End Sub

