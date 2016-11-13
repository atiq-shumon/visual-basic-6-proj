VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   Caption         =   "Main Menu Screen"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   -2520
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Label lblUid 
      Caption         =   "Label1"
      Height          =   285
      Left            =   1650
      TabIndex        =   0
      Top             =   1650
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   8370
      Left            =   0
      Picture         =   "frmMain.frx":0442
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   12000
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "[&Setup]"
      Begin VB.Menu submnuClassInfo 
         Caption         =   "Class Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuBlank20 
         Caption         =   "-"
      End
      Begin VB.Menu submnuSubInfo 
         Caption         =   "Subject Information(Main)"
      End
      Begin VB.Menu mnublank21 
         Caption         =   "-"
      End
      Begin VB.Menu SIS 
         Caption         =   "Subject Information(Sub)"
      End
      Begin VB.Menu fgdsgsd 
         Caption         =   "-"
      End
      Begin VB.Menu submnusupplierinfo 
         Caption         =   "Supplier Information"
         Visible         =   0   'False
      End
      Begin VB.Menu mnublank23 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu SubmnuExamtypeSetUp 
         Caption         =   "Exam Term Set Up"
      End
      Begin VB.Menu mnuETS 
         Caption         =   "Exam Type Setup"
      End
      Begin VB.Menu mnublank22 
         Caption         =   "-"
      End
      Begin VB.Menu submnuMarksDistribution 
         Caption         =   "Marks Distribution"
         Begin VB.Menu submnuMarkCategory 
            Caption         =   "Mark Category"
         End
         Begin VB.Menu submnuMarkDistribution 
            Caption         =   "Mark Distribution"
            Shortcut        =   ^D
         End
      End
      Begin VB.Menu mnublank24 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScTypesetup 
         Caption         =   "Scholarship Type Setup"
      End
      Begin VB.Menu mnublank25 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScnameSetup 
         Caption         =   "Scholarship Name Setup"
      End
      Begin VB.Menu mnublank28 
         Caption         =   "-"
      End
      Begin VB.Menu mnufee 
         Caption         =   "&Fee Category Infomation"
      End
      Begin VB.Menu mnuFSI 
         Caption         =   "Fee Setup information"
      End
      Begin VB.Menu fdsfdsadsa 
         Caption         =   "-"
      End
      Begin VB.Menu SubmnuTctypeSetUp 
         Caption         =   "TC Type Set Up"
      End
      Begin VB.Menu fdgfsdgfsd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVS 
         Caption         =   "Vaccine Setup"
      End
      Begin VB.Menu gfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserCreation 
         Caption         =   "User Creation"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuOperation 
      Caption         =   " [&Entry]"
      Begin VB.Menu submnustudentinfromation 
         Caption         =   "Student Information"
         Shortcut        =   ^S
      End
      Begin VB.Menu submnustudentadmission 
         Caption         =   "Student Admission"
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuStuAdMissionApprove 
         Caption         =   "Student Admission Approve"
         Visible         =   0   'False
      End
      Begin VB.Menu fdsafdas 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRAE 
         Caption         =   "Re-Admission Entry"
      End
      Begin VB.Menu mnuSepRAD 
         Caption         =   "-"
      End
      Begin VB.Menu Submnustuattandence 
         Caption         =   "Student Attendance"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu submnuStudentLeave 
         Caption         =   "Student Leave"
      End
      Begin VB.Menu submnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookList 
         Caption         =   "Book List"
      End
      Begin VB.Menu submnusyllabusPreperation 
         Caption         =   "Syllabus Preperation"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu MnuClassRoutine 
         Caption         =   "Class Routine"
      End
      Begin VB.Menu submnuBlank3 
         Caption         =   "-"
      End
      Begin VB.Menu submnulessonplan 
         Caption         =   "Lesson Plan"
         Enabled         =   0   'False
      End
      Begin VB.Menu LPM 
         Caption         =   "Lesson Plan && Student Performance"
         Shortcut        =   ^L
      End
      Begin VB.Menu submnuBlank4 
         Caption         =   "-"
      End
      Begin VB.Menu submnuExamSchedule 
         Caption         =   "Exam Routine Entry"
      End
      Begin VB.Menu submnuExamRoutine 
         Caption         =   "Exam Routine"
         Visible         =   0   'False
      End
      Begin VB.Menu SubmnuExamseatPlan 
         Caption         =   "Exam Seat Plan"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu submnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu submnuexresultentry 
         Caption         =   "Ex. B. P Result Entry"
      End
      Begin VB.Menu submnuresultentry 
         Caption         =   "Result Entry"
         Shortcut        =   ^R
      End
      Begin VB.Menu submnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu submnufeescollection 
         Caption         =   "Fees Collection"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuspace 
         Caption         =   "-"
      End
      Begin VB.Menu submnubrinfo 
         Caption         =   "Book Recieve Information"
         Visible         =   0   'False
      End
      Begin VB.Menu SubmnuBookDisAndReturnInfo 
         Caption         =   "Book Distribution Information"
         Visible         =   0   'False
      End
      Begin VB.Menu SubmnuBookdisapp 
         Caption         =   "Book Distribution Approval"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBookreturnInfo 
         Caption         =   "Book Return Information"
         Visible         =   0   'False
      End
      Begin VB.Menu mnubookreturningapproval 
         Caption         =   "Book Returning Approval"
         Visible         =   0   'False
      End
      Begin VB.Menu submnuBlank5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu SubmnuTCPreperation 
         Caption         =   "TC Preperation"
      End
      Begin VB.Menu submnuTcApproval 
         Caption         =   "TC Preperation Approval"
      End
   End
   Begin VB.Menu mnuScholership 
      Caption         =   "[&Scholarship]"
      Begin VB.Menu mnuScholershipInfo 
         Caption         =   "Scholarship Information"
      End
   End
   Begin VB.Menu mnuothers 
      Caption         =   "[&Utility]"
      Begin VB.Menu submnulogout 
         Caption         =   "Log Out"
      End
      Begin VB.Menu submnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuRpt 
      Caption         =   "[&Report]"
      Begin VB.Menu mnurptAdmission 
         Caption         =   "Student Admission"
      End
      Begin VB.Menu mnublank27 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStdInfo 
         Caption         =   "Student Personal Information"
      End
      Begin VB.Menu drstgfdsgfds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAR 
         Caption         =   "Attendance Report"
      End
      Begin VB.Menu gfdsg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBAR 
         Caption         =   "Birthday Alert Report"
      End
      Begin VB.Menu fdsafdsafdsa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCollection 
         Caption         =   "Collection Report"
      End
      Begin VB.Menu ytey 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMDB 
         Caption         =   "Marksdistribution"
      End
      Begin VB.Menu fgh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMSP 
         Caption         =   "Marksheet Preparation"
      End
      Begin VB.Menu gfdsgsd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSOP 
         Caption         =   "Statement of Progress"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "[Help]"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LPM_Click()
  frmLessonPlanMain.Show 1
End Sub

Private Sub mnuAbout_Click()
 frmAbout.Show 1
End Sub

Private Sub mnuAR_Click()
 frmStudentAttendanceReport.Show 1
End Sub

Private Sub mnuBAR_Click()
  rptBirthalert.Show 1
End Sub

Private Sub mnuBookList_Click()
FrmBookList.Show 1
End Sub

Private Sub mnuBookreturnInfo_Click()
Frmdistributedbookreturn.Show 1
End Sub

Private Sub mnubookreturningapproval_Click()
FrmBookreturnapprove.Show 1
End Sub

Private Sub MnuClassRoutine_Click()
 frmClassRoutine.Show 1
End Sub

Private Sub submnuAdmissionApprove_Click()

End Sub

Private Sub SubmnuAttendance_Click()

End Sub

Private Sub mnuCollection_Click()
  rptCollection.Show 1
End Sub

Private Sub mnuETS_Click()
  frmExamSetUp.Show 1
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnufee_Click()
  frmFeeInfo.Show 1
End Sub

Private Sub mnuFSI_Click()
  frmFeesetupInfo.Show 1
End Sub

Private Sub mnuMDB_Click()
  RptMarksDistribution.Show 1
End Sub

Private Sub mnuMSP_Click()
  RptMarksheetAll.Show 1
End Sub

Private Sub mnuRAE_Click()
  Dim f As New frmstudentREAdmission
  f.Show 1
End Sub

Private Sub mnurptAdmission_Click()
rptAdmissionInfo.Show 1
End Sub

Private Sub mnuScholershipInfo_Click()
FrmScholership.Show 1
End Sub

Private Sub mnuScnameSetup_Click()
frmScholershipNameInfo.Show 1
End Sub

Private Sub mnuScTypesetup_Click()
frmscholershiptypeinfo.Show 1
End Sub

Private Sub mnuSOP_Click()
   Rpt_stdprogress.Show 1
End Sub

Private Sub mnuStdInfo_Click()
rptStdInfo.Show 1
End Sub

Private Sub MnuStuAdMissionApprove_Click()
   frmAdmissionApprove.Show 1
End Sub

Private Sub mnuUserCreation_Click()
    FrmCreateUser.Show 1
End Sub

Private Sub mnuVS_Click()
  frmVaccineSetUp.Show 1
End Sub

Private Sub SIS_Click()
  frmsubjectinfo_sub.Show 1
End Sub

Private Sub SubmnuBookDisAndReturnInfo_Click()
Frmbookdistribution.Show 1
End Sub

Private Sub SubmnuBookdisapp_Click()
frmBookdistributedApprovedInfo.Show 1

End Sub

Private Sub submnubrinfo_Click()
frmbookrecieveentry.Show 1
End Sub

Private Sub submnuClassInfo_Click()
Dim f As New frmclassinfo
f.Show 1
End Sub

Private Sub submnuExamRoutine_Click()
FrmExamRoutine.Show 1
End Sub

Private Sub submnuExamSchedule_Click()
 frmExamSchedule.Show 1
End Sub

Private Sub SubmnuExamseatPlan_Click()
frmExamSeatPlan.Show 1
End Sub

Private Sub SubmnuExamtypeSetUp_Click()
   frmExamTypeSetUp.Show 1
End Sub

Private Sub submnuexit_Click()
End
End Sub

Private Sub submnuexresultentry_Click()
  frmStudentExResult.Show 1
End Sub

Private Sub submnufeescollection_Click()
  frmCollection_info.Show 1
End Sub

Private Sub submnulessonplan_Click()
 frmlectureinfo.Show 1
End Sub

Private Sub submnuMarkCategory_Click()
Dim f As New frmMarksCat
f.Show 1
End Sub

Private Sub submnuMarkDistribution_Click()
Dim f As New frmsubmarksdis
f.Show 1
End Sub

Private Sub submnuresultentry_Click()
  frmStudentResult.Show 1
End Sub

Private Sub Submnustuattandence_Click()
frmStudentAttendance.Show 1
End Sub

Private Sub submnustudentadmission_Click()
  frmstudentadmission.Show 1
End Sub

Private Sub submnustudentinfromation_Click()
   frmstudentInfo.Show vbModal
End Sub

Private Sub submnuStudentLeave_Click()
   frmStudentleave.Show 1
End Sub

Private Sub submnuSubInfo_Click()
Dim f As New frmsubjectinfo
f.Show 1
End Sub

Private Sub submnusupplierinfo_Click()
   FrmSupplierInfo.Show 1
End Sub

Private Sub submnusyllabusPreperation_Click()
Frmsyllabuspreperation.Show 1
End Sub

Private Sub submnuTcApproval_Click()
  frmtcinfoapprove.Show 1
End Sub

Private Sub SubmnuTCPreperation_Click()
frmTCPreperation.Show 1
End Sub

Private Sub SubmnuTctypeSetUp_Click()
 frmTcType.Show 1
End Sub
