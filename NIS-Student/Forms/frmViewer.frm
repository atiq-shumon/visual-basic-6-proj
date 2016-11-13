VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmViewer 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      DragIcon        =   "frmViewer.frx":0442
      Height          =   8445
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   11985
      lastProp        =   500
      _cx             =   21140
      _cy             =   14896
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StdAdmissionInfo As New CryAdminssionInfo
Dim StudentInfo As New CrysSpecificStudentInfoDetails
Dim rptmr As New CrystalReport1
Dim rptbrt As New CrystalReport2
Dim rptExamRoutine As New CrystalReport3
Dim rptAttendance As New CrystalReport4
Dim rptPerformance As New CrystalReport5
Dim rptresult As New CrystalReport6
Dim rptmarksheet As New CrystalReport7
Dim rpt_marksdistribution As New CrystalReport8
Dim Rpt_std_progress As New CrystalReport9
Dim rpt_marksheet_All As New CrystalReport10
Dim CrysClassRoutine As New CrystalReport11
Dim CrysBPMarks As New CrysExBPMarks
Dim CrysAttenReportYear As New CrysAttendenceReportofYearly
Dim CrysAllStudentDetails As New CrysAllStudentDetails
Dim varCrysAttendenceReportDateToDate As New CrysAttendenceReportDateToDate
Dim varCrysAttendenceReportToDay As New CrysAttendenceReportToDay
Dim varCrysSpecStudAttendenceReportDateToDate As New CrysSpecStudAttendenceReportDateToDate
Dim rptCollectionRpt As New CrystalReport12

Private Sub Form_Load()
    CRViewer91.Zoom 100
    Dim Conn As New ADODB.connection
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
Select Case rptMode
    Case 0
            Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
                      
            cmd.CommandText = "{CALL rptStudentAdmisionInfo ('" & rptAdmissionInfo.CboClassID & " ','" & rptAdmissionInfo.DTPicker1(0) & "', '" & rptAdmissionInfo.DTPicker1(1) & "')}"
            Set rs = cmd.Execute
            
            Call Company
            StdAdmissionInfo.Text11.SetText CompanyName
            StdAdmissionInfo.Text12.SetText Address
            
            If rptAdmissionInfo.Frame3.Visible = True Then
                StdAdmissionInfo.Text16.SetText (rptAdmissionInfo.txtClassName.Text)
            Else
            If rptAdmissionInfo.Frame4.Visible = True Then
                StdAdmissionInfo.Text16.SetText (rptAdmissionInfo.DTPicker1(0).Value & "   To  " & rptAdmissionInfo.DTPicker1(1).Value)
            End If
            End If
            
            StdAdmissionInfo.Database.SetDataSource rs
            StdAdmissionInfo.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = StdAdmissionInfo
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
    Case 1
            Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
                      
            cmd.CommandText = "{CALL rptStudentInfo ('" & rptStdInfo.cboStdID & " ')}"
            Set rs = cmd.Execute
            
            Call Company
'            StudentInfo.Text20.SetText CompanyName
            StudentInfo.Text21.SetText Address
            
            StudentInfo.Database.SetDataSource rs
            StudentInfo.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = StudentInfo
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
            
     Case 2
        Conn.Open GConnString
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText
                  
        cmd.CommandText = "SELECT * FROM RPT_MR ('S'," & Val(frmCollection_info.txtfields(6)) & ")"
        Set rs = cmd.Execute
        
        Call Company
        rptmr.Text17.SetText (frmCollection_info.txtfields(6))
'        rptmr.Text20.SetText CompanyName
'        rptmr.Text12.SetText Address
'
        rptmr.Database.SetDataSource rs
        rptmr.DiscardSavedData
        CRViewer91.ReportSource = rptmr
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault

Case 3
            Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
                      
            cmd.CommandText = "{CALL rptBrtalert (0,'" & rptBirthalert.DTPicker1(0) & "', '" & rptBirthalert.DTPicker1(1) & "')}"
            Set rs = cmd.Execute
            
            Call Company
            rptbrt.Text20.SetText CompanyName
            rptbrt.Text12.SetText Address
'
'            If rptAdmissionInfo.Frame3.Visible = True Then
'                StdAdmissionInfo.Text16.SetText (rptAdmissionInfo.txtClassName.Text)
'            Else
'            If rptAdmissionInfo.Frame4.Visible = True Then
'                StdAdmissionInfo.Text16.SetText (rptAdmissionInfo.DTPicker1(0).Value & "   To  " & rptAdmissionInfo.DTPicker1(1).Value)
'            End If
'            End If
'
            rptbrt.Database.SetDataSource rs
            rptbrt.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = rptbrt
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault

      Case 4
            Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
                      
            cmd.CommandText = "{CALL Rpt_exam_routine ('a','" & Trim(Mid(frmExamSchedule.ComboClass, 1, 5)) & "','" & Trim(frmExamSchedule.ComboYear) & "', '" & Trim(Mid(frmExamSchedule.ComboExamName, 1, 2)) & "','" & Trim(Mid(frmExamSchedule.Combo1, 1, 2)) & "')}"
            Set rs = cmd.Execute
            
            Call Company
'            rptExamRoutine.Text2.SetText CompanyName
'            rptExamRoutine.Text3.SetText Address
'''
'            If rptAdmissionInfo.Frame3.Visible = True Then
'                StdAdmissionInfo.Text16.SetText (rptAdmissionInfo.txtClassName.Text)
'            Else
'            If rptAdmissionInfo.Frame4.Visible = True Then
'                StdAdmissionInfo.Text16.SetText (rptAdmissionInfo.DTPicker1(0).Value & "   To  " & rptAdmissionInfo.DTPicker1(1).Value)
'            End If
'            End If
'

           rptExamRoutine.Text7.SetText (Trim(Mid(frmExamSchedule.ComboExamName.Text, 5, 35)) + "  ( " + Trim(Mid(frmExamSchedule.Combo1, 5, 25)) + " )" + "'" + frmExamSchedule.ComboYear)
            rptExamRoutine.Database.SetDataSource rs
            rptExamRoutine.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = rptExamRoutine
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault

 Case 5
            Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
                      
            cmd.CommandText = "{CALL Rpt_attendance ('a','" & Trim(Mid(frmStudentAttendance.List1, 1, 5)) & "','" & Trim(Mid(frmStudentAttendance.Combo1.Text, 1, 5)) & "','" & Trim(Mid(frmStudentAttendance.Combo2, 1, 1)) & "', '" & Format(frmStudentAttendance.MaskEdBox3, "dd mmm yyyy") & "' )}"
            Set rs = cmd.Execute
 
            rptAttendance.Text13.SetText (Mid(frmStudentAttendance.List1, 7, 36))
            rptAttendance.Text14.SetText (Mid(frmStudentAttendance.Combo1, 9, 36))
            rptAttendance.Text15.SetText (frmStudentAttendance.Combo2)
            rptAttendance.Text9.SetText (frmStudentAttendance.MaskEdBox3)
                   
            rptAttendance.Database.SetDataSource rs
            rptAttendance.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = rptAttendance
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault


 Case 6
            Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
                      
            cmd.CommandText = "{CALL rptStudentPerformance (1,'" & Trim(frmLessonPlanMain.Combo1) & "','" & Trim(frmLessonPlanMain.Combo5) & "'," & Trim(frmStudent_performance.txtfields(2)) & ", " & Trim(Val(frmStudent_performance.txtfields(3))) & "," & Trim(Val(frmStudent_performance.CmbDetail)) & ",'" & Trim(frmTopic_details.txtfields(6)) & "' )}"
            Set rs = cmd.Execute
 
            rptPerformance.Text10.SetText (frmLessonPlanMain.Combo3)
            rptPerformance.Text15.SetText (frmStudent_performance.Combo1)
'            rptAttendance.Text14.SetText (Mid(frmStudentAttendance.Combo1, 9, 36))
'            rptAttendance.Text15.SetText (frmStudentAttendance.Combo2)
'            rptAttendance.Text9.SetText (frmStudentAttendance.MaskEdBox3)
'
            rptPerformance.Database.SetDataSource rs
            rptPerformance.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = rptPerformance
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault

Case 7
           Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
                      
            cmd.CommandText = "{CALL Rpt_marks ('a'," & Trim(frmStudentResult.txtSerial) & ")}"
            Set rs = cmd.Execute
 
            rptresult.Text10.SetText (Mid(frmStudentResult.cboClass, 7, 36))
            rptresult.Text13.SetText (Mid(frmStudentResult.CboSection, 7, 36))
            rptresult.Text14.SetText (Mid(frmStudentResult.CboSubject, 7, 36) + "  ( " + Mid(frmStudentResult.CboCategory, 7, 20) + "  )")
            rptresult.Text15.SetText (frmStudentResult.cmdAcaYear)
            rptresult.Text16.SetText (Mid(frmStudentResult.CboExamType, 3, 36))
            rptresult.Text17.SetText (Mid(frmStudentResult.CboExamID, 3, 36))
          
            rptresult.Database.SetDataSource rs
            rptresult.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = rptresult
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault

Case 8
           Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
                      
            cmd.CommandText = "select * from Rpt_marks_Sheet('a','" & Trim(frmStudentResult.txtid) & "','" & Trim(frmStudentResult.cmdAcaYear) & "','" & Mid(Trim(frmStudentResult.cboClass), 1, 5) & "','" & Trim(Mid(frmStudentResult.CboExamType, 1, 2)) & "','" & Trim(Mid(frmStudentResult.CboExamID, 1, 2)) & "','" & Trim(cmdAcaYear) & "','" & Mid(Trim(frmStudentResult.cboClass), 1, 5) & "','" & Mid(Trim(frmStudentResult.cboClass), 1, 5) & "')"
            Set rs = cmd.Execute
 
            rptmarksheet.Text10.SetText (Mid(frmStudentResult.cboClass, 7, 36))
            rptmarksheet.Text13.SetText (Mid(frmStudentResult.CboSection, 7, 36))
           '' rptmarksheet.Text14.SetText (Mid(frmStudentResult.CboSubject, 7, 36) + "  ( " + Mid(frmStudentResult.CboCategory, 7, 20) + "  )")
            rptmarksheet.Text15.SetText (" ( " + frmStudentResult.cmdAcaYear + " ) ")
            rptmarksheet.Text16.SetText (Mid(frmStudentResult.CboExamType, 3, 36))
            rptmarksheet.Text17.SetText (Mid(frmStudentResult.CboExamID, 3, 36))
              
             rptmarksheet.Text12.SetText (Trim(frmStudentResult.lblStdname))
             rptmarksheet.Text11.SetText (Trim(frmStudentResult.txtid))

            rptmarksheet.Database.SetDataSource rs
            rptmarksheet.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = rptmarksheet
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
   Case 9
             Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
   
            cmd.CommandText = " Select * from Rpt_marks_Sheet_all ('a','" & Trim(Mid(RptMarksheetAll.cboClass, 1, 5)) & "','" & Trim(Mid(RptMarksheetAll.CboSection, 1, 5)) & "','" & Trim(Mid(RptMarksheetAll.CboSubject, 1, 5)) & "','" & Trim(RptMarksheetAll.cmdAcaYear) & "','" & Trim(Mid(RptMarksheetAll.CboExamType, 1, 2)) & "','" & Trim(Mid(RptMarksheetAll.CboExamID, 1, 2)) & "')"
            Set rs = cmd.Execute
 
            rpt_marksheet_All.Text42.SetText (Mid(RptMarksheetAll.cboClass, 7, 36))
            rpt_marksheet_All.Text43.SetText (Mid(RptMarksheetAll.CboSection, 7, 36))
            rpt_marksheet_All.Text44.SetText (Mid(RptMarksheetAll.CboSubject, 7, 36))
            rpt_marksheet_All.Text45.SetText (RptMarksheetAll.cmdAcaYear)
            rpt_marksheet_All.Text46.SetText (Mid(RptMarksheetAll.CboExamType, 3, 36))
            rpt_marksheet_All.Text47.SetText (Mid(RptMarksheetAll.CboExamID, 3, 36))

            rpt_marksheet_All.Database.SetDataSource rs
            rpt_marksheet_All.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = rpt_marksheet_All
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault

            
 Case 10
             Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
   
            cmd.CommandText = " Select * from Rpt_marks_distribution ('a','" & Trim(Mid(RptMarksDistribution.cboClass, 1, 5)) & "','" & Trim(Mid(RptMarksDistribution.CboExamType, 1, 2)) & "','" & Trim(Mid(RptMarksDistribution.CboExamID, 1, 2)) & "')"
            Set rs = cmd.Execute
 
             rpt_marksdistribution.Text19.SetText (Mid(RptMarksDistribution.cboClass, 7, 36))
'            rpt_marksheet_All.Text43.SetText (Mid(RptMarksheetAll.CboSection, 7, 36))
'            rpt_marksheet_All.Text44.SetText (Mid(RptMarksheetAll.CboSubject, 7, 36))
'            rpt_marksheet_All.Text45.SetText (RptMarksheetAll.cmdAcaYear)
            rpt_marksdistribution.Text20.SetText (Mid(RptMarksDistribution.CboExamType, 3, 36))
            rpt_marksdistribution.Text21.SetText (Mid(RptMarksDistribution.CboExamID, 3, 36))

            rpt_marksdistribution.Database.SetDataSource rs
            rpt_marksdistribution.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = rpt_marksdistribution
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
         

 Case 11
             Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
   
            cmd.CommandText = " Select * from Rpt_statement_of_prog ('a','" & Trim(Mid(Rpt_stdprogress.cboClass, 1, 5)) & "','" & Trim(Mid(Rpt_stdprogress.CboSection, 1, 5)) & "','" & Trim(Mid(Rpt_stdprogress.CboExamType, 1, 2)) & "','" & Trim(Mid(Rpt_stdprogress.CboExamID, 1, 2)) & "','" & Trim(Rpt_stdprogress.txtfields(0)) & "','" & Trim(Rpt_stdprogress.cmdAcaYear) & "','" & Trim(Mid(Rpt_stdprogress.Combo2, 1, 1)) & "')"
            Set rs = cmd.Execute
            Rpt_std_progress.Text63.SetText (Rpt_stdprogress.txtfields(0))
            Rpt_std_progress.Text65.SetText (Rpt_stdprogress.txtfields(1))
            Rpt_std_progress.Text66.SetText (Mid(Rpt_stdprogress.cboClass, 7, 36))
            Rpt_std_progress.Text68.SetText (Mid(Rpt_stdprogress.CboSection, 7, 36))
            Rpt_std_progress.Text67.SetText (Rpt_stdprogress.cmdAcaYear)
            Rpt_std_progress.Text70.SetText (Mid(Rpt_stdprogress.CboExamType, 3, 36))
            Rpt_std_progress.Text69.SetText (Mid(Rpt_stdprogress.CboExamID, 3, 36))

            Rpt_std_progress.Database.SetDataSource rs
            Rpt_std_progress.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = Rpt_std_progress
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
            
        Case 12
             Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
   
            cmd.CommandText = " select * from Rpt_Class_Routine(1,'" & frmClassRoutine.List1.Text & "', '" & Trim(Get_Code(frmClassRoutine.Combo1)) & "','" & Trim(Mid(frmClassRoutine.Combo2.Text, 1, 1)) & "','" & Trim(Get_Code(frmClassRoutine.Combo3.Text)) & "','" & Trim(frmClassRoutine.Combo5.Text) & "')"
            Set rs = cmd.Execute
'            Rpt_std_progress.Text63.SetText (Rpt_stdprogress.txtfields(0))
'            Rpt_std_progress.Text65.SetText (Rpt_stdprogress.txtfields(1))
             CrysClassRoutine.Text66.SetText (Get_Description(frmClassRoutine.Combo1.Text))
             CrysClassRoutine.Text68.SetText (Get_Description(frmClassRoutine.Combo3))
             CrysClassRoutine.Text67.SetText (frmClassRoutine.Combo5)
             CrysClassRoutine.Text4.SetText (frmClassRoutine.Combo2)
'            Rpt_std_progress.Text69.SetText (Mid(Rpt_stdprogress.CboExamID, 3, 36))

            CrysClassRoutine.Database.SetDataSource rs
            CrysClassRoutine.DiscardSavedData
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = CrysClassRoutine
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
            
            
        ' Ex. B. P. Marks Distribution
        Case 13
            
            
            Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText
            
            Set rs = getdata("SELECT     a.StdID,(SELECT     studentname From StudentInfo WHERE studentid = a.stdid) " _
                & " AS stdname, a.Roll, a.CWObtainedMarks, a.HWObtainedMarks, a.AttentivenessObtainedMarks," _
                & " a.CleannessObtainedMarks, a.MannersObtainedMarks, a.S_Slr_no, a.M_Slr_no, ClassInfo.ClassID," _
                & " ClassInfo.ClassName, EXB_P_Result_Main.AcaYr,EXB_P_Result_Main.ExamType, EXB_P_Result_Main.ExamID," _
                & " Subject_info_sub_History.Teacher_id, Emp_Per_Info.Emp_fna,Emp_Per_Info.Emp_mna, Emp_Per_Info.Emp_lna," _
                & " EXB_P_Result_Main.SubID, EXB_P_Result_Main.ClassID AS ClassID,Subject_info_sub_History.Sub_code," _
                & " Subject_info_sub_History.Class_code, Subject_info_sub.Sub_title, EXB_P_Result_Main.SectionID," _
                & " SectionInfo.ClassTeacher FROM Subject_info_sub_History INNER JOIN EXB_P_Result_Sub a INNER JOIN " _
                & " EXB_P_Result_Main ON a.M_Slr_no = EXB_P_Result_Main.M_Slr_no INNER JOIN ClassInfo ON " _
                & " EXB_P_Result_Main.ClassID = ClassInfo.ClassID ON Subject_info_sub_History.Sub_code = " _
                & " EXB_P_Result_Main.SubID AND Subject_info_sub_History.Class_code = EXB_P_Result_Main.ClassID " _
                & " INNER JOIN Subject_info_sub ON Subject_info_sub_History.Sub_code = Subject_info_sub.Sub_code AND " _
                & " Subject_info_sub_History.Class_code = Subject_info_sub.Class_code INNER JOIN SectionInfo ON " _
                & " EXB_P_Result_Main.SectionID = SectionInfo.SectionID AND EXB_P_Result_Main.ClassID = " _
                & " SectionInfo.ClassID INNER JOIN Emp_Per_Info ON SectionInfo.ClassTeacher = Emp_Per_Info.Emp_id " _
                & " where EXB_P_Result_Main.ClassID='" & Mid(frmStudentExResult.cboClass, 1, 5) & "'" _
                & " and EXB_P_Result_Main.SectionID='" & Mid(frmStudentExResult.CboSection, 1, 5) & "'" _
                & " and EXB_P_Result_Main.Shift='" & Mid(frmStudentExResult.Combo2, 1, 1) & "'" _
                & " and EXB_P_Result_Main.SubID='" & Mid(frmStudentExResult.CboSubject, 1, 5) & "'" _
                & " and EXB_P_Result_Main.AcaYr='" & Trim(frmStudentExResult.cmdAcaYear) & "'" _
                & " and  EXB_P_Result_Main.ExamType='" & Mid(frmStudentExResult.CboExamType, 1, 2) & "'" _
                & " and EXB_P_Result_Main.ExamID='" & Mid(frmStudentExResult.CboExamID, 1, 2) & "' ORDER BY a.StdID")
            
            CrysBPMarks.DiscardSavedData
            CrysBPMarks.Database.SetDataSource rs
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = CrysBPMarks
            
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
         
        '=======================================================================
        '=====================Attendence report of the year=======================
     Case 14
            
            Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText

         
        Set rs = getdata("SELECT     COUNT(StudentAttendanceLeaveInfo.Present) AS TotalPresent, StudentInfo.StudentID," _
            & " StudentInfo.StudentName, StudentInfo.StuFatherName,StudentInfo.StuMotherName, StudentInfo.StuDateOfBirth," _
            & " StudentInfo.StuBloodGroup, StudentAttendanceLeaveInfo.ClassRoll,StudentAttendanceLeaveInfo.Present," _
            & " StudentAttendanceLeaveInfo.aca_yr, ClassInfo.ClassName, ClassInfo.ShIftname FROM StudentInfo INNER JOIN " _
            & " StudentAttendanceLeaveInfo ON StudentInfo.StudentID = StudentAttendanceLeaveInfo.StudentID INNER JOIN " _
            & " ClassInfo ON StudentAttendanceLeaveInfo.ClassID = ClassInfo.ClassID " _
            & " WHERE     (StudentAttendanceLeaveInfo.ClassID = '" & Get_Code(frmStudentAttendanceReport.cmbClass) & "') " _
            & " AND StudentAttendanceLeaveInfo.Shift = '" & Mid(frmStudentAttendanceReport.cmbSpecificShift, 1, 1) & "'" _
            & " AND (StudentAttendanceLeaveInfo.SectionID = '" & Get_Code(frmStudentAttendanceReport.cmdSection) & "')" _
            & " AND (StudentAttendanceLeaveInfo.aca_yr = '" & frmStudentAttendanceReport.cmdAcademicYear & "')" _
            & " AND (StudentAttendanceLeaveInfo.Present = 'P') " _
            & " GROUP BY StudentInfo.StudentName, ClassInfo.ShIftname, ClassInfo.ClassName, StudentInfo.StudentID, " _
            & " StudentInfo.StudentName, StudentInfo.StuFatherName, StudentInfo.StuMotherName, StudentInfo.StuDateOfBirth," _
            & " StudentInfo.StuBloodGroup, StudentAttendanceLeaveInfo.ClassRoll, StudentAttendanceLeaveInfo.Present ," _
            & " StudentAttendanceLeaveInfo.aca_yr, ClassInfo.ClassName, ClassInfo.ShIftname " _
            & " order by StudentAttendanceLeaveInfo.ClassRoll")
        
        CrysAttenReportYear.DiscardSavedData
        CrysAttenReportYear.Database.SetDataSource rs
        
'        Set rs = getdata("SELECT     COUNT(leave) AS TotalLeave, ClassRoll, aca_yr From StudentAttendanceLeaveInfo " _
'            & " WHERE     (ClassID = '" & Get_Code(frmStudentAttendanceReport.cmbClass) & "') " _
'            & " AND (Shift = '" & Mid(frmStudentAttendanceReport.cmbSpecificShift, 1, 1) & "') " _
'            & " AND (SectionID = '" & Get_Code(frmStudentAttendanceReport.cmdSection) & "') " _
'            & " AND (aca_yr = '" & frmStudentAttendanceReport.cmdAcademicYear & "')" _
'            & " AND (leave = 'Y')GROUP BY ClassRoll, leave, aca_yr ORDER BY ClassRoll ")
'
'            CrysAttenReportYear.txtLeave.SetText rs(0)
            Screen.MousePointer = vbHourglass
            CRViewer91.ReportSource = CrysAttenReportYear
            
            CRViewer91.ViewReport
            Screen.MousePointer = vbDefault
        
   Case 15
            Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText

         
        Set rs = getdata(" SELECT     StudentInfo.StudentID, StudentInfo.StudentName, StudentInfo.StuFatherName, StudentInfo.StuMotherName, StudentInfo.StuBloodGroup," _
            & " StudentInfo.StuStreetPAddress, StudentInfo.StuPDistrict, District.DisDistrictName, ClassInfo.ClassName, ClassInfo.ShIftname," _
            & " StudentEvaluation.classId , StudentEvaluation.SectionID, SectionInfo.Sectiondsc, SectionInfo.ClassTeacher,StudentEvaluation.ClassRoll, " _
            & " StudentInfo.StuDateOfBirth, StudentInfo.StuReligion, StudentInfo.StuPhone" _
            & " FROM         StudentInfo INNER JOIN StudentEvaluation ON StudentInfo.StudentID = StudentEvaluation.StudentID INNER JOIN " _
            & " District ON StudentInfo.StuPDistrict = District.DisDistrictCode INNER JOIN ClassInfo ON StudentEvaluation.ClassId = ClassInfo.ClassID INNER JOIN " _
            & " SectionInfo ON StudentEvaluation.ClassId = SectionInfo.ClassID AND StudentEvaluation.SectionId = SectionInfo.SectionID " _
            & " where StudentEvaluation.ClassId='" & rptStdInfo.CboClassID & "' order by StudentEvaluation.ClassRoll ")
            
        CrysAllStudentDetails.DiscardSavedData
        CrysAllStudentDetails.Database.SetDataSource rs
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = CrysAllStudentDetails
        
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
    
    Case 16
        Conn.Open GConnString
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText

         
        Set rs = getdata(" SELECT     COUNT(*) AS Expr1, a.StudentID,(SELECT     studentname From StudentInfo " _
            & " WHERE      studentid = a.studentid) AS stdname, a.ClassRoll, a.Present, a.PresentCancel, a.aca_yr, a.leave, SectionInfo.Sectiondsc," _
            & " ClassInfo.ClassName , a.classId, a.SectionID FROM StudentAttendanceLeaveInfo a INNER JOIN " _
            & " SectionInfo ON a.ClassID = SectionInfo.ClassID AND a.SectionID = SectionInfo.SectionID INNER JOIN " _
            & " ClassInfo ON SectionInfo.ClassID = ClassInfo.ClassID " _
            & " WHERE     a.attn_date BETWEEN '" & Format(frmStudentAttendanceReport.FromDate, "dd MMM yyyy") & "' AND '" & Format(frmStudentAttendanceReport.toDate, "dd MMM yyyy") _
            & "' AND a.Present = 'p' AND a.ClassID = '" & Get_Code(frmStudentAttendanceReport.cmbClass.Text) & "' and a.aca_yr='" & frmStudentAttendanceReport.cmdAcademicYear & "' " _
            & " GROUP BY a.StudentID, a.ClassRoll, a.Present, a.PresentCancel, a.aca_yr, a.leave, SectionInfo.Sectiondsc, a.ClassID, a.SectionID, ClassInfo.ClassName ")

            
        varCrysAttendenceReportDateToDate.DiscardSavedData
        varCrysAttendenceReportDateToDate.Database.SetDataSource rs
        
        varCrysAttendenceReportDateToDate.txtDateBetween.SetText "From " & Format(frmStudentAttendanceReport.FromDate, "dd/MM/yyyy") & " To " & Format(frmStudentAttendanceReport.toDate, "dd/MM/yyyy") & ""
        
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = varCrysAttendenceReportDateToDate
        
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        
    Case 17
            Conn.Open GConnString
            Set cmd.ActiveConnection = Conn
            cmd.CommandType = adCmdText

         
        Set rs = getdata(" SELECT StudentInfo.StudentID, StudentInfo.StudentName, StudentInfo.StuFatherName," _
            & " StudentInfo.StuMotherName, StudentInfo.StuDateOfBirth, StudentInfo.StuBloodGroup, " _
            & " StudentAttendanceLeaveInfo.ClassRoll,StudentAttendanceLeaveInfo.Present, StudentAttendanceLeaveInfo.aca_yr," _
            & " ClassInfo.ClassName, ClassInfo.ShIftname,StudentAttendanceLeaveInfo.attn_date," _
            & " StudentAttendanceLeaveInfo.Present, StudentAttendanceLeaveInfo.Leave " _
            & " FROM         StudentInfo INNER JOIN StudentAttendanceLeaveInfo ON StudentInfo.StudentID = StudentAttendanceLeaveInfo.StudentID INNER JOIN " _
            & " ClassInfo ON StudentAttendanceLeaveInfo.ClassID = ClassInfo.ClassID " _
            & " WHERE     (StudentAttendanceLeaveInfo.ClassID = '" & Get_Code(frmStudentAttendanceReport.cmbClass) & "') " _
            & " AND StudentAttendanceLeaveInfo.Shift = '" & Mid(frmStudentAttendanceReport.cmbSpecificShift, 1, 1) & "'" _
            & " AND (StudentAttendanceLeaveInfo.SectionID = '" & Get_Code(frmStudentAttendanceReport.cmdSection) & "')" _
            & " AND (StudentAttendanceLeaveInfo.aca_yr = '" & frmStudentAttendanceReport.cmdAcademicYear & "')" _
            & " and StudentAttendanceLeaveInfo.attn_date='" & Format(frmStudentAttendanceReport.FromDate, "dd MMM YYYY") & "'")
            
            
        varCrysAttendenceReportToDay.DiscardSavedData
        varCrysAttendenceReportToDay.Database.SetDataSource rs
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = varCrysAttendenceReportToDay
        
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        
    '======================= Specific Student Date to Date======================================================
    Case 18
        Conn.Open GConnString
        Set cmd.ActiveConnection = Conn
        cmd.CommandType = adCmdText

         
        Set rs = getdata(" SELECT     a.StudentID,(SELECT     studentname From StudentInfo " _
            & " WHERE      studentid = a.studentid) AS stdname, a.ClassRoll, a.Present, a.PresentCancel, a.aca_yr, a.leave, SectionInfo.Sectiondsc," _
            & " ClassInfo.ClassName , a.classId, a.SectionID,a.Leave,a.attn_date FROM StudentAttendanceLeaveInfo a INNER JOIN " _
            & " SectionInfo ON a.ClassID = SectionInfo.ClassID AND a.SectionID = SectionInfo.SectionID INNER JOIN " _
            & " ClassInfo ON SectionInfo.ClassID = ClassInfo.ClassID " _
            & " WHERE     a.attn_date BETWEEN '" & Format(frmStudentAttendanceReport.FromDate, "dd MMM yyyy") & "' AND '" & Format(frmStudentAttendanceReport.toDate, "dd MMM yyyy") & "'" _
            & " and a.StudentID='" & Get_Code(frmStudentAttendanceReport.cmdStudentID) & "'" _
            & " AND a.ClassID = '" & Get_Code(frmStudentAttendanceReport.cmbClass.Text) & "' and a.aca_yr='" & frmStudentAttendanceReport.cmdAcademicYear & "' ")
            '& " GROUP BY a.StudentID, a.ClassRoll, a.Present,a.Leave, a.PresentCancel, a.aca_yr, a.leave, SectionInfo.Sectiondsc, a.ClassID, a.SectionID, ClassInfo.ClassName ")

            
        varCrysSpecStudAttendenceReportDateToDate.DiscardSavedData
        varCrysSpecStudAttendenceReportDateToDate.Database.SetDataSource rs
        varCrysSpecStudAttendenceReportDateToDate.txtFromDate.SetText "From " & Format(frmStudentAttendanceReport.FromDate, "dd/MM/yy") & " To " & Format(frmStudentAttendanceReport.toDate, "dd/MM/yy") & ""
        
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = varCrysSpecStudAttendenceReportDateToDate
        
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        
    Case 19
         Set cmd = Nothing
         Conn.Open GConnString
         Set cmd.ActiveConnection = Conn
         cmd.CommandType = adCmdStoredProc
         cmd.CommandText = " rptStudentCollectionInfo(" & Param_mode & ",'" & rptCollection.cboStdID & "','','')"
         
         
        Set rs = cmd.Execute
            
'            StudentInfo.Text21.SetText Address

        rptCollectionRpt.Database.SetDataSource rs
        rptCollectionRpt.DiscardSavedData
        Screen.MousePointer = vbHourglass
        CRViewer91.ReportSource = rptCollectionRpt
        CRViewer91.ViewReport
        Screen.MousePointer = vbDefault
        
End Select
End Sub
Private Sub Form_Resize()
    CRViewer91.Top = 0
    CRViewer91.Left = 0
    CRViewer91.Height = ScaleHeight
    CRViewer91.Width = ScaleWidth
End Sub
''Private Sub Form_Unload(Cancel As Integer)
'''    RS.Close
''End Sub

