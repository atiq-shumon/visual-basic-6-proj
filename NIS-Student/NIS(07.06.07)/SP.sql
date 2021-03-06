if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ SP_Attendance_count]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ SP_Attendance_count]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Access_Log_I_U]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Access_Log_I_U]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AddSQLBranchUser]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AddSQLBranchUser]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Amount_From_Increment_current]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Amount_From_Increment_current]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Approval_Update]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Approval_Update]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AuthorInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AuthorInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BMEntryInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BMEntryInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BonusOT_AtOnce_Unit1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BonusOT_AtOnce_Unit1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Bonus_Preparation_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Bonus_Preparation_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Bonus_Preparation_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Bonus_Preparation_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookAuthor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookAuthor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookDisretInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookDisretInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookDisretInformation1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookDisretInformation1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookDisretInformation2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookDisretInformation2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookDisretInformation3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookDisretInformation3]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookIssueSave]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookIssueSave]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookList1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookList1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookPurchase]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookPurchase]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookPurchaseSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookPurchaseSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookRequisition]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookRequisition]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookRequisitionSub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookRequisitionSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookReturnSave]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BookReturnSave]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Branch_info_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Branch_info_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Change_PW]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Change_PW]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Check_Current_time]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Check_Current_time]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Check_N_Store_Attn_Backup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Check_N_Store_Attn_Backup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClassInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ClassInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClassRoutine1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ClassRoutine1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Collec_master_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Collec_master_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Collec_sub_save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Collec_sub_save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Comp_Det_Info_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Comp_Det_Info_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Company_Name_Address]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Company_Name_Address]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Company_Policy_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Company_Policy_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Company_Policy_View]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Company_Policy_View]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Count_HOLIDAY_LEAVE_ATTENDANCE]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Count_HOLIDAY_LEAVE_ATTENDANCE]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Create_User_AtOnce]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Create_User_AtOnce]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DaleteAuthorInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[DaleteAuthorInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DataBaseBackup_Restore]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[DataBaseBackup_Restore]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Date_Formats]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Date_Formats]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DeleteBookAuthor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[DeleteBookAuthor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DeleteBookInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[DeleteBookInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DeleteBookIssueRule]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[DeleteBookIssueRule]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DeleteBookIssueRule1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[DeleteBookIssueRule1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DeleteLibrarySubInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[DeleteLibrarySubInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DeletePublisherInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[DeletePublisherInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Delete_Attendance_Record]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Delete_Attendance_Record]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Department_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Department_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Discard_Prepared_OT_Benefits]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Discard_Prepared_OT_Benefits]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Discard_Prepared_Salary]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Discard_Prepared_Salary]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ETypeInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ETypeInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[EmpIncrementInformationInfo_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[EmpIncrementInformationInfo_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Att_Info_I_U]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_Att_Info_I_U]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Att_Notes_I_U]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_Att_Notes_I_U]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Discipline_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_Discipline_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Education_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_Education_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Job_End_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_Job_End_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Job_Hist_Current_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_Job_Hist_Current_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Job_Hist_Future_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_Job_Hist_Future_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Leave_Info_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_Leave_Info_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Movement_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_Movement_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Per_Info_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_Per_Info_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Reference_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_Reference_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Salary_Payscale_Hist]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_Salary_Payscale_Hist]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_View]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_View]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_performance_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_performance_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_super_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Emp_super_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExamSeatPlan]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ExamSeatPlan]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Exam_Type_Info_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Exam_Type_Info_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExamguardPlan1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ExamguardPlan1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExaminatioSchedule]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ExaminatioSchedule]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExaminationRoutine]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ExaminationRoutine]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Export_To_Acct_Vou]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Export_To_Acct_Vou]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FacultySchedule_Record_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[FacultySchedule_Record_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FacultySchedule_Record_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[FacultySchedule_Record_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fee_info_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Fee_info_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fee_setup_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Fee_setup_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fetch_Emp_Id]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Fetch_Emp_Id]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fixed_Head_POP]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Fixed_Head_POP]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fixed_Pay_POP]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Fixed_Pay_POP]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fixed_Variable_Gen_Deduc_Payroll]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Fixed_Variable_Gen_Deduc_Payroll]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fx]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Fx]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_All_from_OT_Fix]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_All_from_OT_Fix]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Attn_Summery]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Attn_Summery]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Columns]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Columns]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Day_Consumed]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Day_Consumed]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Emp_Information]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Emp_Information]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Employees]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Employees]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Jobtitle]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Jobtitle]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_New_Id]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_New_Id]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_New_Id_DCL]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_New_Id_DCL]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_OT_Days_Ind]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_OT_Days_Ind]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_OT_Fix_Ind]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_OT_Fix_Ind]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Overtime]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Overtime]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_PF_SETUP_etc]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_PF_SETUP_etc]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Photo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Photo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Rank_Category]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Rank_Category]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Server_Time]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Server_Time]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_Worker_Time]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_Worker_Time]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_title_Code]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Get_title_Code]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Group_NM_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Group_NM_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Group_Shift_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Group_Shift_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Hol_List_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Hol_List_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HrsInMovement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[HrsInMovement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Import_From_Ayman_Acct]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Import_From_Ayman_Acct]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[IncomeTaxBrkup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[IncomeTaxBrkup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Increment_History_I_U]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Increment_History_I_U]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Increment_History_New_I_U]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Increment_History_New_I_U]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Ins_Salary_Head]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Ins_Salary_Head]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Insert_Into_Fixed_Pay]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Insert_Into_Fixed_Pay]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[IssueRules]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[IssueRules]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Job_Hist_Future_POP]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Job_Hist_Future_POP]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Job_Title_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Job_Title_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Job_Type_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Job_Type_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LS_PLAN_DETAILS_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LS_PLAN_DETAILS_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LS_PLAN_MASTER_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LS_PLAN_MASTER_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LS_PLAN_TOPIC_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LS_PLAN_TOPIC_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Leave_Info_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Leave_Info_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Leave_List_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Leave_List_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Leave_validation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Leave_validation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LectureInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LectureInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LibrarySubInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LibrarySubInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Load_Place]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Load_Place]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Load_leave_duration]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Load_leave_duration]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Loan_Sanction_Information_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Loan_Sanction_Information_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Loan_Sanction_Information_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Loan_Sanction_Information_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Loan_Type_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Loan_Type_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Loan_Type_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Loan_Type_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LogIn_LogOut_Modified]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[LogIn_LogOut_Modified]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Log_sub1_I_U]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Log_sub1_I_U]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Login_logout]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Login_logout]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Markscategorydes]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Markscategorydes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Match_U_ID_N_POP]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Match_U_ID_N_POP]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Maternity_Leave_Validation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Maternity_Leave_Validation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Monthly_Attendance_Backup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Monthly_Attendance_Backup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Monthly_Attn_Summery]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Monthly_Attn_Summery]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Move_Sch_For_Move_Out]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Move_Sch_For_Move_Out]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Move_Viewer]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Move_Viewer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OT_Fix_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[OT_Fix_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OT_Fix_IU]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[OT_Fix_IU]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OT_Hr]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[OT_Hr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OT_Summery_All]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[OT_Summery_All]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OT_Summery_SectionWise]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[OT_Summery_SectionWise]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Office_Time_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Office_Time_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Out_of_Office_Del]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Out_of_Office_Del]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Out_of_Office_IU]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Out_of_Office_IU]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Overtime_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Overtime_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Overtime_IU]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Overtime_IU]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Overtime_Preparation_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Overtime_Preparation_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Overtime_Preparation_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Overtime_Preparation_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Closing_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Closing_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Comp_Withdraw_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Comp_Withdraw_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Components_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Components_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Contribution_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Contribution_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Inst_Scm_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Inst_Scm_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Installment_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Installment_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Investment_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Investment_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Ln_Amendment_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Ln_Amendment_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Ln_Application_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Ln_Application_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Ln_Approval_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Ln_Approval_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Ln_Cancellation_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Ln_Cancellation_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Ln_Policy_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Ln_Policy_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Loan_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Loan_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_POP_Tables]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_POP_Tables]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PF_Pay_History_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PF_Pay_History_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_All_About_Group_Shift]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_All_About_Group_Shift]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Discipline]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Discipline]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Emp_Discipline]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Emp_Discipline]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Emp_Education]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Emp_Education]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Emp_NM_Not_Paid]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Emp_NM_Not_Paid]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Emp_NonUser]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Emp_NonUser]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Emp_Ref]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Emp_Ref]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Emp_detail_for_Payroll]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Emp_detail_for_Payroll]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Holidays]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Holidays]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_ID_Salary_Fixing]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_ID_Salary_Fixing]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Id_Sal_Fixed_Or_Not]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Id_Sal_Fixed_Or_Not]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Increment_future]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Increment_future]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Ind_Emp_Sal_Fix]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Ind_Emp_Sal_Fix]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Job_Setup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Job_Setup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Leave_App_Aprv]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Leave_App_Aprv]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_N_Insert_Approval]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_N_Insert_Approval]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Param]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Param]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Payscale_Main]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Payscale_Main]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Payscale_Sub]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Payscale_Sub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Salary_HeadCode]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Salary_HeadCode]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Tax_Slab]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Tax_Slab]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Taxable_Ceiling]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Taxable_Ceiling]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Tier_Emp_id_Sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Tier_Emp_id_Sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Tier_Setup]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Tier_Setup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_Tr_Determine_Move_Pos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_Tr_Determine_Move_Pos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_User]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_User]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[POP_UserName_SPrivilege]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[POP_UserName_SPrivilege]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pay_Struc_POP]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pay_Struc_POP]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Payscale_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Payscale_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Payscale_Main_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Payscale_Main_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Payscale_Sub_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Payscale_Sub_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pop_Approval]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pop_Approval]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pop_Rpt_Holiday_List]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pop_Rpt_Holiday_List]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pro_Param_Entry]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Pro_Param_Entry]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Promotion_Record_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Promotion_Record_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PublisherInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[PublisherInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rank_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rank_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Renew_FixedPay]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Renew_FixedPay]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Result_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Result_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RprLogin_Fail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[RprLogin_Fail]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RptBookStockInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[RptBookStockInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RptPromotion_Info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[RptPromotion_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Bonus_Preparation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Bonus_Preparation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Daily_Attendance]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Daily_Attendance]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Daily_Not_Present]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Daily_Not_Present]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Deptwise_Leave]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Deptwise_Leave]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Emp_Increment_Info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Emp_Increment_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Employee_Info_Various]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Employee_Info_Various]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Faculy_Schedule_Preparation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Faculy_Schedule_Preparation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_FiscalYr_SalStateAll]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_FiscalYr_SalStateAll]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_FiscalYr_Salary_Statement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_FiscalYr_Salary_Statement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Leave_Ind]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Leave_Ind]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Loan_Sanction_Info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Loan_Sanction_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Monthly_Salary_Statement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Monthly_Salary_Statement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Movement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Movement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_OT_Preparation_monthwise]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_OT_Preparation_monthwise]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Overtime]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Overtime]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_PF_Contribution_Monthwise]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_PF_Contribution_Monthwise]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Perm_Attn]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Perm_Attn]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Perm_Leave]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Perm_Leave]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Todays_Movement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_Todays_Movement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_WhoGetsIncrement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_WhoGetsIncrement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_YearlySalary_Statement]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_YearlySalary_Statement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_exam_routine]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_exam_routine]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_marks]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Rpt_marks]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_All_Id_name_POP]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_All_Id_name_POP]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Attendance_count]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Attendance_count]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Check_ID]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Check_ID]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Data_mapping_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Data_mapping_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Designation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Designation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Emp_Desig_Sec_Join]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Emp_Desig_Sec_Join]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Emp_Id_with_OT_pay]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Emp_Id_with_OT_pay]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_FIX_The_Rest_Fixed_Head]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_FIX_The_Rest_Fixed_Head]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Find_Emp_per_info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Find_Emp_per_info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Fixed_Pay_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Fixed_Pay_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Holiday_List_All]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Holiday_List_All]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Job_Detail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Job_Detail]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Job_Detail_Mod]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Job_Detail_Mod]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Job_End_All]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Job_End_All]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Leave_List_All]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Leave_List_All]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Leave_app_Pop]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Leave_app_Pop]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Leave_app_with_id_name]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Leave_app_with_id_name]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Move_In]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Move_In]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Move_In_Mod]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Move_In_Mod]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Move_Out_In_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Move_Out_In_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Move_Schedule_Pop]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Move_Schedule_Pop]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Move_schedule_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Move_schedule_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Movement_POP]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Movement_POP]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_OT_Pay_Pop]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_OT_Pay_Pop]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Office_Time_Worker_Official]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Office_Time_Worker_Official]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Office_time_All]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Office_time_All]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Pay_Struc_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Pay_Struc_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Payroll_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Payroll_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Payroll_sub_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Payroll_sub_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_Performance_eval]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_Performance_eval]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_W_ID_NM_For_J]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_W_ID_NM_For_J]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SP_W_Performance_eval]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_W_Performance_eval]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[STD_PERFORMANCE_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[STD_PERFORMANCE_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Salary_Adjustment]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Salary_Adjustment]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Salary_AtOnce_Unit1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Salary_AtOnce_Unit1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Salary_AtOnce_Unit2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Salary_AtOnce_Unit2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Salry_HeadSetUp_Delete]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Salry_HeadSetUp_Delete]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Salry_HeadSetUp_Save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Salry_HeadSetUp_Save]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SaveLogin_Fail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SaveLogin_Fail]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ScholerShipInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ScholerShipInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ScholerShipNameSetupInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ScholerShipNameSetupInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ScholerShipTypeInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ScholerShipTypeInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sec_Info_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Sec_Info_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SectionInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SectionInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Section_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Section_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Seek_Id_Not_Paid]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Seek_Id_Not_Paid]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Set_Param]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Set_Param]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Shift_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Shift_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Show_PreviousTiers]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Show_PreviousTiers]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Software_Previleges]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Software_Previleges]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sp_Emp_Loan_Info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Sp_Emp_Loan_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sp_Rpt_Emp_Per]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Sp_Rpt_Emp_Per]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sp_rpt_Emp_Att]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Sp_rpt_Emp_Att]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StuAdmissionEvaluationInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[StuAdmissionEvaluationInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StudentAttendance]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[StudentAttendance]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StudentEvaluation1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[StudentEvaluation1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StudentInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[StudentInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StudentInformation1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[StudentInformation1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StudentInformation2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[StudentInformation2]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Studentadmission1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Studentadmission1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SubjectInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SubjectInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SubjectInformation_SUB]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SubjectInformation_SUB]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SubjectInformation_SUB_Teacher]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SubjectInformation_SUB_Teacher]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SubjectInformation_main]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SubjectInformation_main]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Subjectmarksdistribution1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Subjectmarksdistribution1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SupplierInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SupplierInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Syllabuspreperation1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Syllabuspreperation1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TCInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[TCInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TCInfoApprove]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[TCInfoApprove]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TCType]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[TCType]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Tax_Setup_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Tax_Setup_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Tax_Slab_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Tax_Slab_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Taxable_Ceiling_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Taxable_Ceiling_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Tbl]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Tbl]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Tier_setup_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Tier_setup_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Time_Keeper_I_U]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Time_Keeper_I_U]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Uptodate_Job_detail]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Uptodate_Job_detail]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Validate_PW]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Validate_PW]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[VaxinInformation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[VaxinInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Vou_I_Payroll_Main_U]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Vou_I_Payroll_Main_U]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Worker_Group_I_U_D]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Worker_Group_I_U_D]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Yearly_Holidays]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Yearly_Holidays]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[add_scrn]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[add_scrn]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[emp_bank_account_I_U]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[emp_bank_account_I_U]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[generate_position]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[generate_position]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[getStudentInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[getStudentInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[get_max_Eff_date]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[get_max_Eff_date]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[give_pmt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[give_pmt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_pass_entry]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_pass_entry]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_security]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_security]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_security_entry]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_security_entry]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pro_soft_pass]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pro_soft_pass]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rptBrtalert]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rptBrtalert]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rptEducation_Info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rptEducation_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rptStudentAdmisionInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rptStudentAdmisionInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rptStudentInfo]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rptStudentInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_Bank_Statment]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rpt_Bank_Statment]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_Emp_Att_Info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rpt_Emp_Att_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_Emp_Att_Summery]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rpt_Emp_Att_Summery]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_Emp_Leave_Info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rpt_Emp_Leave_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_Emp_Performance]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rpt_Emp_Performance]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_Employee_list]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rpt_Employee_list]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_Hol_list]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rpt_Hol_list]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_Payroll_Ind_All]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rpt_Payroll_Ind_All]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_Payroll_Payslip]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rpt_Payroll_Payslip]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_Provident_Fund_Statment]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rpt_Provident_Fund_Statment]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_SalStat_SendtoBank]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[rpt_SalStat_SendtoBank]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[s_u_d_leave_info]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[s_u_d_leave_info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[size_pmt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[size_pmt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_leaves_only_for_u]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_leaves_only_for_u]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[sp_set_position]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[sp_set_position]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[temp_Collec_save]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[temp_Collec_save]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo. SP_Attendance_count    Script Date: 3/7/02 10:43:42 AM ******/
/****** Object:  Stored Procedure dbo. SP_Attendance_count    Script Date: 10/28/01 4:34:36 PM ******/
/****** Object:  Stored Procedure dbo. SP_Attendance_count    Script Date: 10/22/01 6:24:18 PM ******/
/****** Object:  Stored Procedure dbo. SP_Attendance_count    Script Date: 9/17/00 4:33:55 PM ******/
/****** Object:  Stored Procedure dbo. SP_Attendance_count    Script Date: 9/4/01 6:30:54 PM ******/
/****** Object:  Stored Procedure dbo. SP_Attendance_count    Script Date: 6/20/01 12:00:38 PM ******/
/****** Object:  Stored Procedure dbo. SP_Attendance_count    Script Date: 5/14/01 9:49:53 AM ******/
/****** Object:  Stored Procedure dbo. SP_Attendance_count    Script Date: 1/1/99 3:01:42 AM ******/
/****** Object:  Stored Procedure dbo. SP_Attendance_count    Script Date: 5/9/01 8:13:06 AM ******/
CREATE PROCEDURE [ SP_Attendance_count]
@emp_Id varchar (10)
as
select count(Emp_Login) from Emp_Att_info where emp_id=@emp_Id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE   Procedure  Access_Log_I_U
@opr varchar(5)
,@U_Id varchar(10)

As 

set nocount on

declare @Access_Id int

set @Access_Id=(Select dbo.GetAuto_SlNo('Acs_Log'))

if @opr='In'
	begin
		insert into Access_Log (U_Id,Access_Id,LogIn)
		values (@U_Id,@Access_Id,getdate())		

		select Access_Id=@Access_Id
	end


if @opr='Out'
	begin
		Update Access_Log set LogOut=getdate()
		where Access_Id=(select max(Access_Id)from Access_Log)
	end


set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


































-- Exec sp_dropuser '16163457'
-- Exec sp_droplogin '16163457'
-- Delete from BranchUser Where BranchCode = '16345'

-- exec AddSQLBranchUser '16345','16163457','16345','0h68:','DaffodilBank','Alamgir',''
create                   Procedure AddSQLBranchUser
	(
		@BranchUser varchar(8),
		@Password varchar(15),
		@Database varchar(20)
	)
--WITH ENCRYPTION -- This is security symbol, Please do not remove it.
As
Set xact_abort On
Set nocount on

if exists (select * from master.dbo.syslogins where loginname = @BranchUser)
	begin
		Exec sp_dropuser @BranchUser
		Exec sp_droplogin @BranchUser
	end 

Exec sp_addlogin @BranchUser,@password,@Database
Exec sp_adduser @BranchUser,@BranchUser,db_owner
EXEC sp_addrolemember  db_accessadmin ,@BranchUser
EXEC sp_addrolemember  db_backupoperator ,@BranchUser

-------- Adds a login as a member of a fixed server role-------
EXEC sp_addsrvrolemember @BranchUser, 'sysadmin'
EXEC sp_addsrvrolemember @BranchUser, 'securityadmin'
EXEC sp_addsrvrolemember @BranchUser, 'setupadmin'
EXEC sp_addsrvrolemember @BranchUser, 'processadmin'
EXEC sp_addsrvrolemember @BranchUser, 'serveradmin'
EXEC sp_addsrvrolemember @BranchUser, 'diskadmin'
EXEC sp_addsrvrolemember @BranchUser, 'dbcreator'
EXEC sp_addsrvrolemember @BranchUser, 'bulkadmin '
----------------------------------------------------------
	













GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Amount_From_Increment_current    Script Date: 3/7/02 10:43:28 AM ******/
/****** Object:  Stored Procedure dbo.Amount_From_Increment_current    Script Date: 10/28/01 4:34:20 PM ******/
/****** Object:  Stored Procedure dbo.Amount_From_Increment_current    Script Date: 10/22/01 6:24:21 PM ******/
/****** Object:  Stored Procedure dbo.Amount_From_Increment_current    Script Date: 9/17/00 4:33:55 PM ******/
/****** Object:  Stored Procedure dbo.Amount_From_Increment_current    Script Date: 9/4/01 6:30:43 PM ******/
CREATE PROCEDURE [Amount_From_Increment_current]
@emp_id varchar(10)
 AS
Select amount from increment_current where emp_id=@Emp_id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Approval_Update    Script Date: 3/7/02 10:43:28 AM ******/
/****** Object:  Stored Procedure dbo.Approval_Update    Script Date: 10/28/01 4:34:21 PM ******/
/****** Object:  Stored Procedure dbo.Approval_Update    Script Date: 10/22/01 6:24:21 PM ******/
/****** Object:  Stored Procedure dbo.Approval_Update    Script Date: 9/17/00 4:33:56 PM ******/
/****** Object:  Stored Procedure dbo.Approval_Update    Script Date: 9/4/01 6:30:43 PM ******/
---------------------------------------------------------------------------
CREATE procedure Approval_Update
@opr varchar(1),
@Current_Pos varchar(1),
@App_id int,
@Policy_Code varchar(10),
@Tier_1 varchar (10), 
@Tier_1_Chk varchar (1),
@Tier_1_Remarks varchar(50),
@Tier_2 varchar(10),
@Tier_2_Chk varchar(1),
@Tier_2_Remarks varchar(50),
@Tier_3 varchar(10),
@Tier_3_Chk varchar(1),
@Tier_3_Remarks varchar(50),
@Tier_4 varchar(10),
@Tier_4_Chk varchar(1),
@Tier_4_Remarks varchar(1),
@Final_Tier varchar(10),
@Final_Tier_Chk varchar(1),
@Final_Tier_Remarks varchar(50),
@Move_Pos char(1),
@U_id varchar(10)
as
if @Current_Pos='1'
begin
	update Approval set 
	Tier_1_Chk=@Tier_1_Chk,
	Tier_1_Remarks=@Tier_1_Remarks,
	Move_Pos=@Move_Pos,U_id=@U_id
	where App_id=@App_id
end
if @Current_Pos='2'
begin
	update Approval set 
	Tier_2_Chk=@Tier_2_Chk,Tier_2_Remarks=@Tier_2_Remarks,
	Move_Pos=@Move_Pos,U_id=@U_id
	where App_id=@App_id
end
if @Current_Pos='3'
begin
	update Approval set 
	Tier_3_Chk=@Tier_3_Chk,Tier_3_Remarks=@Tier_3_Remarks,
	Move_Pos=@Move_Pos,U_id=@U_id
	where App_id=@App_id
end
if @Current_Pos='4'
begin
	update Approval set 
	Tier_4_Chk=@Tier_4_Chk,Tier_4_Remarks=@Tier_4_Remarks,
	Move_Pos=@Move_Pos,U_id=@U_id
	where App_id=@App_id
end
if @Current_Pos='5'
begin
	update Approval set 
	Final_Tier_Chk=@Final_Tier_Chk,Final_Tier_Remarks=@Final_Tier_Remarks,
	Move_Pos=@Move_Pos,U_id=@U_id
	where App_id=@App_id
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE    procedure AuthorInformation(
	@AuthorCode	VARCHAR(5),
	@AuthorName	VARCHAR(80),
	@Remarks VARCHAR(80),
	@User VARCHAR (10))
AS
SET NOCOUNT ON

IF NOT EXISTS (SELECT * FROM AuthorInfo WHERE     (AuthorCode = @AuthorCode))
	INSERT INTO AuthorInfo
        (AuthorCode, AuthorName, AuthorNote, AuthorEntryBy, AuthorEntryDate)
		VALUES     (@AuthorCode,@AuthorName,@Remarks,@User,GETDATE())
ELSE
	UPDATE  AuthorInfo SET              
			AuthorName = @AuthorName, AuthorNote = @Remarks
			WHERE     (AuthorCode = @AuthorCode)

SET NOCOUNT OFF








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





--R-00000001
---exec SupplierInformation '','sdfdsf','dsfdsf','2008-12-12','sdfsd'

CREATE procedure    BMEntryInformation
(
	@RecieveNo		Varchar(10),
	@SuppID			varchar(10),
	@RecDate		DateTime,
	@ClassID		varchar(10),
	@SubjectID		varchar(10),
	@Qty			Money,
	@Notes			Varchar(100),
	@EntryDate		DateTime,
	@EntryBY		varchar(10)
	
)

as

Declare 

	@MaxSLNo int,
	--@SuID char(10)
	@RecNo char(10)


if @RecieveNo =''

	begin
	
		



		SELECT @MaxSLNo =   isnull(max(cast(substring(RecieveNo,3,10) as int)),0)FROM BookRecievedInfo
		
		if @MaxSLNo=0
			set @MaxSLNo=1
		else
			set @MaxSLNo=@MaxSLNo+1

		select @RecNo = dbo.PadString(@MaxSLNo,2,'0','L')

				
		set @RecieveNo ='R-'+ @RecNo 
	end


 if exists (select * from BookRecievedInfo where RecieveNo= @RecieveNo ) 		
Update BookRecievedInfo set

	
	RecieveNo 	= 	@RecieveNo		,
	SuppID 		= 	@SuppID			,
	RecDate 	= 	@RecDate		,
	ClassID 	= 	@ClassID		,
	SubjectID	= 	@SubjectID		,
	Qty 		= 	@Qty			,
	Notes 		= 	@Notes			,
	EntryDate	= 	@EntryDate		,
	EntryBY 	= 	@EntryBY		
				
	
	where RecieveNo 	=	@RecieveNo		
else
 insert into BookRecievedInfo

(
	RecieveNo 	,
	SuppID 		,
	RecDate 	,
	ClassID 	,
	SubjectID	,
	Qty 		,
	Notes 		,
	EntryDate	,
	EntryBY 	
		
		
	
)
values
(
	@RecieveNo		,
	@SuppID			,
	@RecDate		,
	@ClassID		,
	@SubjectID		,
	@Qty			,
	@Notes			,
	@EntryDate		,
	@EntryBY		
	
	
)








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/*
delete from payroll_main
exec Salary_AtOnce_Unit1 'may','2002','dsl'
select * from payroll_main
select Head_code, Head_name from pay_struc where mode='B'

select * from payroll_main
select * from payroll_sub

select * from param_tbl

BonusOT_AtOnce_Unit1 15,'July','2002','dsl'

-------------------------------------------------
@Pay_Type---->head code

*/

create   Procedure BonusOT_AtOnce_Unit1
@Pay_Type int			---------Eid Bonus 15, AGM 16 etc.
,@Pay_Month varchar(12)
,@Pay_Year varchar(4)
,@U_id varchar(10)
AS

set nocount on

declare @Emp_ID varchar(10)
,@Job_Type varchar(30)
,@Job_Duration int
,@Latest_A_Gen varchar(8)
,@Bns_OT_Amount money
,@Policy_No int
,@Message varchar(150)
,@Count int
/****************************************************************************/
set @Count=0


if @Pay_Type != 21	---if not 21 (OT)----Eid Bonus, AGM, BATEXPO, Election Bonus etc.
begin
	
	set @Job_Type=(select dbo.GetParamFlag (dbo.GetPayType_PlcNo(@Pay_Type,1)))
	set @Job_Duration=convert(int,(select dbo.GetParamValue (dbo.GetPayType_PlcNo(@Pay_Type,1))))
	
	if @Job_Type='1'	-------Permanent Job

	begin
		set @Job_Type='Permanent'
	
		DECLARE Get_ID_Cursor CURSOR FOR
			select distinct emp_id from fixed_pay where emp_id in 
				(Select emp_id from emp_Job_Hist_Current 
					where Emp_job_type=@Job_Type
					and datediff(month,Emp_join_date,getdate())>=@Job_Duration) 
					and Emp_Id not in 
						(Select emp_id from payroll_main 
							where pay_month=@Pay_Month 
								and pay_year=@Pay_year 
								and pay_Type=@pay_Type)
	end
		
	------------------------------------------------------------------------

	if @Job_Type='0'	------All Employees

	begin
		DECLARE Get_ID_Cursor CURSOR FOR
			select distinct emp_id from fixed_pay where emp_id in 
				(Select emp_id from emp_Job_Hist_Current 
					where datediff(month,Emp_join_date,getdate())>=@Job_Duration) 
					and Emp_Id not in 
						(Select emp_id from payroll_main 
							where pay_month=@Pay_Month 
								and pay_year=@Pay_year 
								and pay_Type=@pay_Type)
	end


END

/****************************************************************************/
if @Pay_Type = 21 		-------OT

BEGIN

set @Job_Type=(select dbo.GetParamValue (dbo.GetPayType_PlcNo(@Pay_Type,0)))

	DECLARE Get_ID_Cursor CURSOR FOR
		select distinct emp_id from fixed_pay 
			where emp_id in 
				(Select emp_id from emp_Job_Hist_Current 
						where Emp_job_type=@Job_Type
							and Pay_type!=1)	
					------0 wage base,1 salary base
			and Emp_Id in 
				(Select emp_id from payroll_main 
					where pay_month=@Pay_Month 
						and pay_year=@Pay_year 
						and pay_Type=0)
					------If salary is prepared

			and Emp_Id not in 
				(Select emp_id from payroll_main 
				where pay_month=@Pay_Month 
					and pay_year=@Pay_year 
					and pay_Type=@pay_Type)
					------not yet prepared
END
/****************************************************************************/

OPEN Get_ID_Cursor 

    FETCH NEXT FROM Get_ID_Cursor  into @Emp_ID
				
	WHILE (@@FETCH_STATUS = 0)

    		BEGIN
			set @Latest_A_Gen=(select dbo.GetAutoGen_No(@Pay_Type))
			
			if @Pay_type!=21 -----Bonus etc.
				set @Bns_OT_Amount=(Select dbo.GetBonusAmount(@Emp_Id,@Pay_Type))
			if @Pay_type =21 -----Overtime
				set @Bns_OT_Amount=(Select dbo.GetOTAmount(@Emp_Id,@Pay_Month,@Pay_year))

			insert into payroll_main(Emp_Id,A_Gen,Pay_Type,pay_Month,pay_Year,U_id,Dt)
			values(@Emp_Id,@Latest_A_Gen,@Pay_Type,@pay_Month,@pay_Year,@U_id,Getdate())

			--- head code becomes pay_type  and vise-varsa

			Insert into Payroll_Sub(A_Gen,Head_Code,Amount)  
			Values (@Latest_A_Gen,@Pay_Type,@Bns_OT_Amount) 	
	
			set @Count=@Count+1
	
			FETCH NEXT FROM Get_ID_Cursor  into @Emp_ID
			
		END

CLOSE Get_ID_Cursor 
DEALLOCATE Get_ID_Cursor 


if @Count=0 set @Message='Specified job has done previously !'
else

set @Message='Specified job has done successfully for '+ rtrim(convert(char(5),@Count))+' employee(s)!'

select Message=@Message

set nocount Off







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE  PROCEDURE Bonus_Preparation_Delete
	@Emp_ID varchar(10),
	@PayMonth  varchar(9),
	@PayYear char(4)
	
	
as 

If Exists (Select * From BonusPreparation 
    Where  emp_id=@Emp_Id and PayYear=@PayYear and PayMonth=@PayMonth )
	begin		
		Delete from  BonusPreparation 
			Where  emp_id=@Emp_Id and PayYear=@PayYear and PayMonth=@PayMonth 
	end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



----select * from BonusPreparation
---- exec Bonus_Preparation_Save 'SMF-01','September','2005',3500,'16/10/2005','dsl'
--exec Bonus_Preparation_Save 'SMF-01','September','2005',3500,'dsl','10/16/2005'

CREATE PROCEDURE Bonus_Preparation_Save
	@Emp_ID varchar(10),
	@PayMonth  varchar(9),
	@PayYear char(4),
	@Amount money,
	@EntryBy varchar(5),
	@EntryDate datetime
	
as 

If Exists (Select * From BonusPreparation 
    Where  emp_id=@Emp_Id and PayMonth=@PayMonth and PayYear=@PayYear )
	begin		
		Update BonusPreparation  Set 
			Amount=@Amount ,
			EntryDate=@EntryDate ,
			EntryBy=@EntryBy 
			 Where  (emp_id=@Emp_Id and PayMonth=@PayMonth and PayYear=@PayYear )
	end
Else

Insert Into BonusPreparation (
	Emp_ID ,
	PayMonth ,
	PayYear ,
	Amount ,
	EntryBy,
	EntryDate
	 
	
) Values (
	@Emp_ID ,
	@PayMonth ,
	@PayYear,
	@Amount,
	@EntryBy,
	@EntryDate 
	
)








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE procedure BookAuthor(
	@BookCode	VARCHAR(4),
	@AuthorName	VARCHAR(80))
AS
SET NOCOUNT ON

DECLARE @AuthorId VARCHAR (5),
		@Msg VARCHAR (100)

IF NOT EXISTS (SELECT * FROM LibraryBookInfo WHERE     (BookCode = @BookCode))
	BEGIN	
		SET @Msg='Invalid Book Code!'
		GOTO Errmsg
	END

SELECT @AuthorId = AuthorCode FROM AuthorInfo WHERE (AuthorName = @AuthorName)

IF NOT EXISTS (SELECT * FROM LibraryBookAuthor WHERE (BookCode = @BookCode) AND (AuthorId = @AuthorId))
	INSERT INTO LibraryBookAuthor
        (BookCode, AuthorId)
		VALUES (@BookCode,@AuthorId)
ELSE
	UPDATE  LibraryBookAuthor SET              
			AuthorId = @AuthorId WHERE (BookCode = @BookCode) AND (AuthorId = @AuthorId)

RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO











--- Exec BookDisretInformation 'Stu-00000000001','','00002','00001','N','25 Sep 2005','Lia'

CREATE        procedure BookDisretInformation
(
	
	@StudentID			varchar(15),
	@Shift		 		varchar(30),
	@ClassId			varchar(15),
	@SubjectId			varchar(10),
	@BookRecievedback		varchar(1),
	@EntryDate			DateTime,
	@EntryBY			varchar(50)
	
)

as



if exists (select * from BookDistributionandReturnInfo where  StudentID= @StudentID and subjectid=@subjectid and classid=@ClassId and shift=@Shift) 		


Update BookDistributionandReturnInfo set  	


	--SubjectId		=	@SubjectId			,
	BookRecievedback	=	@BookRecievedback		,
	EntryDate		=	@EntryDate			,
	EntryBY			=	@EntryBY			


where StudentID= @StudentID and subjectid=@subjectid and classid=@ClassId and shift=@Shift
else


 insert into BookDistributionandReturnInfo 

(

	StudentID			,
	Shift		 		,
	ClassId				,
	SubjectId			,
	BookRecievedback		,
	EntryDate			,
	EntryBY				
	--DeliveryApproved		,
	--DeliveryApprovedby		,
	--DeliveryApproveddate		,
	--ReturnApprovedBy		
			
		
	
)
values
(
	@StudentID			,
	@Shift		 		,
	@ClassId			,
	@SubjectId			,
	@BookRecievedback		,
	@EntryDate			,
	@EntryBY			
	--@DeliveryApproved		,
	--@DeliveryApprovedby		,
	--@DeliveryApproveddate		,
	--@ReturnApprovedBy		
				
)













GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO











--- Exec BookDisretInformation1 'Stu-00000000001','Morning -  Shift','00002','00001','Y','25 Sep 2005','Lia'
CREATE        procedure BookDisretInformation1
(
	
	@StudentID			varchar(15),
	@Shift		 		varchar(30),
	@ClassId			varchar(15),
	@SubjectId			varchar(10),
	--@BookRecievedback		varchar(1),
	--@EntryDate			DateTime,
	--@EntryBY			varchar(50)
	@DeliveryApproved		varchar(1),
	--@DeliveryApprovedby		varchar(10),
	@DeliveryApproveddate		datetime,
	@DeliveryApprovedby		varchar(10)
	
)

as


if exists(select * from BookDistributionandReturnInfo where  StudentID= @StudentID and classId=@ClassId and Shift=@shift AND SubjectId = @SubjectId  ) 		

--Delete from BookDistributionandReturnInfo where StudentID= @StudentID AND SubjectId = @SubjectId	
Update BookDistributionandReturnInfo set  	
	
	SubjectId		=	@SubjectId	,
	DeliveryApproved	=	@DeliveryApproved,
	DeliveryApprovedby	=	@DeliveryApprovedby,
	DeliveryApproveddate	=	@DeliveryApproveddate			
			

where StudentID= @StudentID AND SubjectId = @SubjectId	and classId=@ClassId and Shift=@shift 












GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO











--- Exec BookDisretInformation 'Stu-00000000001','','00002','00001','N','25 Sep 2005','Lia'

CREATE       procedure BookDisretInformation2
(
	
	@StudentID			varchar(15),
	@Shift		 		varchar(30),
	@ClassId			varchar(15),
	@SubjectId			varchar(10),
	@BookRecievedback		varchar(1),
	@EntryDate			DateTime,
	@EntryBY			varchar(50)
	
)

as



if exists (select * from BookDistributionandReturnInfo where  StudentID= @StudentID and subjectid=@subjectid and classId=@ClassId and Shift=@shift) 		


Update BookDistributionandReturnInfo set  	


	--SubjectId		=	@SubjectId			,
	BookRecievedback	=	@BookRecievedback		,
	EntryDate		=	@EntryDate			,
	EntryBY			=	@EntryBY			


where StudentID= @StudentID and subjectid=@subjectid and ClassId=@classId and Shift=@Shift
else


 insert into BookDistributionandReturnInfo 

(

	StudentID			,
	Shift		 		,
	ClassId				,
	SubjectId			,
	BookRecievedback		,
	EntryDate			,
	EntryBY				
	--DeliveryApproved		,
	--DeliveryApprovedby		,
	--DeliveryApproveddate		,
	--ReturnApprovedBy		
			
		
	
)
values
(
	@StudentID			,
	@Shift		 		,
	@ClassId			,
	@SubjectId			,
	@BookRecievedback		,
	@EntryDate			,
	@EntryBY			
	--@DeliveryApproved		,
	--@DeliveryApprovedby		,
	--@DeliveryApproveddate		,
	--@ReturnApprovedBy		
				
)













GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO











--- Exec BookDisretInformation 'Stu-00000000001','','00002','00001','N','25 Sep 2005','Lia'
CREATE        procedure BookDisretInformation3
(
	
	@StudentID			varchar(15),
	@Shift		 		varchar(30),
	@ClassId			varchar(15),
	@SubjectId			varchar(10),
	@ReturnApproved			varchar(1),
	@ReturnApprovedby		varchar(10),
	@ReturnApproveddate		datetime
	
	
)

as


if exists(select * from BookDistributionandReturnInfo where  StudentID= @StudentID AND SubjectId = @SubjectId and classid=@classId and Shift=@shift  ) 		

Update BookDistributionandReturnInfo set  	

	SubjectId		=	@SubjectId	,
	ReturnApproved		=	@ReturnApproved	,
	ReturnApprovedby	=	@ReturnApprovedby,	
	ReturnApproveddate	=	@ReturnApproveddate		

where StudentID= @StudentID AND SubjectId = @SubjectId	and classid=@classId and Shift=@shift












GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE    procedure BookInfo(
	@BookCode VARCHAR(4),
	@BookName VARCHAR (60),
	@Sub VARCHAR (60),
	@Pub VARCHAR (60),
	@Remarks VARCHAR (100),
	@User VARCHAR (10))
AS
SET NOCOUNT ON
DECLARE @SubCode VARCHAR (5),
		@PubCode VARCHAR (6),
		@Msg VARCHAR (100)

SELECT @SubCode = SubSubjectCode FROM LibrarySubjectInfo WHERE     (SubSubjectName = @Sub)
SELECT @PubCode = PubCode FROM PublisherInfo WHERE     (PubPublisherName = @Pub)
IF @SubCode IS NULL
	BEGIN	
		SET @Msg='Invalid Subject Name!'
		GOTO Errmsg
	END
IF @PubCode IS NULL
	BEGIN	
		SET @Msg='Invalid Publisher Name!'
		GOTO Errmsg
	END

IF NOT EXISTS (SELECT * FROM LibraryBookInfo WHERE     (BookCode = @BookCode))
	INSERT INTO LibraryBookInfo
            (BookCode, BookName, SubCode, PubCode, Remarks, EntryBy, EntryDate)
			VALUES     (@BookCode,@BookName,@SubCode,@PubCode,@Remarks,@User,GETDATE())
ELSE
	UPDATE  LibraryBookInfo SET              
			BookName = @BookName, SubCode = @SubCode, PubCode = @PubCode, Remarks = @Remarks
			WHERE     (BookCode = @BookCode)

RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE   procedure BookIssueSave(
	@StudentId	VARCHAR(15),
	@BookId	VARCHAR(15),
	@IssueDate DATETIME,
	@ExpIssueDate DATETIME,
	@ReqNo VARCHAR (7),
	@ReqDate DATETIME,
	@MaxBook INT,
	@User VARCHAR (10))
AS
SET NOCOUNT ON
DECLARE @Msg VARCHAR (100),
		@BookCode VARCHAR (5),
		@TotBookUse INT

IF LEN(@ReqNo) = 0 	SET @ReqNo = NULL
IF @ReqDate = '01/01/1900' SET @ReqDate = NULL
SELECT @BookCode = BookCode FROM LibraryBookList WHERE (BookId = @BookId)
IF EXISTS (SELECT * FROM BookIssueRefund WHERE (StudentId = @StudentId) AND (BookCode = @BookCode) AND (ActualReturnDate IS NULL))
	BEGIN	
		SET @Msg = 'This book already using by the studend!'
		GOTO Errmsg
	END
IF NOT EXISTS (SELECT * FROM LibraryBookList WHERE     (BookId = @BookId) AND (Demage = 0))
	BEGIN	
		SET @Msg = 'Invalid Book Id!'
		GOTO Errmsg
	END
IF EXISTS (SELECT * FROM BookIssueRefund WHERE     (BookId = @BookId) AND (ActualReturnDate IS NULL))
	BEGIN	
		SET @Msg = 'This book already using by another studend!'
		GOTO Errmsg
	END

SELECT @TotBookUse = ISNULL(COUNT(BookId),0) FROM BookIssueRefund WHERE (ActualReturnDate IS NULL) AND (StudentId = @StudentId)
IF @MaxBook = @TotBookUse
	BEGIN	
		SET @Msg = 'Book Issue Limit completed !'
		GOTO Errmsg
	END

INSERT INTO BookIssueRefund
        (StudentId, BookId, BookCode, IssueDate, ExpReturnDate, ReqNo, ReqDate, Remarks, EntryBy, EntryDate)
		VALUES (@StudentId,@BookId,@BookCode,@IssueDate,@ExpIssueDate,@ReqNo,@ReqDate,NULL,@User,GETDATE())

IF @ReqNo IS NOT NULL
	UPDATE RequisetionSub SET Status = 'Issued'  WHERE (RequisitionId = @ReqNo) AND (BookCode = @BookCode)


RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE        procedure BookList1
(
	
	@ClassID			varchar(5),
	@Eyear	 			Int,
	@SubjectId			varchar(5),
	@Book				varchar(50),
	@Writter			varchar(100),	
        @publisher 		        varchar(300) ,
	@EntryBY			varchar(50),
	@EntryDate			DateTime
            
	
)

as



if exists (select * from Booklist where  Classid= @ClassID and subjectid=@subjectid and book=@Book and Eyear=@Eyear) 		


Update Booklist set  	


	--SubjectId		=	@SubjectId			,
	Book			=	@Book				,
	Writter			=	@Writter			,
	EntryDate		=	@EntryDate			,
	EntryBY			=	@EntryBY			,
        publisher              =       @publisher   


where ClassID= @ClassID and subjectid=@subjectid and Book=@Book and Eyear=@Eyear
else


 insert into Booklist

(

	ClassID				,
	Eyear	 			,
	SubjectId			,
	Book				,
	Writter				,	
	EntryBY				,
	EntryDate			,
        publisher
	
			
		
	
)
values
(
	@ClassID			,
	@Eyear	 			,
	@SubjectId			,
	@Book				,
	@Writter			,	
	@EntryBY			,
	@EntryDate			,
        @publisher 
	
)













GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE    procedure BookPurchase(
	@PurId	VARCHAR(7),
	@PurDate DATETIME,
	@Remarks VARCHAR(70),
	@User VARCHAR (10))
AS
SET NOCOUNT ON

DECLARE @Msg VARCHAR (100)
IF EXISTS (SELECT * FROM BookIssueRefund INNER JOIN LibraryBookList ON BookIssueRefund.BookId = LibraryBookList.BookId WHERE     (LibraryBookList.PurchaseId = @PurId))
	BEGIN
		SET @Msg='Update Not Allow!'
		GOTO Errmsg
	END

IF NOT EXISTS (SELECT * FROM LibraryBook WHERE     (PurchaseId = @PurId))
	INSERT INTO LibraryBook
        (PurchaseId, PurchaseDate, Remarks, EntryDate, EntryBy)
		VALUES     (@PurId,@PurDate,@Remarks,GETDATE(),@User)
ELSE
	UPDATE  LibraryBook SET              
			PurchaseDate = @PurDate, Remarks = @Remarks
			WHERE     (PurchaseId = @PurId)

RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE     procedure BookPurchaseSub(
	@PurId VARCHAR (7),
	@BookName VARCHAR(50),
	@URate MONEY,
	@FineAmt MONEY,
	@Qty INT,
	@TrackId INT)
AS
SET NOCOUNT ON
DECLARE @Msg VARCHAR (100),
		@BookCode VARCHAR (4)

IF NOT EXISTS (SELECT * FROM LibraryBook WHERE     (PurchaseId = @PurId))
	BEGIN
		SET @Msg='Invalid Purchase Id!'
		GOTO Errmsg
	END
IF EXISTS (SELECT * FROM BookIssueRefund INNER JOIN LibraryBookList ON BookIssueRefund.BookId = LibraryBookList.BookId WHERE     (LibraryBookList.PurchaseId = @PurId))
	BEGIN
		SET @Msg='Update Not Allow!'
		GOTO Errmsg
	END
SELECT @BookCode = BookCode FROM LibraryBookInfo WHERE     (BookName = @BookName)
IF NOT EXISTS (SELECT * FROM LibraryBookSub WHERE     (TrackId = @TrackId))
	BEGIN
		IF @Qty = 0
			BEGIN
				SET @Msg = 'Quantity Required!'
				GOTO Errmsg
			END
		INSERT INTO LibraryBookSub
	        (PurchaseId, BookCode, UnitRate, Qty, FineAmt)
			VALUES (@PurId,@BookCode,@URate,@Qty,@FineAmt)
	END
ELSE
	BEGIN
		IF @Qty = 0
			BEGIN
				DELETE FROM LibraryBookSub WHERE     (TrackId = @TrackId)
				DELETE FROM LibraryBookList WHERE     (PurchaseId = @PurId) AND (BookCode = @BookCode)
			END
		ELSE
			UPDATE  LibraryBookSub SET              
				BookCode = @BookCode, UnitRate = @URate, Qty = @Qty, FineAmt = @FineAmt
				WHERE     (TrackId = @TrackId)
	END

-- book list
DECLARE @Sub VARCHAR (5),
		@Auth VARCHAR (5),
		@BookId VARCHAR (15),
		@MaxBookId INT
		
SELECT @Sub = SubCode FROM LibraryBookInfo WHERE     (BookCode = @BookCode)
SELECT @Auth = AuthorId FROM LibraryBookAuthor WHERE (Trackid = (SELECT MIN(Trackid)
FROM LibraryBookAuthor WHERE (BookCode = @BookCode))) AND (BookCode = @BookCode)

DELETE FROM LibraryBookList WHERE (PurchaseId = @PurId) AND (BookCode = @BookCode)
SELECT @MaxBookId = ISNULL(MAX(CAST(SUBSTRING(BookId, 10, 3) AS INT)),0) FROM LibraryBookList WHERE     (SUBSTRING(BookId, 1, 5) = @Sub) AND (SUBSTRING(BookId, 6, 4) = @BookCode)
SET @MaxBookId = @MaxBookId + @Qty
WHILE @Qty > 0
	BEGIN
		IF LEN(@MaxBookId) = 1
			SET @BookId = @Sub + @BookCode + '00' + CAST(@MaxBookId AS VARCHAR)
		ELSE IF LEN(@MaxBookId) = 2
			SET @BookId = @Sub + @BookCode + '0' + CAST(@MaxBookId AS VARCHAR)
		ELSE IF LEN(@MaxBookId) = 3
			SET @BookId = @Sub + @BookCode + CAST(@MaxBookId AS VARCHAR)
		INSERT INTO LibraryBookList
        (PurchaseId, BookCode, BookId, UnitRate, FineAmt, Demage)
		VALUES (@PurId,@BookCode,@BookId,@URate,@FineAmt,0)
		SET @Qty = @Qty - 1
		SET @MaxBookId = @MaxBookId - 1
	END



RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE   procedure BookRequisition(
	@ReqId	VARCHAR(7),
	@ReqDate DATETIME,
	@StudentId VARCHAR(15),
	@Remarks VARCHAR (70),
	@User VARCHAR (10))
AS
SET NOCOUNT ON

DECLARE @Msg VARCHAR (100)
IF EXISTS (SELECT * FROM BookIssueRefund WHERE     (ReqNo = @ReqId))
	BEGIN	
		SET @Msg='Update not allowed!'
		GOTO Errmsg
	END

IF NOT EXISTS (SELECT * FROM RequisitionInfo WHERE     (RequisitionId = @ReqId))
	INSERT INTO RequisitionInfo
        (RequisitionId, RequisitionDate, StudentId, Remarks, EntryBy, EntryDate)
		VALUES     (@ReqId,@ReqDate,@StudentId,@Remarks,@User,GETDATE())
ELSE
	UPDATE  RequisitionInfo SET               
			RequisitionDate = @ReqDate, StudentId = @StudentId, Remarks = @Remarks
			WHERE     (RequisitionId = @ReqId)

RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE   procedure BookRequisitionSub(
	@ReqId	VARCHAR(7),
	@Sub VARCHAR(50),
	@Book VARCHAR(50),
	@Status VARCHAR (10),
	@TrackId VARCHAR (10))
AS
SET NOCOUNT ON

DECLARE @Msg VARCHAR (100),
		@SubCode VARCHAR (5),
		@BookCode VARCHAR (4)
IF NOT EXISTS (SELECT * FROM RequisitionInfo WHERE (RequisitionId = @ReqId))
	BEGIN
		SET @Msg='Invalid Requisition Id!'
		GOTO Errmsg
	END
IF EXISTS (SELECT * FROM BookIssueRefund WHERE     (ReqNo = @ReqId))
	BEGIN	
		SET @Msg='Update not allowed!'
		GOTO Errmsg
	END

SELECT @SubCode = SubSubjectCode FROM LibrarySubjectInfo WHERE (SubSubjectName = @Sub)
SELECT @BookCode = BookCode FROM LibraryBookInfo WHERE (BookName = @Book)

--IF NOT EXISTS (SELECT * FROM RequisetionSub WHERE (TrackId = @TrackId))
IF NOT EXISTS (SELECT * FROM RequisetionSub  WHERE (RequisitionId = @ReqId) AND (BookCode  = @BookCode))
	INSERT INTO RequisetionSub
            (RequisitionId, SubCode, BookCode, Status)
			VALUES (@ReqId,@SubCode,@BookCode,@Status)
ELSE
	UPDATE  RequisetionSub SET              
			SubCode = @SubCode, BookCode = @BookCode, Status = @Status
			WHERE (RequisitionId = @ReqId) AND (BookCode  = @BookCode)
RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE   procedure BookReturnSave(
	@StudentId	VARCHAR(15),
	@BookId	VARCHAR(15),
	@ReturnDate DATETIME,
	@DelayDay INT,
	@FineAmt MONEY)
AS
SET NOCOUNT ON
DECLARE @Msg VARCHAR (100),
		@IssueDate DATETIME

SELECT @IssueDate = IssueDate FROM BookIssueRefund WHERE     (ActualReturnDate IS NULL) AND (BookId = @BookId) AND (StudentId = @StudentId)
IF @ReturnDate < @IssueDate
	BEGIN	
		SET @Msg = 'Invalid Return Date!'
		GOTO Errmsg
	END


UPDATE  BookIssueRefund SET              
		ActualReturnDate = @ReturnDate, DelayDay = @DelayDay, FineAmt = @FineAmt
		WHERE     (StudentId = @StudentId) AND (BookId = @BookId) AND (ActualReturnDate IS NULL)

RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Branch_info_I_U_D    Script Date: 3/7/02 10:43:28 AM ******/
/****** Object:  Stored Procedure dbo.Branch_info_I_U_D    Script Date: 10/28/01 4:34:21 PM ******/
/****** Object:  Stored Procedure dbo.Branch_info_I_U_D    Script Date: 10/22/01 6:24:21 PM ******/
/****** Object:  Stored Procedure dbo.Branch_info_I_U_D    Script Date: 9/17/00 4:33:56 PM ******/
/****** Object:  Stored Procedure dbo.Branch_info_I_U_D    Script Date: 9/4/01 6:30:43 PM ******/
CREATE PROCEDURE Branch_info_I_U_D
@opr varchar (1),
@Code varchar (3),
@Title varchar (25),
@Prev_Title varchar(25),
@Address varchar (100),
@Telephone varchar(30),
@U_id varchar (10)
AS
if @opr='I'
begin
insert into Branch_info (
Title,
Code,
Address,
Telephone,
U_id
) 
values (
@Title,
@Code,
@address,
@Telephone,
@U_id
)
end
                    
if @opr='U'
begin
update Branch_info set 
Title=@Title,
Code=@Code,
Address=@Address,
Telephone=@Telephone,
U_id=@U_id
where Title=@Prev_Title
end
    		
if @opr='D'
begin
delete from Branch_info where Title=@Title
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



-----------------------------------------
create Procedure Change_PW 
@User varchar (45)
,@Old_Pass varchar (40)
,@New_Pass varchar (40)

AS

Declare @Message varchar(50)
,@Prev_Pass varchar(40)
,@Can varchar(1)


set @Can=(select Cancel from Soft_Pass where u_id=@User)

if @Can='' or @Can is null
set @Message='Invalid user or password!'

if @Can='1'
set @Message='You can not change your password!'

if @Can='0'

begin
	set @Prev_Pass=(select user_pass from Soft_Pass where u_id=@User) 
	
	if @Prev_Pass=@Old_Pass
	
		begin
			Update Soft_Pass set user_pass=@New_Pass
			where u_id=@User and cancel='0'
	
			set @Message='Password changed successfully!'
	
		end

	else
		
			set @Message='Invalid passwordly!'		

end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Check_Current_time    Script Date: 3/7/02 10:43:28 AM ******/
/****** Object:  Stored Procedure dbo.Check_Current_time    Script Date: 10/28/01 4:34:21 PM ******/
/****** Object:  Stored Procedure dbo.Check_Current_time    Script Date: 10/22/01 6:24:22 PM ******/
/****** Object:  Stored Procedure dbo.Check_Current_time    Script Date: 9/17/00 4:33:56 PM ******/
CREATE Procedure Check_Current_time
as
select Boot_Up,Latest_Entry=Latest_time,Dt from time_keeper  where 
datepart(day,Dt)=datepart(day,getdate())and
datepart(month,Dt)=datepart(month,getdate())and
datepart(year,Dt)=datepart(year,getdate())




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE    Procedure Check_N_Store_Attn_Backup

AS

set nocount on

Declare @Emp_Id varchar(10)
,@Emp_LogIn datetime
,@Emp_logOut datetime
,@Start_dt datetime
,@End_dt datetime
,@Message varchar(150)
,@Entry_Count int

Set @Entry_Count=0
---------------------------------------------------------------
declare Attn_Cursor cursor for
select Emp_Id, Emp_logIn,Emp_logOut from emp_Att_info_Audit
---------------------------------------------------------------
declare Leave_Cursor cursor for
select Emp_Id,Start_dt,End_dt from emp_Leave_info_Audit
---------------------------------------------------------------


	OPEN Attn_Cursor

		Fetch next from Attn_Cursor into @Emp_Id,@Emp_LogIn,@Emp_logOut
	while @@fetch_Status = 0 
	
	Begin
			if not exists (select * from emp_Att_info 
			where Emp_Id=@Emp_Id and  Emp_logIn=@Emp_logIn and Emp_logout=@Emp_logOut)
	
			BEGIN
				Insert Into Emp_Att_Info
				(Emp_id,Emp_login,Emp_logout,In_Status,In_Notes,Out_Status
					,Out_Notes,Entry_Dt,U_id)	

				select Emp_id,Emp_login,Emp_logout,In_Status,In_Notes,Out_Status
					,Out_Notes,Entry_Dt,U_id from Emp_Att_Info_Audit 
					where Emp_Id=@Emp_Id and  Emp_logIn=@Emp_logIn and Emp_logout=@Emp_logOut
				
					---SET @Entry_Count=@Entry_Count+1				
			END

		Fetch next from Attn_Cursor into @Emp_Id,@Emp_LogIn,@Emp_logout
	END


close Attn_Cursor
Deallocate Attn_Cursor

--------------------------------------------Leave--------------------------------------------			

	OPEN Leave_Cursor

		Fetch next from Leave_Cursor into @Emp_Id,@Start_dt,@End_dt
		
		while @@fetch_Status = 0 

	Begin
		if not exists (select * from emp_Leave_info 
			where Emp_Id=@Emp_Id and  Start_dt=@Start_dt and @End_dt=@End_dt)

		BEGIN
			Insert Into emp_Leave_info
			select * from Emp_Leave_Info_Audit where 
					Emp_Id=@Emp_Id and  
					Start_dt=@Start_dt and 
					@End_dt=@End_dt

			--SET @Entry_Count=@Entry_Count+1				
		END


		Fetch next from Leave_Cursor into @Emp_Id,@Start_dt,@End_dt
	END

	delete from Emp_Leave_info_Audit

	close Leave_Cursor
	Deallocate Leave_Cursor	

--SET @Message= convert(char(6),@Entry_Count)+ ' '+ 'records successfully inserted !'
SET @Message='Records successfully inserted !'
-------------------------------------------------------------------------------------------------------

delete from Emp_Att_Info_Audit
delete from Emp_Leave_Info_Audit

select Message=@Message

Set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






CREATE    procedure ClassInformation
(
	@ClassID	varchar(5),
	@ClassName	Varchar(80),
	@ShiftName	varchar(50),
	@StartTime	datetime,
	@endTime	datetime
	
	
)

AS

if exists (select * from ClassInfo where  ClassID = @ClassID)
Update ClassInfo set

	ClassID		=	@ClassID,	
	ClassName	=	@ClassName,
	Shiftname	=	@Shiftname,
	StartTime	=	@StartTime,
	EndTime		=	@endTime	

	where ClassID = @ClassID 

else
 insert into ClassInfo

(
	ClassID,
	ClassName,
	Shiftname,
	StartTime,
	EndTime	
	
)
values
(
 	@ClassID,	
	@ClassName,
	@Shiftname,
	@StartTime,
	@endTime
)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

---exec ClassRoutine1 1,'00001','M','00001','Saturday','00002','01:00:00 AM','02:00:00 AM' ,'BS-002','00001','01/01/2007','2007',2,3

CREATE procedure ClassRoutine1
                 @mode int,
                 @class varchar(20),
                 @shift varchar(20),
                 @section varchar(20),
                 @day  varchar(20),
                 @subject varchar(20),
                 @sttime datetime,
                 @edtime datetime,
                 @teacher varchar(20),
                 @entryby varchar(20),
                 @effectivedate datetime,
                 @acayr  varchar(20),
                 @serial integer,
                 @trackid integer   
                            



as
  begin
    if @mode=1 
       begin
     
     ---  SELECT  @serial=isnull(max(SerialNo),0)+1 
         --  FROM ClassRoutine
     --- where classid=@class  and  SectionId=@section and shift=@shift and academic_yr=@acayr


       insert into ClassRoutine values(@class
                      ,@shift,@section,@day,@subject,@sttime,@edtime,
                @teacher,@entryby,@effectivedate,@acayr,@serial,@trackid)
               
  end 
    if @mode=2 
       begin
     
       SELECT  @trackid=isnull(max(trackid),0)+1 
           FROM ClassRoutine
      where classid=@class  and  SectionId=@section and shift=@shift and academic_yr=@acayr and SerialNo=@serial

        insert into ClassRoutine values(@class,@shift,@section,@day,@subject,@sttime,@edtime,
                   @teacher,@entryby,@effectivedate,@acayr, @serial,@trackid)
  end

if @mode=3 
           begin
     
        delete from  ClassRoutine  where classid=@class  and  SectionId=@section and shift=@shift and academic_yr=@acayr and SerialNo=@serial and trackid=@trackid
        
end       

end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE     procedure Collec_master_Save
(           @mode	   varchar(1),
            @C_srl     integer,
            @seq_no    integer,
            @mon       varchar(4),
            @yr        varchar(5),
            @collec_date datetime,   
            @Remark    varchar(150),
			@EntryBy	varchar(10),
			@Entrydate	datetime,
           @student_id varchar(50),  
           @class_id varchar(20)
           
)
	AS

         declare  @u_id as varchar(20)
         declare  @class_code as varchar(20)
         
         declare  @fee_code as varchar(20)
         declare  @Act_amount as decimal
         declare  @Fine as decimal 
         declare  @Discount as decimal
         declare  @std_id as varchar(15)
        
         declare @loc_C_srl as integer
         declare @srl_no as integer 

    
 
 
       if @mode='s'
               begin
                    set @loc_c_srl=(select isnull(max(C_srl),0)+1 from Collec_master)
                    insert into Collec_master values(@loc_c_srl,@student_id,@class_id,@mon,@yr,@Remark,getdate(),@entryby,@collec_date,0) 
               end
              
/*
                      declare collec_cursor cursor for
                        select Srl_no,u_id,class_code,Fee_code,Act_amount,Fine,Discount,std_id
                               from temp_collect
                      where seq_no=@seq_no


		    			set @loc_c_srl=(select isnull(max(C_srl),0)+1 from Collec_master)

                   Open collec_cursor

 					   Fetch Next From collec_cursor into @Srl_no,@u_id,@class_code,@Fee_code,@Act_amount,@Fine,@Discount,@std_id

                       insert into Collec_master(C_srl,Std_id,class_id,Mon,Yr,Remark,Entry_by,Entry_date,Collec_date) 
				                            values(@loc_c_srl,@Std_id,@class_code,@Mon,@Yr,@Remark,@EntryBy,@Entrydate,@Collec_date)	

                        set @loc_c_srl=(select isnull(max(C_srl),0) from Collec_master)

                     ---  insert into Collec_details(C_Srl,serial_no,Fee_code,Act_Amount,Discount,Fine,Collec_date,Entry_by,Entry_date)
                                ---       values( @loc_c_srl,@Srl_no,@Fee_code,@Act_Amount,@Discount,@Fine,@Collec_date,@Entryby,@Entrydate) 
		            

                      While @@Fetch_Status = 0
                           begin
			                
         				    insert into Collec_details(C_Srl,serial_no,Fee_code,Act_Amount,Discount,Fine,Entry_by,Entry_date)
                                       values( @loc_c_srl,@Srl_no,@Fee_code,@Act_Amount,@Discount,@Fine,@Entryby,@Entrydate) 
		                                                         

                          Fetch Next From collec_cursor into @Srl_no,@u_id,@class_code,@Fee_code,@Act_amount,@Fine,@Discount,@std_id
                   end 
               End

              delete from temp_collect
           where seq_no=@seq_no
       
	        Close collec_cursor
	       Deallocate collec_cursor
      



       if @mode='u'
               begin
		    			if  exists (select C_srl from Collec_master where C_srl=@C_srl)
		
		                 update Collec_master set 
                                             Std_id=@Std_id,
											 Class_id=@class_id, 
                                             Remark=@Remark,
		                                     Entry_by=@EntryBy,
		                                     Entry_date=@Entrydate

                        where C_srl=@C_srl
                        
                       if  exists (select C_srl from Collec_details where C_srl=@C_srl and fee_code=@fee_code)
                            begin
		                         update Collec_details set 
		                                             Fee_code=@Fee_code,
													 amount=@amount, 
		                                             Entry_by=@EntryBy,
				                                     Entry_date=@Entrydate
		
		                        where C_srl=@C_srl and fee_code=@fee_code 
		                   end  



   
                                               
                  end 
 

       if @mode='d'
               begin
		    			if  exists (select C_srl from Collec_master where C_srl=@C_srl)
		
		                 delete from Collec_master where C_srl=@C_srl  
                         delete from Collec_details where C_srl=@C_srl    
                end 


       if @mode='p'
               begin
		    			if  exists (select C_srl from Collec_details where C_srl=@C_srl and fee_code=@fee_code)
		
		                 delete from Collec_details where C_srl=@C_srl and fee_code=@fee_code  
                end 
 
 

*/









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE     procedure Collec_sub_save
(           @mode	   varchar(1),
            @seq_no    integer,
            @C_srl     integer,
            @EntryBy	varchar(10), 
            @student_id varchar(50),  
            @Fee_code      varchar(12),
            @Act_Amount float,
            @Discount   float,
            @Fine      float
			
           
)
	AS

         declare  @u_id as varchar(20)
         declare  @class_code as varchar(20)
         
         
         
        
         declare @loc_C_srl as integer
         declare @srl_no as integer 

    
 
 
       if @mode='s'
               begin
                    set @loc_c_srl=(select isnull(max(serial_no),0)+1 from Collec_details where C_Srl=@C_srl)
                    insert into Collec_details(C_Srl,serial_no,Fee_code,Act_Amount,Discount,Fine,Entry_by,Entry_date)
                                    values( @C_srl,@loc_c_srl,@Fee_code,@Act_Amount,@Discount,@Fine,@Entryby,getdate()) 
		       
               end
              
/*
                      declare collec_cursor cursor for
                        select Srl_no,u_id,class_code,Fee_code,Act_amount,Fine,Discount,std_id
                               from temp_collect
                      where seq_no=@seq_no


		    			set @loc_c_srl=(select isnull(max(C_srl),0)+1 from Collec_master)

                   Open collec_cursor

 					   Fetch Next From collec_cursor into @Srl_no,@u_id,@class_code,@Fee_code,@Act_amount,@Fine,@Discount,@std_id

                       insert into Collec_master(C_srl,Std_id,class_id,Mon,Yr,Remark,Entry_by,Entry_date,Collec_date) 
				                            values(@loc_c_srl,@Std_id,@class_code,@Mon,@Yr,@Remark,@EntryBy,@Entrydate,@Collec_date)	

                        set @loc_c_srl=(select isnull(max(C_srl),0) from Collec_master)

                     ---       

                      While @@Fetch_Status = 0
                           begin
			                
         				    insert into Collec_details(C_Srl,serial_no,Fee_code,Act_Amount,Discount,Fine,Entry_by,Entry_date)
                                       values( @loc_c_srl,@Srl_no,@Fee_code,@Act_Amount,@Discount,@Fine,@Entryby,@Entrydate) 
		                                                         

                          Fetch Next From collec_cursor into @Srl_no,@u_id,@class_code,@Fee_code,@Act_amount,@Fine,@Discount,@std_id
                   end 
               End

              delete from temp_collect
           where seq_no=@seq_no
       
	        Close collec_cursor
	       Deallocate collec_cursor
      



       if @mode='u'
               begin
		    			if  exists (select C_srl from Collec_master where C_srl=@C_srl)
		
		                 update Collec_master set 
                                             Std_id=@Std_id,
											 Class_id=@class_id, 
                                             Remark=@Remark,
		                                     Entry_by=@EntryBy,
		                                     Entry_date=@Entrydate

                        where C_srl=@C_srl
                        
                       if  exists (select C_srl from Collec_details where C_srl=@C_srl and fee_code=@fee_code)
                            begin
		                         update Collec_details set 
		                                             Fee_code=@Fee_code,
													 amount=@amount, 
		                                             Entry_by=@EntryBy,
				                                     Entry_date=@Entrydate
		
		                        where C_srl=@C_srl and fee_code=@fee_code 
		                   end  



   
                                               
                  end 
 

       if @mode='d'
               begin
		    			if  exists (select C_srl from Collec_master where C_srl=@C_srl)
		
		                 delete from Collec_master where C_srl=@C_srl  
                         delete from Collec_details where C_srl=@C_srl    
                end 


       if @mode='p'
               begin
		    			if  exists (select C_srl from Collec_details where C_srl=@C_srl and fee_code=@fee_code)
		
		                 delete from Collec_details where C_srl=@C_srl and fee_code=@fee_code  
                end 
 
 

*/









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Comp_Det_Info_I_U_D    Script Date: 3/7/02 10:43:28 AM ******/
/****** Object:  Stored Procedure dbo.Comp_Det_Info_I_U_D    Script Date: 10/28/01 4:34:21 PM ******/
/****** Object:  Stored Procedure dbo.Comp_Det_Info_I_U_D    Script Date: 10/22/01 6:24:22 PM ******/
/****** Object:  Stored Procedure dbo.Comp_Det_Info_I_U_D    Script Date: 9/17/00 4:33:56 PM ******/
/****** Object:  Stored Procedure dbo.Comp_Det_Info_I_U_D    Script Date: 9/4/01 6:30:43 PM ******/
/****** Object:  Stored Procedure dbo.Comp_Det_Info_I_U_D    Script Date: 6/20/01 12:00:32 PM ******/
/****** Object:  Stored Procedure dbo.Comp_Det_Info_I_U_D    Script Date: 5/14/01 9:49:53 AM ******/
/****** Object:  Stored Procedure dbo.Comp_Det_Info_I_U_D    Script Date: 1/1/99 3:01:39 AM ******/
/****** Object:  Stored Procedure dbo.Comp_Det_Info_I_U_D    Script Date: 5/9/01 8:13:06 AM ******/
/****** Object:  Stored Procedure dbo.Comp_Det_Info_I_U_D    Script Date: 5/8/01 4:09:45 AM ******/
/****** Object:  Stored Procedure dbo.Comp_Det_Info_I_U_D    Script Date: 4/19/01 2:00:10 AM ******/
CREATE PROCEDURE Comp_Det_Info_I_U_D
@opr as varchar(1),
@Name varchar (50),
@Type varchar (15),
@Address varchar (200),
@City varchar (20),
@Country varchar (20),
@Phone1 varchar (20),
@Phone2 varchar (20),
@Fax varchar (20),
@Email varchar (50),
@Notes varchar (500),
@Entry_Dt datetime,
@U_id varchar (10)
AS
if @opr='I'
begin
insert into Comp_Det_Info (Name,Type,Address,City,Country,Phone1,Phone2,
Fax,Email,Notes,Entry_Dt,U_id) values(@Name,@Type,@Address,@City,@Country,@Phone1,@Phone2,
@Fax,@Email,@Notes,@Entry_Dt,@U_id)
end
if @opr='U'
begin
update Comp_Det_Info set  Name=@Name,Type=@Type,Address=@Address,City=@City,
Country=@Country,Phone1=@Phone1,Phone2=@Phone2,Fax=@Fax,Email=@Email,
Notes=@Notes,Entry_Dt=@Entry_Dt,U_id=@U_id
end
if @opr='D'
begin
	delete from Comp_Det_Info
end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE    procedure Company_Name_Address

AS
declare @Delimiter varchar(1)

select [Name],Type,Address=(Address+ case City
				when  null then ''
				when  '' then ''
				else	','+City
				end

					+ case Country
					when  null then ''
					when  '' then ''
					else	','+Country
					end )

						,Phone='Phone :'+Phone1 + case Phone2
							when  null then ''
							when  '' then ''
							else	','+phone2							end

,Fax,Email='Email :'+Email from Comp_Det_Info







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Company_Policy_I_U_D    Script Date: 3/7/02 10:43:29 AM ******/
/****** Object:  Stored Procedure dbo.Company_Policy_I_U_D    Script Date: 10/28/01 4:34:21 PM ******/
/****** Object:  Stored Procedure dbo.Company_Policy_I_U_D    Script Date: 10/22/01 6:24:22 PM ******/
/****** Object:  Stored Procedure dbo.Company_Policy_I_U_D    Script Date: 9/17/00 4:33:56 PM ******/
/****** Object:  Stored Procedure dbo.Company_Policy_I_U_D    Script Date: 9/4/01 6:30:44 PM ******/
CREATE Procedure [Company_Policy_I_U_D] 
@opr varchar(1),
@Stored_Policy_Code varchar(10),
@Policy_code varchar (10) ,
@Policy_type  varchar (50) ,
@Policy_detail  varchar (500)  ,
@U_id  varchar (10)  
as
if @opr='I'
begin
	insert into Company_Policy (Policy_code,Policy_type,Policy_detail,U_id )
	values(@Policy_code,@Policy_type,@Policy_detail,@U_id )
end
if @opr='U'
begin
	update Company_Policy set
		Policy_type =@Policy_type ,
		Policy_detail =@Policy_detail,
		U_id=@U_id
	where Policy_code=@Policy_Code
end
If @opr='D'
begin
	delete from Company_Policy  where Policy_code=@Policy_Code	
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Company_Policy_View    Script Date: 3/7/02 10:43:29 AM ******/
/****** Object:  Stored Procedure dbo.Company_Policy_View    Script Date: 10/28/01 4:34:22 PM ******/
/****** Object:  Stored Procedure dbo.Company_Policy_View    Script Date: 10/22/01 6:24:22 PM ******/
/****** Object:  Stored Procedure dbo.Company_Policy_View    Script Date: 9/17/00 4:33:57 PM ******/
/****** Object:  Stored Procedure dbo.Company_Policy_View    Script Date: 9/4/01 6:30:44 PM ******/
CREATE PROCEDURE Company_Policy_View
@View  varchar(5),
@Policy_code varchar (10) ,
@Policy_type  varchar (50) 
as
if @View='Code'
begin
select  Policy_code,Policy_type,Policy_detail,U_id  From Company_Policy
 
where Policy_code=@Policy_code
end
if @View='Type'
begin
select  Policy_code,Policy_type,Policy_detail,U_id  From Company_Policy
 
where Policy_Type=@Policy_Type
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Count_HOLIDAY_LEAVE_ATTENDANCE    Script Date: 3/7/02 10:43:42 AM ******/
CREATE  PROCEDURE Count_HOLIDAY_LEAVE_ATTENDANCE
@Mode VARCHAR(8),
@Emp_ID VARCHAR(10),
@Mnt VARCHAR(12),				
@Pay_Year VARCHAR(4)				
AS
IF @Mode= 'LEAVE'				----------------01
BEGIN
DECLARE @@Tot_Leave INT
	SELECT @@Tot_Leave=SUM(DATEDIFF(DAY,start_dt,End_dt)+1 )
	FROM EMP_LEAVE_INFO
		WHERE Emp_ID=@Emp_ID AND
		DATEPART(MONTH,start_dt)=@Mnt AND
			DATEPART(YEAR,start_dt)=@Pay_Year and
		DATEPART(MONTH,end_dt)=@Mnt AND 
			DATEPART(YEAR,end_dt)=@Pay_Year
SELECT Tot_Leave=
	    CASE  WHEN  @@Tot_Leave IS NULL THEN 0 ELSE  @@Tot_Leave END
        
Return
END
---------------------------------------------------------------------------------------
IF @Mode='ATTEND'					----------------02
BEGIN
	SELECT Present=COUNT(emp_login),late=
			(SELECT COUNT(emp_login)FROM emp_att_info WHERE Emp_id=@Emp_id AND 
				IN_STATUS =1 AND 
				DATEPART(MONTH,Entry_Dt)=@Mnt AND 
				DATEPART(YEAR,Entry_Dt)=@Pay_Year)	
	FROM emp_att_info 
	WHERE Emp_id=@Emp_id AND 
		DATEPART(MONTH,Entry_Dt)=@Mnt AND 
		DATEPART(YEAR,Entry_Dt)=@Pay_Year
Return
END
-------------------------------------------------------------------------------------
IF @Mode='HOLID'					----------------03
BEGIN
	SELECT  TOT_HOL= SUM(datediff(day,str_date,End_date)+1)
			 FROM HOL_LIST
				WHERE DATEPART(MONTH,str_date)=@Mnt AND 	
					DATEPART(YEAR,str_date)=@Pay_Year or
					DATEPART(MONTH,End_date)=@Mnt AND 
					DATEPART(YEAR,End_date)=@Pay_Year
		
Return
END
-------------------------------------------------------------------------------------
/*IF @Mode='OT'
BEGIN
DECLARE @MAX_hr INT
SET @MAX_hr=208	
SELECT OT=SUM(DATEDIFF(hh,Emp_login,Emp_logout))-@MAX_hr FROM Emp_Att_info 
	WHERE Emp_id=@Emp_id AND 
		DATEPART(MONTH,Entry_Dt)=@Mnt AND 
		DATEPART(YEAR,Entry_Dt)=@Pay_Year
		
Select Amount=Amount/@MAX_hr from Fixed_Pay where Head_code=(Select Head_code from Pay_Struc WHERE Head_name='BASIC')and 
Emp_id=@Emp_id 
END*/
IF @Mode='OT'						----------------04			
BEGIN
DECLARE @MAX_hr INT
DECLARE @Mult INT
DECLARE @OT FLOAT(2)
DECLARE @Sal_phr FLOAT(5)
DECLARE @Hr_Factor FLOAT(2)
--------------------------------
SET  @MAX_hr=208	---Total Work-hour in a month
SET  @Mult=2		---Over time multiplier
SET  @Hr_Factor=0.01667
 ------------------------------------------------------------------------- -------------------------------------------------------------------------
/* 	OT_Pay=@Sal_phr*@Mult*@OT*/
/* 	Here @OT=Overtime hour,@Sal_phr=Salary per hour=(basic salary/@Max_hr)  */
-------------------------------- ------------------------------------------------------------------------- -----------------------------------------
SELECT @OT= ROUND((SUM(DATEDIFF(n,Emp_login,Emp_logout))* @Hr_Factor),1)-@MAX_hr   FROM Emp_Att_info 
WHERE Emp_id=@Emp_id AND 
		DATEPART(MONTH,Entry_Dt)=@Mnt AND 
		DATEPART(YEAR,Entry_Dt)=@Pay_Year
--------------------------------------------------------------------- ------------------------------------------------------------------------------
SELECT @Sal_phr=Amount/@MAX_hr FROM Fixed_Pay WHERE Head_code=(
	SELECT Head_code FROM Pay_Struc WHERE Head_name='BASIC')and 
	Emp_id=@Emp_ID
SELECT OT=@OT,OT_Pay=(@Sal_phr*@Mult)*@OT
Return
END
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
------------------OT for Liberty-------------------------
If @Mode='L_OT'						----------------05
Declare @SF int
BEGIN
	select @SF=count(*)from overtime where 
		Emp_ID=@Emp_ID and Pay_Month=@Mnt and Pay_year=@Pay_year
	IF @SF=1
	
	BEGIN
		SELECT Return_Status='Exists'
		Return
	
	END
	IF @SF=0
	BEGIN
		Declare @Basic MONEY		---Basic/Consolidated Salary
		Declare @Max_Work_Hr int	---Maximum working hour in a month(30 working days)
		Declare @Total_Hr int		---Total hour spent by an individual Worker
		Declare @Days_Att int 		---Total Attendance
		Declare @Daily_Hr float		---Daily working Hour
		Declare @OT_Factor float	---Multiplier (i.e Basic*2)
		Declare @OT_Hr float
		------------------------
		Set @Daily_Hr=8
		Set @Max_Work_Hr=240		----(30*@Daily_Hr)
		Set @OT_Factor=2
----------------------------Basic Salary--------------------------------------------
		Set @Basic=(SELECT Amount FROM Fixed_Pay WHERE Head_code=(
			SELECT Head_code FROM Pay_Struc WHERE Head_name='BASIC') and 
			Emp_id=@Emp_ID)
----------------------------Total Hour Spent----------------------------------------
		Select @Total_Hr=ROUND(SUM(DATEDIFF(n,Emp_login,Emp_logout)* 0.01667),1) FROM Emp_Att_info 
			WHERE Emp_id=@Emp_ID AND 
			DATEPART(MONTH,Entry_Dt)=@Mnt AND 
			DATEPART(YEAR,Entry_Dt)=@pay_year
------------------------------------------------------------------------------------
		Set @Days_Att=(SELECT COUNT(Emp_logout)FROM  Emp_Att_Info WHERE Emp_id=@Emp_ID AND 
			DATEPART(MONTH,Entry_Dt)=@Mnt AND DATEPART(YEAR,Entry_Dt)=@Pay_Year)
------------------------------------------------------------------------------------
		Set @OT_Hr=@Total_Hr-(@Days_Att*@Daily_Hr)
------------------------------------------------------------------------------------
		select Return_Status='null',OT_Hour=@OT_Hr,OT_Pay=ceiling(((@Basic/@Max_Work_Hr)*@OT_Factor)*@OT_Hr)
	END
Return
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE Procedure Create_User_AtOnce

AS


Set nocount on

declare @Emp_Id varchar(10)
,@Emp_Name varchar(45)
,@user_pass varchar(45)
,@uid varchar(15)
,@cancel bit

set @user_pass='45312123456'	------Password-------'new'
Set @uid='dsl'
Set @cancel=0 

---------------------------------------------------------------------------

declare ID_Cursor cursor for
	select Emp_id,Emp_Name=(Emp_fna+' '+Emp_mna+' '+Emp_lna)
	 from emp_per_info where Emp_Id not in (Select U_Id from Soft_Pass)

---------------------------------------------------------------------------

open ID_Cursor
	
	fetch next from ID_Cursor into @Emp_Id,@Emp_Name

		while @@Fetch_Status=0
		
		BEGIN
			insert into soft_Pass(U_Id,u_name,user_pass,uid,cancel )		
			values (@Emp_Id,@Emp_Name,@user_pass,@uid,@cancel)
				
		fetch next from ID_Cursor into @Emp_Id,@Emp_Name

		END
		
Close ID_Cursor
Deallocate ID_Cursor

select Message='User created successfully,Default Password  [New] is set for all !'

Set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO










CREATE   procedure DaleteAuthorInformation(
	@AuthorCode	VARCHAR(5))
AS
SET NOCOUNT ON
DECLARE @Msg VARCHAR (100)
IF EXISTS (SELECT * FROM LibraryBookAuthor WHERE     (AuthorId = @AuthorCode))
	BEGIN	
		SET @Msg='Delete not allowed!'
		GOTO Errmsg
	END

DELETE FROM AuthorInfo WHERE AuthorCode = @AuthorCode

RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





---DataBaseBackup_Restore 'c:\MyBackup.bak',0

create   Procedure DataBaseBackup_Restore 
(
    @Destination varchar(250),
    @Status int
) as
--set @Destination='c:\MyBackup.bak'
if @Status=0
	Begin
		Exec sp_Dropdevice PMIS_Disk_Device
		exec sp_addumpdevice 'disk',PMIS_Disk_Device,@Destination
		Backup Database PMIS to PMIS_Disk_Device
	end
else
	RESTORE DATABASE [PMIS] FROM  DISK = @Destination WITH  FILE = 1,  NOUNLOAD ,  STATS = 10,  RECOVERY









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




create  Procedure Date_Formats

AS

declare @Count int
Set @Count=0

while @Count<=14

	begin
		print 'Format '+ convert(char(2),@Count)+'  '+convert(char(20),getdate(),@Count)
		set @Count=@Count+1

	end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE procedure DeleteBookAuthor(
	@BookCode	VARCHAR(5),
	@AuthorName VARCHAR (50))
AS
SET NOCOUNT ON
DECLARE @Msg VARCHAR (100),
		@AuthorCode VARCHAR (5)
IF EXISTS (SELECT * FROM BookIssueRefund WHERE     (BookCode = @BookCode))
	BEGIN	
		SET @Msg='Delete not allowed!'
		GOTO Errmsg
	END

SELECT @AuthorCode = AuthorCode FROM AuthorInfo WHERE     (AuthorName = @AuthorName)

DELETE FROM LibraryBookAuthor WHERE     (BookCode = @BookCode) AND (AuthorId = @AuthorCode)

RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE procedure DeleteBookInfo(
	@BookCode	VARCHAR(5))
AS
SET NOCOUNT ON
DECLARE @Msg VARCHAR (100)
IF EXISTS (SELECT * FROM BookIssueRefund WHERE     (BookCode = @BookCode))
	BEGIN	
		SET @Msg='Delete not allowed!'
		GOTO Errmsg
	END

DELETE FROM LibraryBookInfo WHERE     (BookCode = @BookCode)
DELETE FROM LibraryBookAuthor WHERE     (BookCode = @BookCode) 

RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO





CREATE procedure DeleteBookIssueRule(
	@ClassCode	VARCHAR(5))
AS
SET NOCOUNT ON
DECLARE @Msg VARCHAR (100)

----SELECT * FROM BookIssueRule, ClassInfo  WHERE  (BookIssueRule.ClassCode =ClassInfo.ClassID) and  (ClassCode = @ClassCode)
	

DELETE FROM BookIssueRule WHERE (ClassCode = @ClassCode)


RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO








CREATE procedure DeleteBookIssueRule1(
	@ClassCode	VARCHAR(5))
AS
SET NOCOUNT ON
DECLARE @Msg VARCHAR (100)
IF EXISTS (SELECT * FROM BookIssueRule WHERE     (ClassCode = @ClassCode))
	BEGIN	
		SET @Msg='Delete not allowed!'
		GOTO Errmsg
	END

DELETE FROM BookIssueRule WHERE     (ClassCode = @ClassCode)


RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE   procedure DeleteLibrarySubInformation(
	@SubCode	VARCHAR(5))
AS
SET NOCOUNT ON
DECLARE @Msg VARCHAR (100)
IF EXISTS (SELECT * FROM LibraryBookInfo WHERE     (SubCode = @SubCode))
	BEGIN	
		SET @Msg='Delete not allowed!'
		GOTO Errmsg
	END

DELETE FROM LibrarySubjectInfo WHERE SubSubjectCode = @SubCode

RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE   procedure DeletePublisherInformation(
	@PubCode	VARCHAR(6))
AS
SET NOCOUNT ON
DECLARE @Msg VARCHAR (100)
IF EXISTS (SELECT * FROM LibraryBookInfo WHERE     (PubCode = @PubCode))
	BEGIN	
		SET @Msg='Delete not allowed!'
		GOTO Errmsg
	END

DELETE FROM PublisherInfo WHERE (PubCode = @PubCode)

RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF












GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/*
Exec Delete_Attendance_Record '1108','2004-01-20'

*/



CREATE  Proc Delete_Attendance_Record

@Emp_Id varchar(10)
,@Date datetime

AS

set nocount on

if exists(select * from emp_att_Info where emp_id=@Emp_Id
			and datepart(day,Emp_LogIn)=datepart(day,@Date)
			and datepart(year,Emp_LogIn)=datepart(year,@Date)
			and datepart(month,Emp_LogIn)=datepart(Month,@Date))
 begin
	
	Delete from Emp_Att_Info where emp_id=@Emp_id 
				and datepart(day,Emp_LogIn)=datepart(day,@Date)
				and datepart(year,Emp_LogIn)=datepart(year,@Date)
				and datepart(month,Emp_LogIn)=datepart(Month,@Date)
	

	Select Message='Record deleted successfully!'

end

else

		Select Message='Attendance record is not available!'


set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Department_I_U_D    Script Date: 3/7/02 10:43:29 AM ******/
/****** Object:  Stored Procedure dbo.Department_I_U_D    Script Date: 10/28/01 4:34:22 PM ******/
/****** Object:  Stored Procedure dbo.Department_I_U_D    Script Date: 10/22/01 6:24:22 PM ******/
/****** Object:  Stored Procedure dbo.Department_I_U_D    Script Date: 9/17/00 4:33:57 PM ******/
/****** Object:  Stored Procedure dbo.Department_I_U_D    Script Date: 9/4/01 6:30:44 PM ******/
--**********************
CREATE PROCEDURE Department_I_U_D
@opr varchar (1),
@Code varchar (3),
@Title varchar (35),
@Prev_Title varchar (35),
@Description varchar (50),
@U_id varchar (10)
AS
if @opr='I'
begin
insert into Department_Info (
Title,
Code,
Description,
U_id
) 
values (
@Title,
@Code,
@Description,
@U_id
)
end
                    
if @opr='U'
begin
update Department_Info set 
Title=@Title,
Code=@Code,
Description=@Description,
U_id=@U_id
where Title=@Prev_Title
end
    		
if @opr='D'
begin
delete from Department_Info where title=@Title
end



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE   Procedure Discard_Prepared_OT_Benefits
@Mode Varchar(5)
,@Pay_type varchar(2)
,@Emp_Id varchar(10)
,@Pay_month varchar(12)
,@Pay_year varchar(4)
AS
set nocount on

If @Mode='All'
Begin

	Delete from Payroll_main where Pay_month=@Pay_month 
	and Pay_year=@Pay_year and Pay_stat=0 and pay_type=@pay_type

	

End


If @Mode='Ind'
Begin

	Delete from payroll_main 
	Where Emp_Id=@Emp_Id and pay_Month=@pay_Month 
	and pay_Year=@pay_Year and Pay_stat=0 and pay_type=@pay_type
End

select message='Data discarded successfully !    '

set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE       Procedure Discard_Prepared_Salary
@Mode Varchar(5)
,@Emp_Id varchar(10)
,@Pay_month varchar(12)
,@Pay_year varchar(4)
AS
set nocount on

declare @Message varchar(150)
,@Num int
,@Num_Processed int
,@Num_Disbursed int

If @Mode='All'
Begin

	set @Num_Processed =(Select count(*)from payroll_main where Pay_month=@Pay_month 
		and Pay_year=@Pay_year and Pay_stat=0 and pay_type=0)

	set @Num_Disbursed=(Select count(*)from payroll_main where Pay_month=@Pay_month 
		and Pay_year=@Pay_year and Pay_stat=1 and pay_type=0)

	if @Num_Processed=0
		set @Message='No data is there to be deleted for '+ @Pay_month +', '+@Pay_year	
	
	if @Num_Disbursed!=0
		set @Message='Data can not be deleted for the month of '+ @Pay_month +', '+@Pay_year		

	if @Num_Processed !=0

		begin
			Delete from Payroll_main where Pay_month=@Pay_month 
			and Pay_year=@Pay_year and Pay_stat=0 and pay_type=0
		
			set @Message= convert(varchar(6),@Num_Processed)+' record(s) for '+ @Pay_month +', '+@Pay_year +' has been deleted !'
		end

End


If @Mode='Ind'
Begin


	set @Num_Processed=(Select count(*)from payroll_main where Emp_Id=@Emp_Id and Pay_month=@Pay_month 
		and Pay_year=@Pay_year and Pay_stat=0 and pay_type=0)
	
	set @Num_Disbursed=(Select count(*)from payroll_main where Emp_Id=@Emp_Id and Pay_month=@Pay_month 
		and Pay_year=@Pay_year and Pay_stat=1 and pay_type=0)


	if @Num_Processed=0
		set @Message='No data is there to be deleted for '+ @Pay_month +' ,'+@Pay_year 

		begin
			Delete from payroll_main 
			Where Emp_Id=@Emp_Id and pay_Month=@pay_Month 
			and pay_Year=@pay_Year and Pay_Stat=0 and pay_type=0

			set @Message='Salary  for '+ @Pay_month +' ,'+@Pay_year + ' for "'+@Emp_Id +'" has been deleted !'
		end

	if @Num_Disbursed=1
		
		set @Message='Specified data can not be deleted !'

End

select message=@message

set nocount off







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO













--- Exec BookDisretInformation 'Stu-00000000001','','00002','00001','N','25 Sep 2005','Lia'
CREATE      procedure ETypeInfo
(
	
	@EtypeID			varchar(5),
	@ETypeName	 		Varchar(50),
	@Note				varchar(50),
	@EntryBY			varchar(10),
	@EntryDate			DateTime
	
)

as



if exists (select * from ExamTypeInfo where  EtypeID= @EtypeID ) 		


Update ExamTypeInfo set  	


			
	ETypeName	=	@ETypeName	 		,
	Note		=	@Note				,
	EntryBY		=	@EntryBY			,
	EntryDate	=	@EntryDate			
	
where EtypeID= @EtypeID 
else


 insert into ExamTypeInfo

(

	EtypeID		,
	ETypeName	,
	Note		,
	EntryBY		,
	EntryDate	
	
			
		
	
)
values
(
	@EtypeID			,
	@ETypeName	 		,
	@Note				,
	@EntryBY			,
	@EntryDate	
	
)















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE   PROCEDURE EmpIncrementInformationInfo_Save
	@Emp_ID varchar(10),
	@IncrementAmount  decimal (9,2),
	@LastIncrementDt datetime,
	@NextIncremntDt datetime,
	@EffectiveDt datetime,
	@Remarks varchar(150),
	@Entry_Dt datetime,
	@Entry_By varchar(50)
		
as 
Declare
	@CurentBasic decimal (9,2)

	set @CurentBasic=(SELECT CBasic FROM  Emp_Job_Hist_Current WHERE  (Emp_id =@Emp_ID))

If Exists (Select * From EmpIncrementInformationInfo 
    Where  emp_id=@Emp_Id and EffectiveDt=@EffectiveDt)
	begin		
		Update EmpIncrementInformationInfo  Set 
		Emp_ID=@Emp_ID ,
		IncrementAmount=@IncrementAmount ,
		LastIncrementDt=@LastIncrementDt ,
		NextIncremntDt=@NextIncremntDt ,
		Remarks=@Remarks 
    Where  emp_id=@Emp_Id and EffectiveDt=@EffectiveDt
	end
Else

Insert Into EmpIncrementInformationInfo (
	Emp_ID ,
	IncrementAmount  ,
	LastIncrementDt ,
	NextIncremntDt ,
	EffectiveDt ,
	Remarks ,
	Entry_Dt ,
	Entry_By 
) Values (
	@Emp_ID ,
	@IncrementAmount  ,
	@LastIncrementDt ,
	@NextIncremntDt ,
	@EffectiveDt ,
	@Remarks ,
	@Entry_Dt ,
	@Entry_By 
)
Update Emp_Job_Hist_Current set CBasic=(@CurentBasic+@IncrementAmount)
		 Where  emp_id=@Emp_Id





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE   PROCEDURE Emp_Att_Info_I_U
	@Emp_id varchar (10),
	@U_Id varchar (10)
AS
set nocount on
------------------------Variable Declaration-------------------------------------------
declare 
	 @Duty_Type varchar(1)
	,@Start_time varchar(8)
	,@End_time varchar(8)
	,@Relaxed varchar(2)
	,@Abs_time varchar(8)
	,@Sp_Start_time varchar(10)
	,@Sp_Start_Day varchar(10)
	,@Sp_End_time varchar(10)
	,@Sp_End_Day varchar(10)
--------------------------
	,@Min varchar(2)
	,@opr varchar (2)
	,@value varchar(10)
	,@status varchar(100)
	,@Out_Status varchar(1)

--------------------------Get Office Time-------------------Begin----------------------------

Set  @Duty_Type=(Select Duty_type from emp_job_hist_current where Emp_Id=@Emp_Id )



if @Duty_Type='0'	--Shifting duty
begin
	
	declare @ID varchar(10)
	set @ID=@Emp_ID
	---------------------------------------------------------
	declare @S_Dt datetime
	declare @E_Dt datetime
	-------------------------
	select @S_Dt=start_dt,@E_Dt=End_Dt from group_Shift where Shift_Code=(
	select Shift_Code from Group_Shift where Gr_Code=(
	select Gr_Code from Worker_Group where emp_ID=@ID))

	--------------------------------------------------------
	select @Start_time=Shift_Start,@End_time=Shift_End,@Relaxed=[Delay],@Abs_time=Abs_time 
	from Shift where Shift_Code=(
		select Shift_Code from Group_Shift where Gr_Code=(
			select Gr_Code from Worker_Group where emp_ID=@ID))and 
				datepart(day,getdate())>=datepart(day,@S_Dt) and
				datepart(month,getdate())=datepart(month,@S_Dt)and
				datepart(year,getdate())=datepart(year,@S_Dt)and 
				datepart(day,getdate())<=datepart(day,@E_Dt) and
				datepart(month,getdate())=datepart(month,@E_Dt)and
				datepart(year,getdate())=datepart(year,@E_Dt)
END


if @Duty_Type='1'	--Fixed duty time

begin
	select @Start_time=Start_time,@End_time=End_time,@Relaxed=Relaxed,@Abs_time=Abs_time,  
		@Sp_Start_time=Sp_Start_time,@Sp_Start_day=Sp_Start_day,
		@Sp_end_time=Sp_end_time,@Sp_End_day=Sp_End_day
	from office_time where Effect_date=(
		select Mdt=(select max(Effect_date)from office_time 
			where CONVERT(char(12), Effect_date, 1) <= CONVERT(char(12),GETDATE(), 1)))
END



if @Duty_Type='9'	--Changble duty time

begin
		SELECT  @Start_time=ScheduleStTime,@End_time=ScheduleEndTime,
			@Abs_time=AbsentTime FROM  
		DriverSchedule  WHERE  (Emp_Id =@Emp_Id) 
		AND (DayoftheSchedule =dbo.GetDayName (getdate()))

		set @Relaxed=15
		

END

-----------------------------What Day is it?----------------------------------------------- 
SET DATEFIRST 5		-----Set Friday(5) as the first day of the week
-------------------------------------------------------------------------------------------
DECLARE 
	@Td varchar(1),
	@Today varchar(10)

SET @Td=DATEPART(dw,getdate())
   Set @Today=(SELECT [Day Name] =
		CASE @Td
	      	WHEN '1' THEN 'Friday'
     	 	WHEN '2' THEN 'Saturday'
		WHEN '3' THEN 'Sunday'
      		WHEN '4' THEN 'Monday'
		WHEN '5' THEN 'Tuesday'
		WHEN '6' THEN 'Wednesday'
		WHEN '7' THEN 'Thursday' END)


set @Today=ltrim(rtrim(@Today))




--------------------------Is today a special time day? ------------------------------------
If @Today=@Sp_Start_Day set @end_time= @Sp_Start_time
If @Today=@Sp_End_Day  set @end_time= @Sp_End_time
--------------------------Modify Office Time to 24th hour format---------------------------



-------------Start Time---------------------------
	set @Start_time=rtrim(ltrim(@Start_time))
	set @Min=(select substring(@Start_time,4,2))			----Cut minutes


set @Min=convert(int,@Min)+ convert(int,@Relaxed)

if len(@Min)=1
	set @Min='0'+convert(char,@Min)
							----Cut and add 12 with hour
if (select substring(@Start_time,7,1))='P' and (select substring(@Start_time,1,2))!='12'			
	Set @Start_time=convert(char,(convert(int,substring(@Start_time,1,2))+12))

else
	Set @Start_time=convert(char,substring(@Start_time,1,2))
	set @Start_time=rtrim(@Start_time)+ ':'+@Min+':00'		----Add seconds	
	set @Start_time=rtrim(ltrim(@Start_time))


-------------End Time---------------------------
declare 
	@EMin  varchar(4)

Set @End_time=rtrim(ltrim(@End_time))
set @EMin=(select substring(@End_time,3,3))			----Cut minutes
								----Cut and add 12 with hour

if (select substring(@End_time,7,1))='P'and (select substring(@End_time,1,2))!='12' 			 
	Set @End_time=convert(char,(convert(int,substring(Rtrim(@End_time),1,2))+12))
else
	Set @End_time=convert(char,substring(@End_time,1,2))

set @End_time=rtrim(@End_time)+ +@EMin+':00'			----Add seconds
set @End_time=ltrim(rtrim(@End_time))


------------Abs Time---------------------------
declare @AMin  varchar(4)
Set @Abs_time=rtrim(ltrim(@Abs_time))
set @AMin=(select substring(@Abs_time,3,3))			----Cut minutes
								----Cut and add 12 with hour 
if (select substring(@Abs_time,7,1))='P' and (select substring(@Abs_time,1,2))!='12'			
	Set @Abs_time=convert(char,(convert(int,substring(Rtrim(@Abs_time),1,2))+12))
else
	Set @Abs_time=convert(char,substring(@Abs_time,1,2))
set @Abs_time=rtrim(@Abs_time)+ +@AMin+':00'			----Add seconds
set @Abs_time=ltrim(rtrim(@Abs_time))


-----------Special Start Time-----------------------

declare 
	@SSMin  varchar(4)

Set @Sp_Start_time=rtrim(ltrim(@Sp_Start_time))
set @SSMin=(select substring(@Sp_Start_time,3,3))			----Cut minutes
									----Cut and add 12 with hour 
if (select substring(@Sp_Start_time,7,1))='P'and (select substring(@Sp_Start_time,1,2))!='12' 			
	Set @Sp_Start_time=convert(char,(convert(int,substring(Rtrim(@Sp_Start_time),1,2))+12))
else
	Set @Sp_Start_time=convert(char,substring(@Sp_Start_time,1,2))
	set @Sp_Start_time=rtrim(@Sp_Start_time)+ +@SSMin+':00'			----Add seconds
	set @Sp_Start_time=ltrim(rtrim(@Sp_Start_time))



-----------Special End Time-------------------------
declare 
	@SEMin  varchar(4)

	Set @Sp_End_time=rtrim(ltrim(@Sp_End_time))
	set @SEMin=(select substring(@Sp_End_time,3,3))			----Cut minutes
								----Cut and add 12 with hour 

if (select substring(@Sp_End_time,7,1))='P' and (select substring(@Sp_End_time,1,2))!='12' 	
	Set @Sp_End_time=convert(char,(convert(int,substring(Rtrim(@Sp_End_time),1,2))+12))
else
	Set @Sp_End_time=convert(char,substring(@Sp_End_time,1,2))
	set @Sp_End_time=rtrim(@Sp_End_time)+ +@SEMin+':00'			----Add seconds
	set @Sp_End_time=ltrim(rtrim(@Sp_End_time))


-------------------------------Attendance Mode--------------Insert or Update-------------------
--If not logged in then @opr='I'		>insert
--If already logged in then @opr='U'		>update
--If already logged out then @opr='X'	
	

if (select status=count(*) from emp_att_info where emp_id=@emp_id and 
		datepart(day,Entry_Dt)=datepart(day,getdate())and 
		datepart(month,Entry_Dt)=datepart(month,getdate())and 
		datepart(year,Entry_Dt)=datepart(year,getdate()))=1
	begin
		select @value=isnull(emp_logout,'1900-01-01')from emp_att_info
		---select @value=emp_logout from emp_att_info
		where Emp_id =@emp_id and 
			datepart(day,Entry_Dt)=datepart(day,getdate())and 
			datepart(month,Entry_Dt)=datepart(month,getdate())and 
			datepart(year,Entry_Dt)=datepart(year,getdate()) 
			
			---if @Value is null  set @Value=ltrim(rtrim('1900-01-01'))
			
			set @value=ltrim(rtrim(@value))
	
		if substring(@value,1,3)='Jan'or @value='1900-01-01'		
			begin
					set  @opr='U'
			end
		else 
			begin		set  @opr='X'
			
			end
	end
else
if (select status=count(*) from emp_att_info where emp_id=@emp_id and 
		datepart(day,Entry_Dt)=datepart(day,getdate())and 
		datepart(month,Entry_Dt)=datepart(month,getdate())and 
		datepart(year,Entry_Dt)=datepart(year,getdate()))=0
	begin
		set  @opr='I'
	end

-----------------------------Movement Tracker-------------------------------------------------

declare 
		@movement_status varchar(1)

set @movement_status=(select count (*) from emp_movement 
Where Emp_id=@Emp_id and Move_Status=1 and Move_Id=
		(select Move_Id=max(Move_Id) from Emp_movement where emp_id=@emp_Id))

-----------------------------Message Flags-------------------------------------------------

---0---On time Arival 
--if @opr='I' and (convert (char(8),getdate(),14)< @Start_time)

if @opr='I' and (convert (char(8),getdate(),14) < cast(@Start_time as datetime))--convert (char(8),@Start_time,14))
	set @status ='0'
	

---select convert (char(8),getdate(),14)



---1---Late Arival
if (convert (char(8),getdate(),14)> @Start_time) and (convert (char(8),getdate(),14)<@Abs_time)
	set @status ='1'						




---7---Absent
if @opr='I' and ((convert (char(8),getdate(),14)>@Start_time)) and ((convert (char(8),getdate(),14)>@Abs_time)) 
and ((convert (char(8),getdate(),14)<@End_time))  
	set @status='7' 	

---8---Logout before end_time
if  @opr='U'and  ((convert (char(8),getdate(),14)<@End_time))
	set @status='8'				

---0---Logout after end_time
if  @opr='U'and  ((convert (char(8),getdate(),14)>@End_time))
	set @status='0'				
---9---Already Logged Out
if @opr='X' 
	set @status='9'					


---2---Return from Movement


-----------------------------Insert / Update------------------------------------------
-----select Status=@Status,Curnt_Time=getdate(),End_time=@End_time

if @Status='0' or @Status='1'

begin
	if @opr='I'
	begin
		insert into Emp_Att_Info (Emp_id,Emp_login,In_Status,U_id) 
		values (@Emp_id,getdate(),@Status,@U_id)

			if @movement_status=1 
			begin
				update emp_movement set Move_Status='9',Rtn_Dt=getdate(),Entry_Dt=getdate()
					Where Emp_id=@Emp_id and Move_Status=1 and Move_Id=
						(select Move_Id=max(Move_Id) from Emp_movement where emp_id=@emp_Id)
			end




	end
                    
	if @opr='U'
	begin
		update Emp_Att_Info set Emp_logout=getdate(),Out_Status=@Status
		where Emp_id=@Emp_id and 
		convert(varchar (10),Entry_Dt,5)=convert(varchar (10),Getdate(),5)
	end
end


if @opr='U' and ltrim(rtrim(@U_Id))='8'
	begin
		update Emp_Att_Info set Emp_logout=getdate(),Out_Status='8'
		where Emp_id=@Emp_id and 
		convert(varchar (10),Entry_Dt,5)=convert(varchar (10),Getdate(),5)	
		set @status='0'
	end

if @status='8' and @movement_status=1 
	begin
		update emp_movement set Move_Status='9',Rtn_Dt=getdate(),Entry_Dt=getdate()
		Where Emp_id=@Emp_id and Move_Status=1 and Move_Id=
				(select Move_Id=max(Move_Id) from Emp_movement where emp_id=@emp_Id)
		set @status='2'
	end


--------------------Custom Messages and message flags--------------------------
if @opr='I' and @status='0' set @status='Welcome! you have logged in at'+ space(2)+ (convert (char(8),getdate(),14))
if @opr='U' and @status='0' set @status='Good Bye! you have logged out at'+ space(2)+ (convert (char(8),getdate(),14))
if @opr='I' and @status='1' set @status='You are late! you have logged in at'+ space(2)+ (convert (char(8),getdate(),14))
if @status='7' set @status='You are considered to be absent to day...!'
if @status='8' set @status='8' 
if @status='9' set @status='You have already logged out...!'
if @status='2' set @status='Welcome back...!' 


select Status=@status,Curnt_Time=getdate(),End_time=@End_time















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE    Procedure Emp_Att_Notes_I_U(
@Emp_Id varchar(10)	  
,@Note_Type int
,@Reason varchar(50)
,@Description varchar(300)

) as 

declare  @Attn_Id int

set @Attn_id=(
select  Attn_Id from emp_att_info where emp_id=@Emp_Id
and (in_status=1 or out_status=1)
and datepart(day,Emp_login)=datepart(day,getdate())
and datepart(month,Emp_login)=datepart(month,getdate())
and datepart(year,Emp_login)=datepart(year,getdate())
)

if @Attn_id!=''or @Attn_id!=null
begin

if exists(select * from emp_Att_Notes where Attn_id=@Attn_Id and Note_Type=@Note_Type)


	BEGIN
		update Emp_Att_Notes Set Reason=@Reason,Description=@Description where
		Attn_id=@Attn_Id and Note_Type=@Note_Type 
	END

	Else
	BEGIN


	Insert Into Emp_Att_Notes(Attn_Id,Note_Type,Reason,Description) 
	Values (@Attn_Id,@Note_Type,@Reason,@Description)
	
		if @Note_Type=1	----In_Notes
		Begin
			Update Emp_Att_Info set In_Notes=1 where Attn_Id=@Attn_Id 
		end	

		if @Note_Type=2	----Out_Notes
		Begin
			Update Emp_Att_Info set Out_Notes=1 where Attn_Id=@Attn_Id 
		end	
	END 


END





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE  Procedure Emp_Discipline_I_U_D(
    @opr varchar(1)
   ,@Emp_Id varchar(10)
   ,@Track_Id int
   ,@Ref_No varchar(40)
   ,@Reason varchar(500)
   ,@Penalty varchar(200)
   ,@U_Id varchar(10)
   ,@Dt datetime
) as 

If @opr='I'
Begin

Insert Into Emp_Discipline(
   Emp_Id
   ,Ref_No
   ,Reason
   ,Penalty
   ,U_Id
   ,Dt
) Values (
   @Emp_Id
   ,@Ref_No
   ,@Reason
   ,@Penalty
   ,@U_Id
   ,@Dt
)

End

if @opr='U'
Begin

update Emp_Discipline
Set
	Ref_No=@Ref_No
	,Reason=@Reason
	,Penalty=@Penalty
	,U_Id=@U_Id
	,Dt=@Dt
	Where Emp_Id=@Emp_ID and Track_Id=@Track_Id

End


if @opr='D'
Begin

Delete from Emp_Discipline
	Where Track_Id=@Track_Id

End





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







CREATE   Procedure Emp_Education_I_U_D
@opr varchar (10),
@Emp_id varchar (10),
@Edu_Count int,
@Exam_Name varchar (40),
@M_Subject varchar (40),
@Pass_Year varchar (4),
@Degree_From varchar (40),
@Result varchar (25),
@U_id varchar (10)
AS
set nocount on

declare @Message varchar(150)

if @opr='I'
	set @Message='Data saved successfully !'
if @opr='U'
	set @Message='Data updated successfully !'
if @opr='D'
	set @Message='Data deleted successfully !'


IF @opr='I'
BEGIN
Insert Into Emp_Education(
Emp_id
,Exam_Name
,M_Subject
,Pass_Year
,Degree_From
,Result
,U_id
)
values(
@Emp_id
,@Exam_Name
,@M_Subject
,@Pass_Year
,@Degree_From
,@Result
,@U_id
)
END
IF @opr='U'
BEGIN
UPDATE Emp_Education SET 
Exam_Name=@Exam_Name
,M_Subject= @M_Subject
,Pass_Year=@Pass_Year
,Degree_From=@Degree_From
,Result=@Result
,U_id=@U_id
WHERE  Edu_Count=@Edu_Count
END
IF @opr='D'
BEGIN
	DELETE FROM Emp_Education WHERE  Edu_Count=@Edu_Count
END

select message=@message

set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






CREATE  PROCEDURE Emp_Job_End_I_U_D
@opr varchar (1),
@Emp_id varchar (10),
@Job_end_dt datetime,
@Type varchar (20),
@Des varchar (100),
@U_id varchar (10)
-----------------------------	
AS
if @opr='I'
begin
delete from Emp_per_info where emp_id=@Emp_Id
insert into Emp_Job_End (
Emp_id,
Job_end_dt,
Type,
Des,
U_id
) 
values (
@Emp_id,
@Job_end_dt,
@Type,
@Des,
@U_id 
)

end
                    
if @opr='U'
begin
update Emp_Job_End set 
Job_end_dt=@Job_end_dt,
Type=@Type,
Des=@Des,
U_id=@U_id  where Emp_id=@Emp_id
end
    		
if @opr='D'
begin
delete from Emp_Job_End where Emp_id=@Emp_id
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




----exec Emp_per_Info_I_U_D 'U','DD01','dfgfd','dfgdfg','dfgdf','dfgdf','dfgdf','1988-01-01',17,'','Married','Male','','','Bangladeshi','None','Christianism',' ',' ',' ','Bangladesh',' ',' ',' ','Bangladesh',' ',' ',' ',' ',' ',' ','Normal',' ',' ','None    ','','','dsl'

CREATE     PROCEDURE Emp_Job_Hist_Current_I_U_D
@opr varchar (1),
@Emp_id varchar (10),
@Super_ID varchar(10),
@Emp_join_date datetime  ,
@Emp_rank varchar (25),
@Emp_desig  varchar (25), 
@Emp_branch  varchar (25),
@Emp_dept varchar (35),
@Emp_section varchar (20),
@Emp_job_type varchar (20), 
@Responsibility  varchar (200), 
@Pay_type varchar (1), 
@Duty_type varchar (1), 
@U_id varchar (10),
@Permdate datetime,
@CBasic int
--------------------------------
AS

set nocount on

declare @Message varchar(150)

if @opr='I'
	set @Message='Data saved successfully !'
if @opr='U'
	set @Message='Data updated successfully !'
if @opr='D'
	set @Message='Data deleted successfully !'

---------------Insert-----------
if @opr='I'
begin

	if Exists (Select * from Emp_Job_Hist_Current Where Emp_Id=@Emp_Id )
	BEGIN
		
		set @Message='Data already exists!'

	END

	else
	BEGIN
		insert into Emp_Job_Hist_Current (
			Emp_id,Emp_join_date,Emp_rank ,Emp_desig,  Emp_branch, Emp_dept ,Emp_section,Emp_job_type,Responsibility,
			Pay_type ,Duty_type,Super_ID,U_id,Permdate,CBasic) 
		values (@Emp_id,@Emp_join_date,@Emp_rank ,@Emp_desig,@Emp_branch, @Emp_dept ,@Emp_section,@Emp_job_type,
			@Responsibility,@Pay_type ,@Duty_type,@Super_ID,@U_id,@Permdate,@CBasic)

	END
end
------------------------------------
                    
if @opr='U'
begin
update Emp_Job_Hist_Current set 
Emp_id=@Emp_id,
Emp_join_date=@Emp_join_date,
Emp_rank =@Emp_rank ,
Emp_desig=@Emp_desig,
Emp_branch=@Emp_branch, 
Emp_dept=@Emp_dept ,
Emp_section=@Emp_section,
Emp_job_type=@Emp_job_type,
Responsibility=@Responsibility,
Pay_type =@Pay_type ,
Duty_type=@Duty_type,
Super_id=@Super_id,
U_id=@U_id,
Permdate=@Permdate,
CBasic=@CBasic

where Emp_id=@Emp_id
end
---------------------------------------------------  
	
if @opr='D'
begin
	delete from  Emp_Job_Hist_Current where Emp_id = @emp_id

end
---------------------------------------------------


select message=@Message

set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Emp_Job_Hist_Future_I_U_D    Script Date: 3/7/02 10:43:43 AM ******/
/****** Object:  Stored Procedure dbo.Emp_Job_Hist_Future_I_U_D    Script Date: 10/28/01 4:34:37 PM ******/
/****** Object:  Stored Procedure dbo.Emp_Job_Hist_Future_I_U_D    Script Date: 10/22/01 6:24:36 PM ******/
/****** Object:  Stored Procedure dbo.Emp_Job_Hist_Future_I_U_D    Script Date: 9/17/00 4:34:09 PM ******/
/****** Object:  Stored Procedure dbo.Emp_Job_Hist_Future_I_U_D    Script Date: 9/4/01 6:30:54 PM ******/
----------------------------------------
CREATE procedure Emp_Job_Hist_Future_I_U_D
@opr varchar(1),
@stored_date datetime,
@Emp_id  varchar(10),
@Effect_date datetime,
@Emp_Rank varchar(25),
@Emp_desig varchar(25),
@Emp_branch varchar(25),
@Emp_dept varchar(20),
@Emp_section varchar(20),
@Emp_job_type varchar(20),
@U_id varchar(10)
as
if @opr='I'
begin
	insert into emp_job_hist_future (
		Emp_id,
		Effect_date,
		Emp_Rank,	
		Emp_desig,
		Emp_branch,
		Emp_dept,
		Emp_section,
		Emp_job_type,
		U_id)
	values(
		@Emp_id,
		@Effect_date,
		@Emp_Rank,@Emp_desig,
		@Emp_branch,
		@Emp_dept,
		@Emp_section,
		@Emp_job_type,
		@U_id
		)
end
if @opr='U'
begin
	Update emp_job_hist_future set
		Emp_id=@Emp_id,
		Effect_date=@Effect_date,
		Emp_Rank=@Emp_Rank,
		Emp_desig=@Emp_desig,
		Emp_branch=@Emp_branch,
		Emp_dept=@Emp_dept,
		Emp_section=@Emp_section,
		Emp_job_type=@Emp_job_type,
		U_id=@U_id
		where emp_id=@emp_id and Effect_date=@Stored_date
end
if @opr='D'
begin
	delete from emp_job_hist_future 
	where emp_id=@emp_id and Effect_date=@Stored_date
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





CREATE      PROCEDURE [Emp_Leave_Info_I_U_D] 
@opr  varchar(1),
@App_ID  int,
@Emp_id varchar (10) ,
@Leave_name varchar (25) ,
@Start_dt datetime ,
@End_dt  datetime ,
@Des varchar (100)  ,
@Address  varchar (100) ,
@Tel1 varchar (20) ,
@App_Type varchar(1),
@Super_ID varchar(10),
@U_id varchar (10) 
AS
set nocount on

DECLARE  @counter int
,@Duration int
,@Message varchar(150)
,@latest_App_Id int
------------------------------------------------------------------------------------------------------------------------------------------
if @opr='I'
	set @Message='Data saved successfully !'	
if @opr='U'			
	set @Message='Data updated successfully !'		
if @opr='D'
	set @Message='Data deleted successfully !'		
------------------------------------------------------------------------------------------------------------------------------------------

if @Leave_name='Maternity Leave' and (select dbo.GetLeaveValidation_Mt(@Emp_id))=0
	set @Message='Only married female employee can enjoy leave of the current category!'

else

begin


	SET @counter=0
	IF DATEPART(MONTH,@Start_dt)!= DATEPART(MONTH,@End_Dt) OR
		DATEPART(YEAR,@Start_Dt)!= DATEPART(YEAR,@End_Dt)	
	
	/*********************************************************************
	If the leave starts this month and ends next month (Example: Str_Date=2001-09-28 
	and End_Date=2001-10-01)a Single entry will then breaks up into several records.
	In this case (Example) it will create 05 records,one for each day within 
	the given date range.
	***********************************************************************/
	BEGIN
	
		Set @duration=(SELECT DATEDIFF(DAY,@Start_dt,@End_Dt))
		set @latest_App_Id =(select dbo.GetAuto_SlNo('Leave'))
		
		WHILE @counter<=@Duration
			
		BEGIN
			DECLARE @Mod_Str_Dt DATETIME
	 		
			SET @Mod_Str_Dt=DATEADD(DAY,@Counter,@Start_Dt)
	
	------------------------------------------------------------------------------------	
	
				IF @opr='I'		--------INSERT----------
					BEGIN
						if (select dbo.GetLeaveRangeMatched(@Emp_id,@Mod_Str_Dt,@Mod_Str_Dt))=1
					
							set @Message='Specified date range ('+convert(char(9),@Start_dt,6)+ ' - '+convert(char(9),@End_dt,6)+') already exists!'					
						else			
							begin
								insert into Emp_Leave_Info (Emp_id,Leave_name ,Start_dt ,End_dt,Des,Address,Tel1,App_Type,Super_ID,App_ID,U_id) 
								values (@Emp_id,@Leave_name ,@Mod_Str_Dt ,@Mod_Str_Dt,@Des,@Address,@Tel1,@App_Type,@Super_ID,@latest_App_Id,@U_id)
							end
					END
	                    
				IF @opr='U'		--------UPDATE----------
					BEGIN
						Update Emp_Leave_Info set Leave_name=@Leave_name ,Start_dt=@Mod_Str_Dt,End_dt=@Mod_Str_Dt,
						Des=@Des,Address=@Address,Tel1=@Tel1,@App_Type=App_Type,Super_ID=@Super_ID,U_id=@U_id 
						where App_ID=@App_ID 
					END
	
				IF @opr='D'		--------DELETE----------
					BEGIN
						delete from  Emp_Leave_Info  
						where App_ID=@App_ID 
					
					END
				-------------------------------------------------------------------------------------
			SET @counter=@counter+1
		END
	
	END
	ELSE
	BEGIN
		IF @opr='I'		--------INSERT----------
		BEGIN
		
			if (select dbo.GetLeaveRangeMatched(@Emp_id,@Start_dt,@End_dt))=1
		
				set @Message='Specified date range ('+convert(char(9),@Start_dt,6)+ ' - '+convert(char(9),@End_dt,6)+') already exists!'					
			else
				begin
					insert into Emp_Leave_Info (Emp_id,Leave_name ,Start_dt ,End_dt,Des,Address,Tel1,App_Type,App_ID,U_id) 
					values (@Emp_id,@Leave_name ,@Start_dt ,@End_dt,@Des,@Address,@Tel1,@App_Type,(dbo.GetAuto_SlNo('Leave')),@U_id)
				end
		END
			
		IF @opr='U'		--------UPDATE----------
		BEGIN
			Update Emp_Leave_Info set Leave_name=@Leave_name ,Start_dt=@Start_dt,End_dt=@End_dt,
				Des=@Des,Address=@Address,Tel1=@Tel1,@App_Type=App_Type,U_id=@U_id 
				where App_ID=@App_ID 
		END
			
		IF @opr='D'		---------DELETE---------
		BEGIN
			delete from  Emp_Leave_Info  
				where App_ID=@App_ID 
		END
	END

END

select Message= @Message

set nocount off


----select * from emp_leave_info






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







--update Emp_Movement set move_status=1
--select * from Emp_Movement

CREATE          procedure [Emp_Movement_I_U_D]
@opr varchar (10),
@Emp_id varchar (10),
@Move_Id varchar(10),
@Mode varchar (20) ,
@Place  varchar (20),
@Move_Out_Dt  datetime ,
@Exp_Rtn_Dt datetime,
@Cont_Tel varchar(20),
@Move_des  varchar(500)
As
set nocount on
declare @Move_Status int
,@Message varchar(150)

----------------------------Make Schedule---------------------------
if @opr='SI'		-----insert Schedule
begin
	insert into Emp_Movement(Emp_id,Mode,Place,Move_Out_Dt,Exp_Rtn_Dt,Cont_Tel,Move_des)
	Values(@Emp_id ,@Mode,@Place,@Move_Out_Dt,@Exp_Rtn_Dt ,@Cont_Tel,@Move_des)
end
----------------------------Schedule Update---------------------------
if @opr='SU'
begin
set @Move_Status=0
	update Emp_Movement set
		Mode=@Mode,Place=@Place,Move_Out_Dt=@Move_Out_Dt,Exp_Rtn_Dt=@Exp_Rtn_Dt,
		Cont_Tel=@Cont_Tel,Move_des=@Move_des,Move_Status=@Move_Status
	Where Emp_id=@Emp_id and Move_Id=convert(int,@Move_Id) and Move_Status=0

	Set @Message='Schedule successfully updated !'

end
----------------------------Non Scheduled Out---------------------------
-----PNS Group July 10 2003

if @opr='NSOut'

begin

	declare @LogInStat int
	,@LogOutStat int
	,@Off_End varchar(8)
	,@Already_In_Movement int

	set @LogInStat =(select	dbo.GetPresentOrNot(@Emp_Id,getdate()))
	set @LogOutStat =(select dbo.GetLogOutOrNot(@Emp_Id,getdate()))
	set @Off_End =(select dbo.GetOfficeEnd (@Emp_Id,getdate()))
	set @Already_In_Movement =(select dbo.GetIf_In_Movement (@Emp_Id,getdate()))
	
	--date_formats
	---------------------------------Checking------------------------------------------

	if @LogInStat=0				---If Absent
		Set @Message='You did not login to office today,movement is not possible in this context!'
	else					---If present
		if @Already_In_Movement=1	---If alreayd in movement
			Set @Message='You must be back from previous movement prior to make another movement!'
		else
			if @LogOutStat=1 	---If already Logged out
				Set @Message='You have already logged out,movement is not possible now!'
			else			----If not logged out	
				if convert(char(8),getdate(),14)>=@Off_End 
					Set @Message='Office time is over official movement is not possible now !'

	----------------------------------------------------------------------------
	if @LogInStat=1 and 				--If present
	   @LogOutStat=0 and 				--If not logged out
	   @Already_In_Movement=0 and			--If not in Movement
	   convert(char(8),getdate(),14)< @Off_End 	--If Office is not over
	----------------------------------------------------------------------------
		begin
			set @Move_Status=1

			insert into Emp_movement(Emp_id ,Mode,Place,Move_Out_Dt,Exp_Rtn_Dt,Cont_Tel,Move_des,Move_Status)
			Values (@Emp_id ,@Mode,@Place,getdate(),@Exp_Rtn_Dt,@Cont_Tel,@Move_des,@Move_Status)
		
			Set @Message='Your movement is recorded successfully, See you again by  '+convert(char(20),@Exp_Rtn_Dt,0)
		end
	
end 

----------------------------Scheduled Out---------------------------
if @opr='SOut'
set @Move_Status=1
begin
	update Emp_Movement set Move_Status=@Move_Status
	Where Emp_id=@Emp_id  and Move_Status=0
end
----------------------------Move In--------------------------------
if @opr='In'
begin
	
declare @In_Movement int
set @In_Movement =(select dbo.GetIf_In_Movement (@Emp_Id,getdate()))

	if @In_Movement=1

	begin
		update Emp_Movement set Rtn_Dt=Getdate(),Move_Status=9
		Where Emp_id=@Emp_id and Move_Status=1 and Move_Id=
			(select Move_Id=max(Move_Id) from Emp_movement where emp_id=@emp_Id)

		Set @Message='Welcome Back !'
	end
	else

		Set @Message='You were not in movement !'

	

end
----------------------------Delete Schedule------------------------
if @opr='SD'
begin
	Delete from Emp_Movement
	Where Emp_id=@Emp_id and Move_Id=convert(int,@Move_Id) and Move_Status =0

	Set @Message='Specified movement removed from schedule !'
end



select Message=@Message

set nocount off







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






CREATE      PROCEDURE Emp_Per_Info_I_U_D

--------------Variable declaration-------------
	@opr varchar (1),
	@Emp_id varchar(10),
	@Emp_fna varchar(15),
	@Emp_mna varchar(50) ,
	@Emp_lna varchar(15) ,
	@Emp_fa_na varchar(45) ,
	@Emp_ma_na varchar(45) ,
	@Emp_d_of_b datetime ,
	@Emp_age tinyint,
	@Blood_group varchar (8) ,
	@Emp_marital_st varchar(10) ,
	@Emp_gender varchar (6),
	@S_Security varchar (20) ,
	@Voter_ID varchar (10) ,
	@Emp_nat varchar (15) ,
	@Emp_sp_qta varchar (25) ,
	@Emp_religion varchar (15) ,
	@Emp_perm_add varchar (200) ,
	@Emp_perm_town varchar (20) ,
	@Emp_post1 varchar (8) ,
	@Emp_country1 varchar(20) ,
	@Emp_tel1 varchar (20) ,
	@Emp_fax1 varchar (20) ,
	@Emp_email1 varchar (25) ,
	@Contact_person varchar (30) ,
	@Contact_add varchar (200),
	@Contact_town varchar (20),
	@Contact_post varchar (8),
	@Contact_tel varchar (50),
	@Contact_fax varchar (20) ,
	@Contact_email varchar (25) ,
	@Emp_eye varchar (6) ,
	@Emp_height varchar (3) ,
	@Emp_weight varchar (3) ,
	@Emp_disable varchar (500)  ,
	@pick_pic_path varchar (200) ,
	@Entry_Dt datetime ,
	@U_id varchar (10)
-----------------------Insert---------------------------
AS

set nocount on
declare 
		@Message varchar(150),
		@MaxSLNo int,
		@EmpIDNew char(8)


if @Emp_id =''

	begin

		SELECT @MaxSLNo =  isnull(max(cast(substring(Emp_Id,4,3) as int)),0)
				FROM Emp_Per_Info 

		
		if @MaxSLNo=0
			set @MaxSLNo=1
		else
			set @MaxSLNo=@MaxSLNo+1

		select @EmpIDNew= dbo.PadString(@MaxSLNo,3,'0','L')


		set @Emp_id='BS-'+@EmpIDNew
	end









if exists (select * from emp_per_info where emp_Id=@Emp_id )

if @opr='I'
	set @Message='Data saved successfully !'
if @opr='U'
	set @Message='Data updated successfully !'
if @opr='D'
	set @Message='Data Deleted successfully !'



if @opr='I'

begin
insert into Emp_Per_Info (
	Emp_id,
	Emp_fna,
	Emp_mna,
	Emp_lna,
	Emp_fa_na,
	Emp_ma_na,
	Emp_d_of_b,
	Emp_age,
	Blood_group,
	Emp_marital_st,
	Emp_gender ,
	S_Security ,
	Voter_ID ,
	Emp_nat ,
	Emp_sp_qta ,
	Emp_religion ,
	Emp_perm_add ,
	Emp_perm_town ,
	Emp_post1 ,
	Emp_country1 ,
	Emp_tel1 ,
	Emp_fax1 ,
	Emp_email1 ,
	Contact_person,
	Contact_add,
	Contact_town,
	Contact_post,
	Contact_tel,
	Contact_fax,
	Contact_email,
	Emp_eye,
	Emp_height,
	Emp_weight,
	Emp_disable,
	pick_pic_path,
	Entry_Dt,
	U_id) 
	values ( 
	@Emp_id,
	@Emp_fna,
	@Emp_mna,
	@Emp_lna,
	@Emp_fa_na,
	@Emp_ma_na,
	@Emp_d_of_b,
	@Emp_age,
	@Blood_group,
	@Emp_marital_st,
	@Emp_gender,
	@S_Security,
	@Voter_ID,
	@Emp_nat,
	@Emp_sp_qta,
	@Emp_religion,
	@Emp_perm_add,
	@Emp_perm_town,
	@Emp_post1,
	@Emp_country1,
	@Emp_tel1,
	@Emp_fax1,
	@Emp_email1,
	@Contact_person ,
	@Contact_add ,
	@Contact_town ,
	@Contact_post ,
	@Contact_tel ,
	@Contact_fax ,
	@Contact_email,
	@Emp_eye,
	@Emp_height,
	@Emp_weight,
	@Emp_disable,
	@pick_pic_path,
	@Entry_Dt,
	@U_id
)
end
-----------------------------Update----------------------------------
                    
if @opr='U'
begin
update Emp_Per_Info set 
	Emp_fna=@Emp_fna,
	Emp_mna=@Emp_mna,
	Emp_lna=@Emp_lna,
	Emp_fa_na=@Emp_fa_na,
	Emp_ma_na=@Emp_ma_na,
	Emp_d_of_b=@Emp_d_of_b,
	Emp_age=@Emp_age,
	Blood_group=@Blood_group,
	Emp_marital_st=@Emp_marital_st,
	Emp_gender=@Emp_gender,
	S_Security=@S_Security,
	Voter_ID=@Voter_ID,
	Emp_nat=@Emp_nat,
	Emp_sp_qta=@Emp_sp_qta,
	Emp_religion=@Emp_religion,
	Emp_perm_add=@Emp_perm_add,
	Emp_perm_town=@Emp_perm_town,
	Emp_post1=@Emp_post1,
	Emp_country1=@Emp_country1,
	Emp_tel1=@Emp_tel1,
	Emp_fax1=@Emp_fax1,
	Emp_email1=@Emp_email1,
	Contact_person=@Contact_person,
	Contact_add=@Contact_add,
	Contact_town=@Contact_town,
	Contact_post=@Contact_post,
	Contact_tel=@Contact_tel,
	Contact_fax=@Contact_fax,
	Contact_email=@Contact_email,
	Emp_eye=@Emp_eye,
	Emp_height=@Emp_height,
	Emp_weight=@Emp_weight,
	Emp_disable=@Emp_disable,
	pick_pic_path=@pick_pic_path,
	Entry_Dt=@Entry_Dt,
	U_id=@U_ID
	where Emp_id=@Emp_id
end
------------------------------------Delele---------------------------    		
if @opr='D'
begin
	delete from  Emp_Per_Info where emp_id=@emp_id
	
end


select message=@Message
set nocount off










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/****** Object:  Stored Procedure dbo.Emp_Reference_I_U_D    Script Date: 3/7/02 10:43:30 AM ******/
/****** Object:  Stored Procedure dbo.Emp_Reference_I_U_D    Script Date: 10/28/01 4:34:22 PM ******/
/****** Object:  Stored Procedure dbo.Emp_Reference_I_U_D    Script Date: 10/22/01 6:24:22 PM ******/
/****** Object:  Stored Procedure dbo.Emp_Reference_I_U_D    Script Date: 9/17/00 4:33:57 PM ******/
/****** Object:  Stored Procedure dbo.Emp_Reference_I_U_D    Script Date: 9/4/01 6:30:44 PM ******/
------------------------------------------------- Insert/Update/Delete Procedure----------
CREATE   procedure Emp_Reference_I_U_D		
@opr varchar(1),
@Emp_id varchar (10),
@Ref_Count int,
@Ref_Name varchar (30),
@Ref_Occup varchar (40),
@Ref_Add varchar(200),
@Ref_town varchar (20),
@Ref_post varchar (8),
@Ref_Country varchar(20),
@Ref_tel varchar(20),
@Ref_fax varchar (30),
@Ref_email varchar (30),
@Ref_Relation varchar (30),
@U_id varchar(10)
as
set nocount on

declare @Message varchar(150)

if @opr='I'
	set @Message='Data saved successfully !'
if @opr='U'
	set @Message='Data updated successfully !'
if @opr='D'
	set @Message='Data deleted successfully !'


if @opr='I'
begin
	insert into Emp_reference (
		Emp_id,	
		Ref_Name, Ref_Occup,Ref_Add,Ref_town,Ref_post,Ref_Country,Ref_tel,Ref_fax,
		Ref_email,Ref_Relation,U_id		

	)
	values
	(
		@Emp_id,@Ref_Name,@Ref_Occup,@Ref_Add,@Ref_town,@Ref_post,@Ref_Country,@Ref_tel,@Ref_fax,
		@Ref_email,@Ref_Relation,@U_id
	)
end
---------------------------------------
if @opr='U'
begin
	Update Emp_reference set
		
		Ref_Name=@Ref_Name,Ref_Occup=@Ref_Occup,Ref_Add=@Ref_Add,Ref_town=@Ref_town,Ref_post=@Ref_post,
		Ref_Country=@Ref_Country,Ref_tel=@Ref_tel,Ref_fax=@Ref_fax,Ref_email=@Ref_email,
		Ref_Relation=@Ref_Relation,
		U_id=@U_id
	Where Ref_Count=@Ref_Count
end
---------------------------------------
if @opr='D'
begin
	Delete from Emp_reference
	where Ref_Count=@Ref_Count
end

select Message=@Message

set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







CREATE          Procedure Emp_Salary_Payscale_Hist
@opr varchar(10)
,@Emp_Id varchar(10)
,@Basic_Amount varchar(10)
,@Scale_Code int
,@Effect_Dt datetime
,@Acc_No varchar(30)		-------Attention:Need to be excluded
,@U_Id char(10)

as 

set nocount on
declare @Message varchar(200)
,@Check_Scale varchar(200)

set @Check_Scale =(select dbo.GetSalFixErrMsg(@Scale_Code,@Basic_Amount))


if @opr='I'
	set @Message='Data saved successfully !'
if @Opr='I'and @Scale_Code = 0
	set @Message='Rank/catagory is not yet set for Employee :'+@Emp_Id + ' plese update Job Detail!'
if @opr='U'
	set @Message='Data updated successfully !'
if @opr='D'
	set @Message='Data deleted successfully !'

if (@Opr='I' or @Opr='U')and @Scale_Code != 0 and @Check_Scale !='Valid'
	set @Message=@Check_Scale

--Begin transaction

if @opr='D'
BEGIN
	if exists (select * from payroll_main where emp_id=@emp_id)
		begin
			set @Message='Data can not be modified or deleted !'	
		end
	else	
		begin
			DELETE FROM Fixed_Pay WHERE Emp_ID=@Emp_ID
			DELETE from emp_bank_account where Emp_Id=@Emp_Id
			delete from Emp_Payscale_Hist where Emp_Id=@Emp_Id
			delete from Increment_History  where Emp_Id=@Emp_Id
		end
END

if (@opr='I' or @opr='U')and @Check_Scale='Valid' and @Scale_Code != 0


BEGIN

	If exists (select * from payroll_main where emp_id=@emp_id)
		begin
			exec emp_bank_account_I_U @Emp_Id,'',@Acc_No,@U_id
			set @Message='Bank account number is updated !'	
		end
	else	

	begin

		Exec Insert_Into_Fixed_Pay  @Emp_Id ,@Scale_Code,@Basic_Amount,@Effect_Dt,@U_Id
		exec Increment_History_I_U @Emp_Id,@Scale_Code,@Basic_Amount,@Effect_Dt 
		if not exists (Select * from Emp_Payscale_Hist where 
			Emp_id=@Emp_ID and Scale_Code=@Scale_Code)
	
		BEGIN
			Insert Into Emp_Payscale_Hist(Emp_Id,Scale_Code,Effect_Dt,U_Id,Dt)
			 Values (@Emp_Id,@Scale_Code,@Effect_Dt,@U_Id,getdate())
		end
	
		exec emp_bank_account_I_U @Emp_Id,'',@Acc_No,@U_id
	end
END

/*
IF (@@error <> 0)
BEGIN
  ROLLBACK TRANSACTION
	
	set @Message='Error occured during saving data !'	

END
*/
select Message=@Message

set nocount off








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Emp_View    Script Date: 3/7/02 10:43:24 AM ******/
/****** Object:  Stored Procedure dbo.Emp_View    Script Date: 10/28/01 4:34:22 PM ******/
/****** Object:  Stored Procedure dbo.Emp_View    Script Date: 10/22/01 6:24:19 PM ******/
/****** Object:  Stored Procedure dbo.Emp_View    Script Date: 9/17/00 4:33:57 PM ******/
/****** Object:  Stored Procedure dbo.Emp_View    Script Date: 9/4/01 6:30:44 PM ******/
/****** Object:  Stored Procedure dbo.Emp_View    Script Date: 6/20/01 12:00:33 PM ******/
CREATE PROCEDURE [Emp_View]
@Find varchar(10)
 AS
select * from emp_per_info where emp_id=@find




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Emp_performance_I_U_D    Script Date: 3/7/02 10:43:43 AM ******/
/****** Object:  Stored Procedure dbo.Emp_performance_I_U_D    Script Date: 10/28/01 4:34:37 PM ******/
/****** Object:  Stored Procedure dbo.Emp_performance_I_U_D    Script Date: 10/22/01 6:24:36 PM ******/
/****** Object:  Stored Procedure dbo.Emp_performance_I_U_D    Script Date: 9/17/00 4:34:09 PM ******/
/****** Object:  Stored Procedure dbo.Emp_performance_I_U_D    Script Date: 9/4/01 6:30:55 PM ******/
CREATE PROCEDURE Emp_performance_I_U_D
@opr as varchar(1),
@Emp_id  varchar (10),
@Ev_dt_from  datetime,
@Ev_dt_to datetime,
@Job varchar(3),
@Skill varchar(3),
@Comn varchar(3),
@Sincere varchar(3),
@Discip varchar(3),
@Cooper varchar(3),
@Laeader varchar(3),
@Motiv varchar (3),
@Plann varchar (3),
@Initiate varchar(3),
@Att varchar(3),
@App varchar(3),
@U_id char(10) 
AS
if @opr='I'
begin
	insert into Emp_performance (Emp_id,Ev_dt_from,Ev_dt_to,Job,Skill,Comn,Sincere,
		Discip,Cooper,Laeader,Motiv,Plann,Initiate,Att,App,U_id) 
	values (@Emp_id,@Ev_dt_from,@Ev_dt_to,@Job,@Skill,@Comn,@Sincere,@Discip,@Cooper,
		@Laeader,@Motiv,@Plann,@Initiate,@Att,@App,@U_id)
end
                    
if @opr='U'
begin
	update Emp_performance set Ev_dt_from=@Ev_dt_from,Ev_dt_to=@Ev_dt_to,Job=@Job,
		Skill=@Skill,Comn=@Comn,Sincere=@Sincere,Discip=@Discip,Cooper=@Cooper,
		Laeader=@Laeader,Motiv=@Motiv,Plann=@Plann,Initiate=@Initiate,Att=@Att,
		App=@App,U_id=@U_id
		 where Emp_id=@Emp_id and  Ev_dt_from=@Ev_dt_from and Ev_dt_to=@Ev_dt_to
end
    		
if @opr='D'
begin
	delete from  Emp_performance 
		Where  Emp_id=@Emp_id and  Ev_dt_from=@Ev_dt_from and Ev_dt_to=@Ev_dt_to
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Emp_super_I_U_D    Script Date: 3/7/02 10:43:43 AM ******/
/****** Object:  Stored Procedure dbo.Emp_super_I_U_D    Script Date: 10/28/01 4:34:38 PM ******/
/****** Object:  Stored Procedure dbo.Emp_super_I_U_D    Script Date: 10/22/01 6:24:37 PM ******/
/****** Object:  Stored Procedure dbo.Emp_super_I_U_D    Script Date: 9/17/00 4:34:10 PM ******/
/****** Object:  Stored Procedure dbo.Emp_super_I_U_D    Script Date: 9/4/01 6:30:55 PM ******/
/****** Object:  Stored Procedure dbo.Emp_super_I_U_D    Script Date: 6/20/01 12:00:39 PM ******/
/****** Object:  Stored Procedure dbo.Emp_super_I_U_D    Script Date: 5/14/01 9:49:53 AM ******/
/****** Object:  Stored Procedure dbo.Emp_super_I_U_D    Script Date: 1/1/99 3:01:44 AM ******/
/****** Object:  Stored Procedure dbo.Emp_super_I_U_D    Script Date: 5/9/01 8:13:09 AM ******/
/****** Object:  Stored Procedure dbo.Emp_super_I_U_D    Script Date: 5/8/01 4:09:48 AM ******/
CREATE PROCEDURE Emp_super_I_U_D
@opr varchar (1),
@Emp_id varchar (10),
@Sup_id varchar (10),
@U_id varchar (10)
AS
if @opr='I'
begin
insert into Emp_super (
Emp_id,
Sup_id,
U_id
) 
values (
@Emp_id,
@Sup_id,
@U_id
)
end
                    
if @opr='U'
begin
update Emp_super set 
Emp_id=@Emp_id,
Sup_id=@Sup_id,
U_id=@U_id
end
    		
if @opr='D'
begin
	delete from Emp_super
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO















--- Exec BookDisretInformation 'Stu-00000000001','','00002','00001','N','25 Sep 2005','Lia'
CREATE           procedure ExamSeatPlan
(
	
	@ExamDate			Datetime,
	@RoomNo 			varchar(20),
	@StartTime			Varchar(50),
	@Shift				varchar(30),
	@ClassID			varchar(5),	
	@SectionID			varchar(5),
	@StartRoll			int,
	@EndRoll			int,
	@EntryBY			varchar(50),
	@EntryDate			DateTime
	
)

as



if exists (select * from ExamSitPlan where  Examdate=@ExamDate and RoomNo=@RoomNo and	StartTime=@StartTime	) 		


Update ExamSitPlan set  	


	ExamDate			=	@ExamDate			,
	RoomNo				=	@RoomNo 			,
	StartTime			=	@StartTime			,
	Shift				=	@Shift				,
	ClassID				=	@ClassID			,	
	SectionID			=	@SectionID			,
	StartRoll			=	@StartRoll			,
	EndRoll				=	@EndRoll			,
	EntryBY				=	@EntryBY			,
	EntryDate			=	@EntryDate			
	

where  Examdate=@ExamDate and RoomNo=@RoomNo and	StartTime=@StartTime		
else


 insert into ExamSitPlan

(

	ExamDate			,
	RoomNo				,
	StartTime			,
	Shift				,
	ClassID				,	
	SectionID			,
	StartRoll			,
	EndRoll				,
	EntryBY				,
	EntryDate						
	
			
		
	
)
values
(
	@ExamDate			,
	@RoomNo 			,
	@StartTime			,
	@Shift				,
	@ClassID			,
	@SectionID			,
	@StartRoll			,
	@EndRoll			,
	@EntryBY			,
	@EntryDate			
				
	
)

















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





--- Exec BookDisretInformation 'Stu-00000000001','','00002','00001','N','25 Sep 2005','Lia'
CREATE   procedure Exam_Type_Info_Save
(
	
    @mode               varchar(1),
    @groupID			varchar(5),
	@EtypeID			varchar(5),
	@ETypetitle	 		Varchar(50),
	@Note				varchar(50),
	@EntryBY			varchar(10),
	@EntryDate			DateTime
	
)

as

   if @mode='S' 
      begin      
          if  not exists (select Exam_code from Exam_setup where  Group_code=@groupID and Exam_code= @EtypeID ) 
             BEGIN	

               SET @EtypeID =(select isnull(max(CAST(Exam_code AS INT)),0)+1 from Exam_setup where  Group_code=@groupID )


                   IF @EtypeID IN ('1','2','3','4','5','6','7','8','9') 

                      SET @EtypeID ='0'+ @EtypeID 

         

        
                
                 insert into Exam_setup
                           (
							Group_code,
							Exam_code,
                            Exam_title,
							Remarks,
							Entry_by,
                            Entry_date 
	
						   )
				values
							(
							@groupID,
							@EtypeID,
							@ETypetitle	,
                            @Note, 
							@EntryBY,
							@EntryDate	
							)

        end       
    END 

   if @mode='U' 
      begin      
          if  exists (select Exam_code from Exam_setup where  Group_code=@groupID and Exam_code= @EtypeID ) 
          
                 update Exam_setup 
                        set Exam_title=@ETypetitle,
							Remarks=@Note
				 where  Group_code=@groupID and Exam_code= @EtypeID
				
        end       

    if @mode='D' 
      begin      
          if  exists (select Exam_code from Exam_setup where  Group_code=@groupID and Exam_code= @EtypeID ) 
         
                 delete from  Exam_setup where  Group_code=@groupID and Exam_code= @EtypeID
				
        end       

 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO















--- Exec BookDisretInformation 'Stu-00000000001','','00002','00001','N','25 Sep 2005','Lia'
CREATE           procedure ExamguardPlan1
(
	
	@ExamDate			Datetime,
	@RoomNo 			varchar(20),
	@TeacherID			varchar(5),
	@Responsibility			varchar(30),
	@StartTime			Varchar(50),
	@EntryBY			varchar(50),
	@EntryDate			DateTime
	
)

as



if exists (select * from ExamGuardPlan where  Examdate=@ExamDate and RoomNo=@RoomNo and StartTime=@StartTime) 		


Update ExamGuardPlan set  	


	ExamDate			=	@ExamDate			,
	RoomNo				=	@RoomNo 			,
	TeacherID			=	@TeacherID			,
	Responsibility			=	@Responsibility			,
	StartTime			=	@StartTime			,
	EntryBY				=	@EntryBY			,
	EntryDate			=	@EntryDate			
	

where  Examdate=@ExamDate and RoomNo=@RoomNo and StartTime=@StartTime	
else


 insert into ExamGuardPlan

(

	ExamDate			,
	RoomNo				,
	TeacherID			,
	Responsibility			,
	StartTime			,
	EntryBY				,
	EntryDate						
	
			
		
	
)
values
(
	@ExamDate			,
	@RoomNo 			,
	@TeacherID			,
	@Responsibility			,
	@StartTime			,
	@EntryBY			,
	@EntryDate			
				
	
)

















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






--- Exec ExaminatioSchedule 1,2000,'EX-01','25 Sep 2005','N','aaaaaaaa','Lia','25 Sep 2005'

CREATE procedure ExaminatioSchedule
(
    @mode               varchar(1),
	@serial_no          int,
    @ClassId            varchar(10), 
    @ExamYear	 		int,
	@ExamID				varchar(10),
    @ExamTypeID			varchar(10),
	@ExamDate			datetime,
	@MarksCataApplied		varchar(1),
    @sub_id             varchar(10),
    @ExamStartTime      datetime,
    @ExamEndTime        datetime,  
	@EntryBY			varchar(10)
)

as

declare @max_serial as int

if @mode='s' 
  begin
if exists(select serial_no from ExamSchedule 
      where ExamYear=  @ExamYear  and
            ExamTypeID = @ExamTypeID and
            ExamId=@ExamId   and
            ClassId=@ClassId
            )

begin
 Select  @max_serial=(select isnull(max(serial_no),0)+1 from examschedule 
            where ExamYear=  @ExamYear  and
            ExamTypeID = @ExamTypeID and
            ExamId=@ExamId   and
            ClassId=@ClassId)
   insert into ExamSchedule

(
    serial_no,
    ClassId,
   	ExamYear,
    ExamId,
	ExamTypeID			,
	ExamDate			,
	MarksCataApplied		,
    Sub_id,
	ExamStartTime	, 
    ExamEndTime,
	EntryBY				,
	EntryDate			
)
values
(
	@max_serial,
    @ClassId	,
	@ExamYear	,
    @ExamID,
	@ExamTypeID	,
	@ExamDate	,
	@MarksCataApplied,
    @sub_id, 
    @ExamStartTime,
    @ExamEndTime,
	@EntryBY	,	
	getdate()	
	
)
end          


if not exists(select serial_no from ExamSchedule 
      where ExamYear=  @ExamYear  and
            ExamTypeID = @ExamTypeID and
            ExamId=@ExamId   and
            ClassId=@ClassId
            )

begin

  		Select  @max_serial=1
		


 insert into ExamSchedule

(
    serial_no,
    ClassId,
   	ExamYear,
    ExamId,
	ExamTypeID			,
	ExamDate			,
	MarksCataApplied		,
    Sub_id,
	ExamStartTime	, 
    ExamEndTime,
	EntryBY				,
	EntryDate			
	
	
	
)
values
(
	@max_serial,
    @ClassId	,
	@ExamYear	,
    @ExamID,
	@ExamTypeID	,
	@ExamDate	,
	@MarksCataApplied,
    @sub_id, 
    @ExamStartTime,
    @ExamEndTime,
	@EntryBY	,	
	getdate()	
	
)
 end
end

if @mode='u' 
   begin
     Update ExamSchedule 
       set  	
		ExamDate	   =   @ExamDate	,
	    ExamStartTime  =   @ExamStartTime,
	    ExamEndTime    =   @ExamEndTime
		
where ExamYear=  @ExamYear  and
      ExamTypeID = @ExamTypeID and
      ExamId=@ExamId   and
      ClassId=@ClassId and
      serial_no  =@serial_no 

end 
if @mode='d'
  begin
    delete from ExamSchedule where ExamYear=  @ExamYear  and
            ExamTypeID = @ExamTypeID and
            ExamId=@ExamId   and
            ClassId=@ClassId and
            serial_no  =@serial_no  
  end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







--- Exec ExaminatioSchedule 1,2000,'EX-01','25 Sep 2005','N','aaaaaaaa','Lia','25 Sep 2005'
CREATE          procedure ExaminationRoutine
(
	
	
	@ExamYear	 		int,
	@ClassId			varchar(5),
	@ExamID				int,
	@ExamStartDate			datetime,
	@MarksCataApplied		varchar(1),
	@SubjectId			varchar(5),
	@CatagoryId			varchar(5),
	@TotalMarks			int,
	@ExamDate			datetime,
	@StartTime			varchar(50),
	@Note				varchar(80), 
	@EntryBY			varchar(10)
	
)

as


/*

if exists (select * from ExamRoutine where  ExamID= @ExamID ) 		

Update ExamRoutine set  	

	
			
	
	ExamYear		=		@ExamYear	 		,
	ClassId			=		@ClassId			,
	ExamID 			=		@ExamID				,
	StartDate		=		@ExamStartDate			,
	MarksCataApplied	=		@MarksCataApplied		,
	SubjectId		=		@SubjectId			,
	CategoryId		=		@CatagoryId			,
	TotalMarks		=		@TotalMarks			,
	ExamDate		=		@ExamDate			,
	ExamStartTime		=		@StartTime			,
	Note			=		@Note				, 
	EntryBY			=		@EntryBY			,
	EntryDate		=		getdate()	



where ExamID= @ExamID
else
*/

 insert into ExamRoutine

(

	
	ExamYear		,
	ClassId			,
	ExamID 			,
	StartDate		,
	MarksCataApplied	,
	SubjectId		,
	CategoryId		,
	TotalMarks		,
	ExamDate		,
	ExamStartTime		,
	Note			, 
	EntryBY			,
	EntryDate		
	
	
			
		
	
)
values
(
	@ExamYear	 		,
	@ClassId			,
	@ExamID				,
	@ExamStartDate			,
	@MarksCataApplied		,
	@SubjectId			,
	@CatagoryId			,
	@TotalMarks			,
	@ExamDate			,
	@StartTime			,
	@Note				, 
	@EntryBY			,
	getdate()	


	
)









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.Export_To_Acct_Vou    Script Date: 3/7/02 10:43:24 AM ******/
/****** Object:  Stored Procedure dbo.Export_To_Acct_Vou    Script Date: 10/28/01 4:34:22 PM ******/
/****** Object:  Stored Procedure dbo.Export_To_Acct_Vou    Script Date: 10/22/01 6:24:23 PM ******/
CREATE  procedure Export_To_Acct_Vou
as
if (@@Error=0)  
begin
begin transaction
insert into Ayman_Acct..Vou (vou_no,vou_date,vou_slno,cost_code,cust_ord_no,vou_dr,vou_cr,vou_amt,vou_type,uid,dt)
select vou_no,vou_date,vou_slno=vou_sl_no,cost_code,cust_ord_no,vou_dr,vou_cr,vou_amt,vou_type,uid,getdate() 
		from Ayman_PMIS..Vou Where Send='0'
	
	update Ayman_PMIS..Vou set Send='1'
commit transaction
end
else
begin
		Rollback Transaction
		update Ayman_PMIS..Vou set Send='0'
end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE   PROCEDURE FacultySchedule_Record_Delete
	@Emp_ID varchar(10),
	@DayoftheSchedule varchar(50),
	@ScheduleStTime varchar(10)
	
as 

If Exists (Select * From FacultySchedule 
    Where  emp_id=@Emp_Id and DayoftheSchedule=@DayoftheSchedule and ScheduleStTime=@ScheduleStTime )
	begin		
		Delete from  FacultySchedule 
			Where  emp_id=@Emp_Id and DayoftheSchedule=@DayoftheSchedule and ScheduleStTime=@ScheduleStTime
			
	end 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE   PROCEDURE FacultySchedule_Record_Save
	@Emp_ID varchar(10),
	@ScheduleMonth  varchar(20),
	@ScheduleYear varchar(6),
	@ScheduleStTime varchar(10),
	@ScheduleEndTime varchar(10),
	@DayoftheSchedule varchar(50),
	@Entry_Dt datetime,
	@Entry_By varchar(50)
	
as 
set nocount on
If Exists (Select * From FacultySchedule 
    Where  emp_id=@Emp_Id and DayoftheSchedule=@DayoftheSchedule )
	begin		
		Update FacultySchedule  Set 
        ScheduleStTime=@ScheduleStTime,
	ScheduleMonth=@ScheduleMonth ,
	ScheduleYear=@ScheduleYear ,
	ScheduleEndTime=@ScheduleEndTime

	Where  emp_id=@Emp_Id and DayoftheSchedule=@DayoftheSchedule 
	end
Else

Insert Into FacultySchedule (
	Emp_ID ,
	ScheduleMonth ,
	ScheduleYear ,
	ScheduleStTime ,
	ScheduleEndTime ,
	DayoftheSchedule,
	Entry_Dt ,
	Entry_By 
) Values (
	@Emp_ID ,
	@ScheduleMonth ,
	@ScheduleYear ,
	@ScheduleStTime ,
	@ScheduleEndTime ,
	@DayoftheSchedule,
	@Entry_Dt ,
	@Entry_By 
)

set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





create  procedure Fee_info_Save
(
		    @mode	   varchar(1),
			@Fee_code	varchar(10),
			@Fee_title  varchar(75),
			@EntryBy	varchar(10),
			@Entrydate	datetime
)
	AS


         if @mode='s'
               begin
		    			if not exists (select Fee_code from fee_info where Fee_code=@Fee_code )
		
		                insert into fee_info(Fee_code,Fee_title,Entry_by,Entry_date) 
		                            values(@Fee_code,@Fee_title,@EntryBy,@Entrydate)
		            
               end 



       if @mode='u'
               begin
		    			if  exists (select Fee_code from fee_info where Fee_code=@Fee_code )
		
		                 update fee_info set Fee_title=@Fee_title,
		                                     Entry_by=@EntryBy,
		                                     Entry_date=@Entrydate


                        where fee_code=@fee_code
   
                                               
                end 
 

       if @mode='d'
               begin
		    			if  exists (select Fee_code from fee_info where Fee_code=@Fee_code )
		
		                 delete from  fee_info  where fee_code=@fee_code
   
                                               
                end 
 
 







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE   procedure Fee_setup_Save
(
		    @mode	   varchar(1),
            @srl_no    integer,
			@class_id  varchar(5),
			@Fee_Code  varchar(10),
            @acc_Code  varchar(20),
            @fee_amt  varchar(20),
			@EntryBy	varchar(10),
			@Entrydate	datetime
)
	AS
        

         if @mode='s'
               begin
		    			if not exists (select srl_no from fee_setup where srl_no=@srl_no)
                         
                        set @Srl_No=(select isnull(max(Srl_No),0)+1 from fee_setup)
		
		                insert into fee_setup(Srl_No,Class_id,Fee_Code,Acc_code,Fee_amt,Entry_by,Entry_date) 
		                            values(@Srl_No,@Class_id,@Fee_code,@Acc_code,@Fee_amt,@EntryBy,@Entrydate)
		            
               end 


       if @mode='u'
               begin
		    			if  exists (select srl_no from fee_setup where srl_no=@srl_no)
		
		                 update fee_setup set 
                                             Fee_code=@Fee_code,
											 Class_id=@class_id, 
                                             Acc_code=@Acc_code,
                                             Fee_amt=@Fee_amt,
		                                     Entry_by=@EntryBy,
		                                     Entry_date=@Entrydate


                        where srl_no=@srl_no
   
                                               
                end 
 

       if @mode='d'
               begin
		    			if  exists (select srl_no from fee_setup where srl_no=@srl_no)
		
		                 delete from  fee_setup  where srl_no=@srl_no
   
                                               
                end 
 
 








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Fetch_Emp_Id    Script Date: 3/7/02 10:43:24 AM ******/
CREATE PROCEDURE Fetch_Emp_Id
@Mode VARCHAR(3),	
@ID_Prefix VARCHAR(3)	
AS
--------------------------------------------------------------------
/*  'NEW'----->Creats new ID having the specified prefix  */
/*  'All'----->Shows all existing IDs having the specified prefix */
--------------------------------------------------------------------
IF @Mode='NEW'
BEGIN
DECLARE @New_ID int
DECLARE @Latest_Id varchar(10)
	SELECT @New_ID= max(convert(int,emp_id))+1
	FROM Emp_per_Info WHERE emp_id LIKE @ID_Prefix+'%'
	SELECT @Latest_Id=convert(varchar(10),@New_ID)
	SELECT New_Id=
	CASE  
		WHEN len(@Latest_Id)= 5 then @Latest_Id
		WHEN  len(@Latest_Id)< 5 then  '0'+@Latest_Id
		WHEN  @ID_Prefix='0' AND @Latest_Id is null then @ID_Prefix+'101'
		WHEN  @Latest_Id is null then @ID_Prefix+'-001'
		
		
	END
END
IF @Mode='All'
BEGIN
	SELECT emp_id,Emp_Name=(Emp_fna+' '+Emp_mna+' '+Emp_lna)
	FROM Emp_per_Info WHERE emp_id LIKE @ID_Prefix+'%'
	ORDER BY emp_id
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Fixed_Head_POP    Script Date: 3/7/02 10:43:30 AM ******/
/****** Object:  Stored Procedure dbo.Fixed_Head_POP    Script Date: 10/28/01 4:34:23 PM ******/
/****** Object:  Stored Procedure dbo.Fixed_Head_POP    Script Date: 10/22/01 6:24:23 PM ******/
/****** Object:  Stored Procedure dbo.Fixed_Head_POP    Script Date: 9/17/00 4:33:57 PM ******/
/****** Object:  Stored Procedure dbo.Fixed_Head_POP    Script Date: 9/4/01 6:30:45 PM ******/
----------Fixed_Head_POP--------------
create procedure Fixed_Head_POP
as
	select distinct  Head_Name from Pay_Struc 
	where mode='F'




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Fixed_Pay_POP    Script Date: 3/7/02 10:43:43 AM ******/
/****** Object:  Stored Procedure dbo.Fixed_Pay_POP    Script Date: 10/28/01 4:34:38 PM ******/
/****** Object:  Stored Procedure dbo.Fixed_Pay_POP    Script Date: 10/22/01 6:24:37 PM ******/
/****** Object:  Stored Procedure dbo.Fixed_Pay_POP    Script Date: 9/17/00 4:34:10 PM ******/
/****** Object:  Stored Procedure dbo.Fixed_Pay_POP    Script Date: 9/4/01 6:30:55 PM ******/
----------Fixed_Pay_POP--------------
create procedure Fixed_Pay_POP
as
select Emp_id,Head_code,amount from fixed_Pay




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






CREATE  procedure Fixed_Variable_Gen_Deduc_Payroll
@Mode varchar (10),
@Emp_ID varchar (10)
as
if @Mode='Fixed_Add' 			----Fixed additional head
begin
	select   A.Head_code, A.Head_Name,B.Amount from Pay_Struc A,Fixed_pay B
	where A.Mode='F'and A.Operation='+' and B.Head_code = A.Head_code and B.Emp_id=@Emp_Id
	order by A.Head_code
end
---------------------------------------------------------------------------------------------------------------------------------------------------------
if @Mode='Var_Add' 			----Varirabale additional head
begin
	select Head_code,Head_Name from Pay_Struc
	where Mode='V'and Operation='+'
	order by Head_code
end
---------------------------------------------------------------------------------------------------------------------------------------------------------
if @Mode='Fixed_Ded' 			----Fixed additional head
begin
	select   A.Head_code, A.Head_Name,B.Amount from Pay_Struc A,Fixed_pay B
	where A.Mode='F'and A.Operation='-' and B.Head_code = A.Head_code and B.Emp_id=@Emp_Id
	order by A.Head_code
end
---------------------------------------------------------------------------------------------------------------------------------------------------------
if @Mode='Var_Ded' 			----Varirabale deductionl head
begin
	select Head_code, Head_Name from Pay_Struc
	where Mode='V'and Operation='-'
	order by Head_code
end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



create Procedure Fx
AS
select Name from dbo.sysobjects  where 
type='FN' and xtype='FN'





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE Proc Get_All_from_OT_Fix

AS

Set nocount on

select 
a.Emp_Id
,Emp_Nm=(select a.emp_fna+' '+a.emp_mna+' '+a.emp_lna)
,b.Emp_Desig,b.Emp_Dept
,c.Gen_Rate,c.Out_Rate,c.Hol_Rate

from emp_per_info a ,emp_job_hist_current b, OT_Fix c
	where a.Emp_id=b.Emp_id
	and a.Emp_id=c.Emp_id

Set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







--Get_Attn_Summery '9-023','may','2002'  

CREATE    procedure Get_Attn_Summery  
@Emp_id varchar (35),
@pay_Month varchar(12),
@pay_Year varchar(4)
AS
declare @M_Days varchar(12)
, @Febru_days varchar(2)

select @Febru_days=dbo.GetFebruary_Days (convert(int,@pay_Year))

Set @M_Days=(CASE @pay_Month 
		when 'January' THEN '31'
		when 'February' THEN @Febru_days
		when 'March' THEN '31'
		when 'April' THEN '30'
		when 'May' THEN '31'
		when 'June' THEN '30'
		when 'July' THEN '31'
		when 'August' THEN '31'
		when 'September' THEN '30'
		when 'October' THEN '31'
		When 'November' THEN '30'
		When 'December' THEN '31'
	end)

declare @Tot_Hol  int
, @Tot_Weekend int
, @Tot_Leave int
, @Tot_Abs int
, @Late int
, @Late_Abs int
, @Abs_Deduc int
, @Tot_Att int
,@Total_Fixed_Pay money
---------------------------------------------------------
,@Late_Flag int
,@Late_Value int
set @Late_Flag =(select dbo.GetParamFlag(62))
set @Late_Value =(select dbo.GetParamValue(62))
---------------------------------------------------------

set @Tot_Hol = (select count(*) from hol_list where 
(datepart(year,Str_date)=@pay_Year and datename(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datename(month,End_date)=@Pay_month)
)

set @Tot_Weekend = (select count(*) from hol_list where Category ='Weekend'and
(datepart(year,Str_date)=@pay_Year and datename(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datename(month,End_date)=@Pay_month)
)


set @Tot_Leave = (select count(*) from Emp_Leave_Info where emp_Id=@Emp_Id and 
(datepart(year,Start_dt)=@pay_Year and datename(month,Start_dt)=@Pay_month)and
(datepart(year,End_dt)=@pay_Year and datename(month,End_dt)=@Pay_month)
)		


Set @Late = (Select count(*) from emp_Att_info where In_status='1'and emp_id=@Emp_id
		and datename(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@pay_Month)

Set @Tot_Att=(Select count (Emp_Id)from emp_att_info where Emp_ID=@Emp_id 
		and datepart(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@Pay_month)

set @Total_Fixed_Pay=(select sum(Amount)from fixed_pay where emp_Id=@emp_Id)


if @Late_Flag=1
	begin
		set @Late_Abs=(@Late/@Late_Value)
		Set @Tot_Abs=((@M_Days-@Tot_Hol)-(@Tot_Att+@Tot_Leave))+ @Late_Abs
	end

if @Late_Flag=0
	begin
		Set @Tot_Abs=((@M_Days-@Tot_Hol)-(@Tot_Att+@Tot_Leave))
	end






select distinct a.emp_id,Emp_fna=(a.Emp_fna + " " + a.Emp_mna +" "+ a.Emp_lna),
	b.emp_desig,b.emp_dept

	,Working_Days=(@M_Days-@Tot_Hol)
	,Weekend=@Tot_Weekend
	,Other_holidays=(@Tot_Hol-@Tot_Weekend)
        ,Present=@Tot_Att
	,Absent=(@M_Days-@Tot_Hol)-@Tot_Att-@Tot_Leave
	,Late=@Late
	,Late_Abs=@Late_Abs
	,Tot_Absent=@Tot_Abs 
	,Leave=@Tot_Leave
	,Tot_Fixed_Pay=@Total_Fixed_Pay
	,Abs_Deduct=(@Total_Fixed_Pay/convert(int,@M_Days))* @Tot_Abs


	from emp_Per_info a,Emp_Job_Hist_Current b,Emp_Att_info c where 
	a.Emp_ID=B.Emp_ID and a.Emp_ID=c.Emp_ID and a.Emp_ID=@Emp_id 
	and datepart(year,c.Entry_dt)=@pay_Year
	and datename(month,c.Entry_dt)=@Pay_month










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/*	Procedure for 8.0 server */

--Get_Columns 'payscale_main'

CREATE  PROCEDURE Get_Columns (
				 @table_name		nvarchar(384),
				 @table_owner		nvarchar(384) = null,
				 @table_qualifier	sysname = null,
				 @column_name		nvarchar(384) = null,
				 @ODBCVer			int = 2)
AS
	DECLARE @full_table_name	nvarchar(769)
	DECLARE @table_id int

	if @ODBCVer <> 3
		select @ODBCVer = 2
	if @column_name is null /*	If column name not supplied, match all */
		select @column_name = '%'
	if @table_qualifier is not null
	begin
		if db_name() <> @table_qualifier
		begin	/* If qualifier doesn't match current database */
			raiserror (15250, -1,-1)
			return
		end
	end
	if @table_name is null
	begin	/*	If table name not supplied, match all */
		select @table_name = '%'
	end
	if @table_owner is null
	begin	/* If unqualified table name */
		SELECT @full_table_name = quotename(@table_name)
	end
	else
	begin	/* Qualified table name */
		if @table_owner = ''
		begin	/* If empty owner name */
			SELECT @full_table_name = quotename(@table_owner)
		end
		else
		begin
			SELECT @full_table_name = quotename(@table_owner) +
				'.' + quotename(@table_name)
		end
	end

	
		/* this block is for the case where there IS pattern
			matching done on the table name */

		if @table_owner is null /*	If owner not supplied, match all */
			select @table_owner = '%'

		SELECT
			[DATABASE] = convert(sysname,DB_NAME()),
			TABLE_NAME = convert(sysname,o.name),
			COLUMN_NAME = convert(sysname,c.name),
			PK=case dbo.Get_pkeys (@table_name)			
			when convert(sysname,c.name) then ' (PK)' else '  ' end,
			DATA_TYPE =convert (sysname,case
				when t.xusertype > 255 then t.name
				else d.TYPE_NAME collate database_default
			end),			
			convert(int,case
				when type_name(d.ss_dtype) IN ('numeric','decimal') then	/* decimal/numeric types */
					OdbcPrec(c.xtype,c.length,c.xprec)+2
				else
					isnull(d.length, c.length)
			end) LENGTH,
			DEFAULT_VALUE = text,
			NULLABLE = convert(varchar(254),
				rtrim(substring('NO YES',(ColumnProperty (c.id, c.name, 'AllowsNull')*3)+1,3)))

		FROM
			sysobjects o,
			master.dbo.spt_datatype_info d,
			systypes t,
			syscolumns c
			LEFT OUTER JOIN syscomments m on c.cdefault = m.id
				AND m.colid = 1
		WHERE
			o.name like @table_name
			AND user_name(o.uid) like @table_owner
			AND o.id = c.id
			AND t.xtype = d.ss_dtype
			AND c.length = isnull(d.fixlen, c.length)
			AND (d.ODBCVer is null or d.ODBCVer = @ODBCVer)
			AND (o.type not in ('P', 'FN', 'TF', 'IF') OR (o.type in ('TF', 'IF') and c.number = 0))
			AND isnull(d.AUTO_INCREMENT,0) = isnull(ColumnProperty (c.id, c.name, 'IsIdentity'),0)
			AND c.xusertype = t.xusertype
			AND c.name like @column_name


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Get_Day_Consumed    Script Date: 3/7/02 10:43:24 AM ******/
CREATE procedure Get_Day_Consumed
@Emp_ID varchar(10),
@Leave_name varchar(25)
as
/*
select Day_Consumed=sum(datediff(day,a.Start_dt,a.End_dt)+1 )from Emp_leave_Info_odit a
where exists(select * from Approval_audit b where a.App_Id=b.App_Id)
	and a.Emp_id=@Emp_ID and a.Leave_name=@Leave_name
	    and datepart (year,a.Start_dt)=datepart (year,getdate())
                            and datepart (year,a.End_dt)=datepart (year,getdate())
*/
select Day_Consumed=sum(datediff(day,Start_dt,End_dt)+1 )from Emp_leave_Info a
where  Emp_id=@Emp_ID and Leave_name=@Leave_name
	    and datepart (year,Start_dt)=datepart (year,getdate())
                            and datepart (year,End_dt)=datepart (year,getdate())




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Get_Emp_Information    Script Date: 3/7/02 10:43:30 AM ******/
CREATE PROCEDURE Get_Emp_Information
@Mode varchar(5),
@Emp_id varchar(10)
AS
if @Mode ='All'
Begin
	select * from Emp_per_Info where Job_Stat !=9 
end
if @Mode ='Ind'
Begin
	select * from Emp_per_Info where Emp_id=@Emp_id
end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





Create   PROCEDURE Get_Employees

as
SELECT 
a.Emp_id,
NM=(b.Emp_fna+' '+b.Emp_mna+' '+b.Emp_lna), 
a.Emp_join_date,
a.Emp_rank,
a.Emp_desig,
a.Emp_branch,
a.Emp_dept,
a.Emp_section,
a.Emp_job_type,
a.Responsibility,
a.Super_ID,
P_type=a.Pay_type,
D_type=a.Duty_type


from  Emp_Job_Hist_Current a,Emp_Per_Info b
where a.Emp_id=b.Emp_id  
and b.Job_Stat=0





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.Get_Jobtitle    Script Date: 3/7/02 10:43:30 AM ******/
/****** Object:  Stored Procedure dbo.Get_Jobtitle    Script Date: 10/28/01 4:34:23 PM ******/
/****** Object:  Stored Procedure dbo.Get_Jobtitle    Script Date: 10/22/01 6:24:23 PM ******/
/****** Object:  Stored Procedure dbo.Get_Jobtitle    Script Date: 9/17/00 4:33:58 PM ******/
CREATE  procedure Get_Jobtitle
@Mode varchar(5)
as
if @mode='1'		---All Designations  from job_title table
begin
select title from job_title
end 
if @mode='2'		--- Designations for which payscale has not been defined yet
begin
	select a.title,a.Code from job_title a where a.code not in(select Desig_Code from payscale_main)
end 
if @mode='3'		--- Designations for which allowable loan has not been defined yet
begin
	select a.title,a.Code from job_title a where a.title not in(select emp_desig from PF_Ln_Policy)
end 





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





CREATE  procedure Get_New_Id 

as

if (select distinct Vacant=count(Emp_ID) from Emp_Job_End where Emp_ID not in 
			(select distinct Emp_ID from Emp_Per_Info))=0


Begin
		select Latest_Id =max(convert (int,Emp_ID))+1 from Emp_Per_Info
End

/*
else

begin
	select distinct Latest_Id =min(convert(int,Emp_ID)) from Emp_Job_End
		 where Emp_ID not in (select distinct Emp_ID from Emp_Per_Info)
end

*/


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/*
exec Get_New_Id_DCL '9'
*/

CREATE procedure Get_New_Id_DCL 

@Prefix varchar(2)

as

select Latest_Id=(dbo.GetNewID_DCL(@Prefix))



/*
if (select distinct Vacant=count(Emp_ID) from Emp_Job_End where Emp_ID not in 
			(select distinct Emp_ID from Emp_Per_Info))=0
Begin
		select Latest_Id =max(convert (int,Emp_ID))+1 from Emp_Per_Info
End
else
begin
	select distinct Latest_Id =min(convert(int,Emp_ID)) from Emp_Job_End
		 where Emp_ID not in (select distinct Emp_ID from Emp_Per_Info)
end
***/


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/*

Get_OT_Fix_Ind '1103'

*/
Create   PROCEDURE Get_OT_Days_Ind
@Emp_ID varchar(10)
,@pay_month varchar(12)
,@pay_Year varchar(4)
as

declare @Gen_OT money
	,@Out_OT  money
	,@Hol_OT  money

Select @Gen_OT=isnull(sum(Gen_OT_Duration),0)
	,@Out_OT=isnull(sum(Out_OT_Duration),0)
	,@Hol_OT=isnull(sum(Hol_OT_Duration),0) 
	from Overtime
		where Emp_id=@Emp_id
		and Pay_Month=@Pay_Month    
		and pay_Year=@pay_Year

SELECT 
a.Emp_Id
,Emp_Nm=(select a.emp_fna+' '+a.emp_mna+' '+a.emp_lna)
,b.Emp_Desig,b.Emp_Dept
,Gen_OT=@Gen_OT
,Out_OT=@Out_OT
,Hol_OT=@Hol_OT

from  Emp_Per_Info a,Emp_Job_Hist_Current b
where a.Emp_id=b.Emp_id  
and a.Job_Stat=0
and a.Emp_Id=@Emp_Id





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/*

Get_OT_Fix_Ind '1103'

*/
CREATE  PROCEDURE Get_OT_Fix_Ind
@Emp_ID varchar(10)
as

declare @Gen_Rate money
	,@Out_Rate  money
	,@Hol_Rate  money

Select @Gen_Rate=isnull(sum(Gen_Rate),0)
	,@Out_Rate=isnull(sum(Out_Rate),0)
	,@Hol_Rate=isnull(sum(Hol_Rate),0) 
	from OT_Fix
		where Emp_id=@Emp_id

SELECT 
a.Emp_Id
,Emp_Nm=(select a.emp_fna+' '+a.emp_mna+' '+a.emp_lna)
,b.Emp_Desig,b.Emp_Dept
,Gen_Rate=@Gen_Rate 
,Out_Rate=@Out_Rate
,Hol_Rate=@Hol_Rate

from  Emp_Per_Info a,Emp_Job_Hist_Current b
where a.Emp_id=b.Emp_id  
and a.Job_Stat=0
and a.Emp_Id=@Emp_Id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  Proc Get_Overtime
@Pay_Month varchar(12)
,@Pay_Year varchar(4)

AS


Set nocount on

select 
a.Emp_Id
,Emp_Nm=(select a.emp_fna+' '+a.emp_mna+' '+a.emp_lna)
,b.Emp_Desig,b.Emp_Dept
,c.Gen_OT_Duration
,c.Out_OT_Duration
,c.Hol_OT_Duration
from emp_per_info a ,emp_job_hist_current b, Overtime c
	where a.Emp_id=b.Emp_id
	and a.Emp_id=c.Emp_id
	and c.Pay_Month=@Pay_Month
	and c.Pay_Year=@Pay_Year

Set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Get_PF_SETUP_etc    Script Date: 3/7/02 10:43:44 AM ******/
/****** Object:  Stored Procedure dbo.Get_PF_SETUP_etc    Script Date: 10/28/01 4:34:38 PM ******/
/****** Object:  Stored Procedure dbo.Get_PF_SETUP_etc    Script Date: 10/22/01 6:24:38 PM ******/
---DROP PROCEDURE Get_PF_SETUP_etc
CREATE PROCEDURE Get_PF_SETUP_etc
@mode varchar(15),
@param1 varchar(25),
@param2 varchar(25)
AS
--------------------------------------------------------------------
IF @mode='lOAN_CAT'
BEGIN
	select Ln_Code,Ln_NM,Max_Cel,Min_Int_Rate from pf_loan
END
--------------------------------------------------------------------
IF @mode='lOAN_CAT_1'
BEGIN
	select Ln_Code,Ln_NM,Max_Cel,Min_Int_Rate from pf_loan where Ln_NM=@param1
END
--------------------------------------------------------------------
IF @mode='INST_SCM'
BEGIN
	SELECT Inst_Scm_Code,Inst_Scm_NM  FROM PF_INST_SCM
END
--------------------------------------------------------------------
IF @mode='INST_SCM_1'
BEGIN
	SELECT Inst_Scm_Code  FROM PF_INST_SCM WHERE Inst_Scm_NM=@param1
END
--------------------------------------------------------------------
IF @mode='LN_PLC'
BEGIN
	SELECT Emp_Desig,Max_Ln_Allow FROM PF_LN_POLICY
END
--------------------------------------------------------------------
IF @mode='LN_PLC_DSG'
BEGIN
	SELECT Emp_Desig,Max_Ln_Allow FROM PF_LN_POLICY WHERE Emp_Desig=(
		SELECT EMP_DESIG FROM EMP_JOB_HIST_CURRENT WHERE EMP_ID=@param1)
END
--------------------------------------------------------------------




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE procedure Get_Photo
@Emp_Id varchar(10)
AS
	select Location=pick_pic_path from emp_per_info where emp_id=@Emp_id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




--select title from rank


Create procedure Get_Rank_Category
@Mode varchar(5)
as
if @mode='1'		---All Designations  from job_title table
begin
select title from rank
end 
if @mode='2'		--- Rank for which payscale has not been defined yet
begin			--- Rank is fetched rather than Designation for BGMEA
	select a.title from rank a where a.title not in(select Desig_Code from payscale_main)
end 
if @mode='3'		--- Rank for which allowable loan has not been defined yet
begin
	select a.title,a.Code from job_title a where a.title not in(select emp_desig from PF_Ln_Policy)
end 






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




create Procedure Get_Server_Time
AS
Select Server_Time =getdate()




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/****** Object:  Stored Procedure dbo.Get_Worker_Time    Script Date: 3/7/02 10:43:31 AM ******/
/****** Object:  Stored Procedure dbo.Get_Worker_Time    Script Date: 10/28/01 4:34:23 PM ******/
/****** Object:  Stored Procedure dbo.Get_Worker_Time    Script Date: 10/22/01 6:24:24 PM ******/
CREATE  procedure Get_Worker_Time --'9-031'
@Emp_Id varchar(10)
as
declare @ID varchar(10)
set @ID=@Emp_ID
---------------------------------------------------------
declare @S_Dt datetime
declare @E_Dt datetime
-------------------------
select @S_Dt=start_dt,@E_Dt=End_Dt from group_Shift where Shift_Code=(
select Shift_Code from Group_Shift where Gr_Code=(
select Gr_Code from Worker_Group where emp_ID=@ID))
--select @S_Dt
--select @E_Dt
--------------------------------------------------------
select Shift_Start,Shift_End,Relax=[Delay],Abs_time from Shift where Shift_Code=(
select Shift_Code from Group_Shift where Gr_Code=(
select Gr_Code from Worker_Group where emp_ID=@ID))and 
datepart(day,getdate())>=datepart(day,@S_Dt) and
datepart(month,getdate())=datepart(month,@S_Dt)and
datepart(year,getdate())=datepart(year,@S_Dt)and 
datepart(day,getdate())<=datepart(day,@E_Dt) and
datepart(month,getdate())=datepart(month,@E_Dt)and
datepart(year,getdate())=datepart(year,@E_Dt)
 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE procedure Get_title_Code
@Mode varchar(10),
@Title varchar(30)
as
if @Mode='Code'
begin
select Code from job_title Where title=@Title
end
if @Mode='POP'
begin
select a.title, b.Sal_Scale,b.H_R,b.Medical,b.Conv,b.Tel from job_title a,payscale b where a.code=b.code	
end
-----------------------------------------------------------------------------------------------------------
if @Mode='Scle'
begin
select Scale_Code,Payscale=convert(char(5),SB)+'-'+space(1)+ convert(char(5),Incr)+ '-'+space(1)+ convert(char(5),EB) from payscale_Main 
where Desig_Code=(select code from job_Title where Title=@Title)
end


if @Mode='RankScle'
begin
declare @Emp_Id varchar(10)
set @Emp_Id=Rtrim(Ltrim(@Title))

select Scale_Code,
Scale_Code,Category=Desig_Code,Payscale=convert(char(5),SB)+'-'+space(1)+ convert(char(5),Incr)
+ '-'+space(1)+ convert(char(5),EB) + space(1)+'EB'+space(1)+ convert(char(5),EBIncr)+ '-'+
+convert(char(5),EBMax)from payscale_Main 
where Desig_Code=(select Emp_rank from Emp_job_Hist_current where Emp_Id=@Emp_Id)
end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Group_NM_I_U_D    Script Date: 3/7/02 10:43:31 AM ******/
/****** Object:  Stored Procedure dbo.Group_NM_I_U_D    Script Date: 10/28/01 4:34:24 PM ******/
/****** Object:  Stored Procedure dbo.Group_NM_I_U_D    Script Date: 10/22/01 6:24:24 PM ******/
/****** Object:  Stored Procedure dbo.Group_NM_I_U_D    Script Date: 9/17/00 4:33:58 PM ******/
--------------------------------
create procedure Group_NM_I_U_D
@opr varchar (1),
@Gr_Code varchar(5),
@Gr_Name varchar(15),
@U_id    varchar(10)
as
if @opr='I'
begin
	insert into Group_NM(Gr_Code,Gr_Name,U_id)
	values(@Gr_Code,@Gr_Name,@U_id)
end 
if @opr='U'
begin
	Update Group_NM Set Gr_Name=@Gr_Name,U_id=@U_id
	Where Gr_Code=@Gr_Code
end 
if @opr='D'
begin
	Delete from Group_NM
	Where Gr_Code=@Gr_Code
end 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Group_Shift_I_U_D    Script Date: 3/7/02 10:43:31 AM ******/
/****** Object:  Stored Procedure dbo.Group_Shift_I_U_D    Script Date: 10/28/01 4:34:24 PM ******/
/****** Object:  Stored Procedure dbo.Group_Shift_I_U_D    Script Date: 10/22/01 6:24:24 PM ******/
/****** Object:  Stored Procedure dbo.Group_Shift_I_U_D    Script Date: 9/17/00 4:33:58 PM ******/
----------------------------------------
CREATE procedure Group_Shift_I_U_D
@opr varchar (1),
@Gr_Code varchar (5),
@Shift_Code varchar (5),
@Start_Dt datetime,
@Stored_Start_Dt datetime,		---------------useful during editing start date
@End_Dt datetime,
@U_id   varchar(10)
as
if @opr='I'
begin
	insert into Group_Shift(Gr_Code,Shift_Code,Start_Dt,End_Dt,U_id)
	values(@Gr_Code,@Shift_Code,@Start_Dt,@End_Dt,@U_id)
end 
if @opr='U'
begin
	Update  Group_Shift Set Shift_Code=@Shift_Code,Start_Dt=@Start_DT,End_Dt=@End_Dt,U_id=@U_id
	Where Gr_Code=@Gr_Code and Start_Dt=@Stored_Start_Dt 
end 
if @opr='D'
begin
	Delete from Group_Shift
	Where Gr_Code=@Gr_Code and Start_Dt=@Start_DT
end 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Hol_List_I_U_D    Script Date: 3/7/02 10:43:31 AM ******/
/****** Object:  Stored Procedure dbo.Hol_List_I_U_D    Script Date: 10/28/01 4:34:24 PM ******/
/****** Object:  Stored Procedure dbo.Hol_List_I_U_D    Script Date: 10/22/01 6:24:24 PM ******/
CREATE PROCEDURE Hol_List_I_U_D
@opr varchar (1),
@Hol_name varchar (50),
@Hol_desc varchar (200),
@Str_date datetime,
@End_date datetime ,
@Category varchar (25),
@U_id varchar (10)
AS
-----+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
DECLARE  @counter int
DECLARE  @Duration int
SET @counter=0
SELECT @duration=DATEDIFF(DAY,@str_Date,@End_Date)
IF DATEPART(MONTH,@Str_Date)< DATEPART(MONTH,@End_Date) OR
	DATEPART(YEAR,@Str_Date)< DATEPART(YEAR,@End_Date)	
/*********************************************************************
	If the holiday starts this month and ends next month 
	(Example: Str_Date=2001-09-28 and End_Date=2001-10-01)
	a Single entry will then breaks up into several records.
	In this case (Example) it will create 05 records,one for 
	each day within the given date range.
***********************************************************************/
BEGIN
Start_Insert:
	IF @counter<=@Duration
		
	BEGIN
		DECLARE @Mod_Str_Dt DATETIME
 		
		SET @Mod_Str_Dt=DATEADD(DAY,@Counter,@Str_Date)
			-------------------------------------------------------------------------------------	
			IF @opr='I'		--------INSERT----------
			BEGIN
				INSERT INTO  Hol_List (Hol_name,Hol_desc,Str_date,End_date,Category,U_id ) 
				VALUES (@Hol_name,@Hol_desc,@Mod_Str_Dt,@Mod_Str_Dt,@Category,@U_id)
			END
                    
			IF @opr='U'		--------UPDATE----------
			BEGIN
				UPDATE  Hol_List  SET 
					Hol_name=@Hol_name,Hol_desc=@Hol_desc,Str_date=@Mod_Str_Dt,End_date=@Mod_Str_Dt,
					Category=@Category,U_id=@U_id 
	 			WHERE  Hol_name=@Hol_name
			END
			IF @opr='D'		--------DELETE----------
			BEGIN
				DELETE FROM  Hol_List WHERE Hol_name=@Hol_name
			END
			-------------------------------------------------------------------------------------
		SET @counter=@counter+1
	END
	IF @counter<=@duration GOTO Start_Insert
END
ELSE
BEGIN
	IF @opr='I'		--------INSERT----------
	BEGIN
		INSERT INTO Hol_List (Hol_name,Hol_desc,Str_date,End_date,Category,U_id ) 
		VALUES (@Hol_name,@Hol_desc,@Str_Date,@End_Date,@Category,@U_id)
	END
		
	IF @opr='U'		--------UPDATE----------
	BEGIN
		UPDATE Hol_List SET 
		Hol_name=@Hol_name,Hol_desc=@Hol_desc,Str_date=@Str_Date,End_date=@End_Date,
		Category=@Category,U_id=@U_id 
 		WHERE  Hol_name=@Hol_name
	END
		
	IF @opr='D'		---------DELETE---------
	BEGIN
		DELETE FROM Hol_List WHERE Hol_name=@Hol_name
	END
END
	
-----++++++++++++++++++++++++++++++++  Send_To_Table  ++++++++++++++++++++++++++++++++++++++++++
/*
	IF @opr='I'		--------INSERT----------
	BEGIN
		INSERT INTO  Hol_List (Hol_name,Hol_desc,Str_date,End_date,Category,U_id ) 
		VALUES (@Hol_name,@Hol_desc,@Mod_Str_Dt,@Mod_Str_Dt,@Category,@U_id)
	END
                    
	IF @opr='U'		--------UPDATE----------
	BEGIN
		UPDATE  Hol_List  SET 
			Hol_name=@Hol_name,Hol_desc=@Hol_desc,Str_date=@Mod_Str_Dt,End_date=@Mod_Str_Dt,
			Category=@Category,U_id=@U_id 
 		WHERE  Hol_name=@Hol_name
	END
	IF @opr='D'		--------DELETE----------
	BEGIN
		DELETE FROM  Hol_List WHERE Hol_name=@Hol_name
	END
*/
-----++++++++++++++++++++++++++++++++  Send_To_Table  ++++++++++++++++++++++++++++++++++++++++++




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



/*
	select Move_Out_Dt,Entry_Dt from Emp_Movement where emp_id='9-015'
	and datepart(month,Move_Out_Dt)=5 and datepart(year,Move_Out_Dt)=2002
	and datepart(day,Move_Out_Dt)=7

 HrInMovement '9-015','2002-05-7'

*/

create Procedure HrsInMovement 

@Emp_Id varchar(10)
,@Dt datetime

AS


Declare @Hr Float
--,@Emp_Id varchar(10)
--,@Dt datetime
,@Move_Out datetime
,@Move_Back datetime

,@Office_End varchar(8) 
,@MDt varchar(24)

set @Hr=0

---------------------------------------------------------------------------------

declare Movement_Cursor Cursor for

select Move_Out_Dt,Entry_Dt from Emp_Movement where emp_id=@emp_id
		and datepart(Year,Move_Out_Dt)=datepart(Year,@Dt)
			and datepart(month,Move_Out_Dt)=datepart(month,@Dt)
				and datepart(Day,Move_Out_Dt)=datepart(Day,@Dt)

---------------------------------------------------------------------------------

OPEN Movement_Cursor

	fetch next from Movement_Cursor into @Move_Out,@Move_Back
	
	while @@Fetch_Status=0

	BEGIN

		if datepart(day,@Move_Back) != datepart(day,@Move_Out) 		


			Begin

				set @Office_End=(Select dbo.GetOfficeEnd(@Emp_Id,@Move_Out))
							
				set @MDt=convert(char(4),datepart(Year,@Move_Out))+'-'+
					convert(char(4),datepart(Month,@Move_Out))+'-'+
					convert(char(4),datepart(day,@Move_Out))+ space(1)+ @Office_End
				
				set @Move_Back=(select convert(datetime,@MDt))

			end	
			
			Set @Hr=(Select convert(char(6),(ROUND(convert(float,DATEDIFF(n,@Move_Out,@Move_Back)* 0.01667),1))))
			print @Hr
			---Set @Hr=@Hr+(select DATEDIFF(n,@Move_Out,@Move_Back)* 0.01667)
			
			fetch next from Movement_Cursor into @Move_Out,@Move_Back

	END


CLOSE Movement_Cursor
DEALLOCATE Movement_Cursor



--END







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Import_From_Ayman_Acct    Script Date: 3/7/02 10:43:25 AM ******/
/****** Object:  Stored Procedure dbo.Import_From_Ayman_Acct    Script Date: 10/28/01 4:34:20 PM ******/
/****** Object:  Stored Procedure dbo.Import_From_Ayman_Acct    Script Date: 10/22/01 6:24:19 PM ******/
CREATE Procedure Import_From_Ayman_Acct
@Mode varchar(15)
as
-----------------------------------------------------------------------------------------
if @Mode='Cre_Acc_Bank'		/*** returns code for Credit Account (Cash at bank) ***/
Begin
	select Vou_Cr=Acc_Code,Acc_Name from Ayman_Acct..acct Where acc_head=
	(select bank from Ayman_Acct..Fixed_ac)
End
-----------------------------------------------------------------------------------------
if @Mode='Cre_Acc_Cash'		/*** returns code for Credit Account (Cash in hand) ***/
Begin
	select Vou_Cr=Acc_Code,Acc_Name from Ayman_Acct..acct Where Acc_Code=
	(select Cash from Ayman_Acct..Fixed_ac)
End
-----------------------------------------------------------------------------------------
if @Mode='Dr_Acc'		/*** returns code for Debit Account (Salary) ***/
Begin
	select Vou_Dr=dept_trad from Ayman_Acct..Fixed_Ac
End
-----------------------------------------------------------------------------------------
If @Mode='Cost_Code'
Begin
	select cost_code,cost_name from Ayman_Acct..Cost_Centre
End




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/*

IncomeTaxBrkup 100000,50000,150000,10,18,25,530000


*/

CREATE Procedure IncomeTaxBrkup
@Slab1 money
,@Slab2 money
,@Slab3 money
,@TxP2 float
,@TxP3 float
,@TxP4 float
,@Sal money
AS

set nocount on

--------------------------------------------
Create table #Tax(

Des varchar(60)
,Prcnt varchar(50)
,TaxAmount varchar(50)

)
--------------------------------------------
declare @Tax money
,@Tx1 money
,@Tx2 money
,@Tx3 money
,@Tx4 money

,@TotalIncome money

set @Tx1=0
set @TotalIncome =@Sal

insert into #Tax (Des,Prcnt,TaxAmount)
values('On the first Tk. '+ convert(char(12),@slab1),'Nil',' Tk.'+ convert(char(12),@Tx1))
-------------------------------------------

if @Sal>@Slab1
	BEGIN
		set @Sal=@Sal-@Slab1
	
		if @Sal>0
			BEGIN
				if @Sal<@Slab2
					begin
						set @Tx2=@Sal*(@TxP2/100)
						set @Tax=@Tx2
					end
				else
					begin
						set @Tx2=@Slab2*(@TxP2/100)
						set @Tax=@Tx2
					end
		
				set @Sal=@Sal-@Slab2
			END
	END


insert into #Tax (Des,Prcnt,TaxAmount)
values('On the next Tk. '+convert(char(12),@slab2),'@'+convert(char(12),@Txp2)+'%',' Tk.'+ convert(char(12),@Tx2))
----------------------------------------------------------------------------------------
 
if @Sal>0
	BEGIN
		if @Sal<@Slab3
			begin
				
				set @Tx3=(@Sal*(@TxP3/100))
				set @Tax=@Tax+@Tx3
			end
		else
			begin
				set @Tx3=(@Slab3*(@TxP3/100))
				set @Tax=@Tax+@Tx3
			end
		set @Sal=@Sal-@Slab3
	END


insert into #Tax (Des,Prcnt,TaxAmount)
values('On the next Tk. '+ convert(char(12),@slab3),'@'+convert(char(12),@Txp3)+'%',' Tk.'+ convert(char(12),@Tx3))
----------------------------------------------------------------------------------------

if @Sal>0
	BEGIN	
		set @Tx4=(@Sal*(@TxP4/100))
		set @Tax=@Tax+@Tx4
	END
insert into #Tax (Des,Prcnt,TaxAmount)
values('On the rest ---','@'+convert(char(12),@Txp4)+'%',' Tk.'+ convert(char(12),@Tx4))
----------------------------------------------------------------------------------------

insert into #Tax (Des,Prcnt,TaxAmount)
values('Tax on '+convert(char(12),@TotalIncome)+ '(total income)','',' Tk.'+ convert(char(12),@Tax))

--set @Tax=round(@Tax,0)
set @Tax=floor(@Tax)







select * from #Tax


set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE   Procedure Increment_History_I_U
@Emp_Id varchar(10)
,@Scale_Code int
,@Basic_Amount varchar(10)
,@Effect_Dt datetime
AS

Declare 
@SB int
,@Incr int
,@Num_Incr int

select @SB=SB,@Incr=Incr from Payscale_Main 
where Scale_Code=@Scale_Code
Set @Num_Incr=(convert(money,@Basic_Amount)-@SB)/@Incr

if @Num_Incr>=0
Begin 

	If Exists (Select * from Increment_History where 
		Emp_Id=@Emp_Id and Scale_Code=@Scale_Code 
		and Effect_Dt=@Effect_Dt)
	Begin
		Update Increment_History SET Num_Incr=@Num_Incr Where 
		Emp_Id=@Emp_Id and Scale_Code=@Scale_Code 
		and Effect_Dt=@Effect_Dt
	End

	Else

	Begin
		Insert Into Increment_History(Emp_Id,Scale_Code,Num_Incr,Effect_Dt)
		Values(@Emp_Id,@Scale_Code,@Num_Incr,@Effect_Dt)
	End

End







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





/*

select * from increment_history
select * from fixed_pay where emp_id='9-005'

delete from increment_history

delete from emp_salary

delete from payroll_main

*/

----Increment_History_I_U '9-026',4,'15600','2002-07-03'

CREATE   Procedure  Increment_History_New_I_U

@Emp_Id varchar(10)
,@Num_Incr int 
,@Effect_Dt datetime
,@U_Id varchar(10)

AS
set nocount on

Declare 
@Scale_Code int
,@SB int
,@Incr_Amount int
,@Num_Prev_Incr int
,@Num_Latest_Incr int
,@Amount money

------------------------------------------------------------
-----Get present payscale code
set @Scale_Code=(select Scale_Code from Emp_payscale_hist
			where Emp_Id=@Emp_Id and Track_Id= 
				(select max(Track_Id)from Emp_payscale_hist where emP_Id=@Emp_Id))

-----Get Starting Basic and Increment Amount of the present payscale
select @SB=SB,@Incr_Amount=Incr from Payscale_Main 
where Scale_Code=@Scale_Code
		
-----Get present Basic Pay
select @Amount=Amount from fixed_pay 
where  Emp_id=@Emp_id and Head_code='01' 


-----Calculation

----@Num_Prev_Incr---->Number of Increment getting at the moment
set @Num_Prev_Incr= convert(int,(@Amount-convert(money,@SB))/@Incr_Amount)

----@Num_Latest_Incr---->Number of total increment upto date
set @Num_Latest_Incr=@Num_Incr+@Num_Prev_Incr 

----@Amount-----> Becomes the current Basic including today's increment
set @Amount=@SB+(@Incr_Amount*@Num_Latest_Incr)

------------------------------------------------------------
-----Insert Number of today's increment 
if exists (select * from Increment_History where
	Emp_Id=@Emp_Id and Scale_Code=@Scale_Code and 
	CONVERT(CHAR(8),Effect_Dt,5)=CONVERT(CHAR(8),@Effect_Dt,5) )

BEGIN
	update Increment_History set Num_Incr=@Num_Incr where 
	Emp_Id=@Emp_Id and Scale_Code=@Scale_Code and 
	CONVERT(CHAR(8),Effect_Dt,5)=CONVERT(CHAR(8),@Effect_Dt,5) 
END
ELSE
BEGIN
	Insert Into Increment_History(Emp_Id,Scale_Code,Num_Incr,Effect_Dt)
	    		Values(@Emp_Id,@Scale_Code,@Num_Incr,@Effect_Dt)
END

------------------------------------------------------------
-----Delete current Basic Pay and others
---Delete from Fixed_Pay where Emp_Id=@Emp_Id

-----Insert Latest Basic Pay including today's increment 
/*
declare @VarAmount varchar(10)
set @VarAmount = (select convert(varchar(10),@Amount))
Exec Insert_Into_Fixed_Pay @Emp_Id,@Scale_Code,@VarAmount,@Effect_Dt,@U_Id
*/
------------------------------------------------------------

select message='Data saved successfully ! '


set nocount off







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE  Procedure Ins_Salary_Head
AS
insert into Pay_struc (Head_code,Head_name,Operation,Mode,U_id,Dt)values('01','Basic Pay','+','F','dsl','2002-03-02')
insert into Pay_struc (Head_code,Head_name,Operation,Mode,U_id,Dt)values('02','House Rent','+','F','dsl','2002-03-02')
insert into Pay_struc (Head_code,Head_name,Operation,Mode,U_id,Dt)values('03','Conveyance','+','F','dsl','2002-03-02')
insert into Pay_struc (Head_code,Head_name,Operation,Mode,U_id,Dt)values('04','Medical Allowance','+','F','dsl','2002-03-02')
insert into Pay_struc (Head_code,Head_name,Operation,Mode,U_id,Dt)values('05','Telephone Allowance','+','F','dsl','2002-03-02')
insert into Pay_struc (Head_code,Head_name,Operation,Mode,U_id,Dt)values('06','Others(+)','+','V','dsl','2002-03-02')
insert into Pay_struc (Head_code,Head_name,Operation,Mode,U_id,Dt)values('07','Absent Deduction','-','V','dsl','2002-03-02')
insert into Pay_struc (Head_code,Head_name,Operation,Mode,U_id,Dt)values('08','Punitive','-','V','dsl','2002-03-02')
insert into Pay_struc (Head_code,Head_name,Operation,Mode,U_id,Dt)values('09','Advance Deduction','-','V','dsl','2002-03-02')
insert into Pay_struc (Head_code,Head_name,Operation,Mode,U_id,Dt)values('10','Others(-)','-','V','dsl','2002-03-02')




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





CREATE     Procedure  Insert_Into_Fixed_Pay
		 @Emp_Id varchar(10)
		,@Amount money
		,@U_Id varchar(10)
		,@AccountNo varchar(15)
AS

Set nocount on 

declare  @Message varchar(100)


	if exists (Select * from fixed_Pay where emp_id=@emp_id )
	begin
		Update fixed_Pay set 
					Amount=@Amount,
					AccountNo=@AccountNo
			where Head_code='01' and emp_Id=@emp_Id
			set @Message='Data updated successfully !'		
	end 
	else
	begin	
		insert into Fixed_Pay
		values(@Emp_Id,'01',@Amount,@u_Id,getdate(),@AccountNo)
		set @Message='Data saved successfully !'

	end

select Message=@Message

set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE    procedure IssueRules(
	@ClassName VARCHAR(50),
	@MaxBook INT,
	@MaxDay INT,
	@Fine MONEY,
	@User VARCHAR (10))
AS
SET NOCOUNT ON
DECLARE @ClassCode VARCHAR (5)
SELECT @ClassCode = ClassID FROM ClassInfo WHERE     (ClassName = @ClassName)

IF NOT EXISTS (SELECT * FROM BookIssueRule WHERE     (ClassCode = @ClassCode))
	INSERT INTO BookIssueRule
                (ClassCode, MaxNumberOfBook, MaxDayOfUse, FineAmount, Remarks, IsRuEntryDate, ISRuEntryBy)
				VALUES     (@ClassCode,@MaxBook,@MaxDay,@Fine,NULL,GETDATE(),@User)
ELSE
	UPDATE  BookIssueRule SET              
			MaxNumberOfBook = @MaxBook, MaxDayOfUse = @MaxDay, FineAmount = @Fine
			WHERE     (ClassCode = @ClassCode)

SET NOCOUNT OFF









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Job_Hist_Future_POP    Script Date: 3/7/02 10:43:44 AM ******/
/****** Object:  Stored Procedure dbo.Job_Hist_Future_POP    Script Date: 10/28/01 4:34:39 PM ******/
/****** Object:  Stored Procedure dbo.Job_Hist_Future_POP    Script Date: 10/22/01 6:24:38 PM ******/
/****** Object:  Stored Procedure dbo.Job_Hist_Future_POP    Script Date: 9/17/00 4:34:11 PM ******/
/****** Object:  Stored Procedure dbo.Job_Hist_Future_POP    Script Date: 9/4/01 6:30:56 PM ******/
CREATE procedure Job_Hist_Future_POP
@emp_id varchar(10)
as
select Effect_date,Emp_Rank,Emp_desig,Emp_branch,Emp_dept,Emp_section,Emp_job_type 
from emp_job_hist_future   where emp_ID=@emp_id  and Effect_date> getdate()
/*	 and 
	datepart(day, Effect_date) >datepart(day, getdate()) and 
	datepart(month, Effect_date) >datepart(month, getdate()) and
	datepart(year, Effect_date) >datepart(year, getdate()) */




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Job_Title_I_U_D    Script Date: 3/7/02 10:43:32 AM ******/
/****** Object:  Stored Procedure dbo.Job_Title_I_U_D    Script Date: 10/28/01 4:34:24 PM ******/
/****** Object:  Stored Procedure dbo.Job_Title_I_U_D    Script Date: 10/22/01 6:24:25 PM ******/
/****** Object:  Stored Procedure dbo.Job_Title_I_U_D    Script Date: 9/17/00 4:33:59 PM ******/
/****** Object:  Stored Procedure dbo.Job_Title_I_U_D    Script Date: 9/4/01 6:30:46 PM ******/
CREATE PROCEDURE Job_Title_I_U_D
@opr varchar (1),
@Code varchar(3),
@Title varchar (25),
@Prev_Title varchar (25),
@Description varchar (50),
@U_id varchar (10)
AS
if @opr='I'
begin
insert into Job_Title (
Title,
Code,
Description,
U_id
) 
values (
@Title,
@Code,
@Description,
@U_id
)
end
                    
if @opr='U'
begin
update Job_Title set 
Title=@Title,
Code=@Code,
Description=@Description,
U_id=@U_id
where Title=@Prev_Title
end
    		
if @opr='D'
begin
delete from Job_Title where title=@Title
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Job_Type_I_U_D    Script Date: 3/7/02 10:43:32 AM ******/
/****** Object:  Stored Procedure dbo.Job_Type_I_U_D    Script Date: 10/28/01 4:34:25 PM ******/
/****** Object:  Stored Procedure dbo.Job_Type_I_U_D    Script Date: 10/22/01 6:24:25 PM ******/
/****** Object:  Stored Procedure dbo.Job_Type_I_U_D    Script Date: 9/17/00 4:33:59 PM ******/
/****** Object:  Stored Procedure dbo.Job_Type_I_U_D    Script Date: 9/4/01 6:30:46 PM ******/
--*****************************
CREATE PROCEDURE Job_Type_I_U_D
@opr varchar (1),
@Code varchar(3),
@Title varchar (25),
@Prev_Title varchar (25),
@Description varchar (50),
@U_id varchar (10)
AS
if @opr='I'
begin
insert into Job_Type(
Title,
Code,
Description,
U_id
) 
values (
@Title,
@Code,
@Description,
@U_id
)
end
                    
if @opr='U'
begin
update Job_Type set 
Title=@Title,
Code=@Code,
Description=@Description,
U_id=@U_id
where Title=@Prev_Title
end
    		
if @opr='D'
begin
delete from Job_Type where title=@Title
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Encrypted object is not transferable, and script can not be generated. ******/

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO









CREATE    procedure LS_PLAN_MASTER_Save
(
		    @mode	   varchar(1),
            @srl_no    integer,
			@Class_id	varchar(15),
			@Section_id  varchar(15),
			@Term_id	varchar(5),
            @exam_id	varchar(5),
			@Sub_id	    varchar(5)
)
	AS
     DECLARE @SRL_NO_local AS INTEGER  
 
         if @mode='S'
               begin
                       
		    			if not exists (select Srl_no from LS_PLAN_MASTER where Srl_no=@Srl_no )
		                SET @SRL_NO_local=(SELECT isnull(MAX(SRL_NO),0)+1 FROM LS_PLAN_MASTER)
		                insert into LS_PLAN_MASTER(SRL_NO,Class_id,Section_id,Term_id,exam_id,Sub_id) 
		                            values(@SRL_NO_local,@Class_id,@Section_id,@Term_id,@exam_id,@Sub_id)
		            
               end 
 

        if @mode='U'
               begin
                       
		    			if  exists (select Srl_no from LS_PLAN_MASTER where Srl_no=@Srl_no )
		                  UPDATE LS_PLAN_MASTER SET Class_id=@Class_id,
                                                     Section_id=@Section_id,
                                                     Term_id=@Term_id,
                                                     Sub_id=@Sub_id,exam_id=@exam_id
                          WHERE   SRL_NO=@SRL_NO
		            
               end 

        if @mode='D'
               begin
                       
		    			if  exists (select Srl_no from LS_PLAN_MASTER where Srl_no=@Srl_no )
		                 DELETE FROM  LS_PLAN_MASTER WHERE   SRL_NO=@SRL_NO
		            
               end 










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE     procedure LS_PLAN_TOPIC_Save
(
		    @mode	   varchar(1),
            @Srl_no    INTEGER,
            @TOPIC_SRL_NO integer,
			@Topic_title	varchar(200),
			----@LS_Week varchar(5),
			@Entry_by varchar(10),
			@Entry_date	datetime,
                        @font_indicator integer,
                        @aca_yr  varchar(10)
)

	AS

     DECLARE @TOPIC_SRL_NO_LOC AS INTEGER  
 
         if @mode='S'
               begin
                       
		    			---if not exists (select Srl_no from LS_PLAN_TOPIC where Srl_no=@Srl_no and Topic_srl=@TOPIC_SRL_NO )
		                SET @TOPIC_SRL_NO_LOC=(SELECT isnull(MAX(Topic_srl),0)+1 FROM LS_PLAN_TOPIC where  LS_Week=@LS_Week)
		                insert into LS_PLAN_TOPIC(SRL_NO,Topic_srl,Topic_title,LS_Week,Entry_by,Entry_date,font_indicator,AcademicYr) 
		                            values(@SRL_NO,@TOPIC_SRL_NO_LOC,@Topic_title,@LS_Week,@Entry_by,@Entry_date, @font_indicator,@aca_yr)
		            
               end 

         
          if @mode='U'
               begin
                       
		    			if  exists (select Srl_no from LS_PLAN_TOPIC where Srl_no=@Srl_no  and Topic_srl=@TOPIC_SRL_NO  and LS_Week=@LS_Week )
		                
		                UPDATE LS_PLAN_TOPIC  SET Topic_title=@Topic_title,LS_Week=@LS_Week,font_indicator=@font_indicator
                               where Srl_no=@Srl_no  and Topic_srl=@TOPIC_SRL_NO and LS_Week=@LS_Week
		                           
		            
               end 


        if @mode='D'
               begin
                       
		    			if  exists (select Srl_no from LS_PLAN_TOPIC where Srl_no=@Srl_no  and Topic_srl=@TOPIC_SRL_NO  and LS_Week=@LS_Week )
		                
		                DELETE FROM  LS_PLAN_TOPIC   where Srl_no=@Srl_no  and Topic_srl=@TOPIC_SRL_NO  and LS_Week=@LS_Week
		                           
		            
               end
 

       







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Leave_Info_I_U_D    Script Date: 3/7/02 10:43:32 AM ******/
/****** Object:  Stored Procedure dbo.Emp_Leave_Info_I_U_D    Script Date: 10/28/01 4:34:37 PM ******/
/****** Object:  Stored Procedure dbo.Emp_Leave_Info_I_U_D    Script Date: 10/22/01 6:24:36 PM ******/
/****** Object:  Stored Procedure dbo.Emp_Leave_Info_I_U_D    Script Date: 9/17/00 4:34:09 PM ******/
/****** Object:  Stored Procedure dbo.Emp_Leave_Info_I_U_D    Script Date: 9/4/01 6:30:55 PM ******/
/****** Object:  Stored Procedure dbo.Emp_Leave_Info_I_U_D    Script Date: 6/20/01 12:00:38 PM ******/
CREATE PROCEDURE Leave_Info_I_U_D
@opr  varchar(1),
@App_Id varchar(10),
@Emp_id varchar (10) ,
@Leave_name varchar (25) ,
@Start_dt datetime ,
@End_dt  datetime ,
@Des varchar (100)  ,
@Address  varchar (100) ,
@Tel1 varchar (20) ,
@App_Type varchar(1),
@U_id varchar (10) 
AS
if @opr='I'
begin
insert into Emp_Leave_Info (Emp_id,Leave_name ,Start_dt ,End_dt,Des,Address,Tel1,App_Type,U_id) 
values (@Emp_id,@Leave_name ,@Start_dt ,@End_dt,@Des,@Address,@Tel1,@App_Type,@U_id)
end
-------------------------------------------------------------------------------------
                    
if @opr='U'
begin
update Emp_Leave_Info set Leave_name=@Leave_name ,Start_dt=@Start_dt,End_dt=@End_dt,
	Des=@Des,Address=@Address,Tel1=@Tel1,@App_Type=App_Type,U_id=@U_id where App_Id=@emp_id
end
-------------------------------------------------------------------------------------
                   
if @opr='M'
begin
update Emp_Leave_Info set Start_dt=@Start_dt,End_dt=@End_dt,
	 U_id=@U_id where App_Id=@App_Id
end
-------------------------------------------------------------------------------------
    	
	
if @opr='D'
begin
delete from  Emp_Leave_Info  where App_Id=@Emp_id
delete from  Emp_Leave_Info_Odit  where App_Id=@Emp_id
delete from  Approval  where App_Id=@Emp_id
end
if @opr='A'
begin
delete from  Emp_Leave_Info  where App_Id=@App_Id
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Leave_List_I_U_D    Script Date: 3/7/02 10:43:32 AM ******/
CREATE PROCEDURE Leave_List_I_U_D
@opr varchar (1),
@Prev_leave_name varchar (25),
@Leave_name varchar (25),
@Duration varchar(3),
@Des varchar(100),
@U_id varchar (10)
AS
if @opr='I'
begin
insert into Leave_List (
Leave_name,
Duration,
Des ,
U_id
) 
values (
@Leave_name,
@Duration,
@Des ,
@U_id
)
end
                    
if @opr='U'
begin
update Leave_List set 
Leave_name=@Leave_name,
Duration=@Duration,
Des=@Des,
U_id=@U_id   
where Leave_name=@Prev_leave_name
end
    		
if @opr='D'
begin
delete from Leave_List  where leave_name= @leave_name
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Leave_validation    Script Date: 3/7/02 10:43:25 AM ******/
/****** Object:  Stored Procedure dbo.Leave_validation    Script Date: 10/28/01 4:34:25 PM ******/
/****** Object:  Stored Procedure dbo.Leave_validation    Script Date: 10/22/01 6:24:19 PM ******/
/****** Object:  Stored Procedure dbo.Leave_validation    Script Date: 9/17/00 4:34:00 PM ******/
/****** Object:  Stored Procedure dbo.Leave_validation    Script Date: 9/4/01 6:30:46 PM ******/
create procedure Leave_validation
@Emp_ID varchar(10)
as
declare @Result varchar(1)
if (select Emp_gender from emp_per_Info where emp_ID='002')='Female'and (select Emp_marital_st
from emp_per_Info  where emp_ID='003')='Married'
begin 
set @Result='Y'
select @Result
end
else
begin
set @Result='N'
select @Result
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE  procedure LectureInformation
(
	
	@LectureID		varchar(5),
	@LectureDsc		varchar(50),
	@ClassID		varchar(5),
	@SubjectID		varchar(5),
	@LectureDetail		text,
	@LecLessonPrepareBY	varchar(50),
	@LIOpenForStu		varchar(1),
	@EntryDate		DateTime,
	@EntryBY		varchar(50)
	
	
	
	
)

AS if exists (select * from LectureInfo where  ClassID = @ClassID  and SubjectID= @SubjectID and LectureID=@LectureID	)	
Update LectureInfo set

	
	LectureID		=	@LectureID,		
	LectureDsc		=	@LectureDsc,		
	ClassID			=	@ClassID,
	SubjectID		=	@SubjectID,
	LectureDetail		=	@LectureDetail,
	LecLessonPrepareBY	=	@LecLessonPrepareBY,
	LIOpenForStu		=	@LIOpenForStu,									
	EntryDate		=	@EntryDate,
	EntryBY			=	@EntryBY	
	
	where ClassID = @ClassID and SubjectID= @SubjectID and LectureID=@LectureID		
else
 insert into LectureInfo

(
	LectureID		,
	LectureDsc		,
	ClassID			,
	SubjectID		,
	LectureDetail		,
	LecLessonPrepareBY	,
	LIOpenForStu		,
	EntryDate		,
	EntryBY			
		
	
)
values
(
	@LectureID,		
	@LectureDsc,		
 	@ClassID,
	@SubjectID,
	@LectureDetail,
	@LecLessonPrepareBY,
	@LIOpenForStu,									
	@EntryDate,
	@EntryBY	
	
)





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE   procedure LibrarySubInformation(
	@SubCode	VARCHAR(5),
	@SubName	VARCHAR(80),
	@Remarks VARCHAR(80),
	@User VARCHAR (10))
AS
SET NOCOUNT ON

IF NOT EXISTS (SELECT * FROM LibrarySubjectInfo WHERE     (SubSubjectCode = @SubCode))
	INSERT INTO LibrarySubjectInfo
        (SubSubjectCode, SubSubjectName, SubNote, SubEntryDate, SubEntryBy)
		VALUES     (@SubCode,@SubName,@Remarks,GETDATE(),@User)
ELSE
	UPDATE    LibrarySubjectInfo SET              
		SubSubjectName = @SubName, SubNote = @Remarks
		WHERE     (SubSubjectCode = @SubCode)

SET NOCOUNT OFF









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  procedure Load_Place
as
	select Distinct  place from Emp_Movement	




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Load_leave_duration    Script Date: 3/7/02 10:43:32 AM ******/
/****** Object:  Stored Procedure dbo.Load_leave_duration    Script Date: 10/28/01 4:34:25 PM ******/
/****** Object:  Stored Procedure dbo.Load_leave_duration    Script Date: 10/22/01 6:24:25 PM ******/
/****** Object:  Stored Procedure dbo.Load_leave_duration    Script Date: 9/17/00 4:34:00 PM ******/
/****** Object:  Stored Procedure dbo.Load_leave_duration    Script Date: 9/4/01 6:30:40 PM ******/
/****** Object:  Stored Procedure dbo.Load_leave_duration    Script Date: 6/20/01 12:00:33 PM ******/
create procedure Load_leave_duration
@Leave_name varchar(10)
as
select * from leave_list
where leave_name=@leave_name




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



----exec Loan_Sanction_Information_Delete '9-023','01',2005-07-14'
CREATE   PROCEDURE Loan_Sanction_Information_Delete
	@Emp_id varchar(10),
	@Loan_Type  varchar(10),
	@SanctionDate datetime
as 
set nocount on
If Exists (Select * From LoanSanction_Info 
    Where  emp_id=@Emp_Id and Loan_Type=@Loan_Type and SanctionDate=@SanctionDate)
	begin		
		Delete from LoanSanction_Info
		Where  emp_id=@Emp_Id and Loan_Type=@Loan_Type and SanctionDate=@SanctionDate
	end


set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE     PROCEDURE Loan_Sanction_Information_Save
	@Emp_id varchar(10),
	@Loan_Type  varchar(10),
	@InstalmentNo int,
	@GracePeriod int,
	@SanctionAmount decimal(9,2),
	@InstallmentAmount decimal(9,2),
	@SanctionDate datetime,
	@Entry_By varchar(50),
	@Entry_Dt datetime,
	@IstInstallmentDatePrin datetime,
	@IstInstallmentDateInt datetime,
	@InstallmentAmountInt decimal(9,2),
	@InterestInstallmetNo int,
	@PaidInstallment int

as 
set nocount on
If Exists (Select * From LoanSanction_Info 
    Where  emp_id=@Emp_Id and Loan_Type=@Loan_Type and SanctionDate=@SanctionDate)
	begin		
		Update LoanSanction_Info  Set 
			InstalmentNo=@InstalmentNo ,
			GracePeriod=@GracePeriod ,
			SanctionAmount=@SanctionAmount ,
			InstallmentAmount=@InstallmentAmount,
			IstInstallmentDatePrin=@IstInstallmentDatePrin ,
			IstInstallmentDateInt=@IstInstallmentDateInt ,
			InstallmentAmountInt =@InstallmentAmountInt ,
			InterestInstallmetNo =@InterestInstallmetNo, 
			PaidInstallment=@PaidInstallment

		Where  emp_id=@Emp_Id and Loan_Type=@Loan_Type and SanctionDate=@SanctionDate
	end
Else

Insert Into LoanSanction_Info (
	Emp_id,
	Loan_Type ,
	InstalmentNo ,
	GracePeriod ,
	SanctionAmount ,
	InstallmentAmount ,
	SanctionDate ,
	Entry_By ,
	Entry_Dt,
	IstInstallmentDatePrin,
	IstInstallmentDateInt,
	InstallmentAmountInt ,
	InterestInstallmetNo ,
	PaidInstallment
) Values (
	@Emp_id,
	@Loan_Type ,
	@InstalmentNo ,
	@GracePeriod ,
	@SanctionAmount ,
	@InstallmentAmount ,
	@SanctionDate ,
	@Entry_By ,
	@Entry_Dt,
	@IstInstallmentDatePrin,
	@IstInstallmentDateInt,
	@InstallmentAmountInt ,
	@InterestInstallmetNo,
	@PaidInstallment
)

set nocount off








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE PROCEDURE Loan_Type_Delete
	@Loan_Type varchar(10)
as 
set nocount on
If Exists (Select * From Loan_Type 
    Where  Loan_Type=@Loan_Type )
	begin		
		Delete from  Loan_Type  
			where  Loan_Type=@Loan_Type 
	end
set nocount off



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE   PROCEDURE Loan_Type_Save
	@Loan_Type varchar(10),
	@Loan_Description varchar(150), 
	@Entry_By varchar(50),
	@Entry_Dt datetime,
	@IntRate decimal (9,2)
		
as 

If Exists (Select * From Loan_Type 
    Where  Loan_Type=@Loan_Type )
	begin		
		Update Loan_Type  Set 
			Loan_Description=@Loan_Description ,
			IntRate=@IntRate
		where  Loan_Type=@Loan_Type 
	end
Else

Insert Into Loan_Type (
	Loan_Type,
	Loan_Description ,
	Entry_By ,
	Entry_Dt ,
	IntRate
) Values (
	@Loan_Type,
	@Loan_Description ,
	@Entry_By ,
	@Entry_Dt,
	@IntRate 	
)





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.LogIn_LogOut_Modified    Script Date: 3/7/02 10:43:45 AM ******/
/****** Object:  Stored Procedure dbo.LogIn_LogOut_Modified    Script Date: 10/28/01 4:34:39 PM ******/
/****** Object:  Stored Procedure dbo.LogIn_LogOut_Modified    Script Date: 10/22/01 6:24:19 PM ******/
/****** Object:  Stored Procedure dbo.LogIn_LogOut_Modified    Script Date: 9/17/00 4:34:11 PM ******/
--select* from emp_att_info where emp_id ='002'
	
--drop  procedure LogIn_LogOut_Modified
--exec login_logout '003'
CREATE procedure LogIn_LogOut_Modified
@emp_id varchar(10)
as
	
	declare @Log_out varchar(10)
	declare @@Value datetime
if (select status=count(*) from emp_att_info where emp_id=@emp_id and 
		datepart(day,Entry_Dt)=datepart(day,getdate())and 
		datepart(month,Entry_Dt)=datepart(month,getdate())and 
		datepart(year,Entry_Dt)=datepart(year,getdate()))=1
begin
	select @@value=isnull(emp_logout,'1900-12-12')from emp_att_info
	where Emp_id =@emp_id and 
		datepart(day,Entry_Dt)=datepart(day,getdate())and 
		datepart(month,Entry_Dt)=datepart(month,getdate())and 
		datepart(year,Entry_Dt)=datepart(year,getdate()) 
	if @@value='1900-12-12'
		begin
			set  @Log_out='Not_LogOut'
		end
	else 
		begin
			set  @Log_out='Yes_LogOut'
		end		
	select Status=@Log_out
end
else
if (select status=count(*) from emp_att_info where emp_id=@emp_id and 
		datepart(day,Entry_Dt)=datepart(day,getdate())and 
		datepart(month,Entry_Dt)=datepart(month,getdate())and 
		datepart(year,Entry_Dt)=datepart(year,getdate()))=0
begin
	set  @Log_out='LogIn'
	
	select Status=@Log_out
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



create procedure Log_sub1_I_U
@opr varchar(5)
,@Access_Id int
,@Access_Area varchar(25)

AS


if @opr='In'

begin
	insert into Log_Sub1(Access_Id,Access_Area,Entry_Time)
	values(@Access_Id,@Access_Area,getdate())
end


if @opr='Out'

begin

	Update Log_Sub1 set Exit_Time=getdate()
	where track_Id=(select max(Track_Id)from Log_Sub1) 
		
end



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Login_logout    Script Date: 3/7/02 10:43:44 AM ******/
/****** Object:  Stored Procedure dbo.Login_logout    Script Date: 10/28/01 4:34:39 PM ******/
/****** Object:  Stored Procedure dbo.Login_logout    Script Date: 10/22/01 6:24:19 PM ******/
/****** Object:  Stored Procedure dbo.Login_logout    Script Date: 9/17/00 4:34:11 PM ******/
/****** Object:  Stored Procedure dbo.Login_logout    Script Date: 9/4/01 6:30:56 PM ******/
/****** Object:  Stored Procedure dbo.Login_logout    Script Date: 6/20/01 12:00:39 PM ******/
/****** Object:  Stored Procedure dbo.Login_logout    Script Date: 5/14/01 9:49:53 AM ******/
/****** Object:  Stored Procedure dbo.Login_logout    Script Date: 1/1/99 3:01:44 AM ******/
/****** Object:  Stored Procedure dbo.Login_logout    Script Date: 5/9/01 8:13:09 AM ******/
/****** Object:  Stored Procedure dbo.Login_logout    Script Date: 5/8/01 4:09:48 AM ******/
CREATE PROCEDURE Login_logout
@Emp_id varchar (10) 
as
select status=count(*) from emp_att_info where emp_id=@Emp_id and 
cast(Year(Entry_dt) as varchar(5)) + '-' + 
cast(month(Entry_dt) as varchar(2)) + '-' + 
cast(day(Entry_dt) as varchar(2)) =
cast(Year(getdate()) as varchar(5)) + '-' + 
cast(month(getdate()) as varchar(2)) + '-' + 
cast(day(getdate()) as varchar(2))




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




Create procedure Markscategorydes
(
	@McategoryID	varchar(5),
	@mcategoryDsc	varchar(80),
	@Note		varchar(80),
	@EntryBy	varchar(10),
	@Entrydate	datetime
)

AS if exists (select * from Markscategory where McategoryId = @McategoryId )
Update Markscategory set

	McategoryID 	= 	@McategoryID,
	mcategoryDsc	=	@mcategoryDsc,
	Note		=	@Note,
	EntryBy		=	@EntryBy,
	Entrydate	=	@Entrydate


	where McategoryId = @McategoryId 

else
 insert into Markscategory 

(
	McategoryID,
	mcategoryDsc,
	Note,
	EntryBy,
	Entrydate
)
values
(
	@McategoryID,
	@mcategoryDsc,
	@Note,
	@EntryBy,
	@Entrydate
)




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Match_U_ID_N_POP    Script Date: 3/7/02 10:43:51 AM ******/
/****** Object:  Stored Procedure dbo.Match_U_ID_N_POP    Script Date: 10/28/01 4:34:48 PM ******/
/****** Object:  Stored Procedure dbo.Match_U_ID_N_POP    Script Date: 10/22/01 6:24:46 PM ******/
/****** Object:  Stored Procedure dbo.Match_U_ID_N_POP    Script Date: 9/17/00 4:34:19 PM ******/
/****** Object:  Stored Procedure dbo.Match_U_ID_N_POP    Script Date: 9/4/01 6:31:02 PM ******/
CREATE procedure Match_U_ID_N_POP
@U_ID varchar(10)
as
declare @Pos varchar (1)
declare @S_ID varchar (10)
declare @Chk varchar (1)
select @Pos=Move_Pos from Approval
if @Pos=1
begin
select	@S_ID=Tier_1,@Chk=Tier_1_Chk  from Approval
end
if @Pos=2
begin
select 	@S_ID=Tier_2,@Chk=Tier_2_Chk from Approval
end
if @Pos=3
begin
select 	@S_ID=Tier_3,@Chk=Tier_3_Chk from Approval
end
if @Pos=4
begin
select 	@S_ID=Tier_4,@Chk=Tier_4_Chk from Approval
end
if @Pos=5
begin
select 	@S_ID=Final_tier,@Chk=Final_tier_Chk from Approval
end
if @s_ID=@U_ID
begin
exec  Pop_Approval
end
else
select App_ID=0




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Maternity_Leave_Validation    Script Date: 3/7/02 10:43:25 AM ******/
/****** Object:  Stored Procedure dbo.Maternity_Leave_Validation    Script Date: 10/28/01 4:34:26 PM ******/
/****** Object:  Stored Procedure dbo.Maternity_Leave_Validation    Script Date: 10/22/01 6:24:19 PM ******/
/****** Object:  Stored Procedure dbo.Maternity_Leave_Validation    Script Date: 9/17/00 4:34:00 PM ******/
/****** Object:  Stored Procedure dbo.Maternity_Leave_Validation    Script Date: 9/4/01 6:30:46 PM ******/
create procedure Maternity_Leave_Validation
@Emp_ID varchar(10)
as
declare @Result varchar(1)
if (select Emp_gender from emp_per_Info where emp_ID=@Emp_ID)='Female'and (select Emp_marital_st
from emp_per_Info  where emp_ID=@Emp_ID)='Married'
begin 
	set @Result='1'
end
else
begin
	set @Result='0'
end
select Validity=@Result




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/*
	Monthly_Attendance_Backup 'May','2002'
	
	select * from Emp_Att_Info_Audit

*/
CREATE  Procedure Monthly_Attendance_Backup
@Attn_Month varchar(12)
,@Attn_year Varchar(4)
AS

set nocount on
declare @Message varchar(150)
------------------------------------------------------------------------------------------
	delete From Emp_Att_Info_Audit

	insert into Emp_Att_Info_Audit
	select * from Emp_Att_Info where 
		datename(month,Entry_Dt)=@Attn_Month 
		and datepart(year,Entry_Dt)=@Attn_year
------------------------------------------------------------------------------------------
	delete From Emp_Leave_Info_Audit

	insert into Emp_Leave_Info_Audit
	select * from Emp_Leave_Info where 
		datename(month,Start_dt)=@Attn_Month
		and datename(month,End_dt)=@Attn_Month
		and datename(year,Start_dt)=@Attn_year 
		and datename(year,End_dt)=@Attn_year
------------------------------------------------------------------------------------------
	SET @Message='Process going on please wait........'


select Message=@Message

set nocount off



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/*
--exec Monthly_Attn_Summery 'September','2002','9-001'

exec Monthly_Attn_Summery 'December','2003','1104'

select * from emp_Att_notes

*/

CREATE       Procedure Monthly_Attn_Summery

@Pay_Month varchar(12)
,@Pay_Year int
,@Emp_id varchar(10)

AS

set nocount on
---------------------------
declare @M_Days int
,@Days int
,@Year int 	
,@Lp_Count int
--,@Pay_Month varchar(12)
--,@Pay_Year int
---------------------------
--,@Emp_id varchar(10)

,@Emp_Nm varchar(50)
,@Emp_Desig varchar(50)
,@Emp_Dept varchar(50)
,@Date varchar(20)
,@Emp_login varchar(8)
,@In_Status varchar(4)
,@Notes  varchar(50)
,@Emp_logout varchar(8)
,@Category varchar(40)
---------------------------
,@Leave int
,@Hol_Day int
,@Weekend int
,@Work_Day int
---------------------------Temp.Table
create table #Emp_Attn
(Emp_id varchar(10)null
,Emp_Nm varchar(50)
,Emp_Desig varchar(50)
,Emp_Dept varchar(50)
,[Date]varchar(20)
,Emp_login varchar(8)null  
,In_Status varchar(4)null   
,Notes  varchar(50)null   
,Emp_logout varchar(8)null
,Leave int
,Hol_Day int
,Weekend int
,Work_Day int
,Late int
,Attn int
,Absent int
,Out_Office int
,Month_Year varchar(20)
)

set @Lp_Count=1
--set @Pay_Month='January'
--set @Pay_Year='2003'
--set @Emp_id='9-001'
---------------------------------------Find LeapYear---------------
set @Year=convert(int,@Pay_Year)

	if @Year % 4 = 0 
		Begin
			if @Year % 100 <> 0 or @Year % 400 = 0 
				set @Days = 29
			else
				set @Days = 28
		End	
	else
	set @Days = 28
---------------------------------------Days of the month-----------
Set @M_Days=(CASE @pay_Month 
		when 'January' THEN 31 
		when 'February' THEN @Days 
		when 'March' THEN 31	
		when 'April' THEN 30
		when 'May' THEN 31 
		when 'June' THEN 30 
		when 'July' THEN 31 
		when 'August' THEN 31
		when 'September' THEN 30 
		when 'October' THEN 31 
		When 'November' THEN 30 
		When 'December' THEN 31
		end)
------------------------------------------------------------------------------
select @Emp_Nm=(select emp_fna+' '+emp_mna+' '+emp_lna from emp_per_info where emp_id=@emp_id) 
,@Emp_Desig=(select emp_desig from emp_job_hist_current where emp_id=@emp_id)  
,@Emp_Dept=(select emp_dept from emp_job_hist_current where emp_id=@emp_id)  
,@Leave=dbo.GetLeaveDuration(@Emp_ID,@Pay_Month,@Pay_Year)
,@Hol_Day = (select dbo.GetTot_Holday(@Pay_month,@pay_Year))
,@Weekend = (select count(*) from hol_list where Category ='Weekend'and
(datepart(year,Str_date)=@pay_Year and datename(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datename(month,End_date)=@Pay_month))
,@Work_Day=(@M_Days-@Hol_Day)

------------------------------------------------------------------------------

set @Emp_login=''
set @Emp_logout=''
set @In_Status=''
set @Notes=''

while @Lp_Count<=@M_Days

begin

insert into #Emp_Attn
values ('',@Emp_Nm,@Emp_Desig,@Emp_Dept,convert(varchar(2),@Lp_Count),'','','',''
,'','','','','','','','','')


select @Emp_Id=Emp_Id,@Emp_login=(convert(varchar(8),Emp_login,8))
	,@In_Status=case In_Status when 0 then ' 'when 1 then 'Late'end
	,@Notes='',@Emp_logout=convert(varchar(8),Emp_logout,8) 
	from emp_att_info where emp_id=@Emp_id
	and datepart(day,emp_login)=@Lp_Count
	and datepart(year,emp_login)=@pay_year 
	and datename(month,emp_login)=@Pay_month

----------Late Notes-----------------------------------------------
set @Notes=(select Reason+'  ('+[Description]+')' from emp_Att_notes where Attn_ID=
			(select Attn_ID from emp_Att_Info where emp_Id=@Emp_ID
				and datepart(day,emp_login)=@Lp_Count
				and datepart(year,emp_login)=@pay_year 
				and datename(month,emp_login)=@Pay_month))

----------Out of Office Notes-----------------------------------------------
set @Notes=(select [Description] from Out_of_Office where Emp_ID=@Emp_Id
				and datepart(day,Str_Date)=@Lp_Count
				and datepart(year,Str_Date)=@pay_year 
				and datename(month,Str_Date)=@Pay_month)

--------------------------------------------------------------------

update #Emp_Attn set emp_id=@Emp_id,Emp_login=@Emp_login
		,In_Status=@In_Status
		,Notes=@Notes,Emp_logout=@Emp_logout 
	where rtrim(left([Date],2))=@Lp_Count

---------------------Weekend or Holiday---------------------------------
if exists (select * from hol_list 
	where @Lp_Count between datepart(day,Str_date)and datepart(day,End_date)
	and datepart(year,Str_date)=@pay_year 
	and datename(month,Str_date)=@Pay_month)

update #Emp_Attn set Notes=(select top 1 Category= 
		case Category when 'Weekend'then Category
			      when 'Hartal'then Category
			      else Hol_Name end
		from hol_list where datepart(day,Str_date)=@Lp_Count
		and datepart(year,Str_date)=@pay_year 
		and datename(month,Str_date)=@Pay_month)
		,In_Status=''	
	where rtrim(left([Date],2))=@Lp_Count

----------------------Leave---------------------------------------------
if exists (select * from Emp_Leave_Info 
	where @Lp_Count between datepart(day,Start_dt)and datepart(day,End_dt)
	and datepart(year,Start_dt)=@pay_year 
	and datename(month,Start_dt)=@Pay_month and emp_id=@Emp_id)

update #Emp_Attn set Notes='Leave'where rtrim(left([Date],2))=@Lp_Count


------------------------------Unauthorized leave-----------------------
---Modified on January 17 2004 by Shameem Ferdous
---Absent 
if (select  count(*) from Emp_Leave_Info 
			where @Lp_Count between datepart(day,Start_dt)and datepart(day,End_dt)
				and datepart(year,Start_dt)=@pay_year 
				and datename(month,Start_dt)=@Pay_month and emp_id=@Emp_id)=0

if  (select  count(*) from emp_att_info where emp_id=@Emp_id
	and datepart(day,emp_login)=@Lp_Count
	and datepart(year,emp_login)=@pay_year 
	and datename(month,emp_login)=@Pay_month)=0

and (select   count(*) from Hol_List where @Lp_Count between datepart(day,Str_date)
	and datepart(day,End_date)
	and datepart(year,Str_date)=@pay_year 
	and datename(month,Str_date)=@Pay_month)=0

update #Emp_Attn set Notes='Absent'where rtrim(left([Date],2))=@Lp_Count


------------------------------Out of office----------------------------

if exists (select * from Out_of_Office 
	where @Lp_Count between datepart(day,Str_date)and datepart(day,End_date)
	and datepart(year,Str_date)=@pay_year 
	and datename(month,Str_date)=@Pay_month and emp_id=@Emp_id)

update #Emp_Attn set Notes='Out of Office '+'('+ @Notes +')' where rtrim(left([Date],2))=@Lp_Count


----------------------------------------------------------------------
set @Emp_login=''
set @Emp_logout=''
set @In_Status=''
set @Notes=''

	set @Lp_Count=@Lp_Count+1
end


update #Emp_Attn set Notes='' 
where rtrim(left([Date],2))>datepart(day,getdate())
and	datename(month,getdate())= @Pay_month
and datename(year,getdate())=@Pay_Year
----------------------------------------------------------------------
update #Emp_Attn set Leave=@Leave,Hol_Day=@Hol_Day
,Weekend=@Weekend,Work_Day=@Work_Day
,Late=(select  count(In_Status)from #Emp_Attn where In_Status='Late')
,Attn =(select  count(Emp_Login)from #Emp_Attn where Emp_Login !='')
,Absent=(select count(Notes)from #Emp_Attn where Notes='Absent')
,Out_Office=(select  count(Notes)from #Emp_Attn where rtrim(ltrim(left(Notes,13)))='Out of Office')
,Month_Year=@Pay_month+', '+convert(varchar(4),@Pay_Year)
----------------------------------------------------------------------


select * from  #Emp_Attn

drop table #Emp_Attn


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Move_Sch_For_Move_Out    Script Date: 3/7/02 10:43:33 AM ******/
/****** Object:  Stored Procedure dbo.Move_Sch_For_Move_Out    Script Date: 10/28/01 4:34:26 PM ******/
/****** Object:  Stored Procedure dbo.Move_Sch_For_Move_Out    Script Date: 10/22/01 6:24:26 PM ******/
/****** Object:  Stored Procedure dbo.Move_Sch_For_Move_Out    Script Date: 9/17/00 4:34:01 PM ******/
/****** Object:  Stored Procedure dbo.Move_Sch_For_Move_Out    Script Date: 9/4/01 6:30:47 PM ******/
/****** Object:  Stored Procedure dbo.Move_Sch_For_Move_Out    Script Date: 6/20/01 12:00:32 PM ******/
CREATE procedure Move_Sch_For_Move_Out
@Emp_id  varchar (10)
as
Select
	Mode,
	Place,
	Move_Out_Dt,
	Exp_Rtn_Dt,
	Cont_Tel,
	Move_des 
	from Move_Schedule
	Where Emp_id=@Emp_id 
	and  datepart(day,Move_Out_Dt)=datepart(day,getdate())   ----a pecularity is found here 
	and datepart(month,Move_Out_Dt)=datepart(month,getdate()) 
	and datepart(year,Move_Out_Dt)=datepart(year,getdate()) 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Move_Viewer    Script Date: 3/7/02 10:43:45 AM ******/
/****** Object:  Stored Procedure dbo.Move_Viewer    Script Date: 10/28/01 4:34:39 PM ******/
/****** Object:  Stored Procedure dbo.Move_Viewer    Script Date: 10/22/01 6:24:38 PM ******/
/****** Object:  Stored Procedure dbo.Move_Viewer    Script Date: 9/17/00 4:34:11 PM ******/
/****** Object:  Stored Procedure dbo.Move_Viewer    Script Date: 9/4/01 6:30:47 PM ******/
/****** Object:  Stored Procedure dbo.Move_Viewer    Script Date: 6/20/01 12:00:40 PM ******/
CREATE procedure   [Move_Viewer]
@View_mode varchar(10)
as
if @View_mode='Today'
begin
	select	
		A.Emp_id,
		A.Mode,
		A.Place,
		A.Move_Out_Dt,
		A.Exp_Rtn_Dt,
		A.Cont_Tel ,
		A.Move_des ,
		NM=(B.Emp_Fna+' '+B.Emp_Mna+' '+B.Emp_Lna),
		C.Emp_Section
		From Move_Out_In A, Emp_per_info B,Emp_job_hist_current C
		Where A.Emp_id=B.Emp_id and A.Emp_id=C.Emp_id 
		and datepart(day,A.Move_out_Dt)=DATEPART(day, GETDATE()) 
		and datepart(month,A.Move_out_Dt)=DATEPART(month, GETDATE())
		and datepart(year,A.Move_out_Dt)=DATEPART(year, GETDATE())
end
if @View_mode='Yesterday'
begin
	select	
		A.Emp_id,
		A.Mode,
		A.Place,
		A.Move_Out_Dt,
		A.Exp_Rtn_Dt,
		A.Cont_Tel ,
		A.Move_des ,
		NM=(B.Emp_Fna+' '+B.Emp_Mna+' '+B.Emp_Lna),
		C.Emp_Section
		From Move_Out_In A, Emp_per_info B,Emp_job_hist_current C
		Where A.Emp_id=B.Emp_id and A.Emp_id=C.Emp_id 
		and datepart(day,A.Move_out_Dt)=DATEPART(day, GETDATE()) -1
		and datepart(month,A.Move_out_Dt)=DATEPART(month, GETDATE())
		and datepart(year,A.Move_out_Dt)=DATEPART(year, GETDATE())
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





Create   Proc OT_Fix_Delete 
@Emp_id   varchar(10)    

AS

set nocount on
declare @Message varchar(100)

	
		Delete from OT_Fix  
	
		where emp_Id=@emp_Id
		set @Message='Data deleted successfully !'		
	
select Message=@Message

set nocount off







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  Proc OT_Fix_IU 
@Emp_id   varchar(10)    
,@Gen_Rate money             
,@Out_Rate money             
,@Hol_Rate money
,@U_Id varchar(10)
AS

set nocount on
declare @Message varchar(100)

	if exists (Select * from OT_Fix  where emp_id=@emp_id )
	begin
		Update OT_Fix  
		set Gen_Rate=@Gen_Rate             
			,Out_Rate=@Out_Rate
			,Hol_Rate=@Hol_Rate
			,U_Id=@U_Id
		where emp_Id=@emp_Id
		set @Message='Data updated successfully !'		
	end 
	else
	begin	

		insert into OT_Fix
		values(@Emp_Id,@Gen_Rate,@Out_Rate,@Hol_Rate,@U_Id)
		set @Message='Data saved successfully !'
	end

select Message=@Message

set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/*
OT_Hr '9-015','may','2002'

*/



CREATE  Procedure OT_Hr 

@Emp_Id varchar(10)
,@pay_Month varchar(12)
,@pay_year varchar(4)


AS

set nocount on
Declare @Emp_LogIn datetime
,@Emp_LogOut  datetime
,@Total_Hr float
,@Hr float

set @Total_Hr=0

Declare LogInDt_Cursor Cursor for

	select Emp_LogIn from emp_Att_info where Emp_Id=@Emp_Id
		and DATENAME(MONTH,Entry_Dt)=@pay_Month 
		and DATEPART(YEAR,Entry_Dt)=@pay_year
	order by Emp_LogIn	

-------------------------------------------------------------------

open LogInDt_Cursor

	fetch next from LogInDt_Cursor into @Emp_LogIn

	while @@fetch_Status=0

	BEGIN
		

		set @Emp_LogOut=(select dbo.GetLogOut(@Emp_Id,@Emp_LogIn))
		
		Set @Hr=ROUND(SUM(DATEDIFF(n,@Emp_logIn,@Emp_logOut)* 0.01667),1) 
		
		print 	'        '+convert(char(18),@Emp_logIn)+' ,  '+ convert(char(20),@Hr)                       

		SET @Total_Hr=@Total_Hr+@Hr
		print '__________________________________	'+convert(char(6),@Total_Hr)
		fetch next from LogInDt_Cursor into @Emp_LogIn

	END

	

close LogInDt_Cursor
deallocate LogInDt_Cursor

set nocount off

--select Total_Office_Stay_Hour=@Total_Hr
--return @Total_Hr


/*
select Emp_LogIn,Emp_LogOut from Emp_Att_Info where Emp_Id='9-015'
and datename(month,Emp_LogIn)='june'and datepart(Year,Emp_LogIn)='2002'
order by Emp_LogIn
*/




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/****** Object:  Stored Procedure dbo.OT_Summery_All    Script Date: 3/7/02 10:43:45 AM ******/
---OT_Summery_All 'April', '2002'

CREATE   PROCEDURE OT_Summery_All
@pay_Month varchar(12),
@pay_year varchar(4)
AS
set nocount on
create table #OT_Rpt(
	E_ID varchar(10),
	[Name]varchar(45),
	Desig varchar(35),
	Dept varchar(35),
	OT_Hour Float,
	OT_Pay Float,
	Pay_Mon_Yr varchar(25)
)
-----------------------Declare Local variable------------------------------------------------

DECLARE @Mnt varchar(4)
,@Emp_ID varchar(10)	
,@Pay_Mon_Yr varchar(25)
set @Pay_Mon_Yr=(@pay_Month + ', '+@pay_year)


----------------------------------------------------------

Declare @Basic MONEY		---Basic Salary
,@Salary money 			---Total Salary
,@Max_Work_Hr int		---Maximum working hour in a month(30 working days)
,@Total_Hr int			---Total hour spent by an individual Worker
,@Days_Att int 			---Total Attendance
,@Daily_Hr float		---Daily working Hour
,@OT_Factor float		---Multiplier (i.e Basic*2)
,@OT_Hr float


Set @Daily_Hr=8
Set @Max_Work_Hr=240
Set @OT_Factor=2
	
-----------------------Declare Cursor--------------------------------------------------------

DECLARE Paymain_Cursor_All CURSOR FOR		--OT_Summery Section specific
SELECT Emp_Id from Emp_Job_Hist_Current 
	where pay_type=0 and Emp_ID in (select distinct Emp_ID from fixed_Pay)
	
--( Select emp_ID from payroll_main 
--		where pay_Month=@Pay_Month and pay_Year=@Pay_Year)
-----------------------Open Cursor-----------------------------------------------------------

OPEN Paymain_Cursor_All

	FETCH NEXT FROM Paymain_Cursor_All INTO @Emp_ID

	WHILE @@FETCH_STATUS = 0

	BEGIN

----------------------------Basic Salary--------------------------------------------
	Set @Basic=(SELECT dbo.GetPresentBasic (@Emp_ID))

----------------------------Total Salary--------------------------------------------
	Set @Salary=(SELECT dbo.GetPresentSalary(@Emp_ID))

----------------------------Total Hour Spent----------------------------------------
		Select @Total_Hr=ROUND(SUM(DATEDIFF(n,Emp_login,Emp_logout)* 0.01667),1) FROM Emp_Att_info 
			WHERE Emp_id=@Emp_ID AND 
			DATENAME(MONTH,Entry_Dt)=@pay_Month AND 
			DATEPART(YEAR,Entry_Dt)=@pay_year
------------------------------------------------------------------------------------
		Set @Days_Att=(SELECT COUNT(Emp_logout)FROM  Emp_Att_Info WHERE Emp_id=@Emp_ID AND 
			DATENAME(MONTH,Entry_Dt)=@pay_Month AND DATEPART(YEAR,Entry_Dt)=@Pay_Year)
------------------------------------------------------------------------------------
		Set @OT_Hr=@Total_Hr-(@Days_Att*@Daily_Hr)
------------------------------------------------------------------------------------
		DECLARE @ID varchar(10)
		DECLARE @MN varchar(45)
		DECLARE @Dept varchar(35)
		DECLARE @Desig varchar(35)
		SELECT @Id=Emp_ID,@MN=(Emp_fna+' '+Emp_mna+' '+Emp_lna),
		@Desig=(select Emp_desig from emp_JOB_Hist_Current where Emp_ID=@emp_ID ),
		@Dept=(select Emp_Dept from emp_JOB_Hist_Current where Emp_ID=@emp_ID)
 
			FROM Emp_per_info WHERE Emp_ID=@emp_ID
--------------------------------------------------------------------------------------------
		Insert into #OT_Rpt(E_ID,[Name],Desig,Dept,OT_Hour,OT_Pay,Pay_Mon_Yr) 
		SELECT @Id,@MN,Desig=@Desig,Dept=@Dept,
			OT_Hour=@OT_Hr,OT_Pay=@Basic,Pay_Mon_Yr=@Pay_Mon_Yr
			
		FETCH NEXT FROM Paymain_Cursor_All INTO @Emp_ID
END
CLOSE Paymain_Cursor_All
DEALLOCATE Paymain_Cursor_All
 
Select  E_ID,[Name],Desig,Dept,OT_Hour,OT_Pay,Pay_Mon_Yr from  #OT_Rpt
order by E_ID


set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.OT_Summery_SectionWise    Script Date: 3/7/02 10:43:45 AM ******/
CREATE PROCEDURE OT_Summery_SectionWise
@pay_Month varchar(12),
@pay_year varchar(4),
@Param varchar(20)
AS
set nocount on
create table #OT_Rpt_Sec(
	E_ID varchar(10),
	[Name]varchar(25),
	Desig varchar(25),
	[Section]varchar(25),
	OT_Hour Float,
	OT_Pay Float,
	Pay_Mon_Yr varchar(25)
)
-----------------------Declare Local variable------------------------------------------------
DECLARE @Emp_ID varchar(10)	
DECLARE @Pay_Mon_Yr varchar(25)
set @Pay_Mon_Yr=(@pay_Month + '   '+@pay_year)
Declare @Basic MONEY		---Basic/Consolidated Salary
Declare @Max_Work_Hr int	---Maximum working hour in a month(30 working days)
Declare @Total_Hr int		---Total hour spent by an individual Worker
Declare @Days_Att int 		---Total Attendance
Declare @Daily_Hr float		---Daily working Hour
Declare @OT_Factor float	---Multiplier (i.e Basic*2)
Declare @OT_Hr float
Set @Daily_Hr=8
Set @Max_Work_Hr=240
Set @OT_Factor=2
-----------------------Declare Cursor--------------------------------------------------------
DECLARE Paymain_Cursor CURSOR FOR		--OT_Summery Section specific
SELECT Emp_Id from Emp_Job_Hist_Current where Emp_section=@Param --AND
--	Emp_ID in ( Select emp_ID from payroll_main 
	--	where pay_Month=@Pay_Month and pay_Year=@Pay_Year)
-----------------------Open Cursor-----------------------------------------------------------
OPEN Paymain_Cursor
FETCH NEXT FROM Paymain_Cursor INTO @Emp_ID
WHILE @@FETCH_STATUS = 0
BEGIN
----------------------------Basic Salary--------------------------------------------
		Set @Basic=(SELECT Amount FROM Fixed_Pay WHERE Head_code=(
			SELECT Head_code FROM Pay_Struc WHERE Head_name='BASIC') and 
			Emp_id=@Emp_ID)
----------------------------Total Hour Spent----------------------------------------
		Select @Total_Hr=ROUND(SUM(DATEDIFF(n,Emp_login,Emp_logout)* 0.01667),1) FROM Emp_Att_info 
			WHERE Emp_id=@Emp_ID AND 
			DATENAME(MONTH,Entry_Dt)=@pay_Month AND 
			DATEPART(YEAR,Entry_Dt)=@pay_year
------------------------------------------------------------------------------------
		Set @Days_Att=(SELECT COUNT(Emp_logout)FROM  Emp_Att_Info WHERE Emp_id=@Emp_ID AND 
			DATENAME(MONTH,Entry_Dt)=@pay_Month AND DATEPART(YEAR,Entry_Dt)=@Pay_Year)
------------------------------------------------------------------------------------
		Set @OT_Hr=@Total_Hr-(@Days_Att*@Daily_Hr)
------------------------------------------------------------------------------------
		DECLARE @ID varchar(10)
		DECLARE @MN varchar(25)
		DECLARE @Sec varchar(25)
		DECLARE @Desig varchar(25)
		SELECT @MN=(Emp_fna+''+Emp_mna+''+Emp_lna),
		@Desig=(select Emp_desig from emp_JOB_Hist_Current where Emp_ID=@emp_ID ),
		@Sec=(select Emp_section from emp_JOB_Hist_Current where Emp_ID=@emp_ID)
 			FROM Emp_per_info WHERE Emp_ID=@emp_ID
--------------------------------------------------------------------------------------------
		Insert into #OT_Rpt_Sec(E_ID,[Name],Desig,[Section],OT_Hour,OT_Pay,Pay_Mon_Yr) 
		SELECT @Emp_ID,@MN,@Desig,@Sec,
			@OT_Hr,OT_Pay=CEILING(((@Basic/@Max_Work_Hr)*@OT_Factor)*@OT_Hr),@Pay_Mon_Yr
			
		FETCH NEXT FROM Paymain_Cursor INTO @Emp_ID
END
CLOSE Paymain_Cursor
DEALLOCATE Paymain_Cursor
 
Select * from  #OT_Rpt_Sec




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Office_Time_I_U_D    Script Date: 3/7/02 10:43:33 AM ******/
/****** Object:  Stored Procedure dbo.Office_Time_I_U_D    Script Date: 10/28/01 4:34:26 PM ******/
/****** Object:  Stored Procedure dbo.Office_Time_I_U_D    Script Date: 10/22/01 6:24:26 PM ******/
CREATE  PROCEDURE Office_Time_I_U_D
@opr varchar (1),
@Start_time varchar(10) ,
@End_time varchar(10),
@Relaxed  varchar(2),
@Abs_time varchar(10) ,
@Sp_Start_time varchar(10),
@Sp_Start_day varchar(10),
@Sp_End_time varchar(10),
@Sp_End_day varchar(10),
@Effect_date datetime,
@U_id char (10) ,
@Ref_Key varchar(10)
	
AS
if @opr='I'
begin
insert into  Office_Time(
Start_time ,
End_time ,
Relaxed ,
Abs_time,
Sp_Start_time,
Sp_Start_day,
Sp_End_time,
Sp_End_day ,
Effect_date,
U_id  
 ) 
values (
@Start_time ,
@End_time ,
@Relaxed ,
@Abs_time,
@Sp_Start_time ,
@Sp_Start_day ,
@Sp_End_time ,
@Sp_End_day ,
@Effect_date,
@U_id  
)
end
                    
if @opr='U'
begin
update  Office_Time set 
Start_time =@Start_time,
End_time=@End_time,
Effect_date =@Effect_date, 
Relaxed=@Relaxed ,
Abs_time=@Abs_time,
Sp_Start_time=@Sp_Start_time ,
Sp_Start_day=@Sp_Start_day ,
Sp_End_time=@Sp_End_time ,
Sp_End_day=@Sp_End_day ,
U_id  =@U_id  where  Ref_Key=convert(int,@Ref_Key)
end
    		
if @opr='D'
begin
delete from  Office_Time where    Ref_Key=convert(int,@Ref_Key)
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



Create Proc Out_of_Office_Del 
@Track_ID int

AS

Delete from Out_of_Office 
	where track_id=@track_id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



/*

select * from out_of_office

*/

Create Procedure Out_of_Office_IU
@Emp_Id varchar(10),
@Str_Date datetime,
@End_Date datetime,
@Description varchar(100),
@Track_Id int

AS

If exists (Select * from Out_of_Office where Track_id=@Track_Id)

		update Out_of_Office set Emp_Id=@Emp_Id,Str_Date=@Str_Date
			,End_Date=@End_Date,[Description]=@Description
			where Track_Id=@Track_Id

Else
		Insert into out_of_office (Emp_Id,Str_Date,End_Date,[Description])
		values(@Emp_Id,@Str_Date,@End_Date,@Description)




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






Create    Proc Overtime_Delete

@Emp_ID varchar(10)     
,@Pay_Month varchar(12)    
,@pay_Year varchar(4)

as
 Set nocount on

declare @Message varchar(100)

	
	if exists (select * from Overtime where 
			emp_id =@emp_id
			and pay_month=@pay_month
			and pay_year=@pay_year)
	begin


		Delete from Overtime

		where 
			emp_id =@emp_id
			and pay_month=@pay_month
			and pay_year=@pay_year

		set @Message='Data deleted successfully !'		
	end



select Message=@Message

Set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE   Proc Overtime_IU

@Emp_ID varchar(10)     
,@Pay_Month varchar(12)    
,@pay_Year varchar(4)
,@Gen_OT_Duration int
,@Out_OT_Duration int
,@Hol_OT_Duration int
,@U_Id varchar(10)
as
 Set nocount on
declare @Gen_Rate money 
	,@Out_Rate money 
	,@Hol_Rate money
	,@Gen_OT_Pay money
	,@Out_OT_Pay money
	,@Hol_OT_Pay money

	,@Message varchar(100)
-----------------------------------------------
	select @Gen_Rate=Gen_Rate 
		,@Out_Rate=Out_Rate
		,@Hol_Rate=Hol_Rate 
	from OT_Fix 
	where Emp_Id=@Emp_Id
-----------------------------------------------
set @Gen_OT_Pay=(@Gen_OT_Duration * @Gen_Rate)        
set @Out_OT_Pay=(@Out_OT_Duration * @Out_Rate)
set @Hol_OT_Pay=(@Hol_OT_Duration * @Hol_Rate)
-----------------------------------------------
	
	if exists (select * from Overtime where 
			emp_id =@emp_id
			and pay_month=@pay_month
			and pay_year=@pay_year)
	begin

		update Overtime

		set Gen_OT_Duration=@Gen_OT_Duration 
		   ,Gen_OT_Pay=@Gen_OT_Pay          
  
	 	   ,Out_OT_Duration =@Out_OT_Duration
		   ,Out_OT_Pay=@Out_OT_Pay            

        	   ,Hol_OT_Duration=@Hol_OT_Duration 
		   ,Hol_OT_Pay=@Hol_OT_Pay
			where 
			emp_id =@emp_id
			and pay_month=@pay_month
			and pay_year=@pay_year
			and P_Status =0

		set @Message='Data updated successfully !'		
	end

	else

	begin
		insert into Overtime(Emp_ID ,Pay_Month	,pay_Year
				,Gen_OT_Duration	,Gen_OT_Pay
				,Out_OT_Duration	,Out_OT_Pay
				,Hol_OT_Duration	,Hol_OT_Pay
				,U_Id)
			values(@Emp_ID	,@Pay_Month	,@pay_Year
				,@Gen_OT_Duration	,@Gen_OT_Pay
				,@Out_OT_Duration	,@Out_OT_Pay
				,@Hol_OT_Duration	,@Hol_OT_Pay
				,@U_Id)

			set @Message='Data saved successfully !'		
	end


select Message=@Message

Set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE   PROCEDURE Overtime_Preparation_Delete
	@Emp_Id varchar(10),
	@Date_of_OT  datetime,
	@Ot_St_Time varchar(25)
	
		
as 

Delete from Overtime_Preparation where 
		Emp_Id=@Emp_Id 
		and Date_of_OT=@Date_of_OT 
		and Ot_St_Time=@Ot_St_Time 
		







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  PROCEDURE Overtime_Preparation_Save
	@Emp_ID varchar(10),
	@Date_of_OT  datetime,
	@Ot_St_Time varchar(25),
	@OT_End_Time varchar(25),
	@Remarks varchar(150),
	@OT_Amount decimal,
	@Entry_Dt datetime,
	@Entry_By varchar(50),
	@P_Status int,
	@NOOfHr decimal
		
as 

If Exists (Select * From Overtime_Preparation 
    Where  emp_id=@Emp_Id and Date_of_OT=@Date_of_OT and Ot_St_Time=@Ot_St_Time and @P_Status=0)
	begin		
		Update Overtime_Preparation  Set 
			Ot_St_Time=@Ot_St_Time ,
			OT_End_Time=@OT_End_Time ,
			Remarks=@Remarks ,
			OT_Amount=@OT_Amount ,
			NOOfHr =@NOOfHr
			where emp_id=@Emp_Id and Date_of_OT=@Date_of_OT and Ot_St_Time=@Ot_St_Time and @P_Status=0
	end
Else

Insert Into Overtime_Preparation (
	Emp_ID,
	Date_of_OT,
	Ot_St_Time ,
	OT_End_Time ,
	Remarks ,
	OT_Amount ,
	Entry_Dt ,
	Entry_By ,
	P_Status ,
	NOOfHr
) Values (
	@Emp_ID ,
	@Date_of_OT  ,
	@Ot_St_Time ,
	@OT_End_Time ,
	@Remarks ,
	@OT_Amount ,
	@Entry_Dt ,
	@Entry_By ,
	@P_Status ,
	@NOOfHr 
)






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Closing_I_U_D    Script Date: 3/7/02 10:43:33 AM ******/
/****** Object:  Stored Procedure dbo.PF_Closing_I_U_D    Script Date: 10/28/01 4:34:26 PM ******/
/****** Object:  Stored Procedure dbo.PF_Closing_I_U_D    Script Date: 10/22/01 6:24:26 PM ******/
----------***************************************
Create procedure PF_Closing_I_U_D
@opr varchar(1),
@Cur_dt datetime,
@GPF_C_Trac_No Int,
@Emp_ID Varchar,
@Emp_Tot_Cont_Amt money,
@Comp_Tot_Cont_Amt money,
@Int_Acc money,
@Tot_Amt money,
@Amt_Nego money,
@U_id varchar (10)
	
as
if @opr='I'
begin
	insert into PF_Closing (Cur_dt,GPF_C_Trac_No,Emp_ID,Emp_Tot_Cont_Amt,Comp_Tot_Cont_Amt,
		Int_Acc,Tot_Amt,Amt_Nego,U_id)
	values(@Cur_dt,@GPF_C_Trac_No,@Emp_ID,@Emp_Tot_Cont_Amt,@Comp_Tot_Cont_Amt,
		@Int_Acc,@Tot_Amt,@Amt_Nego,@U_id)
end
if @opr='U'
begin
	Update  PF_Closing set Cur_dt=@Cur_dt,GPF_C_Trac_No=@GPF_C_Trac_No,Emp_ID=@Emp_ID,
		Emp_Tot_Cont_Amt=@Emp_Tot_Cont_Amt,Comp_Tot_Cont_Amt=@Comp_Tot_Cont_Amt,
		Int_Acc=@Int_Acc,Tot_Amt=@Tot_Amt,Amt_Nego=@Amt_Nego,U_id=@U_id
	where Emp_ID=@Emp_ID
end
if @opr='D'
begin
	Delete from PF_Closing where Emp_ID=@Emp_ID
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Comp_Withdraw_I_U_D    Script Date: 3/7/02 10:43:34 AM ******/
/****** Object:  Stored Procedure dbo.PF_Comp_Withdraw_I_U_D    Script Date: 10/28/01 4:34:27 PM ******/
/****** Object:  Stored Procedure dbo.PF_Comp_Withdraw_I_U_D    Script Date: 10/22/01 6:24:27 PM ******/
----**************************************
create procedure [PF_Comp_Withdraw_I_U_D]
@opr varchar(1),
@Cur_dt datetime,
@Comp_Withd_Code int,
@Invst_code int,
@Com_With_Mat_V money,
@Int_Dist_Flag varchar(1),
@U_id varchar (10) 
as
if @opr='I'
begin
	insert into PF_Comp_Withdraw (Cur_dt,Comp_Withd_Code,Invst_code,Com_With_Mat_V,Int_Dist_Flag,U_id)
	values(@Cur_dt,@Comp_Withd_Code,@Invst_code,@Com_With_Mat_V,@Int_Dist_Flag,@U_id)
end
if @opr='U'
begin
	Update PF_Comp_Withdraw set Cur_dt=@Cur_dt,Comp_Withd_Code=@Comp_Withd_Code,
		Invst_code=@Invst_code,Com_With_Mat_V=@Com_With_Mat_V,
		Int_Dist_Flag=@Int_Dist_Flag,U_id=@U_id
	where Comp_Withd_Code=@Comp_Withd_Code
end
if @opr='D'
begin
	Delete from PF_Comp_Withdraw where Comp_Withd_Code=@Comp_Withd_Code
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Components_I_U_D    Script Date: 3/7/02 10:43:34 AM ******/
/****** Object:  Stored Procedure dbo.PF_Components_I_U_D    Script Date: 10/28/01 4:34:27 PM ******/
/****** Object:  Stored Procedure dbo.PF_Components_I_U_D    Script Date: 10/22/01 6:24:27 PM ******/
--****************************************
create procedure [PF_Components_I_U_D]
@opr varchar(1),
@Cur_dt datetime,
@Com_code varchar(5),
@Com_NM Varchar(50), 
@U_id varchar (10) 
as
if @opr='I'
begin
	insert into PF_Components(Cur_dt,Com_code,Com_NM,U_id) 
	values(@Cur_dt,@Com_code,@Com_NM,@U_id)
end
if @opr='U'
begin
	Update PF_Components set Cur_dt=@Cur_dt,Com_code=@Com_code,Com_NM=@Com_NM,U_id=@U_id
	where com_code=@com_code
end
if @opr='D'
begin
	Delete from PF_Components where com_code=@com_code
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Contribution_I_U_D    Script Date: 3/7/02 10:43:34 AM ******/
/****** Object:  Stored Procedure dbo.PF_Contribution_I_U_D    Script Date: 10/28/01 4:34:27 PM ******/
/****** Object:  Stored Procedure dbo.PF_Contribution_I_U_D    Script Date: 10/22/01 6:24:27 PM ******/
---**********************************************
create procedure [PF_Contribution_I_U_D]
@opr varchar(1),
@Cur_dt datetime,
@Cont_ID int ,
@Emp_ID Varchar (10),
@Emp_cont_amt money,
@Comp_cont_amt money ,
@U_id varchar (10) 
as
if @opr='I'
begin
	insert into PF_Contribution 
		(Cur_dt,Cont_ID,Emp_ID,Emp_cont_amt,Comp_cont_amt,U_id )
	values(@Cur_dt,@Cont_ID,@Emp_ID,@Emp_cont_amt,@Comp_cont_amt,@U_id )
end
if @opr='U'
begin
	Update PF_Contribution set 
	Cur_dt=@Cur_dt,Cont_ID=@Cont_ID,Emp_ID=@Emp_ID,
	Emp_cont_amt=@Emp_cont_amt,Comp_cont_amt=@Comp_cont_amt,
	U_id=@U_id where Emp_ID=@Emp_ID
end
if @opr='D'
begin
	Delete from PF_Contribution  where Emp_ID=@Emp_ID
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Inst_Scm_I_U_D    Script Date: 3/7/02 10:43:34 AM ******/
/****** Object:  Stored Procedure dbo.PF_Inst_Scm_I_U_D    Script Date: 10/28/01 4:34:27 PM ******/
/****** Object:  Stored Procedure dbo.PF_Inst_Scm_I_U_D    Script Date: 10/22/01 6:24:27 PM ******/
create procedure [PF_Inst_Scm_I_U_D]
@opr varchar(1),
@Old_Scm_Code  varchar (10),
@Inst_Scm_Code  varchar (10),
@Inst_Scm_NM varchar (25),
@U_id varchar (10)
as
if @opr='I'
begin
	insert into PF_Inst_Scm (Inst_Scm_Code,Inst_Scm_NM,U_id)
	values(@Inst_Scm_Code,@Inst_Scm_NM,@U_id)
end
if @opr='U'
begin
	Update PF_Inst_Scm set Inst_Scm_Code=@Inst_Scm_Code,Inst_Scm_NM=@Inst_Scm_NM,
	       U_id=@U_id
	where Inst_Scm_Code=@Old_Scm_Code
end
if @opr='D'
begin
	Delete from PF_Inst_Scm where Inst_Scm_Code=@Old_Scm_Code
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Installment_I_U_D    Script Date: 3/7/02 10:43:34 AM ******/
/****** Object:  Stored Procedure dbo.PF_Installment_I_U_D    Script Date: 10/28/01 4:34:27 PM ******/
/****** Object:  Stored Procedure dbo.PF_Installment_I_U_D    Script Date: 10/22/01 6:24:27 PM ******/
----********************************************************
Create procedure [PF_Installment_I_U_D]
@opr varchar(1),
@Cur_dt datetime,
@Ln_App_No Int,
@Ins_Amt money,
@P_By Varchar(10),
@R_By Varchar(10),
@U_id varchar (10)
	
as
if @opr='I'
begin
	insert into PF_Installment(Cur_dt,Ln_App_No,Ins_Amt,P_By,R_By,U_id)
	values(@Cur_dt,@Ln_App_No,@Ins_Amt,@P_By,@R_By,@U_id)
end
if @opr='U'
begin
	Update  PF_Installment set Cur_dt=@Cur_dt,Ln_App_No=@Ln_App_No,
		Ins_Amt=@Ins_Amt,P_By=@P_By,R_By=@R_By,U_id=@U_id
	where Ln_App_No=@Ln_App_No
end
if @opr='D'
begin
	Delete from PF_Installment where Ln_App_No=@Ln_App_No
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Investment_I_U_D    Script Date: 3/7/02 10:43:34 AM ******/
/****** Object:  Stored Procedure dbo.PF_Investment_I_U_D    Script Date: 10/28/01 4:34:28 PM ******/
/****** Object:  Stored Procedure dbo.PF_Investment_I_U_D    Script Date: 10/22/01 6:24:28 PM ******/
-----*************************************
create procedure [PF_Investment_I_U_D]
@opr varchar(1),
@Cur_dt datetime,
@Invst_code int,
@Com_code Varchar(10),
@Com_Sr_No varchar(30),
@Com_Value money,
@Year_in_rate int,
@Invst_Dt datetime,
@Mat_Dt datetime,
@Notes varchar(500),
@U_id varchar (10) 
as
if @opr='I'
begin
	insert into PF_Investment (Cur_dt,Invst_code,Com_code,Com_Sr_No,Com_Value,
	Year_in_rate,Invst_Dt,Mat_Dt,Notes,U_id)
	values(@Cur_dt,@Invst_code,@Com_code,@Com_Sr_No ,@Com_Value,
	@Year_in_rate,@Invst_Dt,@Mat_Dt,@Notes,@U_id)
end
if @opr='U'
begin
	Update PF_Investment set Cur_dt=@Cur_dt,Invst_code=@Invst_code,Com_code=@Com_code,
	Com_Sr_No=@Com_Sr_No,Com_Value=@Com_Value,
	Year_in_rate=@Year_in_rate,Invst_Dt=@Invst_Dt,Mat_Dt=@Mat_Dt,Notes=@Notes,U_id=@U_id
	where Invst_code=@Invst_code
end
if @opr='D'
begin
	Delete from PF_Investment where Invst_code=@Invst_code
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Ln_Amendment_I_U_D    Script Date: 3/7/02 10:43:35 AM ******/
/****** Object:  Stored Procedure dbo.PF_Ln_Amendment_I_U_D    Script Date: 10/28/01 4:34:28 PM ******/
/****** Object:  Stored Procedure dbo.PF_Ln_Amendment_I_U_D    Script Date: 10/22/01 6:24:28 PM ******/
---------******************************************
Create procedure [PF_Ln_Amendment_I_U_D]
@opr varchar(1),
@Cur_dt datetime,
@Ln_App_No Int,
@Ln_Amt money,
@Notes Varchar(500),
@Status Varchar(1),
@U_id varchar (10)
	
as
if @opr='I'
begin
	insert into PF_Ln_Amendment(Cur_dt,Ln_App_No,Ln_Amt,Notes,Status,U_id)
	values(@Cur_dt,@Ln_App_No,@Ln_Amt,@Notes,@Status,@U_id)
end
if @opr='U'
begin
	Update  PF_Ln_Amendment set Cur_dt=@Cur_dt,Ln_App_No=@Ln_App_No,Ln_Amt=@Ln_Amt,
		Notes=@Notes,Status=@Status,U_id=@U_id
	where Ln_App_No=@Ln_App_No
end
if @opr='D'
begin
	Delete from PF_Ln_Amendment where Ln_App_No=@Ln_App_No
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Ln_Application_I_U_D    Script Date: 3/7/02 10:43:35 AM ******/
/****** Object:  Stored Procedure dbo.PF_Ln_Application_I_U_D    Script Date: 10/28/01 4:34:28 PM ******/
/****** Object:  Stored Procedure dbo.PF_Ln_Application_I_U_D    Script Date: 10/22/01 6:24:28 PM ******/
-------************************************
CREATE procedure [PF_Ln_Application_I_U_D]
@opr varchar(1),
@Emp_ID Varchar (10),
@Ln_Code Varchar(10),	
@Ln_Amt money,
@Rfnd_Code varchar (15) ,
@Inst_Amount money,
@Notes Varchar(500),
@Ln_App_No int,
@U_id varchar (10) 
as
if @opr='I'
begin
	insert into PF_Ln_Application (Emp_ID,Ln_Code,Ln_Amt,Rfnd_Code,Inst_Amount,Notes,U_id)
	values(@Emp_ID,@Ln_Code,@Ln_Amt,@Rfnd_Code,@Inst_Amount,@Notes,@U_id)
end
/*
if @opr='U'
begin
	Update  PF_Ln_Application set Emp_ID=@Emp_ID,
		Ln_Code=@Ln_Code,Ln_Amt=@Ln_Amt,Rfnd_Code=@Rfnd_Code,Inst_Amount=@Inst_Amount,Notes=@Notes,U_id=@U_id
	where Ln_App_No=@Ln_App_No
end
if @opr='D'
begin
	Delete from PF_Ln_Application where Ln_App_No=@Ln_App_No
end
*/




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Ln_Approval_I_U_D    Script Date: 3/7/02 10:43:25 AM ******/
/****** Object:  Stored Procedure dbo.PF_Ln_Approval_I_U_D    Script Date: 10/28/01 4:34:20 PM ******/
/****** Object:  Stored Procedure dbo.PF_Ln_Approval_I_U_D    Script Date: 10/22/01 6:24:20 PM ******/
------**************************************
create procedure [PF_Ln_Approval_I_U_D]
@opr varchar(1),
@Cur_dt datetime,
@Ln_App_No Int,
@Sanc_Amt Money,
@With_Int_Rate money,
@Ref_Amt money,
@Notes Varchar(500),
@Status Varchar(1),
@U_id varchar (10) 
as
if @opr='I'
begin
	insert into PF_Ln_Approval(Cur_dt,Ln_App_No,Sanc_Amt,With_Int_Rate,
		    Ref_Amt,Notes,Status,U_id)
	values(@Cur_dt,@Ln_App_No,@Sanc_Amt,@With_Int_Rate,@Ref_Amt,@Notes,
		@Status,@U_id)
end
if @opr='U'
begin
	Update  PF_Ln_Approval set Cur_dt=@Cur_dt,Sanc_Amt=@Sanc_Amt,
		With_Int_Rate=@With_Int_Rate,Ref_Amt=@Ref_Amt,Notes=@Notes,
		Status=@Status,U_id=@U_id
	where Ln_App_No=@Ln_App_No
end
if @opr='D'
begin
	Delete from PF_Ln_Approval where Ln_App_No=@Ln_App_No
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Ln_Cancellation_I_U_D    Script Date: 3/7/02 10:43:35 AM ******/
/****** Object:  Stored Procedure dbo.PF_Ln_Cancellation_I_U_D    Script Date: 10/28/01 4:34:28 PM ******/
/****** Object:  Stored Procedure dbo.PF_Ln_Cancellation_I_U_D    Script Date: 10/22/01 6:24:28 PM ******/
------------****************************
Create procedure [PF_Ln_Cancellation_I_U_D]
@opr varchar(1),
@Cur_dt datetime,
@Ln_App_No Int,
@Notes Varchar(500),
@U_id varchar (10)
	
as
if @opr='I'
begin
	insert into PF_Ln_Cancellation(Cur_dt,Ln_App_No,Notes,U_id)
	values(@Cur_dt,@Ln_App_No,@Notes,@U_id)
end
if @opr='U'
begin
	Update  PF_Ln_Cancellation set Cur_dt=@Cur_dt,Ln_App_No=@Ln_App_No,
		Notes=@Notes,U_id=@U_id
	where Ln_App_No=@Ln_App_No
end
if @opr='D'
begin
	Delete from PF_Ln_Cancellation where Ln_App_No=@Ln_App_No
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Ln_Policy_I_U_D    Script Date: 3/7/02 10:43:35 AM ******/
/****** Object:  Stored Procedure dbo.PF_Ln_Policy_I_U_D    Script Date: 10/28/01 4:34:28 PM ******/
/****** Object:  Stored Procedure dbo.PF_Ln_Policy_I_U_D    Script Date: 10/22/01 6:24:28 PM ******/
-----------**********************************************
Create procedure [PF_Ln_Policy_I_U_D]
@opr varchar(1),
@Cur_dt datetime,
@Emp_Desig varchar(25),
@Max_Ln_Allow money,
@U_id varchar (10)
	
as
if @opr='I'
begin
	insert into PF_Ln_Policy(Cur_dt,Emp_Desig,Max_Ln_Allow,U_id)
	values(@Cur_dt,@Emp_Desig,@Max_Ln_Allow,@U_id)
end
if @opr='U'
begin
	Update  PF_Ln_Policy set Cur_dt=@Cur_dt,Emp_Desig=@Emp_Desig,
		Max_Ln_Allow=@Max_Ln_Allow,U_id=@U_id
	where Emp_Desig=@Emp_Desig
end
if @opr='D'
begin
	Delete from PF_Ln_Policy where Emp_Desig=@Emp_Desig
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Loan_I_U_D    Script Date: 3/7/02 10:43:35 AM ******/
/****** Object:  Stored Procedure dbo.PF_Loan_I_U_D    Script Date: 10/28/01 4:34:28 PM ******/
/****** Object:  Stored Procedure dbo.PF_Loan_I_U_D    Script Date: 10/22/01 6:24:29 PM ******/
create procedure [PF_Loan_I_U_D]
@opr varchar(1),
@Cur_dt datetime,
@Ln_Code Varchar(10),
@Ln_NM Varchar(30),
@Max_Cel money,
@Min_Int_Rate money,
@U_id varchar (10) 
as
if @opr='I'
begin
	insert into PF_Loan (Cur_dt,Ln_Code,Ln_NM,Max_Cel,Min_Int_Rate,U_id)
	values(@Cur_dt,@Ln_Code,@Ln_NM,@Max_Cel,@Min_Int_Rate,@U_id)
end
if @opr='U'
begin
	Update PF_Loan set Cur_dt=@Cur_dt,Ln_Code=@Ln_Code,Ln_NM=@Ln_NM,
	       Max_Cel=@Max_Cel,Min_Int_Rate=@Min_Int_Rate,
	       U_id=@U_id
	where Ln_Code=@Ln_Code
end
if @opr='D'
begin
	Delete from PF_Loan where Ln_Code=@Ln_Code
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_POP_Tables    Script Date: 3/7/02 10:43:36 AM ******/
/****** Object:  Stored Procedure dbo.PF_POP_Tables    Script Date: 10/28/01 4:34:29 PM ******/
/****** Object:  Stored Procedure dbo.PF_POP_Tables    Script Date: 10/22/01 6:24:29 PM ******/
--select* from PF_loan
CREATE PROCEDURE [PF_POP_Tables] 
@mode VARCHAR(10),
@param1 VARCHAR(20)
AS
-----------------------------------------------------------------------------------1
IF @mode='CONTRIB'
BEGIN
	SELECT Cur_dt,Cont_ID,Emp_ID,Emp_cont_amt,Comp_cont_amt FROM PF_Contribution						/* table-01*/
END
-----------------------------------------------------------------------------------2
IF @mode='COMPO'
BEGIN
	SELECT Cur_dt,Com_code,Com_NM FROM PF_COMPONENTS
END
-----------------------------------------------------------------------------------3
IF @mode='INVEST'
BEGIN
	SELECT Cur_dt,Invst_code,Com_code,Com_Sr_No,Com_Value,
		Year_in_rate,Invst_Dt,Mat_Dt,Notes 
	FROM PF_Investment 
END
-----------------------------------------------------------------------------------4
IF @mode='COMP_WD'
BEGIN
	SELECT Cur_dt,Comp_Withd_Code,Invst_code,Com_With_Mat_V,Int_Dist_Flag FROM PF_Comp_Withdraw 
END
-----------------------------------------------------------------------------------5
IF @mode='LOAN'
BEGIN
	SELECT Cur_dt,Ln_Code,Ln_NM,Max_Cel,Min_Int_Rate FROM PF_Loan
END
-----------------------------------------------------------------------------------6
IF @mode='LN_APPL'
BEGIN
	SELECT Cur_dt,Ln_App_No,Emp_ID,Ln_Code,Ln_Amt,Notes,Status FROM PF_Ln_Application 
END
-----------------------------------------------------------------------------------7
IF @mode='LN_APRV'
BEGIN
	SELECT Cur_dt,Ln_App_No,Sanc_Amt,With_Int_Rate,Ref_Amt,Notes,Status FROM PF_Ln_Approval
END
-----------------------------------------------------------------------------------8
IF @mode='LN_AMND'
BEGIN
	SELECT Cur_dt,Ln_App_No,Ln_Amt,Notes,Status FROM PF_Ln_Amendment
END
-----------------------------------------------------------------------------------9
IF @mode='LN_CNCL'
BEGIN
	SELECT Cur_dt,Ln_App_No,Notes FROM PF_Ln_Cancellation
END
-----------------------------------------------------------------------------------10
IF @mode='LN_COL'
BEGIN
	SELECT Cur_dt,Ln_App_No,P_Amt,P_By,R_By FROM PF_Ln_Collection
END
-----------------------------------------------------------------------------------11
IF @mode='INSTALL'
BEGIN
	SELECT Cur_dt,Ln_App_No,Ins_Amt,P_By,R_By FROM PF_Installment
END
-----------------------------------------------------------------------------------12
IF @mode='LN_PLC'
BEGIN
	SELECT  Cur_dt,Emp_Desig,Max_Ln_Allow FROM PF_Ln_Policy
END
-----------------------------------------------------------------------------------13
IF @mode='PF_CLS'
BEGIN
	SELECT Cur_dt,GPF_C_Trac_No,Emp_ID,Emp_Tot_Cont_Amt,Comp_Tot_Cont_Amt,
		Int_Acc,Tot_Amt,Amt_Nego FROM PF_Closing 
END
-----------------------------------------------------------------------------------14
IF @mode='PF_HIST'
BEGIN
	SELECT Cur_dt,GPF_C_Trac_No,P_Amt,P_By,R_By FROM PF_Pay_History 
END
-----------------------------------------------------------------------------------15
IF @mode='INST_SCM'
BEGIN
	SELECT *  FROM PF_Inst_Scm
END
--------------------------------------------------------------------------------------16
IF @MODE='LN_APP'
BEGIN
	
	SELECT Ln_App_No,Emp_ID,Category=(select Ln_NM from PF_Loan where Ln_Code=PF_Ln_Application.Ln_Code),
	Ln_Amt,Refund_Mode=(Select Inst_Scm_NM from PF_inst_scm where Inst_Scm_Code=PF_Ln_Application.Rfnd_Code),
	Inst_Amount,Notes FROM PF_Ln_Application where Emp_ID=@param1
END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.PF_Pay_History_I_U_D    Script Date: 3/7/02 10:43:35 AM ******/
/****** Object:  Stored Procedure dbo.PF_Pay_History_I_U_D    Script Date: 10/28/01 4:34:29 PM ******/
/****** Object:  Stored Procedure dbo.PF_Pay_History_I_U_D    Script Date: 10/22/01 6:24:29 PM ******/
Create procedure PF_Pay_History_I_U_D
@opr varchar(1),
@Cur_dt datetime,
@GPF_C_Trac_No Int,
@P_Amt money,
@P_By Varchar(10),
@R_By Varchar(10),
@U_id varchar (10)
	
as
if @opr='I'
begin
	insert into PF_Pay_History (Cur_dt,GPF_C_Trac_No,P_Amt,P_By,R_By,U_id)
	values(@Cur_dt,@GPF_C_Trac_No,@P_Amt,@P_By,@R_By,@U_id)
end
if @opr='U'
begin
	Update  PF_Pay_History set Cur_dt=@Cur_dt,GPF_C_Trac_No=@GPF_C_Trac_No,
		P_Amt=@P_Amt,P_By=@P_By,R_By=@R_By,U_id=@U_id
	where GPF_C_Trac_No=@GPF_C_Trac_No
end
if @opr='D'
begin
	Delete from PF_Pay_History where GPF_C_Trac_No=@GPF_C_Trac_No
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







CREATE   procedure POP_All_About_Group_Shift
@mode varchar(15),
@Criteria_1 varchar(15),
@Criteria_2 varchar(15)
as
---------------------Grouping------------------------"exec POP_All_About_Group_Shift 'Group'"
if @mode='Group'
begin
	select Gr_Code,Gr_Name from Group_NM
end 
-------------------------------
if @mode='Rest_Workers'
begin
	
	select Emp_Id,Emp_Name=(Emp_fna+' '+Emp_mna+' '+Emp_lna) 
		from Emp_per_info where Emp_Id in 
		(select emp_id from emp_job_Hist_Current where Duty_type='0') and 
			Emp_Id not in 
				(select Emp_Id from worker_Group)
end 
-------------------------------
if @mode='Worker_Group'
begin
	select a.emp_id,Emp_Name=(select Emp_fna+" "+Emp_mna+ " "+Emp_lna from Emp_per_info where emp_id=a.emp_id ),
	Gr_Name=(select gr_name from Group_nm where gr_code=a.gr_code )
	from worker_group a
end 
-------------------------------
if @mode='Shift'
begin
	select Shift_Code,Shift_Name,Shift_Start,Shift_End,[Delay],Abs_time from shift
end
-------------------------------
if @mode='Group_Shift'
begin
select Shift_Name=(select Shift_name from Shift where shift_Code=a.Shift_Code),
Gr_name=(select Gr_name from Group_NM where Gr_Code=a.Gr_Code),Start_Dt,End_Dt
from group_Shift a where  a.End_Dt > getdate()
end
-----------------------------------
if @mode='Group_Worker'
begin
	select a.emp_id,Emp_Name=(select Emp_fna+" "+Emp_mna+ " "+Emp_lna from Emp_per_info where emp_id=a.emp_id ),
	Gr_Name=(select gr_name from Group_nm where gr_code=a.gr_code ) from worker_group a where gr_code=@Criteria_1
end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




Create Procedure POP_Discipline 
@Emp_Id Varchar(10)
AS

SELECT Emp_Id,Track_Id,Ref_No,Reason,Penalty,U_Id,Dt 
from Emp_Discipline where emp_Id=@emp_Id





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




Create Procedure POP_Emp_Discipline
@Emp_Id varchar(10)

as 
Select Track_Id,Ref_No,Reason,Penalty,U_Id,Dt From Emp_Discipline
    Where Emp_Id=@Emp_Id order by Track_Id





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.POP_Emp_Education    Script Date: 3/7/02 10:43:36 AM ******/
CREATE PROCEDURE POP_Emp_Education
@Emp_ID varchar(10)
 AS
select * from Emp_Education where Emp_ID=@Emp_ID




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/***

select * from payroll_main
exec POP_Emp_NM_Not_Paid 'Id','','18','August','2002'

exec POP_Emp_NM_Not_Paid 'pop','9-001','18','August','2002'
***/

CREATE      Procedure POP_Emp_NM_Not_Paid
@Mode varchar(10),
@Emp_Id varchar(10),
@Pay_Type varchar(2),
@pay_Month varchar(12),
@Pay_Year varchar(4)
AS

if @Mode='ID'
Begin
	select emp_Id from payroll_main where Pay_month=@pay_Month 
		and Pay_year=@Pay_Year 
		and Pay_Type=@Pay_Type 
		and Pay_Stat='0'
End

if @Mode='POP'
Begin
	select Emp_Name=(select Emp_Fna+' '+Emp_Mna+' '+Emp_Lna from Emp_Per_Info where emp_Id=@Emp_ID)
	,Emp_Dept=(select Emp_Dept from Emp_Job_Hist_Current where emp_Id=@Emp_ID)
	,Emp_Desig=(select Emp_Desig from Emp_Job_Hist_Current where emp_Id=@Emp_ID)
	,Emp_join_date=(select Emp_join_date from Emp_Job_Hist_Current where emp_Id=@Emp_ID)
	,Amount= case @pay_type
		when 0 then (select dbo.GetNetPay (@Emp_ID,@Pay_Month,@Pay_Year))
		else (select dbo.GetNetBonusPay (@Emp_ID,@Pay_Type,@Pay_Month,@Pay_Year))
		end
	,Pay_month=@Pay_month,Pay_Year=@Pay_Year
	,A_Gen=(Select A_Gen from payroll_main where emp_id=@Emp_Id and 
			Pay_month=@Pay_month and Pay_Year=@Pay_Year and Pay_Type=@Pay_Type)
end









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO











/****** Object:  Stored Procedure dbo.POP_Emp_NonUser    Script Date: 1/1/02 10:42:33 AM ******/

/****** Object:  Stored Procedure dbo.POP_Emp_NonUser    Script Date: 12/30/01 6:57:03 PM ******/

/****** Object:  Stored Procedure dbo.POP_Emp_NonUser    Script Date: 12/27/01 6:38:38 PM ******/

CREATE  procedure POP_Emp_NonUser

as
	Select distinct
	a.Emp_ID,NM=(a.Emp_Fna+' '+a.Emp_Mna+' '+a.Emp_Lna),b.Emp_desig	
	from Emp_per_info a,Emp_Job_Hist_Current b 
	where a.emp_id=b.emp_id and b.emp_id not in (select u_id from soft_pass)
	
















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.POP_Emp_Ref    Script Date: 3/7/02 10:43:36 AM ******/
/****** Object:  Stored Procedure dbo.POP_Emp_Ref    Script Date: 10/28/01 4:34:29 PM ******/
/****** Object:  Stored Procedure dbo.POP_Emp_Ref    Script Date: 10/22/01 6:24:29 PM ******/
/****** Object:  Stored Procedure dbo.POP_Emp_Ref    Script Date: 9/17/00 4:34:01 PM ******/
/****** Object:  Stored Procedure dbo.POP_Emp_Ref    Script Date: 9/4/01 6:30:47 PM ******/
CREATE PROCEDURE POP_Emp_Ref
@Emp_ID varchar(10)
 AS
select * from Emp_Reference where Emp_ID=@Emp_ID




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.POP_Emp_detail_for_Payroll    Script Date: 3/7/02 10:43:46 AM ******/
/****** Object:  Stored Procedure dbo.POP_Emp_detail_for_Payroll    Script Date: 10/28/01 4:34:40 PM ******/
/****** Object:  Stored Procedure dbo.POP_Emp_detail_for_Payroll    Script Date: 10/22/01 6:24:39 PM ******/
/****** Object:  Stored Procedure dbo.POP_Emp_detail_for_Payroll    Script Date: 9/17/00 4:34:12 PM ******/
/****** Object:  Stored Procedure dbo.POP_Emp_detail_for_Payroll    Script Date: 9/4/01 6:30:57 PM ******/
CREATE PROCEDURE POP_Emp_detail_for_Payroll
as
Select distinct
	A.Emp_ID,
	NM=(A.Emp_Fna+' '+A.Emp_Mna+' '+A.Emp_Lna),
	B.Emp_desig,
	B.Emp_dept,
	B.Emp_section,
	B.Emp_rank,
	B.Emp_branch,
	B.Emp_dept,
	B.Emp_section,
	B.Emp_job_type
	
	from	Emp_per_info A,Emp_Job_Hist_Current B
	where A.emp_id= B.emp_id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO




/*
POP_Holidays '2003'
*/
CREATE   Procedure POP_Holidays
@Year varchar(4)
AS

select Hol_name,Hol_Desc,Str_date=min(Str_date),End_date=max(End_date),
	DURATION=DATEDIFF(DAY,min(Str_date),max(End_date))+1,Category from Hol_list
where datepart(year,Str_date)=@Year
 group by Hol_name,Hol_Desc,Category
order by Str_date





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE       Procedure POP_ID_Salary_Fixing 

@Mode varchar (10)
,@Emp_dept varchar (45)
,@Emp_desig varchar (45)
As

If @Mode='Fixed'

BEGIN


	select a.emp_id,[Name]=a.Emp_fna+' '+a.Emp_mna + ' '+ a.Emp_lna,
	Basic_Amount=c.Amount, 
	AccountNo=C.AccountNo
	from emp_per_info a,emp_Job_hist_current b, fixed_Pay c
	where b.Emp_desig=@Emp_desig 
		and b.Emp_dept=@Emp_dept 
		and a.emp_id=b.emp_id 
		and a.emp_id=c.emp_id 
		and b.emp_id in(select distinct Emp_ID from fixed_Pay)

END

If @Mode='NotFixed'

BEGIN

	select a.emp_id,[Name]=a.Emp_fna+' '+a.Emp_mna + ' '+ a.Emp_lna,
	Basic_Amount=0--=c.Amount	
	from emp_per_info a,emp_Job_hist_current b--, fixed_Pay c
	where b.Emp_desig=@Emp_desig 
		and b.Emp_dept=@Emp_dept 
		and a.emp_id=b.emp_id 
		--and a.emp_id=c.emp_id 
		and b.emp_id not in(select distinct Emp_ID from fixed_Pay)

END











GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




Create Procedure POP_Id_Sal_Fixed_Or_Not
@Mode varchar(10)
AS

IF @Mode='NotFixed'
Begin
	select Emp_Id from Emp_Per_Info where Emp_Id not in 
			(select distinct Emp_Id from fixed_pay)
End

IF @Mode='Fixed'

Begin
	select Emp_Id from Emp_Per_Info where Emp_Id in
			(select distinct Emp_Id from fixed_pay)
End







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.POP_Increment_future    Script Date: 3/7/02 10:43:36 AM ******/
/****** Object:  Stored Procedure dbo.POP_Increment_future    Script Date: 10/28/01 4:34:29 PM ******/
/****** Object:  Stored Procedure dbo.POP_Increment_future    Script Date: 10/22/01 6:24:30 PM ******/
/****** Object:  Stored Procedure dbo.POP_Increment_future    Script Date: 9/17/00 4:34:02 PM ******/
/****** Object:  Stored Procedure dbo.POP_Increment_future    Script Date: 9/4/01 6:30:48 PM ******/
CREATE procedure POP_Increment_future
@emp_id varchar(10)
as
select Title,Amount,Effect_date,uniq  from increment_future
where effect_date>getdate() and emp_ID=@emp_id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




Create Procedure POP_Ind_Emp_Sal_Fix
@Emp_Id varchar(10)
AS

select a.emp_id,[Name]=a.Emp_fna+' '+a.Emp_mna + ' '+ a.Emp_lna
,b.Emp_dept,b.Emp_desig,c.Basic_Amount,d.Acc_No 
from emp_per_info a,emp_Job_hist_current b,Emp_Salary c,Emp_bank_Account d
where 
c.emp_id=d.emp_id and a.emp_id=b.emp_id  and a.emp_id=C.emp_id and 
b.emp_id =@Emp_Id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.POP_Job_Setup    Script Date: 3/7/02 10:43:36 AM ******/
/****** Object:  Stored Procedure dbo.POP_Job_Setup    Script Date: 10/28/01 4:34:29 PM ******/
/****** Object:  Stored Procedure dbo.POP_Job_Setup    Script Date: 10/22/01 6:24:30 PM ******/
/****** Object:  Stored Procedure dbo.POP_Job_Setup    Script Date: 9/17/00 4:34:02 PM ******/
/****** Object:  Stored Procedure dbo.POP_Job_Setup    Script Date: 9/4/01 6:30:48 PM ******/
create procedure [POP_Job_Setup]
@POP_Table varchar(15)
as
if @POP_Table='Job_Title'
	begin
		Select* from Job_title
	end
------------------------
if @POP_Table='Job_Type'
	begin
		Select* from Job_Type
	end
------------------------
if @POP_Table='Department_Info'
	begin
		Select* from Department_Info
	end
------------------------
if @POP_Table='Branch_Info'
	begin
		Select* from Branch_Info
	end
------------------------
if @POP_Table='Rank'
	begin
		Select* from Rank
	end
-------------------------
if @POP_Table='Sec_Info'
	begin
		Select* from Sec_Info
	end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.POP_Leave_App_Aprv    Script Date: 3/7/02 10:43:25 AM ******/
/****** Object:  Stored Procedure dbo.POP_Leave_App_Aprv    Script Date: 10/28/01 4:34:41 PM ******/
/****** Object:  Stored Procedure dbo.POP_Leave_App_Aprv    Script Date: 10/22/01 6:24:40 PM ******/
/****** Object:  Stored Procedure dbo.POP_Leave_App_Aprv    Script Date: 9/17/00 4:34:13 PM ******/
/****** Object:  Stored Procedure dbo.POP_Leave_App_Aprv    Script Date: 9/4/01 6:30:58 PM ******/
CREATE PROCEDURE [POP_Leave_App_Aprv]
@Emp_ID varchar (10),
@App_id int
 AS
Select 
	Leave_name,Start_dt,End_dt,App_Type,Dt from Emp_Leave_Info
	where Emp_ID=@Emp_ID and App_id=@App_id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.POP_N_Insert_Approval    Script Date: 3/7/02 10:43:46 AM ******/
/****** Object:  Stored Procedure dbo.POP_N_Insert_Approval    Script Date: 10/28/01 4:34:41 PM ******/
/****** Object:  Stored Procedure dbo.POP_N_Insert_Approval    Script Date: 10/22/01 6:24:40 PM ******/
/****** Object:  Stored Procedure dbo.POP_N_Insert_Approval    Script Date: 9/17/00 4:34:13 PM ******/
/****** Object:  Stored Procedure dbo.POP_N_Insert_Approval    Script Date: 9/4/01 6:30:58 PM ******/
CREATE PROCEDURE POP_N_Insert_Approval
@App_ID varchar(10),
@Policy varchar(10)
AS
-------------Collecting Supervisors------------------
declare @S1 as varchar(10)
declare @S2 as varchar(10)
declare @S3 as varchar(10)
declare @S4 as varchar(10)
declare @S5 as varchar(10)
--------------------------------------------------------
declare @Policy_Code varchar(10) 
declare @Emp_ID varchar(10)
-------------CollectingTier policy(1,0,0,1,'001')------------------
declare @T1 as varchar(1)
declare @T2 as varchar(1)
declare @T3 as varchar(1)
declare @T4 as varchar(1)
declare @T5 as varchar(10)
---------------------------Insert values for approval regarding supervisor-----------------------------
declare @I1 as varchar(10)
declare @I2 as varchar(10)
declare @I3 as varchar(10)
declare @I4 as varchar(10)
declare @I5 as varchar(10)
-----------------------------Checking for type of application--------------------------------------------------
if @Policy='leave'
begin
------------------------------------Collecting Employee ID for selected App_id--------------------
	select @Emp_ID=emp_id from EMP_LEAVE_INFO where app_id=@app_id	
end
---------------------------Collecting herarchi for selected employee------------------------------
set @s5=(select sup_id from emp_super where emp_id=
(select sup_id from emp_super where emp_id=
(select sup_id from emp_super where emp_id=
(select sup_id from emp_super where emp_id=
(select sup_id from emp_super where emp_id=@emp_id)))))
set @s4=(select sup_id from emp_super where emp_id=
(select sup_id from emp_super where emp_id=
(select sup_id from emp_super where emp_id=
(select sup_id from emp_super where emp_id=@emp_id))))
set @s3=(select sup_id from emp_super where emp_id=
(select sup_id from emp_super where emp_id=
(select sup_id from emp_super where emp_id=@emp_id)))
set @s2=(select sup_id from emp_super where emp_id=
(select sup_id from emp_super where emp_id=@emp_id))
set @s1=(select sup_id from emp_super where emp_id=@emp_id)
-------------------------------------------------------------------------
-----------------------------Collecting Appoval tiers for selected Emp_id--------------------------------------------
select  @Policy_Code=Policy_Code,@T1=Tier_1,@T2=Tier_2,@T3=Tier_3,@T4=Tier_4,@T5=Final_Tier from tier_setup where Policy_Code=
		(select Policy_Code from Company_policy where Policy_type like + '%' + @Policy + '%' and code=
			(select code from department_info where title=
				(Select Emp_dept from emp_job_hist_current where Emp_ID=@Emp_ID)))
------------Prepaing Insert values--------------------------------------------
IF @T1=1 AND @S1 is not NULL
BEGIN
	SET @I1=@S1
END
ELSE
BEGIN
	SET @I1='NA'
END
IF @T2=1 AND @S2 is not NULL
BEGIN
	SET @I2=@S2
END
ELSE
BEGIN
	SET @I2='NA'
END
IF @T3=1 AND @S3 is not NULL
BEGIN
	SET @I3=@S3
END
ELSE
BEGIN
	SET @I3='NA'
END
IF @T4=1 AND @S4 is not NULL
BEGIN
	SET @I4=@S4
END
ELSE
BEGIN
	SET @I4='NA'
END
IF @T5 is not NULL	
BEGIN
	SET @I5=@T5	--------<<Attention
END
ELSE
BEGIN
	SET @I5='NA'
END
-------------------Finally Inserting values-------------------------------------
INSERT INTO APPROVAL ( App_ID,Policy_Code,TIER_1,TIER_2,TIER_3,TIER_4,FINAL_TIER) VALUES
(@App_ID,@Policy_Code,@I1,@I2,@I3,@I4,@I5)
-------------------Initializing starting position---------------------------------
IF @T1=1 exec sp_set_position @app_id,1
else IF @T2=1 exec sp_set_position @app_id,2
else IF @T3=1 exec sp_set_position @app_id,3
else IF @T4=1 exec sp_set_position @app_id,4
else IF @T5=1 exec sp_set_position @app_id,5
else exec sp_set_position @app_id,0




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



--select * from param_tbl
create procedure POP_Param
@Policy_No int
AS

select Policy=@Policy_No,Flag=dbo.GetParamFlag(@Policy_No)
,Value=dbo.GetParamValue(@Policy_No)



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE Procedure POP_Payscale_Main
@Mode varchar(10)
as

IF @Mode='Desig'

BEGIN

	Select Designation=a.Title,a.Code,b.SB,b.Incr,b.EB,b.EBIncr,b.EBMax,b.Scale_Code
	from job_title a,Payscale_Main b where 
	a.Code=b.desig_Code
END

IF @Mode='Rank'

BEGIN

	Select Designation=a.Title,b.SB,b.Incr,b.EB,b.EBIncr,b.EBMax,b.Scale_Code,EndBasic='EB'
	from Rank a,Payscale_Main b where 
	a.Title=b.desig_Code
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE Procedure POP_Payscale_Sub
@Scale_Code int
AS

Select a.Scale_Code,a.Head_Code,b.Head_name,a.Calc_Type,a.Prcnt_of_Basic  
from Payscale_sub a,pay_struc b where a.Head_code=b.Head_code 
and a.Scale_Code=@Scale_Code



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




----POP_Salary_HeadCode'Code','Basic pay'

CREATE Procedure POP_Salary_HeadCode
@Mode varchar(5)
,@Param varchar(30)
as

if @Mode='Head'

BEGIN
	select Head_code, Head_name from pay_struc where 
		Operation='+'and Head_Code not in
			(select head_code from Taxable_Ceiling)
END

if @Mode='Code'

BEGIN
	select Head_code from pay_struc where Head_name=@Param
END







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE Procedure POP_Tax_Slab
@Fin_Year varchar(9)
AS

Select * from Tax_Slab where Fin_year=@Fin_Year


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE  Procedure POP_Taxable_Ceiling

AS

Select Head_Name=(dbo.GetSalaryHead (Head_Code)),Ceiling_Amount,Track_id 

from Taxable_Ceiling






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.POP_Tier_Emp_id_Sp    Script Date: 3/7/02 10:43:47 AM ******/
/****** Object:  Stored Procedure dbo.POP_Tier_Emp_id_Sp    Script Date: 10/28/01 4:34:41 PM ******/
/****** Object:  Stored Procedure dbo.POP_Tier_Emp_id_Sp    Script Date: 10/22/01 6:24:40 PM ******/
/****** Object:  Stored Procedure dbo.POP_Tier_Emp_id_Sp    Script Date: 9/17/00 4:34:13 PM ******/
/****** Object:  Stored Procedure dbo.POP_Tier_Emp_id_Sp    Script Date: 9/4/01 6:30:58 PM ******/
create procedure POP_Tier_Emp_id_Sp
@Emp_ID varchar(10),
@Policy varchar(10)
as
select  Policy_Code,Tier_1,Tier_2,Tier_3,Tier_4,Final_Tier from tier_setup where Policy_Code=
		(select Policy_Code from Company_policy where Policy_type like @Policy and code=
			(select code from department_info where title=
				(Select Emp_dept from emp_job_hist_current where Emp_ID=@Emp_ID)))




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.POP_Tier_Setup    Script Date: 3/7/02 10:43:37 AM ******/
/****** Object:  Stored Procedure dbo.POP_Tier_Setup    Script Date: 10/28/01 4:34:30 PM ******/
/****** Object:  Stored Procedure dbo.POP_Tier_Setup    Script Date: 10/22/01 6:24:30 PM ******/
/****** Object:  Stored Procedure dbo.POP_Tier_Setup    Script Date: 9/17/00 4:34:02 PM ******/
/****** Object:  Stored Procedure dbo.POP_Tier_Setup    Script Date: 9/4/01 6:30:48 PM ******/
CREATE Procedure POP_Tier_Setup
as
Select distinct a.title ,b.Policy_type,c.Tier_1,c.Tier_2,c.Tier_3,c.Tier_4,c.Final_Tier from 
department_Info a,Company_Policy b,tier_setup c where a.code=c.code and 
b.Policy_code=c.Policy_code
/*select Department=(select Title from Department_Info where Code in (select Code from Tier_Setup)),
	Policy_name=(select  Policy_type from Company_Policy
		 where Company_Policy.Policy_Code=Tier_Setup.Policy_code),
	Tier_1,Tier_2,Tier_3,Tier_4,Final_Tier from Tier_Setup*/




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.POP_Tr_Determine_Move_Pos    Script Date: 3/7/02 10:43:47 AM ******/
/****** Object:  Stored Procedure dbo.POP_Tr_Determine_Move_Pos    Script Date: 10/28/01 4:34:42 PM ******/
/****** Object:  Stored Procedure dbo.POP_Tr_Determine_Move_Pos    Script Date: 10/22/01 6:24:41 PM ******/
/****** Object:  Stored Procedure dbo.POP_Tr_Determine_Move_Pos    Script Date: 9/17/00 4:34:13 PM ******/
/****** Object:  Stored Procedure dbo.POP_Tr_Determine_Move_Pos    Script Date: 9/4/01 6:30:58 PM ******/
Create Procedure POP_Tr_Determine_Move_Pos
@App_ID varchar(10)
as
select  Current_Pos=Aprv.move_Pos,Tr_S.Tier_1,Tr_S.Tier_2,Tr_S.Tier_3,Tr_S.Tier_4,Tr_S.Final_Tier 
from Approval Aprv,tier_setup Tr_S where Tr_S.Policy_Code =
	(select CP.Policy_Code from Company_policy CP where Policy_type like '%Leave%' and code=
		(select DI.code from department_info DI where title=
			(Select JHC.Emp_dept from emp_job_hist_current JHC where Emp_ID=
				(select LI.Emp_ID from Emp_leave_info LI where App_ID=@App_ID))))




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



/*




DELETE FROM soft_pass where u_id <>'DSL'
DELETE FROM Emp_Att_Info
DELETE FROM  [dbo].[Emp_Per_Info]
DELETE FROM  [dbo].[Emp_Job_Hist_Current]



*/
CREATE   procedure POP_User

AS

select u_id,u_name,
permit = case 
	when cancel=0 then 'Yes' 
	else 'No' 
	end
	,cancel 
from soft_pass where u_name <>'DBA'





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO










/****** Object:  Stored Procedure dbo.POP_UserName_SPrivilege    Script Date: 1/1/02 10:42:34 AM ******/

/****** Object:  Stored Procedure dbo.POP_UserName_SPrivilege    Script Date: 12/30/01 6:57:03 PM ******/

/****** Object:  Stored Procedure dbo.POP_UserName_SPrivilege    Script Date: 12/27/01 6:38:32 PM ******/
CREATE PROCEDURE [POP_UserName_SPrivilege]

@Emp_ID varchar(10)
 AS



select u_name from soft_pass where u_id =@Emp_ID













GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Pay_Struc_POP    Script Date: 3/7/02 10:43:33 AM ******/
/****** Object:  Stored Procedure dbo.Pay_Struc_POP    Script Date: 10/28/01 4:34:26 PM ******/
/****** Object:  Stored Procedure dbo.Pay_Struc_POP    Script Date: 10/22/01 6:24:26 PM ******/
/****** Object:  Stored Procedure dbo.Pay_Struc_POP    Script Date: 9/17/00 4:34:01 PM ******/
/****** Object:  Stored Procedure dbo.Pay_Struc_POP    Script Date: 9/4/01 6:30:47 PM ******/
CREATE PROCEDURE [Pay_Struc_POP]
@Purpose varchar(10)
 AS
if @Purpose='Structure'
begin
	select	 Head_code,Head_name,Operation,Mode from Pay_Struc
end
-------------------------------------------------------------------------------------------------------------
declare @Blank_Field varchar (25)			--returns a blank field which does not exists
if @Purpose='Mapping'
begin
	
	select	 Head_code,Head_name,Custom_name = @Blank_Field from Pay_Struc
	where Mode='V'
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Payscale_I_U_D    Script Date: 3/7/02 10:43:33 AM ******/
/****** Object:  Stored Procedure dbo.Payscale_I_U_D    Script Date: 10/28/01 4:34:26 PM ******/
/****** Object:  Stored Procedure dbo.Payscale_I_U_D    Script Date: 10/22/01 6:24:26 PM ******/
/****** Object:  Stored Procedure dbo.Payscale_I_U_D    Script Date: 9/17/00 4:34:01 PM ******/
CREATE procedure Payscale_I_U_D
@opr varchar (1),
@Code varchar (10),
@Sal_Scale varchar (30),
@H_R varchar (4),
@Medical varchar (4) ,
@Conv Varchar (10),
@Tel Varchar (4),
@U_id varchar (10)
AS
if @opr='I'
begin
	insert into Payscale
		(Code,
		Sal_Scale,
		H_R,
		Medical,
		Conv,
		Tel,
		U_id)
	values
		(@Code,
		@Sal_Scale,
		@H_R,
		@Medical,
		@Conv,
		@Tel,
		@U_id)
end
if @opr='U'
begin
	Update Payscale set Sal_Scale=@Sal_Scale,H_R=@H_R,
		Medical=@Medical,Conv=@Conv,Tel=@Tel,U_id=@U_id 
		Where Code=@Code
end 
if @opr='D'
begin
	Delete from  Payscale 
		Where Code=@Code
end 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE     Procedure Payscale_Main_I_U_D(
    @opr varchar(1)   
   ,@Scale_Code int
   ,@Desig_Code varchar(10)
   ,@SB int
   ,@Incr int
   ,@EB int
   ,@EBIncr int
   ,@EBMax int
   ,@U_Id varchar(10)
) as

set nocount on

declare @Message varchar(150)

if @opr='I'

Begin


Set @Scale_Code=(select dbo.GetAuto_SlNo('Payscale'))

Insert Into Payscale_Main
(Scale_Code,Desig_Code,SB,Incr,EB,EBIncr,EBMax,U_Id) 
Values (@Scale_Code,@Desig_Code,@SB,@Incr,@EB,@EBIncr,@EBMax,@U_Id)

set @Message='Data saved successfully!'	

END

if @opr='U'

Begin

if not exists (select *  from increment_history
			 where scale_code=@scale_code)
	begin
		UPDATE Payscale_Main Set SB=@SB,Incr=@Incr,EB=@EB,EBIncr=@EBIncr,EBMax=@EBMax,U_Id=@U_Id
		where Scale_Code=@Scale_Code
		set @Message='Update  done successfully!'	
	end
	else
	
		set @Message='You can not update  payscale!'
	
END

if @opr='D'

Begin

	if not exists (select * from increment_history
				 where scale_code=@scale_code)
	begin

		Delete from Payscale_Main where @Desig_Code=Desig_Code 
		set @Message='Data deleted successfully!'	
	end

	else
	
		set @Message='You can not delete  payscale!'
	

END

Select Scale_Code=@Scale_Code,Message=@Message

set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE    Procedure Payscale_Sub_I_U_D(

   @Scale_Code int
   ,@Head_Code varchar(10)
   ,@Calc_Type int	
   ,@Prcnt_of_Basic int
) as 

if exists (select * from Payscale_Sub where Scale_Code=@Scale_Code
and Head_Code=@Head_Code)
begin
update Payscale_Sub set
	Calc_Type=@Calc_Type,Prcnt_of_Basic=@Prcnt_of_Basic 
	where Scale_Code=@Scale_Code and Head_Code=@Head_Code

end 
Else
begin
Insert Into Payscale_Sub(
   Scale_Code
   ,Head_Code
   ,Calc_Type
   ,Prcnt_of_Basic
) Values (
   @Scale_Code
   ,@Head_Code
   ,@Calc_Type	
   ,@Prcnt_of_Basic
)

end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Pop_Approval    Script Date: 3/7/02 10:43:46 AM ******/
/****** Object:  Stored Procedure dbo.Pop_Approval    Script Date: 10/28/01 4:34:40 PM ******/
/****** Object:  Stored Procedure dbo.Pop_Approval    Script Date: 10/22/01 6:24:39 PM ******/
/****** Object:  Stored Procedure dbo.Pop_Approval    Script Date: 9/17/00 4:34:12 PM ******/
/****** Object:  Stored Procedure dbo.Pop_Approval    Script Date: 9/4/01 6:30:57 PM ******/
CREATE Procedure Pop_Approval
as
	select 
	
	b.App_ID,b.Leave_name,b.App_Type,b.Start_dt,b.End_dt,Total=datediff(day,b.Start_dt,b.End_dt),b.Des,b.Address,b.Tel1,b.Dt,
	a.emp_id,
	nm=(a.emp_fna+" " + a.emp_mna + " "+ a.emp_lna),
	c.Emp_desig,
	c.Emp_dept,
	c.Emp_section,
	c.Emp_job_type,
	d.Duration
	
	
	from emp_per_info a,Emp_Leave_info b,Emp_Job_hist_current c,Leave_list d
		where exists(select * from emp_leave_info
			where a.emp_id=b.emp_id and a.emp_id=c.emp_id 
			and b.emp_id=c.emp_id)
			and d.Leave_name=b.Leave_name order by Start_dt




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Pop_Rpt_Holiday_List    Script Date: 3/7/02 10:43:37 AM ******/
CREATE PROCEDURE [Pop_Rpt_Holiday_List]
@Id as varchar(10)
 AS
if @Id='All'
	begin
		
		select Hol_name,Hol_Desc,Str_date=min(Str_date),End_date=max(End_date),
			DURATION=DATEDIFF(DAY,min(Str_date),max(End_date))+1,Category from Hol_list
			 group by Hol_name,Hol_Desc,Category
		
	end
if @Id='Rpt'
	begin
	
		select Hol_name,Str_date=min(Str_date),End_date=max(End_date),
		DURATION=DATEDIFF(DAY,min(Str_date),max(End_date))+1,Category from Hol_list
		where datepart (year,str_date)=datepart (year,getdate()) and category !='Week end' 
		group by Hol_name,Category
	end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO









/*

Pro_Param_Entry 'all'
select * from param_tbl
Module,description
*/


CREATE         Procedure Pro_Param_Entry 
@Mode varchar(5)
AS

set nocount on
if @Mode='All'

Begin

declare @Module varchar(20)

delete from param_tbl

-------------------Security---------------------------------------------------------
set @Module='Security'

insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,10,'Access_Log',1,'0','')
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,11,'Event_Log',1,'0','')
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,12,'Erase_Log',1,'7','')
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,13,'PW_Cng_Notice',1,'21','')
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,14,'Min_PW_Len',1,'3','')
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,15,'Max_PW_Len',1,'8','')
-------------------Attendance---------------------------------------------------------
set @Module='Attendance'

insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,20,'TodaysAttnRpt',1,'0','')
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,21,'Attn_PW',1,'0','')

-------------------Discount--(for HMS only)-------------------------------------------------------
set @Module='Discount'

insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,30,'FO_Discount',0,'0','')
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,31,'Rst_Gst_Discount',0,'0','')
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,32,'Rst_Stf_Discount',0,'0','')

-------------------Payroll------------------------------------------------------------------------

set @Module='Payroll'

insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,60,'Ignore_Att_Record',0,'0','')
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,61,'GetFrom_Preceding_Month',1,'0','')
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,62,'AbsentForLate',1,'3','')
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,63,'Eid_Bonus_Prcnt',0,'50','')---0--Percent of Basic,1--Percent of gross sal
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,64,'Eid_Bonus_Deservs',1,'12','')---1--Permanent and job duration =12 months
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,65,'AGM_Prcnt',0,'50','') ---0--Percent of Basic,1--Percent of gross sal
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,66,'AGM_Deservs',1,'12','')--0 all employees--1--Permanent and job duration =12 months
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,67,'BTEXPO_Prcnt',0,'50','')---0--Basic,1 gross sal
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,68,'BTEXPO_Deservs',1,'12','')--0 all employees--1--Permanent and job duration =12 months 
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,69,'OT_Deservs',0,'Permanent','')---0 Wage drawer,
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,70,'Election_Alwnc_Prcnt',0,'20','')---0 Wage drawer,
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,71,'Election_Alwnc_Deserves',1,'12','')---0 Wage drawer,
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,72,'Travel_tour_Alwnc',1,'1','')---0 Wage drawer,
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,73,'Incentive',1,'1','')--0 all employees--1--Permanent
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,74,'OT_Tm_Breakup',1,'15','')---0 no break up,1 has break up,'15'minutes break up,
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,75,'OT_Pay_Prcnt',0,'200','')---0-- percent of basic,1-- percent of gross,200% 
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,76,'1_Day_Salary_factor',0,'0','')---0 Total salary divided by days of the month
											--1--Total salary divided by total Working days,
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,77,'Send_Payment_to_Accounts',0,'0','')---
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,78,'InformationAutoUpdate',0,'0','')---
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,79,'AutoDisbursment',1,'7','')---

-------------------Printer-------------------------------------------------------------------------
set @Module='Printing'
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,90,'PrinterPreference',1,'0','')
insert into Param_Tbl (Module,Policy_No,Policy,Flag,value,description)values(@Module,91,'ReportPrining',1,'0','')


end

Select Message='Default settings have been restored !'

set nocount off












GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  PROCEDURE Promotion_Record_Save
	@Emp_ID varchar(10),
	@Pre_Desig  varchar(50),
	@Pre_Dept varchar(50),
	@Pre_Scale varchar(25),
	@Last_Pro_Dt datetime,
	@C_Desig varchar(50),
	@C_Dept varchar(50),
	@C_Scale varchar(25),
	@Pro_Effe_Dt datetime,
	@Entry_Dt datetime
		
as 

If Exists (Select * From Promotion_Record 
    Where  emp_id=@Emp_Id and Pro_Effe_Dt=@Pro_Effe_Dt)
	begin		
		Update Promotion_Record  Set 
			Pre_Desig=@Pre_Desig ,
			Pre_Dept=@Pre_Dept ,
			Pre_Scale=@Pre_Scale ,
			Last_Pro_Dt=@Last_Pro_Dt,
			C_Desig=@C_Desig,
			C_Dept=@C_Dept,
			C_Scale=@C_Scale
			   Where  emp_id=@Emp_Id and Pro_Effe_Dt=@Pro_Effe_Dt
	end
Else

Insert Into Promotion_Record (
	Emp_ID ,
	Pre_Desig ,
	Pre_Dept ,
	Pre_Scale ,
	Last_Pro_Dt ,
	C_Desig ,
	C_Dept ,
	C_Scale ,
	Pro_Effe_Dt ,
	Entry_Dt 
) Values (
	@Emp_ID ,
	@Pre_Desig ,
	@Pre_Dept ,
	@Pre_Scale ,
	@Last_Pro_Dt ,
	@C_Desig ,
	@C_Dept ,
	@C_Scale ,
	@Pro_Effe_Dt ,
	@Entry_Dt 
)

Update Promotion_Record set C_Scale=@C_Scale
		 Where  emp_id=@Emp_Id








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







CREATE    procedure PublisherInformation(
	@PubCode	VARCHAR(6),
	@PubName	VARCHAR(50),
	@Address VARCHAR (50),
	@Country VARCHAR(100),
	@Remarks VARCHAR (100),
	@User VARCHAR (10))
AS
SET NOCOUNT ON
DECLARE
@CountryCode VARCHAR (5),
@Msg VARCHAR (100)


SELECT @CountryCode = CntCoutryCode FROM Country WHERE (CntCountryName = @Country)

IF @CountryCode IS NULL
	BEGIN	
		SET @Msg='Invalid Country Name!'
		GOTO Errmsg
	END

IF NOT EXISTS (SELECT * FROM PublisherInfo WHERE     (PubCode = @PubCode))
	INSERT INTO PublisherInfo
        (PubCode, PubPublisherName, PublisherAdress, PubCountryCode, Remarks, PubEntryDate, PubEntryBy)
		VALUES (@PubCode,@PubName,@Address,@CountryCode,@Remarks,GETDATE(),@User)
ELSE
	UPDATE 	PublisherInfo SET              
			PubPublisherName = @PubName, PublisherAdress = @Address, PubCountryCode = @CountryCode, Remarks = @Remarks
			WHERE (PubCode = @PubCode)

RETURN
Errmsg:
	RAISERROR(@Msg,16,1)
	RETURN
SET NOCOUNT OFF



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Rank_I_U_D    Script Date: 3/7/02 10:43:37 AM ******/
/****** Object:  Stored Procedure dbo.Rank_I_U_D    Script Date: 10/28/01 4:34:31 PM ******/
/****** Object:  Stored Procedure dbo.Rank_I_U_D    Script Date: 10/22/01 6:24:31 PM ******/
/****** Object:  Stored Procedure dbo.Rank_I_U_D    Script Date: 9/17/00 4:34:03 PM ******/
/****** Object:  Stored Procedure dbo.Rank_I_U_D    Script Date: 9/4/01 6:30:48 PM ******/
--***********************
CREATE PROCEDURE Rank_I_U_D
@opr varchar (1),
@Title varchar (25),
@Prev_Title varchar (25),
@Description varchar (50),
@U_id varchar (10)
AS
if @opr='I'
begin
insert into Rank(
Title,
Description,
U_id
) 
values (
@Title,
@Description,
@U_id
)
end
                    
if @opr='U'
begin
update Rank set 
Title=@Title,
Description=@Description,
U_id=@U_id
where Title=@Prev_Title
end
    		
if @opr='D'
begin
delete from Rank where title=@Title
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/*
select Tot_Num_Incr= sum(Num_Incr)from  increment_history 
where Emp_Id='9-027' and Effect_Dt <=getdate()

Renew_FixedPay 'dsl'
select * from increment_history
select * from fixed_pay where emp_id='9-026'
update increment_history set done=0 where  track_id>=25

delete from emp_salary
delete from payroll_main


*/
----Increment_History_I_U '9-026',4,'15600','2002-07-03'
CREATE    Procedure Renew_FixedPay
@U_Id varchar(10)
AS
set nocount on

Declare 
@Emp_Id varchar(10) 
,@Effect_dt datetime
,@Scale_Code int
,@SB int
,@Incr_Amount int
,@Num_Prev_Incr int
,@Num_Total_Incr int
,@Amount money
,@Count int

set @Count=0

declare Incr_hist Cursor for

select  Emp_Id,Effect_dt from increment_history 
where Effect_Dt <= dbo.getdateonly(getdate()) and Done=0

-------------------------------------------------
open Incr_hist 

	fetch next from Incr_hist into @Emp_Id,@Effect_dt
	
	while @@Fetch_Status=0
	begin
		-----Get present payscale code
		set @Scale_Code=(select Scale_Code from Emp_payscale_hist
			where Emp_Id=@Emp_Id and Track_Id= 
				(select max(Track_Id)from Emp_payscale_hist where emP_Id=@Emp_Id))

		-----Get Starting Basic and Increment Amount of the present payscale
		select @SB=SB,@Incr_Amount=Incr from Payscale_Main 
		where Scale_Code=@Scale_Code

		-----Calculation
		---@Num_Total_Incr---->Number of total increment upto date
		set @Num_Total_Incr= (select sum(Num_Incr)from 
		increment_history where Scale_Code=@Scale_Code and Emp_ID=@Emp_ID)

		----@Amount-----> Becomes the current Basic including today's increment
		set @Amount=@SB+(@Incr_Amount* @Num_Total_Incr)

		-----Insert Latest Basic Pay including today's increment 

		declare @VarAmount varchar(10)
		set @VarAmount = (select convert(varchar(10),@Amount))
		
		Exec Insert_Into_Fixed_Pay @Emp_Id,@Scale_Code,@VarAmount,@Effect_Dt,@U_Id

		set @Count=@Count+1

		fetch next from Incr_hist into @Emp_Id,@Effect_dt
	end
	
close Incr_hist 
deallocate Incr_hist 

Declare @Done varchar(40) 

if @Count=0 
	set @Done=' '
else
	set @Done= '('+ rtrim(convert(char(5),@Count)) + ') Salary information updated !'

select Done=@Done

set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





---select studentname,studentid from studentinfo 



/****** Object:  Stored Procedure dbo.BookIssueReturn_Save    Script Date: 29/04/2006 12:41:03 ******/


CREATE     Procedure Result_Save
(
		@mode 					varchar(1),
	    @M_Slr_no 				int,
		@S_Slr_no				int,
	    @ClassID 				varchar(10),
		@SectionID				varchar(10),
		@Shift					varchar(1),
		@SubID					varchar(10),
		@AcaYr					varchar(5),
		@ExamType				varchar(5),
		@ExamID					varchar(5),
        @categoryid             varchar(7),
		@StdID					varchar(10),
		@Roll						int,
		@Obtain_Marks				decimal(5,2),
        @pass_Marks					int,
        @full_Marks					int,
		@EntryBy					varchar(5),
		@EntryDate				datetime
) 

as 

SET XACT_ABORT On

Declare @Max_M_Slr_no as int,
		  @Max_S_Slr_no as int

Begin Tran

if @mode='S'-- Here - Means New Data is to be save or updated
	Begin
	
		If (@M_Slr_no=0)
	
			begin
						
				select @Max_M_Slr_no = isnull(max(M_Slr_no),0) from Result_Main
				set @Max_M_Slr_no = @Max_M_Slr_no + 1
				
				Insert Into Result_Main
					(
					  	M_Slr_no,
						ClassID,
						SectionID,
						Shift,
						SubID,
						AcaYr,
					   ExamType,
						ExamID,
                        categoryid
					 ) 
				Values 
					(
					   @Max_M_Slr_no,
					   @ClassID,
					   @SectionID,
						@Shift,
						@SubID,
						@AcaYr,
						@ExamType,
						@ExamID,
                        @categoryid
					 )

				select @Max_S_Slr_no = isnull(max(S_Slr_no),0) from Result_Sub
				set @Max_S_Slr_no = @Max_S_Slr_no + 1
						
				Insert Into Result_Sub
					(
					   M_Slr_no,	
                       S_Slr_no,
						StdID,
						Roll,
						ObtainedMarks,
                        PassMarks,
                        Fullmarks, 
						EntryBy,
						EntryDate
					) 
				Values 
				  (
						@Max_M_Slr_no,
                        @Max_S_Slr_no,
						@StdID,
						@Roll,
						@Obtain_Marks,
                       @pass_Marks	,
                       @full_Marks	,
						@EntryBy,
						@EntryDate
					)
		end

		if (@M_Slr_no<>0)
			begin
				select @Max_S_Slr_no = isnull(max(S_Slr_no),0) from Result_Sub
				set @Max_S_Slr_no = @Max_S_Slr_no + 1
						
				Insert Into Result_Sub
					(
                        M_Slr_no,
						S_Slr_no,
						StdID,
						Roll,
						ObtainedMarks,
                        PassMarks,
                        Fullmarks, 
						EntryBy,
						EntryDate
					) 
				Values 
				  (
						@M_Slr_no,
						@Max_S_Slr_no,
						@StdID,
						@Roll,
						@Obtain_Marks,
                       @pass_Marks	,
                       @full_Marks	,
						@EntryBy,
						@EntryDate
					)		
			end

end

if @mode='U'-- Here - Means previous data is to be deleated

	begin
		If Exists (Select S_Slr_no From Result_Sub Where S_Slr_no=@S_Slr_no and M_Slr_no=@M_Slr_no)

			Update Result_Sub 
  
			Set  obtainedMarks      = @Obtain_Marks,
				 EntryBy	= @EntryBy,
			 	EntryDate	= @EntryDate
 	     Where M_Slr_no=@M_Slr_no and S_Slr_no=@S_Slr_no
	end


if @mode='D'-- Here - Means previous data is to be deleated

	begin
		If Exists (Select S_Slr_no From Result_Sub Where S_Slr_no=@S_Slr_no and M_Slr_no=@M_Slr_no)

		Delete From Result_Sub where S_Slr_no=@S_Slr_no and M_Slr_no=@M_Slr_no
	end


if @@Error=0
	Commit
else
	Rollback






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE PROCEDURE RprLogin_Fail
@mode int,
@emp_id varchar(40),
@month varchar(20),
@year int,
@date datetime
 AS
if @mode=1 ----for all and specific month and year
begin
	select a.u_id,a.u_name,a.user_pass,a.udt,b.emp_desig,b.emp_dept
	from login_fail a, Emp_Job_Hist_Current b where a.u_id=b.emp_id and datepart(year,a.udt)=@year
	and datename(month,a.udt)= @month
end
if @mode=2  -----for specific emp and specific month and year
begin
	select a.u_id,a.u_name,a.user_pass,a.udt,b.emp_desig,b.emp_dept
	from login_fail a, Emp_Job_Hist_Current b where a.u_id=b.emp_id and a.u_id=@emp_id and datepart(year,a.udt)=@year
	and datename(month,a.udt)= @month


end
if @mode=3 ----for all emp and specific day
begin
	select a.u_id,a.u_name,a.user_pass,a.udt,b.emp_desig,b.emp_dept
	from login_fail a, Emp_Job_Hist_Current b where a.u_id=b.emp_id  and datepart(year,a.udt)=@year
	and datename(month,a.udt)= @month and day(udt)=day(@date)
end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




Create  PROCEDURE RptBookStockInfo

@mode int,
@bookname varchar(50)
AS

------for specific book
if @mode=1
	Select count(BookCode) from LibraryBookList where BookCode =(Select BookCode from librarybookinfo where BookName=@bookname)
------for all book
if @mode=1
	
	select distinct BookCode,count(BookCode)from LibraryBookList group by BookCode



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




----exec RptPromotion_Info '9-027'

CREATE PROCEDURE RptPromotion_Info
	@Emp_Id varchar (10)
	
AS
begin

set nocount on

	SELECT Promotion_Record.Emp_Id, Emp_Per_Info.Emp_fna+ ' ' + Emp_Per_Info.Emp_mna+ ' ' +
	Emp_Per_Info.Emp_lna as Employee_Name, 
	Promotion_Record.Last_Pro_Dt as Last_Prom_Date,Promotion_Record.Pre_Desig as Prev_Desig,
	Promotion_Record.C_Desig as Curr_Desig, Promotion_Record.Pre_Dept as Prev_Dept,
	Promotion_Record.C_Dept as Curr_Dept, Promotion_Record.Pro_Effe_Dt as Pro_Effective_Dt, 
	Promotion_Record.Pre_Scale as Prev_Scale, Promotion_Record.C_Scale as Curr_Scale
	FROM  Promotion_Record INNER JOIN
	Emp_Per_Info ON Promotion_Record.Emp_Id = Emp_Per_Info.Emp_id
	where Promotion_Record.Emp_Id=@Emp_Id

end



set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







----exec Rpt_Bonus_Preparation '','May','2005'


CREATE   PROCEDURE Rpt_Bonus_Preparation
	@Emp_Id varchar (10),
	@pay_Month varchar(9),
	@pay_Year char(4)
	
	
AS
begin
if  @Emp_Id <> ''
			
			SELECT  BonusPreparation.Emp_id, Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS emp_Name, 
			Emp_Job_Hist_Current.Emp_desig, BonusPreparation.PayMonth, BonusPreparation.PayYear, BonusPreparation.Amount
			FROM  BonusPreparation INNER JOIN
			Emp_Per_Info ON BonusPreparation.Emp_id = Emp_Per_Info.Emp_id INNER JOIN
			Emp_Job_Hist_Current ON Emp_Per_Info.Emp_id = Emp_Job_Hist_Current.Emp_id
			where  (BonusPreparation.Emp_id=@Emp_Id)
					and (BonusPreparation.PayMonth=@pay_Month) 
					and (BonusPreparation.PayYear=@pay_Year)
					
else
		SELECT  BonusPreparation.Emp_id, Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS emp_Name, 
		Emp_Job_Hist_Current.Emp_desig, BonusPreparation.PayMonth, BonusPreparation.PayYear, BonusPreparation.Amount
		FROM BonusPreparation INNER JOIN Emp_Per_Info ON BonusPreparation.Emp_id = Emp_Per_Info.Emp_id INNER JOIN
		Emp_Job_Hist_Current ON Emp_Per_Info.Emp_id = Emp_Job_Hist_Current.Emp_id
		where  (BonusPreparation.PayMonth=@pay_Month) 
		and (BonusPreparation.PayYear=@pay_Year) 

end



set nocount off










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







--Rpt_Daily_Attendance '2002-04-08' 
/****** Object:  Stored Procedure dbo.Rpt_Daily_Attendance    Script Date: 3/7/02 10:43:47 AM ******/
CREATE     PROCEDURE Rpt_Daily_Attendance 
@Entry_Dt varchar(12)
AS
SELECT DISTINCT a.emp_id,Emp_name=(a.Emp_fna + " " + a.Emp_mna +" "+ a.Emp_lna),
		b.emp_desig,b.emp_dept,Emp_login=convert(char(8),c.Emp_login,14),
		In_Remarks=case when c.In_Status='1' then '*'else null end,
		Emp_logout=convert(char(8),c.Emp_logout,14)
		,Notes=(select dbo.GetAttnNotes(c.Emp_login))
		,Out_Remarks=case when c.Out_Status='8' then '*'else null end,
		c.Entry_Dt				
		from emp_Per_info a,Emp_Job_Hist_Current b, Emp_Att_Info c where 
		a.Emp_ID=B.Emp_ID and a.Emp_ID=c.Emp_ID and 
		datepart(day,c.Entry_Dt)=datepart(day,@Entry_Dt)and
		datepart(Month,c.Entry_Dt)=datepart(Month,@Entry_Dt)and
		datepart(year,c.Entry_Dt)=datepart(year,@Entry_Dt)
		







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO








/*

	exec Rpt_Daily_Not_Present '2003-12-04'

*/


CREATE      PROCEDURE Rpt_Daily_Not_Present
@Entry_Dt varchar(12)
AS

SELECT a.emp_id,Emp_name=(a.Emp_fna + ' ' + a.Emp_mna +' '+ a.Emp_lna)
	,b.emp_desig,b.emp_dept
	,Remarks=dbo.Get_If_In_Leave(b.emp_id,@Entry_Dt) 
	,[Date]=convert(datetime,(@Entry_Dt))
	from	emp_Per_info a,Emp_Job_Hist_Current b
		where a.emp_Id=b.emp_Id 		
			and b.emp_Id not in 
			(select Emp_id from Emp_Att_Info
				Where datepart(day,Entry_Dt)=datepart(day,@Entry_Dt)
					and datepart(Month,Entry_Dt)=datepart(Month,@Entry_Dt)
					and datepart(year,Entry_Dt)=datepart(year,@Entry_Dt))


	order by a.Emp_id







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





----exec Rpt_Deptwise_Leave 'Daffodil Software Ltd.','12/12/1995','12/12/2005'
---exec Rpt_Deptwise_Leave  'Corporate Sales','01/26/2004','26/02/2009'

CREATE   PROCEDURE Rpt_Deptwise_Leave
	@Dept_Nm varchar (50),
	@LeaveStarTm datetime,
	@LeaveEndTm  datetime
	
	
AS
begin
if  @Dept_Nm = ''
			
		SELECT Emp_Leave_Info.Emp_id, Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Emp_Name, 
		Emp_Leave_Info.Leave_name, Emp_Job_Hist_Current.Emp_desig, Emp_Job_Hist_Current.Emp_dept, Emp_Leave_Info.Start_dt, 
		/*No_Of_Emp=(SELECT count(Emp_id) AS No_Of_Emp FROM  Emp_Job_Hist_Current WHERE     (Emp_dept = @Dept_Nm)	),*/
		Emp_Leave_Info.End_dt, Emp_Leave_Info.Super_ID
		FROM Emp_Leave_Info INNER JOIN Emp_Job_Hist_Current ON Emp_Leave_Info.Emp_id = Emp_Job_Hist_Current.Emp_id INNER JOIN
		Emp_Per_Info ON Emp_Leave_Info.Emp_id = Emp_Per_Info.Emp_id
		WHERE (Emp_Leave_Info.Start_dt BETWEEN @LeaveStarTm AND @LeaveEndTm) 
					
else
		SELECT Emp_Leave_Info.Emp_id, Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Emp_Name, 
		Emp_Leave_Info.Leave_name, Emp_Job_Hist_Current.Emp_desig, Emp_Job_Hist_Current.Emp_dept, Emp_Leave_Info.Start_dt,
		/*No_Of_Emp=(SELECT count(Emp_id) AS No_Of_Emp FROM  Emp_Job_Hist_Current WHERE     (Emp_dept = @Dept_Nm)	),	 */
		Emp_Leave_Info.End_dt, Emp_Leave_Info.Super_ID
		FROM Emp_Leave_Info INNER JOIN Emp_Job_Hist_Current ON Emp_Leave_Info.Emp_id = Emp_Job_Hist_Current.Emp_id INNER JOIN
		Emp_Per_Info ON Emp_Leave_Info.Emp_id = Emp_Per_Info.Emp_id
		WHERE (Emp_Leave_Info.Start_dt BETWEEN @LeaveStarTm AND @LeaveEndTm) 
		AND (Emp_Job_Hist_Current.Emp_dept = @Dept_Nm)

end



set nocount off














GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





----exec Rpt_Emp_Increment_Info '9-023','1995-12-01','2008-12-12'


CREATE   PROCEDURE Rpt_Emp_Increment_Info
	@Emp_Id varchar (10),
	@Fromdate datetime,
	@Todate datetime
	
AS
begin
if  @Emp_Id <> ''
		
		SELECT EmpIncrementInformationInfo.Emp_ID, Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Emp_Name, 
		Emp_Job_Hist_Current.Emp_desig, Emp_Job_Hist_Current.Emp_dept, Emp_Job_Hist_Current.CBasic, EmpIncrementInformationInfo.IncrementAmount, 
		EmpIncrementInformationInfo.LastIncrementDt, EmpIncrementInformationInfo.NextIncremntDt, EmpIncrementInformationInfo.EffectiveDt, 
		EmpIncrementInformationInfo.Track_Id
		FROM EmpIncrementInformationInfo INNER JOIN
		Emp_Per_Info ON EmpIncrementInformationInfo.Emp_ID = Emp_Per_Info.Emp_id INNER JOIN
		Emp_Job_Hist_Current ON Emp_Per_Info.Emp_id = Emp_Job_Hist_Current.Emp_id
		WHERE (EmpIncrementInformationInfo.Emp_ID =@Emp_Id)
		and EmpIncrementInformationInfo.EffectiveDt between @Fromdate and @Todate
		order by EmpIncrementInformationInfo.Track_Id desc 

else
		SELECT EmpIncrementInformationInfo.Emp_ID, Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS Emp_Name, 
		Emp_Job_Hist_Current.Emp_desig, Emp_Job_Hist_Current.Emp_dept, Emp_Job_Hist_Current.CBasic, EmpIncrementInformationInfo.IncrementAmount, 
		EmpIncrementInformationInfo.LastIncrementDt, EmpIncrementInformationInfo.NextIncremntDt, EmpIncrementInformationInfo.EffectiveDt, 
		EmpIncrementInformationInfo.Track_Id
		FROM  EmpIncrementInformationInfo INNER JOIN
		Emp_Per_Info ON EmpIncrementInformationInfo.Emp_ID = Emp_Per_Info.Emp_id INNER JOIN
		Emp_Job_Hist_Current ON Emp_Per_Info.Emp_id = Emp_Job_Hist_Current.Emp_id
		WHERE  EmpIncrementInformationInfo.EffectiveDt between @Fromdate and @Todate
				order by EmpIncrementInformationInfo.Track_Id desc 
end



set nocount off








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




--exec Rpt_Employee_Info_Various 'Ind','9-031'


CREATE   procedure Rpt_Employee_Info_Various 
	@Mode varchar (10),				/*@Mode='Ind'	/ 'All' / ' Desig' / 'Dept' / 'BldGr'/ 'JbTp'	*/
	@Param_1 Varchar(40)				/*@Param_1=Emp_ID / Emp_Desig / Emp_Dept / Blood_Group / Emp_Job_Type*/
as
/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/
if @Mode='Ind'
begin

	select  a.Emp_id,
		Emp_fna=(a.Emp_fna+" "+a.Emp_mna+" "+a.Emp_lna),
		a.Emp_mna,a.Emp_lna,	
		a.Emp_fa_na,a.Emp_ma_na,a.Emp_d_of_b,a.Emp_age,
		a.Blood_group,a.Emp_marital_st,a.Emp_gender ,a.Emp_nat ,
		a.Emp_sp_qta ,a.Emp_religion ,a.Emp_perm_add ,a.Emp_perm_town ,
		a.Emp_post1 ,a.Emp_tel1 ,Country=a.Contact_person,a.Contact_add,a.Contact_town,
		a.Contact_post,	a.Contact_tel,a.Contact_fax,a.Contact_email,a.Emp_eye,
		a.Emp_height,a.Emp_weight,a.Emp_disable,b.Emp_join_date,b.Emp_desig,b.Emp_Dept
		,c.Exam_Name,c.M_Subject,c.Pass_Year,c.Degree_From,c.Result

	from Emp_Per_info a,Emp_Job_Hist_current b
		, emp_education c

		Where a.Emp_id=b.Emp_id 
			and a.Emp_id=c.Emp_id 	
			and (c.Exam_Name !=null 
			or c.Exam_Name !='')
			and a.Emp_id=@Param_1
				
end		




if @mode='Edu'

begin
	select * from emp_education --where 
	--emp_id='1108' 
	--and (Exam_Name !=null 
		--or Exam_Name !='')
	order by Pass_Year desc
end

if @mode='Ref'

begin
	select * from emp_reference
	order by Ref_Count
end



/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/
if @Mode='All'
begin
	select
		a.Emp_id,Emp_Name=(Emp_fna+" "+Emp_mna+" "+Emp_lna),
		a.Emp_fa_na,a.Emp_d_of_b,a.Emp_age,	a.Blood_group,	a.Emp_gender ,
		a.Emp_religion ,a.Emp_perm_add ,a.Emp_perm_town ,a.Emp_tel1 ,a.Contact_person,
		a.Contact_tel,b.Emp_desig,Emp_Dept=b.Emp_Dept,b.Emp_join_date,
		b.Emp_section

		from Emp_Per_info a,Emp_Job_Hist_current b
		Where a.Emp_id=b.Emp_id
		order by a.Emp_ID

end
/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/
if @Mode='Desig'
begin
	select
		a.Emp_id,Emp_Name=(a.Emp_fna+" "+a.Emp_mna+" "+a.Emp_lna),	
		a.Emp_perm_add ,a.Emp_perm_town ,a.Emp_tel1,
		b.Emp_join_date,'Designation: '+b.Emp_desig,b.Emp_Dept
		from emp_per_info a,emp_job_hist_current b
		where a.emp_id=b.emp_id and b.Emp_desig=@param_1
ORDER BY a. emp_id 
end
/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/
if @Mode='Dept'
begin
	select
		a.Emp_id,Emp_Name=(a.Emp_fna+" "+a.Emp_mna+" "+a.Emp_lna),	
		a.Emp_perm_add ,a.Emp_perm_town ,a.Emp_tel1,b.Emp_join_date,
		a.Emp_d_Of_b,b.Emp_desig,'Department: '+b.Emp_Dept
		
		from emp_per_info a,emp_job_hist_current b
		where a.emp_id=b.emp_id and b.Emp_dept=@Param_1
ORDER BY a. emp_id 
end
/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/
if @Mode='BldGr'
begin
	select
		a.Emp_id,Emp_Name=(a.Emp_fna+" "+a.Emp_mna+" "+a.Emp_lna),	
		a.Emp_perm_add ,a.Emp_perm_town ,a.Emp_tel1 ,
		a.Blood_group,b.Emp_desig,b.Emp_Dept
		from emp_per_info a,emp_job_hist_current b
		
		where a.emp_id=b.emp_id and a.Blood_group=@Param_1
ORDER BY a. emp_id 
end
/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/
if @Mode='JbTp'
begin
	select
		a.Emp_id,Emp_Name=(a.Emp_fna+" "+a.Emp_mna+" "+a.Emp_lna),	
		a.Emp_perm_add ,a.Emp_perm_town ,a.Emp_tel1 ,
		b.Emp_desig,b.Emp_Dept,b.Emp_join_date,b.Emp_job_type
		from emp_per_info a,emp_job_hist_current b
		where a.emp_id=b.emp_id and b.Emp_job_type=@Param_1
ORDER BY a. emp_id 
end

/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/

if @Mode='Sal'
begin
	SELECT a.emp_id,
	Emp_name=(select Emp_fna+' '+Emp_mna+' '+Emp_lna from Emp_Per_info where emp_id=a.emp_id),
	Dept='Department: '+(select emp_dept from emp_Job_hist_current where emp_id=a.emp_id),
	Desig=(select emp_desig from emp_Job_hist_current where emp_id=a.emp_id),
	Joining_Dt=(select Emp_join_date from emp_Job_hist_current where emp_id=a.emp_id),
	catagory=(select Emp_rank from emp_Job_hist_current where emp_id=a.emp_id),
	
	sum(CASE b.head_name WHEN 'Basic Pay' THEN a.amount ELSE 0 END) AS  Basic,
	sum(CASE b.head_name WHEN 'House Rent' THEN a.amount ELSE 0 END) AS H_Rent,
	sum(CASE b.head_name WHEN 'Conveyance' THEN a.amount ELSE 0 END) AS  Conveyance,
	sum(CASE b.head_name WHEN 'Medical Allowance' THEN a.amount  ELSE 0 END) AS Med_Allow

	from fixed_Pay a,pay_struc b ,emp_Job_hist_current C 
	where b.head_code=a.head_code and A.Emp_Id=C.Emp_Id and c.emp_dept=@param_1
		
	group by a.emp_id
	order by a.emp_id
end



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




----exec Rpt_Faculy_Schedule_Preparation '9-031'


CREATE     PROCEDURE Rpt_Faculy_Schedule_Preparation
	@Emp_Id varchar (10)
	
AS
begin
if  @Emp_Id <> ''

		SELECT   distinct (a.emp_id) as Emp_Id,
		emp_name=(select  Emp_fna + ' ' + Emp_mna + ' ' + Emp_lna as Emp_Name FROM Emp_Per_Info where a.Emp_Id =Emp_Per_Info.emp_id),
		Saturday=(select ScheduleStTime+ ' - ' + ScheduleEndTime FROM  FacultySchedule where a.Emp_Id=FacultySchedule.emp_id
		and FacultySchedule.DayoftheSchedule='Saturday'),
		Sunday=(select ScheduleStTime+ ' - ' + ScheduleEndTime FROM  FacultySchedule where a.Emp_Id=FacultySchedule.emp_id
		and FacultySchedule.DayoftheSchedule='Sunday'),
		Monday=(select ScheduleStTime+ ' - ' + ScheduleEndTime FROM  FacultySchedule where a.Emp_Id=FacultySchedule.emp_id
		and FacultySchedule.DayoftheSchedule='Monday'),
		Tuesday=(select ScheduleStTime+ ' - ' + ScheduleEndTime FROM  FacultySchedule where a.Emp_Id=FacultySchedule.emp_id
		and FacultySchedule.DayoftheSchedule='Tuesday'),
		Wednesday=(select ScheduleStTime+ ' - ' + ScheduleEndTime FROM  FacultySchedule where a.Emp_Id=FacultySchedule.emp_id
		and FacultySchedule.DayoftheSchedule='wednesday'),
		Thursday=(select ScheduleStTime+ ' - ' + ScheduleEndTime FROM  FacultySchedule where a.Emp_Id=FacultySchedule.emp_id
		and FacultySchedule.DayoftheSchedule='Thursday')
		FROM  Emp_Per_Info ,FacultySchedule a
		where a.emp_id=@Emp_Id
		
else
		SELECT   distinct (a.emp_id) as Emp_Id,
		emp_name=(select  Emp_fna + ' ' + Emp_mna + ' ' + Emp_lna as Emp_Name FROM Emp_Per_Info where a.Emp_Id =Emp_Per_Info.emp_id),
		Saturday=(select ScheduleStTime+ ' - ' + ScheduleEndTime FROM  FacultySchedule where a.Emp_Id=FacultySchedule.emp_id
		and FacultySchedule.DayoftheSchedule='Saturday'),
		Sunday=(select ScheduleStTime+ ' - ' + ScheduleEndTime FROM  FacultySchedule where a.Emp_Id=FacultySchedule.emp_id
		and FacultySchedule.DayoftheSchedule='Sunday'),
		Monday=(select ScheduleStTime+ ' - ' + ScheduleEndTime FROM  FacultySchedule where a.Emp_Id=FacultySchedule.emp_id
		and FacultySchedule.DayoftheSchedule='Monday'),
		Tuesday=(select ScheduleStTime+ ' - ' + ScheduleEndTime FROM  FacultySchedule where a.Emp_Id=FacultySchedule.emp_id
		and FacultySchedule.DayoftheSchedule='Tuesday'),
		Wednesday=(select ScheduleStTime+ ' - ' + ScheduleEndTime FROM  FacultySchedule where a.Emp_Id=FacultySchedule.emp_id
		and FacultySchedule.DayoftheSchedule='wednesday'),
		Thursday=(select ScheduleStTime+ ' - ' + ScheduleEndTime FROM  FacultySchedule where a.Emp_Id=FacultySchedule.emp_id
		and FacultySchedule.DayoftheSchedule='Thursday')
		FROM  Emp_Per_Info ,FacultySchedule a
				
end



set nocount off







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





/*
dbo.Salary_AtOnce_Unit1 'december','2001','dsl'

exec Rpt_YearlySalary_Statement '9-001',2001

exec Rpt_FiscalYr_SalStateAll '2002'

*/
CREATE Procedure Rpt_FiscalYr_SalStateAll
@pay_Year varchar(4) 
AS

set nocount on

Create table #Salary_All(
emp_id varchar(10),Emp_name varchar(45),Dept varchar(35)
,Category varchar(35),Desig varchar(35),Basic money
,H_Rent money,Conveyance money,Med_Allow money,Charge_Allow money
,Tech_Allow money,Others_Add money,Bonus money,Arrear money
,Abs_Deduc money,Adv_Deduc money,Others_Deduct money,IncomeTax money
,pay_Month varchar(12) ,pay_Year varchar(4) 

)

declare @emp_id varchar(10)

declare ID_Cursor cursor for
	select distinct emp_id from fixed_pay

Open  ID_Cursor

	fetch next from Id_Cursor into @emp_id

		WHILE @@FETCH_STATUS = 0

		BEGIN

			insert into #Salary_All
			exec Rpt_FiscalYr_Salary_Statement @Emp_Id,@Pay_Year		
      	                
			fetch next from Id_Cursor into @emp_id
		END
		




select * from #Salary_All 

close ID_Cursor
deallocate  ID_Cursor

set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/*
dbo.Salary_AtOnce_Unit1 'december','2001','dsl'

exec Rpt_YearlySalary_Statement '9-001',2001

exec Rpt_FiscalYr_Salary_Statement '9-001','2002'

*/
CREATE Procedure Rpt_FiscalYr_Salary_Statement
@emp_id varchar(10)
,@pay_Year varchar(4) 
AS

set nocount on

Create table #Salary(
emp_id varchar(10),Emp_name varchar(45),Dept varchar(35)
,Category varchar(35),Desig varchar(35),Basic money
,H_Rent money,Conveyance money,Med_Allow money,Charge_Allow money
,Tech_Allow money,Others_Add money,Bonus money,Arrear money
,Abs_Deduc money,Adv_Deduc money,Others_Deduct money,IncomeTax money
,pay_Month varchar(12) ,pay_Year varchar(4) 

)

declare @Month_Num  int
,@Month_Name varchar(12)
,@Fiscal_Year int

Set @Month_Num=7
Set @Fiscal_Year = convert(int,@pay_Year)-1

while @Month_Num <=12

BEGIN

	set @Month_Name=(select dbo.Month_Name(@Month_Num))

	insert into #Salary
	exec Rpt_Monthly_Salary_Statement @Emp_Id,@Month_Name,@Fiscal_Year 

	Set @Month_Num=@Month_Num+1	

		if @Month_Num>12 
		begin
			set @Month_Num=1
			set @Fiscal_Year=@Fiscal_Year+1
		end 
	
		if @Month_Num>6 and @Fiscal_Year=convert(int,@Pay_year)
	
		break

END


select * from #Salary 


set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







CREATE  PROCEDURE Rpt_Leave_Ind
	@Emp_Id varchar(10),
	@Leave_year varchar(4)
AS
set nocount on
create table #Lv_Rpt(
	Leave_name varchar(30)
	,Start_dt datetime
	,End_dt datetime
	,Duration int
	,Maximum int
	,Des varchar(40)
	,Address varchar(40)
	,Tel1 varchar(15)
	,Leave_year varchar(4)
	,emp_id varchar(10)	
	,Emp_fna varchar(45)	
	,emp_desig varchar(25)	
	,emp_dept varchar(25)
	,Emp_Section varchar(25)	
)
-----------------------Declare Local variable------------------------------------------------
	DECLARE @Maximum int 
	DECLARE @Leave_Name varchar(30)	
	DECLARE @Emp_fna varchar(45)	
	DECLARE @emp_desig varchar(25)	
	DECLARE @emp_dept varchar(25)
	DECLARE @Emp_Section varchar(25)	

select  @Emp_fna=(a.Emp_fna + " " + a.Emp_mna +" "+ a.Emp_lna),
	@emp_desig=b.emp_desig,@emp_dept=b.emp_dept,@Emp_Section=b.Emp_Section
	from emp_Per_info a,Emp_Job_Hist_Current b Where 
	a.Emp_ID=B.Emp_ID and A.Emp_ID=@Emp_ID

-----------------------Declare Cursor--------------------------------------------------------
DECLARE Leave_Cursor CURSOR FOR		
SELECT Distinct Leave_name,Duration from Leave_List
-----------------------Open Cursor-----------------------------------------------------------
OPEN Leave_Cursor
FETCH NEXT FROM Leave_Cursor INTO @Leave_Name,@Maximum
WHILE @@FETCH_STATUS = 0
BEGIN
--------------------------------------------------------------------------------------------
	Insert into #Lv_Rpt(Leave_name,Start_dt,End_dt,Duration,
	Maximum,Des,Address,Tel1,Leave_Year,emp_id,Emp_fna,emp_desig,emp_dept,Emp_Section)
	SELECT Leave_name,Start_dt=min(Start_dt),End_dt=max(End_dt),
	Duration=(datediff(day,min(Start_dt),max(End_dt))+1),Maximum=@Maximum,
	Des,Address,Tel1,Leave_Year=@Leave_Year,emp_id=@emp_id,Emp_fna=@Emp_fna,emp_desig=@emp_desig,emp_dept=@emp_dept,Emp_Section=@Emp_Section
	from Emp_Leave_info where emp_id=@Emp_Id and year(start_dt)=@Leave_year and year(end_dt)=@Leave_year
	and Leave_name=@Leave_name
	group by Leave_name,Des,Address,Tel1,App_Type,App_ID,Leave_name			
		
	FETCH NEXT FROM Leave_Cursor INTO @Leave_Name,@Maximum
END
CLOSE Leave_Cursor
DEALLOCATE Leave_Cursor
 
Select * from  #Lv_Rpt





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




----exec Rpt_Loan_Sanction_Info '9-031'

CREATE PROCEDURE Rpt_Loan_Sanction_Info
	@Emp_Id varchar (5),
	@FromDate datetime,
	@ToDate datetime
AS
begin


if @Emp_Id=''
	SELECT LoanSanction_Info.Emp_id, Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS EMPLOYEE_NAME, 
	LoanSanction_Info.InstalmentNo, LoanSanction_Info.GracePeriod, LoanSanction_Info.SanctionAmount, LoanSanction_Info.InstallmentAmount, 
	LoanSanction_Info.SanctionDate, LoanSanction_Info.IstInstallmentDatePrin, LoanSanction_Info.IstInstallmentDateInt, 
	LoanSanction_Info.InstallmentAmountInt, LoanSanction_Info.InterestInstallmetNo, Loan_Type.Loan_Description
	FROM LoanSanction_Info INNER JOIN
	Loan_Type ON LoanSanction_Info.Loan_Type = Loan_Type.Loan_Type INNER JOIN
	Emp_Per_Info ON LoanSanction_Info.Emp_id = Emp_Per_Info.Emp_id
	WHERE  (LoanSanction_Info.InstalmentNo >= LoanSanction_Info.PaidInstallment)
	and LoanSanction_Info.SanctionDate between @FromDate and  @ToDate
else

	SELECT LoanSanction_Info.Emp_id, Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna AS EMPLOYEE_NAME, 
	LoanSanction_Info.InstalmentNo, LoanSanction_Info.GracePeriod, LoanSanction_Info.SanctionAmount, LoanSanction_Info.InstallmentAmount, 
	LoanSanction_Info.SanctionDate, LoanSanction_Info.IstInstallmentDatePrin, LoanSanction_Info.IstInstallmentDateInt, 
	LoanSanction_Info.InstallmentAmountInt, LoanSanction_Info.InterestInstallmetNo, Loan_Type.Loan_Description
	FROM LoanSanction_Info INNER JOIN
	Loan_Type ON LoanSanction_Info.Loan_Type = Loan_Type.Loan_Type INNER JOIN
	Emp_Per_Info ON LoanSanction_Info.Emp_id = Emp_Per_Info.Emp_id
	WHERE (LoanSanction_Info.InstalmentNo >= LoanSanction_Info.PaidInstallment)
	AND LoanSanction_Info.Emp_id=@Emp_Id
	AND  LoanSanction_Info.SanctionDate between @FromDate and  @ToDate
end








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





--Rpt_Monthly_Salary_Statement '9-027', 'June', '2002'

Create  Procedure Rpt_Monthly_Salary_Statement
@Emp_Id varchar(10)
,@Pay_Month varchar(12)
,@Pay_Year varchar(4)
AS 
 
SELECT a.emp_id,
Emp_name=(select Emp_fna+' '+Emp_mna+' '+Emp_lna from Emp_Per_info where emp_id=a.emp_id),
Dept=(select emp_dept from emp_Job_hist_current where emp_id=a.emp_id),
Category=(select emp_rank from emp_Job_hist_current where emp_id=a.emp_id),
Desig=(select emp_desig from emp_Job_hist_current where emp_id=a.emp_id),
sum(CASE b.head_name WHEN 'Basic Pay' THEN c.amount ELSE 0 END) AS  Basic,
sum(CASE b.head_name WHEN 'House Rent' THEN c.amount ELSE 0 END) AS H_Rent,
sum(CASE b.head_name WHEN 'Conveyance' THEN c.amount ELSE 0 END) AS  Conveyance,
sum(CASE b.head_name WHEN 'Medical Allowance' THEN c.amount  ELSE 0 END) AS Med_Allow,
sum(CASE b.head_name WHEN 'Charge Allowance' THEN c.amount  ELSE 0 END) AS Charge_Allow,
sum(CASE b.head_name WHEN 'Tech Allowance' THEN c.amount ELSE 0 END) AS  Tech_Allow,
sum(CASE b.head_name WHEN 'Others(+)' THEN c.amount  ELSE 0 END) AS [Others(+)],
sum(CASE b.head_name WHEN 'Bonus' THEN c.amount ELSE 0 END) AS Bonus,
sum(CASE b.head_name WHEN 'Arrear' THEN c.amount ELSE 0 END) AS Arrear,
sum(CASE b.head_name WHEN 'Absent Deduction' THEN c.amount ELSE 0 END) AS [Abs_Deduc],
sum(CASE b.head_name WHEN 'Advance Deduction' THEN c.amount ELSE 0 END) AS [Adv_Deduc],
sum(CASE b.head_name WHEN 'Others(-)' THEN c.amount ELSE 0 END) AS [Others(-)],
sum(CASE b.head_name WHEN 'IncomeTax' THEN c.amount ELSE 0 END) AS  IncomeTax,
pay_Month,a.pay_Year
from Payroll_Main a,pay_struc b, payroll_sub c,  emp_Job_hist_current q

where a.pay_Month=@pay_Month AND a.pay_Year=@pay_Year
and a.a_gen=c.a_gen and b.head_code=c.head_code
and a.emp_id=q.emp_id and a.emp_id=@Emp_Id
group by a.emp_id,a.pay_Month,a.pay_Year










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO










--select * from emp_movement
--Rpt_Movement'Ind','9-015','July','2003'
CREATE         procedure Rpt_Movement
@Mode varchar(5)
,@emp_id varchar(10)
,@Move_Month varchar(12)
,@Move_year varchar(12)
as
if @Mode='Ind'
begin
SELECT Emp_ID=@emp_id,
Emp_name=(select Emp_fna+' '+Emp_mna+' '+Emp_lna from Emp_Per_info where emp_id=@emp_id),
Desig=(select emp_desig from emp_Job_hist_current where emp_id=@emp_id),
Dept=(select emp_dept from emp_Job_hist_current where emp_id=@emp_id),
Move_Id,Mode,Place,Cont_Tel,Dt=convert(char(12),Move_Out_Dt,9),Move_Out_Dt,Exp_Rtn_Dt,Rtn_Dt= 
				case Move_Status 
				when 9 then Rtn_Dt else null 
				end
,Hour_On_Movement=(select dbo.GetHrInMovement(@Emp_Id,Move_Out_Dt))
,Cont_Tel,Move_des,[Month]=@Move_Month,[Year]=@Move_Year 
from emp_movement 
where emp_id=@emp_id and datename(month,Move_Out_Dt)=@Move_month and datepart(year,Move_Out_Dt)=@Move_year
end




if @Mode='Out'
begin
	select emp_id,Move_Id,Place,Move_Out_Dt,Exp_Rtn_Dt,Rtn_Dt=
				case Move_Status 
				when 9 then Rtn_dt else null 
				end,
	Cont_Tel,Move_des from emp_movement 
	where Move_Status=1 and convert(char(12),Move_Out_Dt,5)=convert(char(12),getdate(),5)
end
	











GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



----exec Rpt_OT_Preparation_monthwise '9-027','15-dec-2000','15-dec-2005'

CREATE    PROCEDURE Rpt_OT_Preparation_monthwise
	@Emp_Id varchar (5),
	@OT_St_Date datetime,			
	@OT_End_Date datetime

AS
begin

set nocount on

		SELECT     Overtime_Preparation.Emp_ID, Emp_Per_Info.Emp_fna+' '+
		Emp_Per_Info.Emp_mna +' '+ Emp_Per_Info.Emp_lna as Employee_Nm, 
		Overtime_Preparation.Date_of_OT, 
		Overtime_Preparation.Ot_St_Time, Overtime_Preparation.OT_End_Time, Overtime_Preparation.OT_Amount,
		Overtime_Preparation.P_Status,Overtime_Preparation.Remarks, Emp_Job_Hist_Current.CBasic, Overtime_Preparation.NOOfHr
		FROM Overtime_Preparation INNER JOIN
		Emp_Per_Info ON Overtime_Preparation.Emp_ID = Emp_Per_Info.Emp_id INNER JOIN
		Emp_Job_Hist_Current ON Emp_Per_Info.Emp_id = Emp_Job_Hist_Current.Emp_id
		WHERE     (Overtime_Preparation.Emp_ID = @Emp_Id) AND (Overtime_Preparation.Date_of_OT BETWEEN @OT_St_Date AND @OT_End_Date)
		ORDER BY Overtime_Preparation.TraceID
end



set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




--Rpt_Overtime 'August','2003'

CREATE  Proc Rpt_Overtime 

@Pay_Month varchar(12)
,@Pay_Year varchar(4)
AS


Declare @Comp_Name varchar(75)
,@Address varchar(150)

set @Comp_Name=(Select [Name] from Comp_Det_Info)
set @Address=(Select Address+ ', '+ City from Comp_Det_Info)

select 
a.Emp_Id
,Emp_Nm=(select a.emp_fna+' '+a.emp_mna+' '+a.emp_lna)
,b.Emp_Desig,b.Emp_Dept
,Gen_OT_Duration 
,Gen_OT_Pay            
,Out_OT_Duration 
,Out_OT_Pay            
,Hol_OT_Duration 
,Hol_OT_Pay           
,Pay_Month_Year= Pay_Month+ ', '+@Pay_Year
,Comp_Name=@Comp_Name 
,Address=@Address
from  Emp_Per_Info a,Emp_Job_Hist_Current b, Overtime c
where a.Emp_id=b.Emp_id  
and a.Emp_Id=c.Emp_Id
and c.Pay_Month=@Pay_Month
and c.Pay_Year=@Pay_Year









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






----exec Rpt_PF_Contribution_Monthwise '','August','2005'

CREATE   PROCEDURE Rpt_PF_Contribution_Monthwise
	@Emp_Id varchar (5),
	@Paymonth varchar(9),
	@PayYear varchar(4)
AS
begin


if @Emp_Id=''
		SELECT  Payroll_main.Emp_Id, Payroll_main.pay_Month, Payroll_main.pay_Year, Payroll_sub.Amount, 
		Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + '' + Emp_Per_Info.Emp_lna AS Emp_Name, Emp_Job_Hist_Current.Emp_desig
		FROM Payroll_main INNER JOIN
		Payroll_sub ON Payroll_main.A_Gen = Payroll_sub.A_Gen INNER JOIN
		Emp_Per_Info ON Payroll_main.Emp_Id = Emp_Per_Info.Emp_id INNER JOIN
		Emp_Job_Hist_Current ON Emp_Per_Info.Emp_id = Emp_Job_Hist_Current.Emp_id
		WHERE (Payroll_sub.Head_Code = '01') and (Payroll_main.pay_Month=@Paymonth)
		and (Payroll_main.pay_Year=@PayYear)
else

	SELECT Payroll_main.Emp_Id, Payroll_main.pay_Month, Payroll_main.pay_Year, Payroll_sub.Amount, 
	Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + '' + Emp_Per_Info.Emp_lna AS Emp_Name, Emp_Job_Hist_Current.Emp_desig
	FROM  Payroll_main INNER JOIN Payroll_sub ON Payroll_main.A_Gen = Payroll_sub.A_Gen INNER JOIN
	Emp_Per_Info ON Payroll_main.Emp_Id = Emp_Per_Info.Emp_id INNER JOIN
	Emp_Job_Hist_Current ON Emp_Per_Info.Emp_id = Emp_Job_Hist_Current.Emp_id
	WHERE     (Payroll_sub.Head_Code = '01') and ( Payroll_main.Emp_Id=@Emp_Id) and 
	(Payroll_main.pay_Month=@Paymonth)
	and (Payroll_main.pay_Year=@PayYear)




end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  Procedure Rpt_Perm_Attn
@Emp_Id varchar(10)
,@Yr varchar(4)


AS

set nocount on

declare @Mnth int
Set @Mnth=1
-----------------------------------------------
create table #Attn(

emp_id varchar(10)    
,Emp_fna varchar(30)
,emp_desig varchar(20) 
,emp_dept varchar(20)
,Joining_Dt datetime
,Job_type varchar(20)
,Payment_Basis varchar(20)
,Duty_type varchar(20)
,Month_NM varchar(12)
,Mn_Num int
,Yr varchar(4)    
,Work_Day int   
,Tot_Hol  int   
,Weekend  int   
,Present  int   
,Late     int   
,Leave    int   
)


Start_Insert:

	IF @Mnth<=12
		
	BEGIN
	
		insert into #Attn(emp_id,Emp_fna ,emp_desig,emp_dept
		,Joining_Dt,Job_type,Payment_Basis,Duty_type 
		,Month_NM ,Mn_Num,Yr ,Work_Day,Tot_Hol,Weekend,Present,Late,Leave) 
		exec rpt_Emp_Performance @Emp_Id,@Mnth,@Yr

		SET @Mnth=@Mnth+1
	END

	IF  @Mnth<=12 GOTO Start_Insert


	Select DISTINCT * from #Attn order by Mn_Num

set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




Create PROCEDURE [Rpt_Perm_Leave]

@emp_Id varchar(10)
,@yr varchar(4) 
AS
	
	select 
	Leave_name ,Start_dt=min(Start_dt),End_dt=max(End_dt),
	Duration=datediff(day,min(Start_dt),max(End_dt))+1,Des,Address,Tel1,App_Type,App_ID
	from Emp_Leave_info where emp_id=@emp_id and 
	datepart(year,Start_dt)=@yr and datepart(year,End_dt)=@yr
	group by Leave_name,Des,Address,Tel1,App_Type,App_ID			








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






/*select * from emp_att_info where emp_id='9-015'and  
datepart(day,Entry_Dt)=4 and datepart(month,Entry_Dt)=5

select * from Emp_Movement where Move_Status=1
--delete from Emp_Movement where Move_id=20
select * from emp_per_info

Rpt_Todays_Movement 'out'

Rpt_Todays_Movement 'All','2003-7-09'

*/

CREATE   Procedure  Rpt_Todays_Movement 
@Mode varchar(5)
,@Move_Dt datetime

AS
if @Mode='all'

SELECT DISTINCT a.emp_id,Emp_name=(a.Emp_fna + " " + a.Emp_mna +" "+ a.Emp_lna),
b.emp_desig,b.emp_dept,c.Move_Id,c.Mode,c.Place,Move_Out_Time=convert(char(20),c.Move_Out_Dt,0),
Exp_Rtn_Time=convert(char(20),c.Exp_Rtn_Dt,0),c.Cont_Tel,c.Move_des
,Rtn_Time= case  convert(varchar(1),Move_Status)
when '1'  then 'Not returned'
when '9'  then convert(char(20),c.Rtn_Dt,0)
end			
,Move_Dt=@Move_Dt
		from emp_Per_info a,Emp_Job_Hist_Current b, Emp_Movement c where 
		a.Emp_ID=B.Emp_ID and a.Emp_ID=c.Emp_ID and 
--		convert(char(12),c.Move_Out_Dt,5)=convert(char(12),Getdate(),5)
		
		convert(char(12),c.Move_Out_Dt,5)=convert(char(12),@Move_Dt,5)



if @Mode='Out'

SELECT DISTINCT a.emp_id,Emp_name=(a.Emp_fna + " " + a.Emp_mna +" "+ a.Emp_lna),
b.emp_desig,b.emp_dept,c.Move_Id,c.Mode,c.Place,Move_Out_Time=convert(char(20),c.Move_Out_Dt,0),
Exp_Rtn_Time=convert(char(20),c.Exp_Rtn_Dt,0)
,Rtn_Time= convert(char(20),c.Rtn_Dt,0)
,c.Cont_Tel,c.Move_des
from emp_Per_info a,Emp_Job_Hist_Current b, Emp_Movement c 
where 	a.Emp_ID=B.Emp_ID and a.Emp_ID=c.Emp_ID and c.Move_Status=1 and
	convert(char(12),c.Move_Out_Dt,5)=convert(char(12),Getdate(),5)
	





/*


declare Move_Cur cursor for 
select Move_Id,Entry_Dt from emp_movement

open 



*/





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/*
	exec Rpt_WhoGetsIncrement
*/

CREATE Procedure Rpt_WhoGetsIncrement

AS

set nocount on
----------------------------------------
create table #Increment(
emp_Id varchar(10)
,Emp_Name varchar(45)
,Emp_Dept varchar(35)
,Emp_Desig varchar(35)
,emp_join_date datetime
,Last_Dt_Incr datetime
,num_Incr int
,Incr_Amount money
)
----------------------------------------

declare @Emp_Id varchar(10)
,@Emp_Name varchar(45)
,@Emp_Dept varchar(35)
,@Emp_Desig varchar(35)
,@emp_join_date datetime
,@Diff int
,@Last_Incr_Dt datetime
,@num_Incr int
,@Incr_Amount money

--------------------------------------------------------------------------------
declare Id_cursor cursor for
	select distinct emp_id from increment_history
--------------------------------------------------------------------------------

open id_Cursor
fetch next from Id_cursor into @Emp_Id

	while @@Fetch_status = 0

	begin
		
		set @Diff=(select datediff(year,dbo.GetLatestIncrDate (@Emp_Id),getdate()))
	
			if @Diff>=1	
			begin
				set @num_Incr=(select dbo.GetLatestIncrNum(@emp_Id) )
				set @Incr_Amount=(select dbo.GetLatestIncrAmount(@emp_Id))		 
				set @Last_Incr_Dt=(select dbo.GetLatestIncrDate(@emp_Id))
		
				set @Emp_name=(select Emp_fna+' '+Emp_mna+' '+Emp_lna from emp_per_Info where Emp_Id=@Emp_Id)
				set @emp_desig=(select emp_desig from Emp_Job_Hist_Current where Emp_Id=@Emp_Id)
				set @emp_dept=(select emp_dept from Emp_Job_Hist_Current where Emp_Id=@Emp_Id)
				set @emp_join_date=(select emp_join_date from Emp_Job_Hist_Current where Emp_Id=@Emp_Id)
		
				insert into #increment(emp_Id,Emp_Name,Emp_Dept
					,Emp_Desig,emp_join_date,Last_Dt_Incr,num_Incr,Incr_Amount)
				
				values(@Emp_Id,@Emp_name,@emp_dept,
					@emp_desig,@emp_join_date,@Last_Incr_Dt,@num_Incr,@Incr_Amount)
		
			end	

		fetch next from Id_cursor into @Emp_Id	
	end


close Id_cursor
deallocate Id_cursor

select * from #Increment


set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/*

exec Rpt_YearlySalary_Statement '9-001','2002'

*/
CREATE Procedure Rpt_YearlySalary_Statement
@emp_id varchar(10)
,@pay_Year varchar(4) 
AS

set nocount on

Create table #Salary(
emp_id varchar(10),Emp_name varchar(45),Dept varchar(35)
,Category varchar(35),Desig varchar(35),Basic money
,H_Rent money,Conveyance money,Med_Allow money,Charge_Allow money
,Tech_Allow money,Others_Add money,Bonus money,Arrear money
,Abs_Deduc money,Adv_Deduc money,Others_Deduct money,IncomeTax money
,pay_Month varchar(12) ,pay_Year varchar(4) 

)

declare @Month_Num  int
,@Month_Name varchar(12)


Set @Month_Num=1

Start:

set @Month_Name=(select dbo.Month_Name(@Month_Num))

insert into #Salary
exec Rpt_Monthly_Salary_Statement @Emp_Id,@Month_Name,@pay_Year

Set @Month_Num=@Month_Num+1

while @Month_Num<=12 goto Start

select * from #Salary 

set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE   procedure Rpt_exam_routine
(
    @mode               varchar(1),
	@ClassId            varchar(10), 
    @ExamYear	 		int,
	@ExamID				varchar(10),
    @ExamTypeID			varchar(10)
)
as


declare @max_serial as int

if @mode='a' 
  begin
    select a.ExamDate,
           a.Sub_id,(select Sub_title from subject_info_sub b
                    where b.Sub_code=a.Sub_id 
                    and b.Class_code=a.ClassId)as sub_title,
           a.ClassId,(select b.ClassName from classinfo b
                 where b.ClassID=a.ClassID) as class_title,a.ExamStartTime,
           a.ExamEndTime 
       from ExamSchedule a
      where a.ExamYear=  @ExamYear  and
            a.ExamTypeID = @ExamTypeID and
            a.ExamId=@ExamId 
  end 
/*
if @mode='c' 
  begin
    select a.ExamDate,
           a.Sub_id,
           (select Sub_title from subject_info_sub b
                    where b.Sub_code=a.Sub_id 
                    and b.Class_code=a.ClassId)as sub_title,
           a.ClassId,(select b.ClassName from classinfo b
                 where b.ClassID=a.ClassID) as class_title, a.ExamStartTime,
           a.ExamEndTime 
       from ExamSchedule  a
      where a.ExamYear=  @ExamYear  and
            a.ExamTypeID = @ExamTypeID and
            a.ExamId=@ExamId 
            and a.classid=@classid 
  end 
*/





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





CREATE   procedure Rpt_marks
(
    @mode               varchar(1),
	@main_code          int
)
as


declare @max_serial as int

if @mode='a' 
  begin
    select a.StdID,(select studentname from studentinfo where
           studentid=a.StdID) as stdname,
           a.Roll,
           a.obtainedMarks ,
           a.EntryBy,
           a.EntryDate 
       from result_sub a 
      where a.M_Slr_no=@main_code 
          
              
  end 








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE    Procedure SP

AS
select Name from dbo.sysobjects  where 
type='P' and xtype='P' and name!='SP'
and name!='Fx'and name!='Tbl'and name not like 'Dt_%'






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_All_Id_name_POP    Script Date: 3/7/02 10:43:49 AM ******/
/****** Object:  Stored Procedure dbo.SP_All_Id_name_POP    Script Date: 10/28/01 4:34:44 PM ******/
/****** Object:  Stored Procedure dbo.SP_All_Id_name_POP    Script Date: 10/22/01 6:24:42 PM ******/
/****** Object:  Stored Procedure dbo.SP_All_Id_name_POP    Script Date: 9/17/00 4:34:15 PM ******/
/****** Object:  Stored Procedure dbo.SP_All_Id_name_POP    Script Date: 9/4/01 6:30:49 PM ******/
CREATE PROCEDURE [SP_All_Id_name_POP]
AS
	select  a.emp_id,nm=(a.emp_fna+" " + a.emp_mna + " "+ a.emp_lna),b.Emp_Desig
	from Emp_Per_Info a,Emp_Job_Hist_Current b where a.emp_id=b.emp_id
		




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Attendance_count    Script Date: 3/7/02 10:43:49 AM ******/
/****** Object:  Stored Procedure dbo.SP_Attendance_count    Script Date: 10/28/01 4:34:44 PM ******/
/****** Object:  Stored Procedure dbo.SP_Attendance_count    Script Date: 10/22/01 6:24:20 PM ******/
/****** Object:  Stored Procedure dbo.SP_Attendance_count    Script Date: 9/17/00 4:33:55 PM ******/
/****** Object:  Stored Procedure dbo.SP_Attendance_count    Script Date: 9/4/01 6:30:59 PM ******/
/****** Object:  Stored Procedure dbo.SP_Attendance_count    Script Date: 6/20/01 12:00:42 PM ******/
/****** Object:  Stored Procedure dbo.SP_Attendance_count    Script Date: 5/14/01 9:49:54 AM ******/
/****** Object:  Stored Procedure dbo.SP_Attendance_count    Script Date: 1/1/99 3:01:45 AM ******/
/****** Object:  Stored Procedure dbo.SP_Attendance_count    Script Date: 5/9/01 8:13:06 AM ******/
CREATE PROCEDURE  [SP_Attendance_count]
@Emp_Id varchar(10)
 AS
select count(Emp_login) from Emp_Att_info
where emp_id=@emp_id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Check_ID    Script Date: 3/7/02 10:43:49 AM ******/
CREATE PROCEDURE [SP_Check_ID]
@Id as varchar(10)
 AS
if @Id='All'
	begin
		select nm=(emp_fna+" " + emp_mna + " "+ emp_lna) from emp_per_info
	end
else
	begin
		select nm=(a.emp_fna+" " + a.emp_mna + " "+ a.emp_lna),b.Emp_desig,b.Emp_dept
		from emp_per_info a,Emp_Job_Hist_Current b where 
			a.emp_id=b.emp_id and	a.emp_id=@Id
	end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Data_mapping_I_U_D    Script Date: 3/7/02 10:43:38 AM ******/
/****** Object:  Stored Procedure dbo.SP_Data_mapping_I_U_D    Script Date: 10/28/01 4:34:32 PM ******/
/****** Object:  Stored Procedure dbo.SP_Data_mapping_I_U_D    Script Date: 10/22/01 6:24:32 PM ******/
/****** Object:  Stored Procedure dbo.SP_Data_mapping_I_U_D    Script Date: 9/17/00 4:34:04 PM ******/
/****** Object:  Stored Procedure dbo.SP_Data_mapping_I_U_D    Script Date: 9/4/01 6:30:50 PM ******/
CREATE procedure [SP_Data_mapping_I_U_D] 
@opr varchar (10),
@Gen_data varchar(25),
@Head_code varchar(2),
@U_id varchar (10)
as
if @opr='I'
begin
	insert into Data_mapping
		(Gen_Data,
		Head_code,
		U_id)
		Values(	
		@Gen_Data,
		@Head_code,
		@U_id)
end
if @opr='U'
begin
	Update Data_mapping set
		Gen_Data=@Gen_Data,
		Head_code=@Head_code,
		U_id=@U_id
		where Head_code=@Head_code
end
if @opr='D'
begin
	Delete from Data_mapping 
		where Head_code=@Head_code 
	
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Designation    Script Date: 3/7/02 10:43:49 AM ******/
/****** Object:  Stored Procedure dbo.SP_Designation    Script Date: 10/28/01 4:34:44 PM ******/
/****** Object:  Stored Procedure dbo.SP_Designation    Script Date: 10/22/01 6:24:43 PM ******/
/****** Object:  Stored Procedure dbo.SP_Designation    Script Date: 9/17/00 4:34:15 PM ******/
/****** Object:  Stored Procedure dbo.SP_Designation    Script Date: 9/4/01 6:30:41 PM ******/
/****** Object:  Stored Procedure dbo.SP_Designation    Script Date: 6/20/01 12:00:42 PM ******/
/****** Object:  Stored Procedure dbo.SP_Designation    Script Date: 5/14/01 9:49:54 AM ******/
/****** Object:  Stored Procedure dbo.SP_Designation    Script Date: 1/1/99 3:01:46 AM ******/
/****** Object:  Stored Procedure dbo.SP_Designation    Script Date: 5/9/01 8:13:10 AM ******/
/****** Object:  Stored Procedure dbo.SP_Designation    Script Date: 5/8/01 4:09:48 AM ******/
CREATE PROCEDURE [SP_Designation]
@Emp_Id as varchar(10)
 AS
	select Emp_desig from Emp_Job_hist_current where emp_id=@Emp_id
	




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





CREATE   PROCEDURE SP_Emp_Desig_Sec_Join
@Emp_id varchar (10)
as
SELECT 
a.Emp_id,
NM=(b.Emp_fna+' '+b.Emp_mna+' '+b.Emp_lna), 
a.Emp_join_date,
a.Emp_rank,
a.Emp_desig,
a.Emp_branch,
a.Emp_dept,
a.Emp_section,
a.Emp_job_type,
a.Responsibility,
a.Super_ID,
a.Permdate,
a.CBasic,
P_type=a.Pay_type,
D_type=a.Duty_type


from  Emp_Job_Hist_Current a,Emp_Per_Info b
where a.Emp_id=b.Emp_id  and a.Emp_id=@Emp_id



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Emp_Id_with_OT_pay    Script Date: 3/7/02 10:43:26 AM ******/
/****** Object:  Stored Procedure dbo.SP_Emp_Id_with_OT_pay    Script Date: 10/28/01 4:34:45 PM ******/
/****** Object:  Stored Procedure dbo.SP_Emp_Id_with_OT_pay    Script Date: 1/29/01 10:24:55 AM ******/
create procedure SP_Emp_Id_with_OT_pay
@Emp_id varchar(10),
@Pay_month varchar(12),
@Pay_year varchar(4)
as
Select OT_Pay from Emp_Payroll
where Emp_id=@emp_id And Pay_month=@Pay_month and Pay_year=@Pay_year




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_FIX_The_Rest_Fixed_Head    Script Date: 3/7/02 10:43:49 AM ******/
/****** Object:  Stored Procedure dbo.SP_FIX_The_Rest_Fixed_Head    Script Date: 10/28/01 4:34:45 PM ******/
/****** Object:  Stored Procedure dbo.SP_FIX_The_Rest_Fixed_Head    Script Date: 10/22/01 6:24:43 PM ******/
/****** Object:  Stored Procedure dbo.SP_FIX_The_Rest_Fixed_Head    Script Date: 9/17/00 4:34:15 PM ******/
/****** Object:  Stored Procedure dbo.SP_FIX_The_Rest_Fixed_Head    Script Date: 9/4/01 6:30:59 PM ******/
create procedure SP_FIX_The_Rest_Fixed_Head
@Emp_ID varchar (10)
as
select Head_Code,Head_Name,mode from Pay_Struc where Mode='F' and not exists(select * from Fixed_Pay
where Fixed_Pay.Head_code= Pay_Struc.Head_code and  Fixed_Pay.Emp_id=@Emp_id)




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Find_Emp_per_info    Script Date: 3/7/02 10:43:26 AM ******/
/****** Object:  Stored Procedure dbo.SP_Find_Emp_per_info    Script Date: 10/28/01 4:34:32 PM ******/
/****** Object:  Stored Procedure dbo.SP_Find_Emp_per_info    Script Date: 10/22/01 6:24:20 PM ******/
/****** Object:  Stored Procedure dbo.SP_Find_Emp_per_info    Script Date: 9/17/00 4:34:04 PM ******/
/****** Object:  Stored Procedure dbo.SP_Find_Emp_per_info    Script Date: 9/4/01 6:30:50 PM ******/
/****** Object:  Stored Procedure dbo.SP_Find_Emp_per_info    Script Date: 6/20/01 12:00:35 PM ******/
/****** Object:  Stored Procedure dbo.SP_Find_Emp_per_info    Script Date: 5/14/01 9:49:53 AM ******/
/****** Object:  Stored Procedure dbo.SP_Find_Emp_per_info    Script Date: 1/1/99 3:01:41 AM ******/
/****** Object:  Stored Procedure dbo.SP_Find_Emp_per_info    Script Date: 5/9/01 8:13:07 AM ******/
CREATE PROCEDURE [SP_Find_Emp_per_info]
@Mode as varchar(2),					---Mode   ID, FN,MN,LN
@SC as varchar(20)					---SC       Search Criterion ie. Emp_ID , Emp_fna, Emp_fna,  Emp_fna
					
 AS
if @Mode='ID'
	begin
		select * from Emp_per_info  where emp_id=@SC
	end
if @Mode='FN'
	begin
		select * from Emp_per_info  where emp_fna=@SC
	end
if @Mode='MN'
	begin
		select * from Emp_per_info  where emp_mna=@SC
	end
if @Mode='LM'
	begin
		select * from Emp_per_info  where emp_lna=@SC
	end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Fixed_Pay_I_U_D    Script Date: 3/7/02 10:43:50 AM ******/
/****** Object:  Stored Procedure dbo.SP_Fixed_Pay_I_U_D    Script Date: 10/28/01 4:34:45 PM ******/
/****** Object:  Stored Procedure dbo.SP_Fixed_Pay_I_U_D    Script Date: 10/22/01 6:24:43 PM ******/
/****** Object:  Stored Procedure dbo.SP_Fixed_Pay_I_U_D    Script Date: 9/17/00 4:34:16 PM ******/
/****** Object:  Stored Procedure dbo.SP_Fixed_Pay_I_U_D    Script Date: 9/4/01 6:31:00 PM ******/
CREATE procedure [SP_Fixed_Pay_I_U_D]
@opr varchar(1),
@Emp_id varchar (10),
@Head_code varchar(2),
@Amount money,
@U_id varchar(10)
as
if @opr='I'
begin
	insert into Fixed_pay(
	Emp_id,
	Head_code,
	Amount,
	U_id)
	values(
	@Emp_id,
	@Head_code,
	@Amount,
	@U_id)
end
if @opr='U'
begin
	update Fixed_pay set
	Head_code=@Head_code,
	Amount=@Amount,
	U_id=@U_id 	
	where Emp_ID=@Emp_ID
end
if @opr='D'
begin
	delete from fixed_pay
	where Emp_Id=@Emp_Id
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Holiday_List_All    Script Date: 3/7/02 10:43:38 AM ******/
/****** Object:  Stored Procedure dbo.SP_Holiday_List_All    Script Date: 10/28/01 4:34:32 PM ******/
/****** Object:  Stored Procedure dbo.SP_Holiday_List_All    Script Date: 1/29/01 10:24:54 AM ******/
/****** Object:  Stored Procedure dbo.SP_Holiday_List_All    Script Date: 5/14/01 9:49:53 AM ******/
/****** Object:  Stored Procedure dbo.SP_Holiday_List_All    Script Date: 1/1/99 3:01:41 AM ******/
/****** Object:  Stored Procedure dbo.SP_Holiday_List_All    Script Date: 5/9/01 8:13:07 AM ******/
/****** Object:  Stored Procedure dbo.SP_Holiday_List_All    Script Date: 5/8/01 4:09:46 AM ******/
CREATE PROCEDURE [SP_Holiday_List_All]
@Id as varchar(10)
 AS
if @Id='All'
	begin
		select Hol_name,Hol_Desc,str_date,End_date,Duration= datediff(dd,str_date,end_date)+1, category  from Hol_List
		
	end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE   PROCEDURE SP_Job_Detail
as
Select distinct
	Emp_ID,
	Emp_join_date,
	Emp_rank,
	Emp_desig,
	Emp_branch,
	Emp_dept,
	Emp_section,
	Emp_job_type,
	Responsibility,
	P_Type=Pay_type,
	D_Type=Duty_type,
	super_id
 from Emp_Job_Hist_Current


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Job_Detail_Mod    Script Date: 3/7/02 10:43:50 AM ******/
/****** Object:  Stored Procedure dbo.SP_Job_Detail_Mod    Script Date: 10/28/01 4:34:45 PM ******/
/****** Object:  Stored Procedure dbo.SP_Job_Detail_Mod    Script Date: 10/22/01 6:24:44 PM ******/
/****** Object:  Stored Procedure dbo.SP_Job_Detail_Mod    Script Date: 9/17/00 4:34:16 PM ******/
/****** Object:  Stored Procedure dbo.SP_Job_Detail_Mod    Script Date: 9/4/01 6:30:41 PM ******/
/****** Object:  Stored Procedure dbo.SP_Job_Detail_Mod    Script Date: 6/20/01 12:00:43 PM ******/
CREATE PROCEDURE SP_Job_Detail_Mod
@Find Varchar (10)
as
select Emp_id,Emp_join_date,Emp_rank,Emp_desig,Emp_branch,Emp_dept,Emp_section,Emp_job_type from  Emp_Job_Hist_Current 
where  Emp_id=@Find




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Job_End_All    Script Date: 3/7/02 10:43:50 AM ******/
/****** Object:  Stored Procedure dbo.SP_Job_End_All    Script Date: 10/28/01 4:34:45 PM ******/
/****** Object:  Stored Procedure dbo.SP_Job_End_All    Script Date: 10/22/01 6:24:44 PM ******/
/****** Object:  Stored Procedure dbo.SP_Job_End_All    Script Date: 9/17/00 4:34:16 PM ******/
/****** Object:  Stored Procedure dbo.SP_Job_End_All    Script Date: 9/4/01 6:31:00 PM ******/
/****** Object:  Stored Procedure dbo.SP_Job_End_All    Script Date: 6/20/01 12:00:43 PM ******/
/****** Object:  Stored Procedure dbo.SP_Job_End_All    Script Date: 5/14/01 9:49:54 AM ******/
/****** Object:  Stored Procedure dbo.SP_Job_End_All    Script Date: 1/1/99 3:01:46 AM ******/
/****** Object:  Stored Procedure dbo.SP_Job_End_All    Script Date: 5/9/01 8:13:10 AM ******/
/****** Object:  Stored Procedure dbo.SP_Job_End_All    Script Date: 5/8/01 4:09:49 AM ******/
CREATE PROCEDURE [SP_Job_End_All]
@Id as varchar(10)
 AS
if @Id='All'
	begin
		select emp_id,Job_end_dt,type,des from Emp_Job_end
	end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Leave_List_All    Script Date: 3/7/02 10:43:39 AM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_List_All    Script Date: 10/28/01 4:34:32 PM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_List_All    Script Date: 10/22/01 6:24:32 PM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_List_All    Script Date: 9/17/00 4:34:05 PM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_List_All    Script Date: 9/4/01 6:30:42 PM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_List_All    Script Date: 6/20/01 12:00:36 PM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_List_All    Script Date: 5/14/01 9:49:53 AM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_List_All    Script Date: 1/1/99 3:01:41 AM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_List_All    Script Date: 5/9/01 8:13:08 AM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_List_All    Script Date: 5/8/01 4:09:46 AM ******/
CREATE PROCEDURE [SP_Leave_List_All]
@Id as varchar(10)
 AS
if @Id='All'
	begin
		select Leave_name,Duration,des from Leave_List
	end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  PROCEDURE [SP_Leave_app_Pop]

@emp_Id varchar(10)
,@yr varchar(4) 
AS
	
	select 
	Leave_name ,Start_dt=min(Start_dt),End_dt=max(End_dt),
	Duration=datediff(day,min(Start_dt),max(End_dt))+1,Des,Address,Tel1,App_Type,App_ID
	from Emp_Leave_info where emp_id=@emp_id and 
	datepart(year,Start_dt)=@yr and datepart(year,End_dt)=@yr
	group by Leave_name,Des,Address,Tel1,App_Type,App_ID			






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Leave_app_with_id_name    Script Date: 3/7/02 10:43:27 AM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_app_with_id_name    Script Date: 10/28/01 4:34:46 PM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_app_with_id_name    Script Date: 10/22/01 6:24:45 PM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_app_with_id_name    Script Date: 9/17/00 4:34:17 PM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_app_with_id_name    Script Date: 9/4/01 6:30:50 PM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_app_with_id_name    Script Date: 6/20/01 12:00:36 PM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_app_with_id_name    Script Date: 5/14/01 9:49:54 AM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_app_with_id_name    Script Date: 1/1/99 3:01:46 AM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_app_with_id_name    Script Date: 5/9/01 8:13:10 AM ******/
/****** Object:  Stored Procedure dbo.SP_Leave_app_with_id_name    Script Date: 5/8/01 4:09:49 AM ******/
CREATE PROCEDURE [SP_Leave_app_with_id_name]
@emp_Id as varchar(10)
 AS
	
		select a.emp_id,nm=(a.emp_fna+" " + a.emp_mna + " "+ a.emp_lna),
		b.Leave_name,b.Start_dt,b.End_dt,b.Des,b.App_ID
		from emp_per_info a,Emp_Leave_info b where exists(select * from emp_leave_info
		where a.emp_id=b.emp_id and a.emp_id=@emp_id and b.emp_id=@emp_id) 
		group by a.emp_id,a.emp_fna,a.emp_mna,a.emp_lna,b.Leave_name,b.Start_dt,b.End_dt,b.Des,App_ID
/*	end
else
	begin
		select a.emp_id,nm=(a.emp_fna+" " + a.emp_mna + " "+ a.emp_lna),
		b.Leave_name,b.Start_dt,b.End_dt,b.Des,Uniq_id
		from emp_per_info a,Emp_Leave_Info b where exists(select * from emp_Leave_info
		where a.emp_id=b.emp_id and a.emp_id=@Id) 
		group by a.emp_id,a.emp_fna,a.emp_mna,a.emp_lna,b.Leave_name,b.Start_dt,b.End_dt,b.Des,Uniq_id
	end
*/




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Move_In    Script Date: 3/7/02 10:43:39 AM ******/
/****** Object:  Stored Procedure dbo.SP_Move_In    Script Date: 10/28/01 4:34:32 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_In    Script Date: 10/22/01 6:24:32 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_In    Script Date: 9/17/00 4:34:05 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_In    Script Date: 9/4/01 6:30:51 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_In    Script Date: 6/20/01 12:00:36 PM ******/
CREATE PROCEDURE [SP_Move_In]
@Emp_id varchar (10)
as
declare @@Latest_Move Datetime
select @@Latest_Move=max(move_Out_dt) From Move_Out_In
update Move_Out_in set Rtn_dt=getdate()
where Emp_Id=@Emp_Id and  Move_Out_Dt=@@Latest_Move




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Move_In_Mod    Script Date: 3/7/02 10:43:39 AM ******/
/****** Object:  Stored Procedure dbo.SP_Move_In_Mod    Script Date: 10/28/01 4:34:33 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_In_Mod    Script Date: 10/22/01 6:24:33 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_In_Mod    Script Date: 9/17/00 4:34:05 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_In_Mod    Script Date: 9/4/01 6:30:51 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_In_Mod    Script Date: 6/20/01 12:00:36 PM ******/
CREATE procedure SP_Move_In_Mod
@emp_id varchar(10)
as
update Move_Out_In set Rtn_dt=getdate()
where Emp_Id=@Emp_id and Rtn_dt is null




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Move_Out_In_I_U_D    Script Date: 3/7/02 10:43:39 AM ******/
/****** Object:  Stored Procedure dbo.SP_Move_Out_In_I_U_D    Script Date: 10/28/01 4:34:33 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_Out_In_I_U_D    Script Date: 10/22/01 6:24:33 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_Out_In_I_U_D    Script Date: 9/17/00 4:34:05 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_Out_In_I_U_D    Script Date: 9/4/01 6:30:51 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_Out_In_I_U_D    Script Date: 6/20/01 12:00:36 PM ******/
CREATE procedure [SP_Move_Out_In_I_U_D]
@opr varchar (1),
@Emp_id varchar (10),
@Mode varchar (20),
@Place varchar (30),
@Move_Out_Dt datetime,
@Exp_Rtn_Dt datetime,
@Cont_Tel varchar (20),
@Move_des varchar (500)
as
if @opr='I'
begin
insert into Move_Out_In
(
Emp_id ,
Mode,
Place ,
Exp_Rtn_Dt ,
Cont_Tel ,
Move_des
)
Values
(
@Emp_id ,
@Mode,
@Place ,
@Exp_Rtn_Dt ,
@Cont_Tel ,
@Move_des
)
end
if @opr='U'
begin
	update Move_Out_In set
		Mode=@Mode,
		Place=@Place ,
		Move_Out_Dt=@Move_Out_Dt,
		Exp_Rtn_Dt=@Exp_Rtn_Dt ,
		Cont_Tel=@Cont_Tel ,
		Move_des=@Move_des
	Where Emp_id=@Emp_id AND Move_Out_Dt=@Move_Out_Dt
end
if @opr='D'
begin
Delete from Move_Out_In
Where Emp_id=@Emp_id AND Move_Out_Dt=@Move_Out_Dt
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Move_Schedule_Pop    Script Date: 3/7/02 10:43:40 AM ******/
/****** Object:  Stored Procedure dbo.SP_Move_Schedule_Pop    Script Date: 10/28/01 4:34:33 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_Schedule_Pop    Script Date: 10/22/01 6:24:33 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_Schedule_Pop    Script Date: 9/17/00 4:34:06 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_Schedule_Pop    Script Date: 9/4/01 6:30:52 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_Schedule_Pop    Script Date: 6/20/01 12:00:32 PM ******/
CREATE procedure [SP_Move_Schedule_Pop]
@Emp_id  varchar (10)
as
Select 
	Emp_id ,
	Mode,
	Place,
	Move_Out_Dt,
	Exp_Rtn_Dt,
	Cont_Tel,
	Move_des 
	from Move_Schedule
	 where Emp_Id=@Emp_Id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Move_schedule_I_U_D    Script Date: 3/7/02 10:43:40 AM ******/
/****** Object:  Stored Procedure dbo.SP_Move_schedule_I_U_D    Script Date: 10/28/01 4:34:33 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_schedule_I_U_D    Script Date: 10/22/01 6:24:33 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_schedule_I_U_D    Script Date: 9/17/00 4:34:06 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_schedule_I_U_D    Script Date: 9/4/01 6:30:51 PM ******/
/****** Object:  Stored Procedure dbo.SP_Move_schedule_I_U_D    Script Date: 6/20/01 12:00:32 PM ******/
CREATE procedure [SP_Move_schedule_I_U_D]
@opr varchar (1),
@Emp_id varchar (10),
@Mode varchar (20) ,
@Place  varchar (20),
@Move_Out_Dt  datetime ,
@Exp_Rtn_Dt datetime,
@Cont_Tel varchar(20),
@Move_des  varchar(500)
As
if @opr='I'
begin
insert into Move_Schedule
(
Emp_id ,
Mode,
Place,
Move_Out_Dt,
Exp_Rtn_Dt ,
Cont_Tel,
Move_des
)
Values
(
@Emp_id ,
@Mode,
@Place,
@Move_Out_Dt,
@Exp_Rtn_Dt ,
@Cont_Tel,
@Move_des
)
end
if @opr='U'
begin
update Move_Schedule set
Mode=@Mode,
Place=@Place,
Move_Out_Dt=@Move_Out_Dt,
Exp_Rtn_Dt=@Exp_Rtn_Dt,
Cont_Tel=@Cont_Tel,
Move_des=@Move_des
Where Emp_id=@Emp_id and Move_Out_Dt=@Move_Out_Dt
end
if @opr='D'
begin
Delete from Move_Schedule 
Where Emp_id=@Emp_id and Move_Out_Dt=@Move_Out_Dt
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE procedure [SP_Movement_POP]
@Mode varchar(10),
@emp_id varchar(10)
as
if @Mode='Out'
begin
	select * from Emp_Movement where emp_id =@emp_id and move_status=1 
end
if @Mode='Schedule'
begin
	select * from Emp_Movement where emp_id =@emp_id and move_status=0 and convert(char(12),Move_Out_Dt,1)=convert(char(12),getdate (),1)
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_OT_Pay_Pop    Script Date: 3/7/02 10:43:27 AM ******/
/****** Object:  Stored Procedure dbo.SP_OT_Pay_Pop    Script Date: 10/28/01 4:34:46 PM ******/
/****** Object:  Stored Procedure dbo.SP_OT_Pay_Pop    Script Date: 10/22/01 6:24:45 PM ******/
/****** Object:  Stored Procedure dbo.SP_OT_Pay_Pop    Script Date: 9/17/00 4:34:17 PM ******/
/****** Object:  Stored Procedure dbo.SP_OT_Pay_Pop    Script Date: 9/4/01 6:31:00 PM ******/
create procedure SP_OT_Pay_Pop
as
select Emp_ID,Pay_month,Pay_year,OT_pay from Emp_payroll




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Office_Time_Worker_Official    Script Date: 3/7/02 10:43:40 AM ******/
CREATE procedure SP_Office_Time_Worker_Official
@Emp_Id varchar(10)
as
declare @ID_Type varchar(1)
set @ID_Type=left(LTRIM(RTRIM(@Emp_Id)),1)
if @ID_Type=0 		---or @ID_Type=1
begin
	select Start_time,End_time,Relaxed=10,Abs_time  from office_time where Effect_date=(
	 select Mdt=(select max(Effect_date)from office_time where  CONVERT(char(12), Effect_date, 1) <= CONVERT(char(12),GETDATE(), 1)))
END
if @ID_Type=1 or @ID_Type=2 or @ID_Type=3 or @ID_Type=4 or @ID_Type=5 or @ID_Type=6 or @ID_Type=7 or @ID_Type=8 or @ID_Type=9
begin
	select Start_time,End_time,Relaxed,Abs_time  from office_time where Effect_date=(
	select Mdt=(select max(Effect_date)from office_time where  CONVERT(char(12), Effect_date, 1) <= CONVERT(char(12),GETDATE(), 1)))
END
/*++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/
/*if @ID_Type=3 or @ID_Type=5 or @ID_Type=7 or @ID_Type=9        --------- Not applicable for 'Liberty' ,done for 'Ayman'
begin
	declare @ID varchar(10)
	set @ID=@Emp_ID
	---------------------------------------------------------
	declare @S_Dt datetime
	declare @E_Dt datetime
	-------------------------
	select @S_Dt=start_dt,@E_Dt=End_Dt from group_Shift where Shift_Code=(
		select Shift_Code from Group_Shift where Gr_Code=(
			select Gr_Code from Worker_Group where emp_ID=@ID))
	--------------------------------------------------------
	
	select  Relaxed=0,Abs_time =0, Start_time=Shift_Start, End_time =Shift_End from Shift where Shift_Code=(
		select Shift_Code from Group_Shift where Gr_Code=(
			select Gr_Code from Worker_Group where emp_ID=@ID))and 
	
			CONVERT(char(12),GETDATE(), 1)>=CONVERT(char(12),@S_Dt, 1) AND
			CONVERT(char(12),GETDATE(), 1)<=CONVERT(char(12), @E_Dt, 1)
end 
*/




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE  PROCEDURE [SP_Office_time_All]
@Mode as varchar(10)
 AS
if @Mode='All' 
begin
	
select Start_time,End_time,Relaxed,Abs_time,Sp_Start_time,Sp_Start_Day,Sp_End_time,Sp_End_Day ,Effect_date,Ref_Key from office_time
end
if @Mode='Now' 
begin
select Start_time,end_time ,Relaxed,Abs_time from office_time where datepart(day,Effect_date)=(
 select Mdt=(
	select datepart(day,max (Effect_date))from office_time 
	       where  datepart(day,Effect_date)<=datepart(day,getdate())and
	              datepart(month,Effect_date)<=datepart(month,getdate())and
	              datepart(year,Effect_date)<=datepart(year,getdate())))
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Pay_Struc_I_U_D    Script Date: 3/7/02 10:43:40 AM ******/
/****** Object:  Stored Procedure dbo.SP_Pay_Struc_I_U_D    Script Date: 10/28/01 4:34:34 PM ******/
/****** Object:  Stored Procedure dbo.SP_Pay_Struc_I_U_D    Script Date: 10/22/01 6:24:34 PM ******/
/****** Object:  Stored Procedure dbo.SP_Pay_Struc_I_U_D    Script Date: 9/17/00 4:34:06 PM ******/
/****** Object:  Stored Procedure dbo.SP_Pay_Struc_I_U_D    Script Date: 9/4/01 6:30:52 PM ******/
----Procedure-----------
CREATE procedure SP_Pay_Struc_I_U_D
@opr varchar (1),
@Old_value varchar(2),
@Head_code varchar(2),
@Head_name varchar(30),
@Operation varchar(1),
@Mode varchar(1),
@U_id varchar(10)
as
if @opr='I'
begin
	insert into Pay_Struc(
		Head_code,
		Head_name,
		Operation,		
		Mode,
		U_id
)
		Values(	
		@Head_code,
		@Head_name,
		@Operation,
		@Mode,
		@U_id 
)
end
if @opr='U'
begin
	Update Pay_Struc set
		Head_code=@Head_code,
		Head_name=@Head_name,
		Operation=@Operation,
		Mode=@Mode,
		U_id=@U_id
		where Head_code=@Old_value
end
if @opr='D'
begin
	delete from Pay_Struc 
		where Head_code=@Head_code 
	
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE Procedure [SP_Payroll_I_U_D]
	@opr varchar(10),
	@Emp_Id varchar(10),
	@Head_Code varchar (2),
	@pay_Month varchar(12),
	@pay_Year varchar(4),
	@amount money,
	@U_id varchar(10)
	
as
set nocount on
declare
  
	 @A_Gen varchar(8)
	,@pay_Type int
	,@Message varchar(150)


if @opr='I'
	set @Message='Salary preparation done successfully !'
if @opr='U'
	set @Message='Salary modification done successfully !'
if @opr='D'
	set @Message='Salary deleted successfully !'

if @opr='I'


begin
	
	set @pay_Type=0		-------Note: if pay type is changed then send it from front

	set @A_Gen=(select dbo.GetAutoGen_No(@pay_Type))	-------------------------------

	insert into Payroll_main(
		Emp_Id,
		A_Gen,
		pay_Month ,
		pay_Year ,
		pay_Type,
		U_id	)
	values(
		@Emp_Id,
		@A_gen,
		@pay_Month ,
		@pay_Year ,
		@pay_Type,
		@U_id
		)
	
	exec SP_Payroll_sub_I_U_D @opr,@A_Gen,@Head_Code,@amount
end

if @opr='U'
begin
	set @A_Gen=(select a_gen from payroll_main 
		where emp_id=@emp_id and Pay_month=@pay_month 
		and pay_year=@pay_year and Pay_type=0)

	exec SP_Payroll_sub_I_U_D @opr,@A_Gen,@Head_Code,@amount
	
end
if @opr='D'
begin
	delete from Payroll_main
		where 	emp_id=@emp_id 
			and pay_month=@pay_month
			and pay_year=@pay_year 
			and Pay_type=0

end
if @opr='S'
begin
	set @opr='I'
	set @A_Gen=(select a_gen from payroll_main 
		where emp_id=@emp_id 
			and Pay_month=@pay_month 
			and pay_year=@pay_year 
			and pay_type=0)

	exec SP_Payroll_sub_I_U_D @opr,@A_Gen,@Head_Code,@amount
end


select Message=@Message

set nocount off






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE   Procedure  [SP_Payroll_sub_I_U_D]
@opr varchar(1),
@A_Gen varchar(8),
@Head_Code varchar(2),
@Amount money
as
if @opr='I'
begin
	insert into payroll_sub(
		A_Gen,
		Head_Code,
		Amount			
		)
	values(
		@A_Gen,
		@Head_Code,
		@Amount
		)
end
if @opr='U'
begin
	update payroll_sub set Amount=@Amount
	where 	A_Gen=@A_Gen and Head_Code=@Head_Code
		
end
if @opr='D'
begin
	delete from payroll_sub
	where A_Gen=@A_Gen
end



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_Performance_eval    Script Date: 3/7/02 10:43:51 AM ******/
/****** Object:  Stored Procedure dbo.SP_Performance_eval    Script Date: 10/28/01 4:34:46 PM ******/
/****** Object:  Stored Procedure dbo.SP_Performance_eval    Script Date: 10/22/01 6:24:45 PM ******/
/****** Object:  Stored Procedure dbo.SP_Performance_eval    Script Date: 9/17/00 4:34:17 PM ******/
/****** Object:  Stored Procedure dbo.SP_Performance_eval    Script Date: 9/4/01 6:30:42 PM ******/
/****** Object:  Stored Procedure dbo.SP_Performance_eval    Script Date: 6/20/01 12:00:43 PM ******/
/****** Object:  Stored Procedure dbo.SP_Performance_eval    Script Date: 5/14/01 9:49:54 AM ******/
/****** Object:  Stored Procedure dbo.SP_Performance_eval    Script Date: 1/1/99 3:01:47 AM ******/
/****** Object:  Stored Procedure dbo.SP_Performance_eval    Script Date: 5/9/01 8:13:10 AM ******/
CREATE PROCEDURE [SP_Performance_eval] 
AS
select * from Emp_performance 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_W_ID_NM_For_J    Script Date: 3/7/02 10:43:28 AM ******/
/****** Object:  Stored Procedure dbo.SP_W_ID_NM_For_J    Script Date: 10/28/01 4:34:35 PM ******/
/****** Object:  Stored Procedure dbo.SP_W_ID_NM_For_J    Script Date: 10/22/01 6:24:20 PM ******/
/****** Object:  Stored Procedure dbo.SP_W_ID_NM_For_J    Script Date: 9/17/00 4:34:07 PM ******/
/****** Object:  Stored Procedure dbo.SP_W_ID_NM_For_J    Script Date: 9/4/01 6:30:53 PM ******/
/****** Object:  Stored Procedure dbo.SP_W_ID_NM_For_J    Script Date: 6/20/01 12:00:37 PM ******/
/****** Object:  Stored Procedure dbo.SP_W_ID_NM_For_J    Script Date: 5/14/01 9:49:53 AM ******/
/****** Object:  Stored Procedure dbo.SP_W_ID_NM_For_J    Script Date: 1/1/99 3:01:42 AM ******/
CREATE PROCEDURE [SP_W_ID_NM_For_J]
@ID varchar(10)
 AS
	select Emp_id ,Emp_NM=(Emp_fna+" "+Emp_mna+ " "+Emp_lna) from Emp_per_info 
	where Emp_id=@ID




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.SP_W_Performance_eval    Script Date: 3/7/02 10:43:51 AM ******/
/****** Object:  Stored Procedure dbo.SP_W_Performance_eval    Script Date: 10/28/01 4:34:48 PM ******/
/****** Object:  Stored Procedure dbo.SP_W_Performance_eval    Script Date: 10/22/01 6:24:46 PM ******/
/****** Object:  Stored Procedure dbo.SP_W_Performance_eval    Script Date: 9/17/00 4:34:18 PM ******/
/****** Object:  Stored Procedure dbo.SP_W_Performance_eval    Script Date: 9/4/01 6:30:53 PM ******/
/****** Object:  Stored Procedure dbo.SP_W_Performance_eval    Script Date: 6/20/01 12:00:44 PM ******/
/****** Object:  Stored Procedure dbo.SP_W_Performance_eval    Script Date: 5/14/01 9:49:54 AM ******/
/****** Object:  Stored Procedure dbo.SP_W_Performance_eval    Script Date: 1/1/99 3:01:48 AM ******/
CREATE PROCEDURE [SP_W_Performance_eval] 
@ID Varchar(10)
AS
select A.Emp_Id,Emp_NM=(A.Emp_fna +" "+ A.Emp_mna +" "+ A.Emp_lna),A.Pick_Pic_Path,B.* from   Emp_per_info  A, Emp_performance  B
where A.Emp_Id=@ID   and B.Emp_Id=@ID




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO









CREATE        procedure STD_PERFORMANCE_Save
(
		    @mode	   varchar(1),
            @Student_id    VARCHAR(15),
            @classid     varchar(10),
            @sectionid   varchar(10), 
            @Class_roll VARCHAR(15),
            @Srl_no integer,
				@Topic_srl	INTEGER,
				@Details_srl INTEGER,
            @Prfm    VARCHAR(50),
            @Remarks   VARCHAR(50),
			@Entry_by varchar(10),
			@Entry_date	datetime,
            @Academic_yr varchar(10)  
)
	AS
 
 
         if @mode='S'
               begin
                       
		    			
		                insert into Std_Study_performance
(
						Student_id,
                        classid,  
                        sectionid,
						Class_roll,
						Srl_no,	
						Topic_srl,		
						Details_srl,	
						Prfm,			
						Remarks,
						Entry_by,
						Entry_date,
                        Academic_yr
) 
		                values
(
						@Student_id,
                        @classid,
                        @sectionid,
						@Class_roll,
						@Srl_no,	
						@Topic_srl,		
						@Details_srl,	
						@Prfm,			
						@Remarks,
						@Entry_by,
						@Entry_date,
                        @Academic_yr
) 
		            
               end 

         
          if @mode='U'
               begin
                       
		    			if  exists (select * from Std_Study_performance where Student_id=@Student_id and Class_roll=@Class_roll and Srl_no=@Srl_no and Topic_srl=@Topic_srl and Details_srl=@Details_srl  and classid=@classid and sectionid=@sectionid and Academic_yr=@Academic_yr)
		                  
		                UPDATE Std_Study_performance  
						
								SET 
								
										Prfm=@Prfm,
										Remarks=@Remarks
									
                               
                         where (Student_id=@Student_id and Class_roll=@Class_roll and Srl_no=@Srl_no and Topic_srl=@Topic_srl and Details_srl=@Details_srl  and classid=@classid and sectionid=@sectionid and Academic_yr=@Academic_yr)	                           
		           
               end 


        if @mode='D'
               begin
                       
		    			if  exists (select * from Std_Study_performance where Student_id=@Student_id and Class_roll=@Class_roll and Srl_no=@Srl_no and Topic_srl=@Topic_srl and Details_srl=@Details_srl  and classid=@classid and sectionid=@sectionid and Academic_yr=@Academic_yr )
		                
		                DELETE FROM  Std_Study_performance   where Student_id=@Student_id and Class_roll=@Class_roll and Srl_no=@Srl_no and Topic_srl=@Topic_srl and Details_srl=@Details_srl  and classid=@classid and sectionid=@sectionid and Academic_yr=@Academic_yr
		                           
		            
               end







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE PROCEDURE Salary_Adjustment --'February','2002','1001'
@Operation Varchar(1),
@pay_Month varchar(12),
@pay_Year varchar(4),
@Emp_Id varchar(10)
AS
	select a.Head_code,a.Head_Name,b.Amount  from Pay_Struc a,payroll_sub b
	where a.Mode='V'and a.Operation=@Operation and a.Head_code=b.Head_code and b.A_Gen=
	(Select A_Gen From Payroll_main where
		Emp_Id=@Emp_Id and pay_Month=@pay_Month and pay_Year=@pay_Year and Pay_Stat=0) 
	order by  a.Head_code




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO









/*

delete from payroll_main
exec Salary_AtOnce_Unit1 'may','2002','dsl'
select count(*) from payroll_main where pay_month='may'

*/

CREATE       Procedure Salary_AtOnce_Unit1 
@Department varchar(35)
,@Pay_Month varchar(12)
,@Pay_Year varchar(4)
,@U_id varchar(10)
AS


set nocount on
--set @@lock_timeout=60000

declare @Count int
,@Message varchar(200)

set @Count=0


create table #Pay_Sub
(
A_Gen varchar(8)
,Head_Code varchar(10) 
,Amount money
)

Declare @Emp_ID varchar(10)

--------------------------------------------------------------------------------
DECLARE Get_ID_Cursor CURSOR FOR
select distinct emp_id from fixed_pay where emp_Id in 
	(Select emp_id from emp_Job_Hist_Current where emp_dept=@Department)and Emp_Id not in
	 	(Select emp_id from payroll_main where pay_month=@Pay_Month and pay_year=@Pay_year and pay_Type=0)

order by emp_Id
--------------------------------------------------------------------------------

OPEN Get_ID_Cursor 

    FETCH NEXT FROM Get_ID_Cursor  into @Emp_ID
				
	WHILE (@@FETCH_STATUS = 0)
    		BEGIN

			exec Salary_AtOnce_Unit2 @Emp_Id,@Pay_Month,@Pay_Year,@U_id 	

			FETCH NEXT FROM Get_ID_Cursor  into @Emp_ID
			
			set @Count=@Count+1
		END
		


CLOSE Get_ID_Cursor 
DEALLOCATE Get_ID_Cursor 

if @Count=0 
	set @Message='Salary is previously processed for ' + @Pay_month+', '+@Pay_Year
	

if @count >=1
 set @Message='Process done successfully for '+convert(varchar(6),@Count)+' employees of "'+ @Department + '" for '+ @Pay_Month +', ' +@Pay_Year

	select message =@message,Dpt=@Department,Num=@count

set nocount Off








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





CREATE      Procedure  Salary_AtOnce_Unit2

@Emp_Id varchar(10)
,@Pay_Month varchar(12)
,@Pay_Year varchar(4)
,@U_id varchar(10)

AS

set nocount on


declare @Latest_A_Gen varchar(8)
,@Pay_Type int


set @Latest_A_Gen=(select dbo.GetAutoGen_No(0))
set @Pay_Type=0

insert into payroll_main(Emp_Id,A_Gen,Pay_Type,pay_Month,pay_Year,U_id,Dt)
values(@Emp_Id,@Latest_A_Gen,@Pay_Type,@pay_Month,@pay_Year,@U_id,Getdate())

-------------------Insert Fixed Pay into Temporary table---------------------------------------------------

Insert into #Pay_Sub
select A_Gen=@Latest_A_Gen,Head_code,Amount from fixed_pay where emp_id=@Emp_Id

-------------------Insert Variable Pay into Temporary table-----------------------------------------------

Insert into #Pay_Sub               
select A_Gen=@Latest_A_Gen,Head_Code,Amount=0 from pay_struc where mode='V'

-------------------Update Absent Deduction--------------------------------------------------------------------

declare @Abs_Deduc int
,@GetPreced int

set @Abs_Deduc=(select dbo.GetParamFlag (60))
set @GetPreced=(select dbo.GetParamFlag (61))

if @Abs_Deduc=0		----1 means ignore 

	update #Pay_Sub set amount=(select dbo.GetAbsent_Deduction (@Emp_Id,@Pay_Month,@Pay_Year))
	where Head_Code=(select Head_Code from Pay_Struc where Head_Name='Absent Deduction')


-------------------Update Advance Deduction-----------------------------------------------


--Brings info from loan moudle
/*
if @Advance=1

update #Pay_Sub set amount=(select dbo.GetAdvance_Deduction (@Emp_Id,@Pay_Month,@Pay_Year))
where Head_Code=(select Head_Code from Pay_Struc where Head_Name='Advance Deduction')
*/

-------------------Update Telephone Allowance -----------------------------------------------

--Brings data from immeidate preceding month.
--if @Tel=1


if @GetPreced=1


Begin
/*
	update #Pay_Sub set amount=(select dbo.GetVar_PreMonth (@Emp_Id,@Pay_Month,@Pay_Year,'Telephone Allowance'))
	where Head_Code=(select Head_Code from Pay_Struc where Head_Name='Telephone Allowance')
*/
-------------------Update Others (+) -----------------------------------------------
--Brings data from immeidate preceding month.
--if @Other_Add=1

	update #Pay_Sub set amount=(select dbo.GetVar_PreMonth (@Emp_Id,@Pay_Month,@Pay_Year,'Others(+)'))
	where Head_Code=(select Head_Code from Pay_Struc where Head_Name='Others(+)')

-------------------Update Advance Deduction-----------------------------------------------
--Brings data from immeidate preceding month.
--if @Advance=1

	update #Pay_Sub set amount=(select dbo.GetVar_PreMonth(@Emp_Id,@Pay_Month,@Pay_Year,'Advance Deduction'))
	where Head_Code=(select Head_Code from Pay_Struc where Head_Name='Advance Deduction')

---select * from Pay_struc

	update #Pay_Sub set amount=(select dbo.GetVar_PreMonth(@Emp_Id,@Pay_Month,@Pay_Year,'Telephone/Mobile'))
	where Head_Code=(select Head_Code from Pay_Struc where Head_Name='Telephone/Mobile')


	update #Pay_Sub set amount=(select dbo.GetVar_PreMonth(@Emp_Id,@Pay_Month,@Pay_Year,'HR Deduction'))
	where Head_Code=(select Head_Code from Pay_Struc where Head_Name='HR Deduction')


	update #Pay_Sub set amount=(select dbo.GetVar_PreMonth(@Emp_Id,@Pay_Month,@Pay_Year,'Loan Deduction'))
	where Head_Code=(select Head_Code from Pay_Struc where Head_Name='Loan Deduction')

end




Insert into Payroll_Sub
select * from #Pay_Sub
------------------------------------------------------------------

delete from #Pay_Sub

set nocount off


--Select * from pay_struc



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE PROCEDURE Salry_HeadSetUp_Delete
	@Head_code varchar(2)
	
as 
set nocount on
If Exists (Select * From Pay_Struc 
    Where  Head_code=@Head_code)
	begin		
		Delete from  Pay_Struc 			
		Where  Head_code=@Head_code
	end

set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE PROCEDURE Salry_HeadSetUp_Save
	@Head_code varchar(2),
	@Head_name  varchar(30),
	@Operation varchar(1),
	@Mode varchar(1),
	@U_id varchar(100),
	@Dt datetime
as 
set nocount on
If Exists (Select * From Pay_Struc 
    Where  Head_code=@Head_code)
	begin		
		Update Pay_Struc  Set 
			Head_name=@Head_name ,
			Operation=@Operation ,
			Mode=@Mode
		Where  Head_code=@Head_code
	end
Else

Insert Into Pay_Struc (
	Head_code ,
	Head_name  ,
	Operation ,
	Mode ,
	U_id ,
	Dt 
) Values (
	@Head_code ,
	@Head_name  ,
	@Operation ,
	@Mode ,
	@U_id ,
	@Dt 
)

set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE PROCEDURE SaveLogin_Fail
@u_id varchar (10),
@u_name varchar (40),
@user_pass varchar (45)
AS
insert into Login_Fail values(@u_id,@u_name,@user_pass,getdate())


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO










CREATE       procedure ScholerShipInformation
(
	
	@StudentID		varchar(15),
	@ScholerYear	 	int,
	@ClassId		varchar(5),
	@Type			varchar(30),
	@Name			varchar(80),
	@grade			varchar(30),	
	@Amount			float,
	@NoOfvalidYear		int,
	@EntryBy		varchar(10),
	@Entrydate		datetime
	
)

as

 if exists (select * from ScholerShipinfo where StudentID= @StudentID and ClassId=@ClassId and  Type = @Type and scname=@Name ) 		
Update ScholerShipinfo set

	
	StudentID			=	@StudentID		,
	ScholerYear			=	@ScholerYear	 	,
	ClassId				=	@ClassId		,
	Type				=	@Type			,
	ScName				=	@name			,
	grade				=	@grade			,	
	Amount				=	@Amount			,
	NoOfvalidYear			=	@NoOfvalidYear		,
	EntryBY				=	@EntryBY		,
	EntryDate			=	@EntryDate			
	
	where StudentID= @StudentID and ClassId=@ClassId and  Type = @Type and scname=@Name
else
 insert into ScholerShipinfo 

(
	StudentID		,
	ScholerYear		,
	ClassId			,
	Type			,
	Scname			,
	grade			,
	Amount			,
	NoOfvalidYear		,
        EntryBY			,
	EntryDate
	
)
values
(
	@StudentID		,
	@ScholerYear	 	,
	@ClassId		,
	@Type			,
	@name			,
	@grade			,
	@Amount			,
	@NoOfvalidYear		,
	@EntryBy		,
	@entrydate		
	
)










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO










CREATE       procedure ScholerShipNameSetupInformation
(
	
	@ScholerShipNameId	varchar(5),
	@SchType	 	varchar(5),
	@ScName			varchar(80),
	@SchBy			varchar(80),
	@Address		varchar(80),	
	@DesOfSch		text,
	@EntryBy		varchar(10),
	@Entrydate		datetime
	
)

as

 if exists (select * from scholerShipNameSetup where  ScholerShipNameId= @ScholerShipNameId ) 		
Update scholerShipNameSetup set

	
	ScholerShipNameId		=	@ScholerShipNameId,
	SchType				=	@SchType		,
	ScName				=	@ScName	,
	SchBy				=	@SchBy,
	Address				=	@Address,
	DesOfSch			=	@DesOfSch,
	EntryBY				=	@EntryBY		,
	EntryDate			=	@EntryDate			
	
	where ScholerShipNameId= @ScholerShipNameId 
else
 insert into scholerShipNameSetup 

(
	ScholerShipNameId	,
	SchType	 	,
	ScName			,
	SchBy			,
	Address		,	
	DesOfSch		,
	EntryBy		,
	Entrydate		
	
)
values
(
	@ScholerShipNameId	,
	@SchType	 	,
	@ScName			,
	@SchBy			,
	@Address		,	
	@DesOfSch		,
	@EntryBy		,
	@Entrydate	
)










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO










CREATE       procedure ScholerShipTypeInformation
(
	
	@SchTypeId		varchar(5),
	@Name		 	varchar(80),
	@Notes			varchar(80),
	@EntryBy		varchar(10),
	@Entrydate		datetime
	
)

as

 if exists (select * from ScholershipType where  SchTypeId= @SchTypeId ) 		
Update ScholershipType set

	
	SchTypeId			=	@SchTypeId		,
	ScTypeName			=	@Name			,
	Notes				=	@Notes			,
	EntryBY				=	@EntryBY		,
	EntryDate			=	@EntryDate			
	
	where  SchTypeId= @SchTypeId 
else
 insert into ScholershipType 

(
	SchTypeId		,
	ScTypeName		,
	Notes			,
        EntryBY			,
	EntryDate
	
)
values
(
	@SchTypeId		,
	@Name		 	,
	@Notes			,
	@EntryBy		,
	@entrydate		
	
)










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Sec_Info_I_U_D    Script Date: 3/7/02 10:43:38 AM ******/
/****** Object:  Stored Procedure dbo.Sec_Info_I_U_D    Script Date: 10/28/01 4:34:31 PM ******/
/****** Object:  Stored Procedure dbo.Sec_Info_I_U_D    Script Date: 10/22/01 6:24:31 PM ******/
/****** Object:  Stored Procedure dbo.Sec_Info_I_U_D    Script Date: 9/17/00 4:34:03 PM ******/
/****** Object:  Stored Procedure dbo.Sec_Info_I_U_D    Script Date: 9/4/01 6:30:49 PM ******/
CREATE PROCEDURE Sec_Info_I_U_D
@opr varchar (1),
@Code varchar(3),
@Title varchar (25),
@Prev_Title varchar (25),
@Description varchar (50),
@U_id varchar (10)
AS
if @opr='I'
begin
insert into Sec_Info (
Title,
Code,
Description,
U_id
) 
values (
@Title,
@Code,
@Description,
@U_id
)
end
                    
if @opr='U'
begin
update Sec_Info set 
Title=@Title,
Code=@Code,
Description=@Description,
U_id=@U_id
where Title=@Prev_Title
end
    		
if @opr='D'
begin
delete from Sec_Info where title=@Title
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE      procedure SectionInformation
(
             @mode                            integer, 	
	@SectionID		varchar(5),
	@SectionDsc		varchar(50),
	@ClassID		varchar(5),
	@SectionRoomNo		varchar(5),
	@classteacher		varchar(50),
	@SecMonitor1		varchar(50),
	@SecMonitor2		varchar(50),
	@EntryDate		DateTime	,
	@EntryBY		varchar(50),
              @effectiveDate              datetime
)
AS 
declare @trackid as integer
declare @Historytrackid as integer
begin
   if @mode=1 
         begin
           if not exists (select * from SectionInfo where  ClassID = @ClassID  and SectionID = @SectionID)
             begin

	            select @trackid=isnull(max(trackid),0)+1 from  SectionInfo
                    select @Historytrackid=isnull(max(trackid),0)+1 from  SectionHistoryInfo   
        	    insert into SectionInfo

			(
			SectionID	,
			SectionDsc	,
			ClassID		,
			SectionRoomNo	,
			classteacher	,
			SecMonitor1	,
			SecMonitor2	,
			EntryDate	,
			EntryBY		,
	                           EffectiveDate	,
	                           trackid					
			)

			values
			(
			@SectionID,
			@SectionDsc,		
		 	@ClassID,
			@SectionRoomNo,
			@classteacher,
			@SecMonitor1,
			@SecMonitor2,
			@EntryDate,
			@EntryBY,
	                           @effectiveDate,
	                           @trackid
		
	                   )
                insert into SectionHistoryInfo

			(
			SectionID	,
			SectionDsc	,
			ClassID		,
			SectionRoomNo	,
			classteacher	,
			SecMonitor1	,
			SecMonitor2	,
			EntryDate	,
			EntryBY		,
	                           EffectiveDate	,
	                           trackid
			)

			values
			(
			@SectionID,
			@SectionDsc,		
		 	@ClassID,
			@SectionRoomNo,
			@classteacher,
			@SecMonitor1,
			@SecMonitor2,
			@EntryDate,
			@EntryBY,
	                           @effectiveDate,
	                           @Historytrackid
		
	                   )
                 end 
            else
                begin
                 select @Historytrackid=isnull(max(trackid),0)+1 from  SectionHistoryInfo   
                  insert into SectionHistoryInfo

			(
			SectionID	,
			SectionDsc	,
			ClassID		,
			SectionRoomNo	,
			classteacher	,
			SecMonitor1	,
			SecMonitor2	,
			EntryDate	,
			EntryBY		,
	                           EffectiveDate	,
	                           trackid
				
	
			)
			values
			(
			@SectionID,
			@SectionDsc,		
		 	@ClassID,
			@SectionRoomNo,
			@classteacher,
			@SecMonitor1,
			@SecMonitor2,
			@EntryDate,
			@EntryBY,
	                           @effectiveDate,
	                           @Historytrackid
			
	                   )
              
          end 
       end   ----end of mode 1

end ---end of as begin
/*
Update SectionInfo set

	
	SectionID	=	@SectionID,
	SectionDsc	=	@SectionDsc,		
	ClassID		=	@ClassID,
	SectionRoomNo	=	@SectionRoomNo,
	classteacher	=	@classteacher,
	SecMonitor1	=	@SecMonitor1,
	SecMonitor2	=	@SecMonitor2,
	EntryDate	=	@EntryDate,
	EntryBY		=	@EntryBY	
	
	where ClassID = @ClassID and SectionID = @SectionID
else
*/
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Section_I_U_D    Script Date: 3/7/02 10:43:38 AM ******/
/****** Object:  Stored Procedure dbo.Section_I_U_D    Script Date: 10/28/01 4:34:31 PM ******/
/****** Object:  Stored Procedure dbo.Section_I_U_D    Script Date: 10/22/01 6:24:31 PM ******/
/****** Object:  Stored Procedure dbo.Section_I_U_D    Script Date: 9/17/00 4:34:03 PM ******/
/****** Object:  Stored Procedure dbo.Section_I_U_D    Script Date: 9/4/01 6:30:41 PM ******/
CREATE PROCEDURE Section_I_U_D
@opr varchar (1),
@Title varchar (25),
@Prev_Title varchar (25),
@Description varchar (50),
@U_id varchar (10)
AS
if @opr='I'
begin
insert into Sec_Info(
Title,
Description,
U_id
) 
values (
@Title,
@Description,
@U_id
)
end
                    
if @opr='U'
begin
update Sec_Info set 
Title=@Title,
Description=@Description,
U_id=@U_id
where Title=@Prev_Title
end
    		
if @opr='D'
begin
delete from Sec_Info where title=@Title
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




--exec Seek_Id_Not_Paid 'Prep','0','October','2004','Daffodil Software Ltd.','Accountant'

----SELECT emp_id from emp_job_Hist_Current
/**--------------Salary Preparation---------**/

CREATE Procedure Seek_Id_Not_Paid
	@Mode varchar(10),
	@pay_type varchar(2),
	@pay_Month varchar(12),
	@pay_Year varchar(4),
	@Emp_dept varchar(45),
	@Emp_Desig varchar(45)
AS
if @Mode='Prep'
Begin
	select distinct emp_Id from fixed_Pay where emp_id in 
		(select emp_id from emp_job_Hist_Current where Emp_dept=@Emp_dept 
			and Emp_Desig=@Emp_Desig) and emp_Id not in 
				(select emp_Id from payroll_main where 
					pay_Month=@pay_Month and  pay_Year=@pay_Year and pay_type=@pay_type) 
end
if @Mode='Edit'
Begin
	select distinct emp_Id from fixed_Pay where emp_Id in 
		(select emp_id from emp_job_Hist_Current where Emp_dept=@Emp_dept 
					and Emp_Desig=@Emp_Desig) and emp_Id in
				(select emp_Id from payroll_main where 
					pay_Month=@pay_Month and pay_Year=@pay_Year and Pay_Stat=0 and pay_type=@pay_type)
end





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE Procedure Set_Param

@Policy_No int
,@Flag int
,@Value varchar(50)

AS

if @Flag=0
	begin
		set @Value=0
	end

update Param_Tbl set Flag=@Flag,Value=@Value where  Policy_No= @Policy_No 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







/****** Object:  Stored Procedure dbo.Shift_I_U_D    Script Date: 10/22/01 6:24:31 PM ******/

/****** Object:  Stored Procedure dbo.Shift_I_U_D    Script Date: 9/17/00 4:34:04 PM ******/


------------------------------------------

CREATE  procedure Shift_I_U_D

@opr varchar (1),
@Shift_Code varchar (5),
@Shift_Name varchar (15),
@Shift_Start varchar (15),
@Shift_End varchar (15),
@Delay varchar (15),
@Abs_Time varchar (15),
@U_id    varchar(10)

as


if @opr='I'

begin
	insert into Shift(Shift_Code,Shift_Name,Shift_Start,Shift_End,[Delay],Abs_Time,U_id)
	values(@Shift_Code,@Shift_Name,@Shift_Start,@Shift_End,@Delay,@Abs_Time,@U_id)
end 


if @opr='U'

begin
	Update  Shift Set Shift_Name=@Shift_Name,Shift_Start=@Shift_Start,Shift_End=@Shift_End,
	[Delay]=@Delay,Abs_Time=@Abs_Time,U_id=@U_id
	Where Shift_Code=@Shift_Code
end 

if @opr='D'

begin
	Delete from Shift
	Where Shift_Code=@Shift_Code
end 










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Show_PreviousTiers    Script Date: 3/7/02 10:43:38 AM ******/
/****** Object:  Stored Procedure dbo.Show_PreviousTiers    Script Date: 10/28/01 4:34:31 PM ******/
/****** Object:  Stored Procedure dbo.Show_PreviousTiers    Script Date: 10/22/01 6:24:32 PM ******/
/****** Object:  Stored Procedure dbo.Show_PreviousTiers    Script Date: 9/17/00 4:34:04 PM ******/
/****** Object:  Stored Procedure dbo.Show_PreviousTiers    Script Date: 9/4/01 6:30:49 PM ******/
CREATE PROCEDURE Show_PreviousTiers
@Tier varchar(1),
@App_ID varchar(10)
 AS
if @Tier='1'
begin
select Tier_1_Chk,Tier_1_Remarks from Approval_Audit where app_ID=@app_ID
end
if @Tier='2'
begin
	select Tier_2_Chk,Tier_2_Remarks from Approval_Audit where app_ID=@app_ID
end
if @Tier='3'
begin
	select Tier_3_Chk,Tier_3_Remarks from Approval_Audit where app_ID=@app_ID
end
if @Tier='4'
begin
	select Tier_4_Chk,Tier_4_Remarks from Approval_Audit where app_ID=@app_ID
end
if @Tier='5'
begin
	select Final_Tier_Chk Final_Tier_Remarks from Approval_Audit where app_ID=@app_ID
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE Procedure Software_Previleges
@MODE varchar(10),
@U_Id varchar(10),
@Portion varchar(5)
AS

IF @MODE='Not_In'

BEGIN
	IF @Portion='ALL' 

	BEGIN
		select scr_no,descript,code from soft_bag where code not in (select code from permit where u_id=@U_ID)
		order by scr_no
	END
	
	ELSE

	BEGIN
		select scr_no,descript,code from soft_bag where code not in (select code from permit where u_id=@U_ID)
		and portion=@portion order by scr_no
	END
END

IF @MODE='In'

BEGIN

	IF @Portion='ALL' 

	BEGIN
		select scr_no,descript,code from soft_bag where code in (select code from permit where u_id=@U_ID)
		order by scr_no
	END
	
	ELSE

	BEGIN
		select scr_no,descript,code from soft_bag where code in (select code from permit where u_id=@U_ID)
		and portion=@portion order by scr_no
	END



END










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



-----exec Sp_Emp_Loan_Info ''

CREATE    PROCEDURE Sp_Emp_Loan_Info
@Emp_Id varchar(10)
as


if @Emp_Id <>''

		SELECT Emp_id, Loan_Type, InstalmentNo,
		GracePeriod, SanctionAmount, 
		InstallmentAmount, SanctionDate,
		IstInstallmentDatePrin, IstInstallmentDateInt, 
		InstallmentAmountInt , InterestInstallmetNo From LoanSanction_Info
		WHERE     (Emp_id = @Emp_Id)

else
	SELECT Emp_id, Loan_Type, InstalmentNo,
		GracePeriod, SanctionAmount, 
		InstallmentAmount, SanctionDate,
		IstInstallmentDatePrin, IstInstallmentDateInt, 
		InstallmentAmountInt , InterestInstallmetNo From LoanSanction_Info
		





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Sp_Rpt_Emp_Per    Script Date: 3/7/02 10:43:51 AM ******/
/****** Object:  Stored Procedure dbo.Sp_Rpt_Emp_Per    Script Date: 10/28/01 4:34:47 PM ******/
/****** Object:  Stored Procedure dbo.Sp_Rpt_Emp_Per    Script Date: 1/29/01 10:24:55 AM ******/
/****** Object:  Stored Procedure dbo.Sp_Rpt_Emp_Per    Script Date: 5/14/01 9:49:54 AM ******/
/****** Object:  Stored Procedure dbo.Sp_Rpt_Emp_Per    Script Date: 1/1/99 3:01:47 AM ******/
CREATE procedure Sp_Rpt_Emp_Per
@Emp_id varchar(10),
@Mode varchar (10)
as
if @Mode='Ind'
begin
	select
	a.Emp_id,
	Emp_fna=(a.Emp_fna+" "+a.Emp_mna+" "+a.Emp_lna),
	a.Emp_fa_na,
	a.Emp_ma_na,
	a.Emp_d_of_b,
	a.Emp_age,
	a.Blood_group,
	a.Emp_marital_st,
	a.Emp_gender ,
	a.Emp_nat ,
	a.Emp_sp_qta ,
	a.Emp_religion ,
	a.Emp_perm_add ,
	a.Emp_perm_town ,
	a.Emp_post1 ,
	a.Emp_tel1 ,
	a.Contact_person,
	a.Contact_add,
	a.Contact_town,
	a.Contact_post,
	a.Contact_tel,
	a.Contact_fax,
	a.Contact_email,
	a.Emp_eye,
	a.Emp_height,
	a.Emp_weight,
	a.Emp_disable,
	
	b.Emp_desig,
	b.Emp_Section
	from Emp_Per_info a,Emp_Job_Hist_current b
	Where a.Emp_id=b.Emp_id and a.Emp_id=@Emp_id
end
if @Mode='All'
begin
select
	a.Emp_id,
	Emp_fna=(Emp_fna+" "+Emp_mna+" "+Emp_lna),
	a.Emp_fa_na,
	a.Emp_d_of_b,
	a.Emp_age,
	a.Blood_group,
	a.Emp_gender ,
	a.Emp_religion ,
	a.Emp_perm_add ,
	a.Emp_perm_town ,
	a.Emp_tel1 ,
	a.Contact_person,
	a.Contact_tel,
	b.Emp_desig,
	b.Emp_Section
	from Emp_Per_info a,Emp_Job_Hist_current b
	Where a.Emp_id=b.Emp_id
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Sp_rpt_Emp_Att    Script Date: 3/7/02 10:43:51 AM ******/
/****** Object:  Stored Procedure dbo.Sp_rpt_Emp_Att    Script Date: 10/28/01 4:34:47 PM ******/
/****** Object:  Stored Procedure dbo.Sp_rpt_Emp_Att    Script Date: 10/22/01 6:24:45 PM ******/
/****** Object:  Stored Procedure dbo.Sp_rpt_Emp_Att    Script Date: 9/17/00 4:34:17 PM ******/
/****** Object:  Stored Procedure dbo.Sp_rpt_Emp_Att    Script Date: 9/4/01 6:31:01 PM ******/
/****** Object:  Stored Procedure dbo.Sp_rpt_Emp_Att    Script Date: 6/20/01 12:00:43 PM ******/
/****** Object:  Stored Procedure dbo.Sp_rpt_Emp_Att    Script Date: 5/14/01 9:49:54 AM ******/
/****** Object:  Stored Procedure dbo.Sp_rpt_Emp_Att    Script Date: 1/1/99 3:01:47 AM ******/
CREATE procedure Sp_rpt_Emp_Att
@Emp_id varchar(10),
@Mode varchar(10),
@Entry_Dt  datetime
as
if @Mode='Ind'
begin
select 
A.Emp_id,
Emp_fna=(A.Emp_fna+" "+a.Emp_mna+" "+A.Emp_lna),
B.Emp_desig,
B.Emp_Section,
C.Emp_Login,
C.Emp_logout
from Emp_per_info A,Emp_Job_Hist_Current B,Emp_Att_Info c
where A.emp_id=B.emp_id and  C.emp_id=B.emp_id and A.emp_Id=@Emp_id
end
if @Mode='All'
begin
select 
A.Emp_id,
Emp_fna=(A.Emp_fna+" "+a.Emp_mna+" "+A.Emp_lna),
B.Emp_desig,
B.Emp_Section,
C.Emp_Login,
C.Emp_logout
from Emp_per_info A,Emp_Job_Hist_Current B,Emp_Att_Info c
where  C. Entry_Dt=@Entry_Dt  
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

create  procedure StuAdmissionEvaluationInformation
(
	@mode                                   integer,
	@StudentID				varchar(15),
	@AdmissionEvaluationDate		datetime,
	@Shift					varchar(30),
	@ClassID				varchar(5),
	@SectionID				varchar(5),
	@ClassRoll				int,
	@EntryBy				varchar(10),
	@Entrydate				Datetime,
	@Approval				varchar(1),
	@AdmissionCancel			varchar(1),
	@Active					varchar(1),
	@ActiveClass				varchar(1),
        @Aca_yr				        varchar(6) 
)
As
begin

if @mode=1 
   begin
 if  exists(select StudentID from StudentAdmission where StudentID=@StudentID and ClassID=@ClassID and Aca_yr=@Aca_yr )
   begin
update StudentAdmission set
	AdmissionDate		= 	@AdmissionEvaluationDate,
	Shift			= 	@Shift,
	ClassID			= 	@ClassID,
	SectionID		= 	@SectionID,
	ClassRoll		= 	@ClassRoll,
	EntryBy			= 	@EntryBy,
	Entrydate		= 	@Entrydate,
	Approval		= 	@Approval,
	AdmissionCancel		= 	@AdmissionCancel,
        Aca_yr                  =       @Aca_yr

where StudentID=@StudentID and ClassID=@ClassID and  Aca_yr=@Aca_yr
end
if not exists(select StudentID from StudentAdmission where StudentID=@StudentID and ClassID=@ClassID and Aca_yr=@Aca_yr)
  begin
   declare @max_serial as int

  select   @max_serial=isnull(max(serial_no),0)+1 from StudentAdmission

 insert into StudentAdmission
(
	StudentID		,
	AdmissionDate		,
	Shift			,
	ClassID			,
	SectionID		,
	ClassRoll		,
	EntryBy			,
	Entrydate		,
	--AdmitApproveBy		,
	--AdmitApproveDate	,
	Approval		,
	AdmissionCancel,
    serial_no,
     aca_yr,
     active_std
	
)
values
(
	@StudentID		,
	@AdmissionEvaluationDate		,
	@Shift			,
	@ClassID			,
	@SectionID		,
	@ClassRoll		,
	@EntryBy			,
	@Entrydate		,
	--@AdmitApproveBy		,
	--@AdmitApproveDate	,
	@Approval,	
	@AdmissionCancel,
    @max_serial,
    @aca_yr,1)
end 
if exists(select StudentID from StudentEvaluation where StudentID=@StudentID)
  begin
update StudentEvaluation set


	
	EvaluationDate		= @AdmissionEvaluationDate,
	Shift			= @Shift,
	ClassID			= @ClassID,
	SectionID		= @SectionID,
	ClassRoll		= @ClassRoll,
	EntryBy			= @EntryBy,
	Entrydate		= @Entrydate,
	Active			= @Active,
	ActiveClass		=@ActiveClass




where StudentID=@StudentID
end 
else
  begin

 insert into StudentEvaluation

(
	StudentID		,
	EvaluationDate		,
	Shift			,
	ClassID			,
	SectionID		,
	ClassRoll		,
	EntryBy			,
	Entrydate		,
	Active			,
	ActiveClass
		
	
)
values
(
	@StudentID		,
	@AdmissionEvaluationDate		,
	@Shift			,
	@ClassID			,
	@SectionID		,
	@ClassRoll		,
	@EntryBy			,
	@Entrydate		,
	@Active			,
	@ActiveClass
    
)
 end
end -------end of mode =1

if @mode=2
      begin
        delete from StudentAdmission where StudentID=@StudentID and ClassID=@ClassID and Aca_yr=@Aca_yr
  end 




end










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

create           procedure StudentAttendance 
(
        @mode               varchar(1),	
	@StudentId			varchar(15),
	@Shift	 			Varchar(30),
	@ClassID			varchar(5),
	@SectionID			varchar(5),
	@classRoll			int,
	@Present			varchar(1),
	@PresentCancel		       varchar(1),
        @attn_date			datetime,
        @aca_yr                        varchar(10),
        @exam_term                     varchar(10),
        @exam_type                     varchar(10),
        @leave                         varchar(1)
)

as
  declare @var_std_id varchar(15) 
  declare @class_roll_no int
  if @mode='a'
       begin
           declare my_cursor cursor for
           select distinct(StudentId)  from StudentAdmission 
           where ClassId=@ClassID 
                 and SectionId=@SectionID
                 and Shift=@Shift and active_std=1
           open my_cursor
         
          fetch next from my_cursor into @var_std_id
          select @class_roll_no=classRoll 
              from  StudentAdmission
         where StudentId=@var_std_id 
             and serial_no=(select max(serial_no) 
                from StudentAdmission 
                  where StudentId=@var_std_id )

          while @@fetch_status=0
           begin
            insert into StudentAttendanceLeaveInfo
			(

				StudentId	,
				Shift	 	,
				ClassID		,
				SectionID	,
				classRoll	,
				Present		,
				EntryTime	,
                                  attn_date,
				PresentCancel,
                                aca_yr,
                                exam_term,
                                Exam_type,
                                leave
				
			)
			values
			(
				@var_std_id	,
				@Shift	 	,
				@ClassID		,
				@SectionID	,
				@class_roll_no	,
				@Present		,
				getdate()	,
                                @attn_date,
				@PresentCancel	,
                                @aca_yr,
                                @exam_term,
                                @exam_type,
                                @leave 
			)
           
                    
          fetch next from my_cursor into @var_std_id
          select @class_roll_no=classRoll 
              from  StudentAdmission
         where StudentId=@var_std_id 
             and serial_no=(select max(serial_no) 
                from StudentAdmission 
                  where StudentId=@var_std_id ) 

           end
      end
               
if @mode='s'
       begin
         select @class_roll_no=classRoll 
              from  StudentAdmission
         where StudentId=@StudentId 
             and serial_no=(select max(serial_no) 
                from StudentAdmission 
             where StudentId=@StudentId )

           insert into StudentAttendanceLeaveInfo
			(
				StudentId	,
				Shift	 	,
				ClassID		,
				SectionID	,
				classRoll	,
				Present		,
				EntryTime	,
                                attn_date,
				PresentCancel,
                                aca_yr,
                                exam_term,
                                Exam_type,
                                leave
				
			)
			values
			(
				@StudentId	,
				@Shift	 	,
				@ClassID		,
				@SectionID	,
				@class_roll_no	,
				@Present		,
				getdate(),
                                @attn_date,
				@PresentCancel	,
                                @aca_yr,
                                @exam_term,
                                @exam_type,
                                @leave 
			)
          
                 
      end

     if @mode='u'
          begin
             update StudentAttendanceLeaveInfo
		  set Present=@Present,
                    EntryTime=getdate(),
                    PresentCancel=@PresentCancel,
                    leave=@leave	         	
                where  ClassID=@ClassID	
                      and SectionID=@SectionID
                      and Shift=@Shift  
                      and attn_date= @attn_date
          end

  if @mode='p'
          begin
             update StudentAttendanceLeaveInfo
		  set Present=@Present,
                    EntryTime=getdate(),
                    PresentCancel=@PresentCancel,
                    leave=@leave	         	
                where  ClassID=@ClassID	
                      and SectionID=@SectionID
                      and Shift=@Shift  
                      and attn_date= @attn_date
                      and StudentId=@StudentId
          end


close my_cursor
deallocate my_cursor











GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO










CREATE       procedure StudentEvaluation1

(
	@StudentID				varchar(15),
	@EvaluationDate				datetime,
	@Active				varchar(1),
	@ActiveClass				varchar(1),
	@AdmissionCancel			varchar(1),
	@EntryBy				varchar(10),
	@Entrydate				Datetime,
	@ClassId				varchar(5)
	
	
	
	
)

AS if exists (select StudentId  from StudentEvaluation where   StudentId = @StudentID  )	
Update StudentEvaluation set

	EvaluationDate		=	@EvaluationDate,
	Active			=	@Active,
	ActiveClass		=	@ActiveClass
	
	where StudentId=@StudentId and ClassId=@ClassId

IF EXISTS(select StudentId  from StudentAdmission where   StudentId = @StudentID  )	
Update StudentAdmission set
	AdmissionCancel		=	@AdmissionCancel,
	EntryBy			=	@EntryBy,
	Entrydate		=	@Entrydate

where StudentId=@StudentId










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






---exec StudentInformation Auto Number
---exec StudentInformation '454s' ,'dfs','fdsfds','fdsfds','fdsafs',null,'n',1,1,'001','islam','12/12/2004','n','n','12/12/2004','dsl'
----saeffasdf,asdf,asdf,asdf,,N,3,4,18,Islam,12/12/2012,N,N,01/02/2006,DSL,
CREATE        procedure StudentInformation
(
	
	@StudentID		varchar(15),
	@StudentName	 	varchar(80),
	@StuFatherName		varchar(80),
	@StuMotherName		varchar(80),
	@LegalGerdian		varchar(80),	
	@StuMarraigeDate	Datetime,
	@StuMoorFalet		varchar(1),
	@Stubrono		tinyint,
	@Stusisno		tinyint,
	@StuCountryofBirth	varchar(5),
	@StuReligion		varchar(20),
	@StuDateofBirth		Datetime,
	@Computer		varchar(1),
	@Internet		varchar(1),
	@EntryDate		DateTime,
	@EntryBY		varchar(50)
	
)

as

Declare 

	@MaxSLNo int,
	@StID char(15)



if @StudentID is null

	begin
	
		/*SELECT @MaxSLNo = select  isnull(max(cast(StudentID as int)),0)
			FROM StudentInfo */
		
		SELECT  @MaxSLNo = isnull(max(cast(substring(StudentID,5,11) as int)),0)
				FROM StudentInfo 



		if @MaxSLNo=0
			set @MaxSLNo=1
		else
			set @MaxSLNo=@MaxSLNo+1

		select @StID= dbo.PadString(@MaxSLNo,6,'0','L')

				
		set @StudentID='STI-'+@StID
	end

if exists (select StudentID from StudentInfo where  StudentID= @StudentID) 		

Update StudentInfo set

	
	StudentID			=	@StudentID		,
	StudentName	 		=	@StudentName	 	,
	StuFatherName			=	@StuFatherName		,
	StuMotherName			=	@StuMotherName		,
	LegalGerdian			=	@LegalGerdian		,
	StuMarraigeDate			=	@StuMarraigeDate	,
	StuMoorFalet			=	@StuMoorFalet		,
	Stubrono			=	@Stubrono		,
	Stusisno			=	@Stusisno		,
	StuCountryofBirth		=	@StuCountryofBirth	,
	StuReligion			=	@StuReligion		,
	StuDateofBirth			=	@StuDateofBirth		,
	Computer			=	@Computer		,
	Internet			=	@Internet		,
	stuEntryDate			=	@EntryDate		,
	stuEntryBY			=	@EntryBY		
	
	
	where StudentID= @StudentID 	
else
 insert into StudentInfo 

(
	StudentID			,
	StudentName	 		,
	StuFatherName			,
	StuMotherName			,
	LegalGerdian			,
	StuMarraigeDate			,
	StuMoorFalet			,
	Stubrono			,
	Stusisno			,
	StuCountryofBirth		,
	StuReligion			,
	StuDateofBirth			,
	Computer			,
	Internet			,
	stuEntryDate			,
    stuEntryBY			
		
	
)
values
(
	@StudentID			,
	@StudentName	 		,
	@StuFatherName			,
	@StuMotherName			,
	@LegalGerdian			,
	@StuMarraigeDate			,
	@StuMoorFalet			,
	@Stubrono			,
	@Stusisno			,
	@StuCountryofBirth		,
	@StuReligion			,
	@StuDateofBirth			,	
	@Computer			,
	@Internet			,
	@EntryDate			,

        @EntryBY			
)










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE     procedure StudentInformation1
(
	
	@StudentID		varchar(15),
	@Stuhight		Float,
	@StuWeight		Float,
	@StuBloodGroup		varchar(5),
	@NewVaccineDate		Datetime
	
)


 as if exists (select * from StudentInfo where  StudentID= @StudentID) 		
Update StudentInfo set

	
	StudentID			=	@StudentID		,
	Stuhight			=	@Stuhight		,
	StuWeight			=	@StuWeight		,
	StuBloodGroup			=	@StuBloodGroup		,
	Nextvaccinedate			=	@NewVaccineDate
	
	
	where StudentID= @StudentID 	
else
 insert into StudentInfo 

(
	StudentID			,
	Stuhight			,
	StuWeight			,
	StuBloodGroup			,
	Nextvaccinedate
	
)
values
(
	@StudentID			,
	@Stuhight			,
	@StuWeight			,
	@StuBloodGroup			,
	@NewVaccineDate
			
)








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE     procedure StudentInformation2
(
	
	@StudentID		varchar(15),
	@StuStreetPAddress	varchar(150),
	@StuPDistrict		varchar(5),
	@StuPCountry		varchar(5),
	@StuCStreetAddress	varchar(150),
	@StuCDistrict		varchar(5),
	@StuCCountry		varchar(5),
	@StuPhone		varchar(50),
	@StuEmail		varchar(50),
	@ImmAdress		varchar(150),
	@ImmPhone		varchar(50),
	@ImmMob			varchar(50)
	
	
)

as




 if exists (select * from StudentInfo where  StudentID= @StudentID) 		
Update StudentInfo set

	
	StudentID			=	@StudentID		,
	StuStreetPAddress		=	@StuStreetPAddress	,
	StuPDistrict			=	@StuPDistrict		,
	StuPCountry			=	@StuPCountry		,
	StuCStreetAddress		=	@StuCStreetAddress	,
	StuCDistrict			=	@StuCDistrict		,
	StuCCountry			=	@StuCCountry		,
	StuPhone			=	@StuPhone		,
	StuEmail			=	@StuEmail		,
	ImmAddress			=	@ImmAdress		,
	ImmPhone			=	@ImmPhone		,
	ImmMob				=	@ImmMob
	
	where StudentID= @StudentID 	
else
 insert into StudentInfo 

(
	StudentID			,
	StuStreetPAddress		,
	StuPDistrict			,
	StuPCountry			,
	StuCStreetAddress		,
	StuCDistrict			,
	StuCCountry			,		
	StuPhone			,
	StuEmail			,
	ImmAddress			,
	ImmPhone			,
	ImmMob		
	
	
)
values
(
	@StudentID			,
	@StuStreetPAddress		,
	@StuPDistrict			,
	@StuPCountry			,
	@StuCStreetAddress		,
	@StuCDistrict			,
	@StuCCountry			,	
	@StuPhone			,
	@StuEmail			,
	@ImmAdress			,
	@ImmPhone			,
	@ImmMob
)








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO










CREATE       procedure Studentadmission1

(
	@StudentID				varchar(15),
	@Active					varchar(1),
	@AdmitApproveDate			datetime,
	@APProval				varchar(1),
	@AdmitApproveBy				varchar(10)
	
	
	
	
)

AS if exists (select  StudentId from StudentEvaluation where   StudentId = @StudentID  )	
Update StudentEvaluation set



	
	Active			=	@Active,
	ApproveBy		=	@AdmitApproveBy,
	ApproveDate		=	@AdmitApproveDate
	
	where StudentId=@StudentId

IF EXISTS(select StudentId from StudentAdmission where   StudentId = @StudentID  )	
Update StudentAdmission set
	AdmitApproveDate		=	@AdmitApproveDate,
	APProval			=	@APProval,
	AdmitApproveBy			=	@AdmitApproveBy
	
where StudentId=@StudentId










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE    procedure SubjectInformation
(
	@SubjectID	varchar(5),
	@ClassID	Varchar(30),
	@Subjectdsc	Varchar(100),
	@TotalMarks	int,
	@Subjectunit	varchar(3),
	@SubjectType	varchar(30),
	@EntryBy	varchar(10),
	@entryDate	datetime
	
)

AS if exists (select * from SubjectInfo where SubjectID = @SubjectID and ClassID = @ClassID )
Update SubjectInfo set

	SubjectID	= 	@SubjectID,	
	ClassID		=	@ClassID,	
	Subjectdsc	=	@Subjectdsc,
	TotalMarks	=	@TotalMarks,
	Subjectunit	=	@Subjectunit,
	SubjectType	=	@SubjectType,	
	EntryBy		=	@EntryBy,
	Entrydate	=	@Entrydate


	where SubjectID = @SubjectID and ClassID = @ClassID 

else
 insert into SubjectInfo 

(
	SubjectID,
	ClassID,
	Subjectdsc,
	TotalMarks,
	Subjectunit,
	SubjectType,	
	EntryBy,
	Entrydate
)
values
(
	@SubjectID,	
	@ClassID,	
	@Subjectdsc,
	@TotalMarks,
	@Subjectunit,			
	@SubjectType,
	@EntryBy,
	@Entrydate
)






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO








CREATE          procedure SubjectInformation_SUB
(
		@Mode varchar(2),
		@M_code	varchar(10),
                           @Sub_code	varchar(10),
               	@class_code	varchar(10),  
		@Sub_title	Varchar(200),          
		@EntryBy	varchar(10),
		@entryDate	datetime
                
                  
				
)
---with encryption
AS
		
		if @mode='S' 
                       begin
 
   		if not exists (select M_code from Subject_Info_sub where  Sub_code = @Sub_code and Class_code = @class_code )
                           begin

                             insert into  subject_info_sub
                                  (M_code,Sub_code,class_code,Sub_title,Entry_by,Entry_date)

                             values(
                                        @M_code,@Sub_code,@class_code,@Sub_title,
                                	 @EntryBy,@entryDate)
             	           end 
/*
                          update  subject_info_sub
                                  set Sub_title=@Sub_title,
                                     Teacher_id=@techar_id
                         where M_code = @M_code and Sub_code = @Sub_code and  class_code=@class_code 

                          select @trackid=isnull(max(trackid),0)+1 from   subject_info_sub_History

                         insert into  subject_info_sub_History
                                  (M_code,Sub_code,class_code,Sub_title,Teacher_id,Entry_by,Entry_date,trackid)

                             values(
                                    @M_code,@Sub_code,@class_code,@Sub_title,
                                @techar_id ,@EntryBy,@entryDate,@trackid)
*/
			end 



        
            if @mode='U' 
                  begin
 
	            if  exists (select M_code from Subject_Info_sub where M_code = @M_code and Sub_code = @Sub_code and class_code=@class_code )

                         update  subject_info_sub
                                  set Sub_title=@Sub_title
                                    
                         where M_code = @M_code and Sub_code = @Sub_code and  class_code=@class_code 

	     end 


/*

          

            if @mode='D' 
                  begin
 
	          if  exists (select M_code from Subject_Info_sub where M_code = @M_code and Sub_code = @Sub_code and class_code=@class_code and trackid=@trackid)

                         delete from  subject_info_sub
                                where M_code = @M_code and Sub_code = @Sub_code and class_code=@class_code and  trackid=@trackid
                 end
                           
	 


*/



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE  procedure SubjectInformation_SUB_Teacher
(           @mode	   varchar(1),
            @M_code    varchar(10),
            @S_code    varchar(10),
            @Classid	varchar(10), 
            @teacher varchar(50),  
            @effective_date     datetime,
            @entryby varchar(10),
            @entrydate   datetime,
            @trackid     int
			
           
)
	AS

         declare  @u_id as varchar(20)
         declare  @class_code as varchar(20)
         
         
         
        
         declare @loc_C_srl as integer
         declare @srl_no as integer 

    
 
 
       if @mode='s'
               begin
                    set @loc_c_srl=(select isnull(max(trackid),0)+1 from subject_info_sub_history)
                    insert into subject_info_sub_history values
                 ( @M_code,@S_code,@Classid,'',@teacher,@entryby,@entrydate,@loc_c_srl,@effective_date) 
		       
               end
              
/*
                      declare collec_cursor cursor for
                        select Srl_no,u_id,class_code,Fee_code,Act_amount,Fine,Discount,std_id
                               from temp_collect
                      where seq_no=@seq_no


		    			set @loc_c_srl=(select isnull(max(C_srl),0)+1 from Collec_master)

                   Open collec_cursor

 					   Fetch Next From collec_cursor into @Srl_no,@u_id,@class_code,@Fee_code,@Act_amount,@Fine,@Discount,@std_id

                       insert into Collec_master(C_srl,Std_id,class_id,Mon,Yr,Remark,Entry_by,Entry_date,Collec_date) 
				                            values(@loc_c_srl,@Std_id,@class_code,@Mon,@Yr,@Remark,@EntryBy,@Entrydate,@Collec_date)	

                        set @loc_c_srl=(select isnull(max(C_srl),0) from Collec_master)

                     ---       

                      While @@Fetch_Status = 0
                           begin
			                
         				    insert into Collec_details(C_Srl,serial_no,Fee_code,Act_Amount,Discount,Fine,Entry_by,Entry_date)
                                       values( @loc_c_srl,@Srl_no,@Fee_code,@Act_Amount,@Discount,@Fine,@Entryby,@Entrydate) 
		                                                         

                          Fetch Next From collec_cursor into @Srl_no,@u_id,@class_code,@Fee_code,@Act_amount,@Fine,@Discount,@std_id
                   end 
               End

              delete from temp_collect
           where seq_no=@seq_no
       
	        Close collec_cursor
	       Deallocate collec_cursor
      



       if @mode='u'
               begin
		    			if  exists (select C_srl from Collec_master where C_srl=@C_srl)
		
		                 update Collec_master set 
                                             Std_id=@Std_id,
											 Class_id=@class_id, 
                                             Remark=@Remark,
		                                     Entry_by=@EntryBy,
		                                     Entry_date=@Entrydate

                        where C_srl=@C_srl
                        
                       if  exists (select C_srl from Collec_details where C_srl=@C_srl and fee_code=@fee_code)
                            begin
		                         update Collec_details set 
		                                             Fee_code=@Fee_code,
													 amount=@amount, 
		                                             Entry_by=@EntryBy,
				                                     Entry_date=@Entrydate
		
		                        where C_srl=@C_srl and fee_code=@fee_code 
		                   end  



   
                                               
                  end 
 

       if @mode='d'
               begin
		    			if  exists (select C_srl from Collec_master where C_srl=@C_srl)
		
		                 delete from Collec_master where C_srl=@C_srl  
                         delete from Collec_details where C_srl=@C_srl    
                end 


       if @mode='p'
               begin
		    			if  exists (select C_srl from Collec_details where C_srl=@C_srl and fee_code=@fee_code)
		
		                 delete from Collec_details where C_srl=@C_srl and fee_code=@fee_code  
                end 
 
 

*/









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





---exec SubjectInformation_main 'S','00004','mathemaics','1','compulsory,','dsl','12-mar-2006'



CREATE    procedure SubjectInformation_main
(
			    @Mode varchar(2),
				@M_code	varchar(10),
			    @M_title	Varchar(200),
				@Subjectunit	varchar(3),
				@SubjectType	varchar(30),
				@EntryBy	varchar(10),
				@entryDate	datetime
				
)

AS
		
			if @mode='S' 
                  begin
 
					 if not exists (select M_code from SubjectInfomain where M_code = @M_code)

                         insert into  subjectinfomain 
                                  (M_code,M_title,
                                   SubjectUnit,SubjectType,EntryBy,EntryDate)

                             values(
                                    @M_code,@M_title,@subjectunit,
                                    @subjectType,@EntryBy,@entryDate)

			end 




            if @mode='U' 
                  begin
 
					 if  exists (select M_code from SubjectInfomain where M_code = @M_code  )

                         UPDATE  subjectinfomain 
                                  SET M_title=@M_title,
                                   SubjectUnit=@subjectunit,
                                  SubjectType=@subjectType
                        where M_code = @M_code 

                            
			end 



          

            if @mode='D' 
                  begin
 
					 if  exists (select M_code from SubjectInfomain where M_code = @M_code  )

                         DELETE FROM  subjectinfomain 
                               where M_code = @M_code  

                            
			end 
	
                           
	 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE    procedure Subjectmarksdistribution1
(
	@mode       varchar(2),
    @ClassID	varchar(5), 
	@SubjectID	varchar(5),
	@term_code  varchar(5),
    @Exam_code  varchar(5),
	@CategoryID	varchar(5),
	@passmarks   int,
    @Fullmarks   int
)

AS
   if @mode='S' 
      begin
           if  not exists (select * from Subjectmarksdistribution where  ClassID = @ClassID and CategoryID	= @CategoryID and SubjectID = @SubjectID and term_code=@term_code and  Exam_code=@Exam_code )
   
                insert into Subjectmarksdistribution 

				(
                ClassID,
				SubjectID,	
				term_code,
				Exam_code,
                categoryid,
                passmarks,
				fullmarks
	    		)
				values
				(
                @ClassID, 
				@SubjectID,
			 	@term_code,
                @Exam_code,  
				@CategoryID,
				@passmarks,
                @fullmarks	
				)

         end

 if @mode='U' 
      begin
           if   exists (select * from Subjectmarksdistribution where  ClassID = @ClassID and CategoryID	= @CategoryID and SubjectID = @SubjectID and term_code=@term_code and  Exam_code=@Exam_code )
   
               update Subjectmarksdistribution 
                      set passmarks=@passmarks,
                          fullmarks	=@fullmarks	
               where  ClassID = @ClassID and CategoryID	= @CategoryID and SubjectID = @SubjectID and term_code=@term_code and  Exam_code=@Exam_code 
         end

if @mode='D' 
      begin
           if   exists (select * from Subjectmarksdistribution where  ClassID = @ClassID and CategoryID	= @CategoryID and SubjectID = @SubjectID and term_code=@term_code and  Exam_code=@Exam_code )
   
               delete from Subjectmarksdistribution 
                      
               where  ClassID = @ClassID and CategoryID	= @CategoryID and SubjectID = @SubjectID and term_code=@term_code and  Exam_code=@Exam_code 
         end








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








---exec SupplierInformation '','dfgfd','dsfdsf','mn.,n,mn','2008-12-12','sdfsd'

CREATE       procedure SupplierInformation
(
	
	@SuppID			varchar(10),
	@SuppName	 	varchar(30),
	@SuppAddr		varchar(100),
	@Phone			varchar(30),
	@EntryDate		DateTime,
	@EntryBY		varchar(10)
	
)

as

Declare 

	@MaxSLNo int,
	@SuID char(10)



if @SuppID =''

	begin
	
		



		SELECT @MaxSLNo = isnull(max(cast(substring(SuppID,6,10) as int)),0)FROM 
								SupplierInfo 
		
		if @MaxSLNo=0
			set @MaxSLNo=1
		else
			set @MaxSLNo=@MaxSLNo+1
		
		select @SuID = dbo.PadString(@MaxSLNo,5,'0','L')

				
		set @SuppID ='Supp-'+@SuID 
	end


 if exists (select * from SupplierInfo where SuppID = @SuppID) 		
Update SupplierInfo set

	
	SuppID 				=	@SuppID			,
	SuppName	 		=	@SuppName	 	,
	SuppAddr	 		=	@SuppAddr	 	,
	Phone				=	@Phone			,
	EntryDate			=	@EntryDate		,
	EntryBY				=	@EntryBY		
	
	
	where SuppID =	@SuppID 					
else
 insert into SupplierInfo 

(
	SuppID 		,
	SuppName	,
	SuppAddr	,
	Phone		,
	EntryDate	,
	EntryBY						
		
		
	
)
values
(
	@SuppID		,
	@SuppName	,
	@SuppAddr	,
	@Phone		,
	@EntryDate	,
	@EntryBY		
	
	
)










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO












--- Exec BookDisretInformation 'Stu-00000000001','','00002','00001','N','25 Sep 2005','Lia'
CREATE       procedure Syllabuspreperation1
(
	
	@ClassID			varchar(5),
	@Eyear	 			Int,
	@SubjectId			varchar(5),
	@Syllabusdetail			text,
	@PreparedBy			varchar(50),	
	@EntryBY			varchar(50),
	@EntryDate			DateTime
	
)

as



if exists (select * from Syllabuspreperation where  Classid = @ClassID and subjectid=@subjectid and Eyear=@Eyear) 		


Update Syllabuspreperation set  	


	
	
	Syllabusdetail		=	@Syllabusdetail	,
	PreparedBy		=	@PreparedBy	,	
	EntryDate		=	@EntryDate	,
	EntryBY			=	@EntryBY			


where ClassID= @ClassID and subjectid=@subjectid and Eyear=@Eyear
else


 insert into Syllabuspreperation

(

	ClassID				,
	Eyear	 			,
	SubjectId			,
	Syllabusdetail			,
	PreparedBy			,
	EntryBY				,
	EntryDate			
	
			
		
	
)
values
(
	@ClassID			,
	@Eyear	 			,
	@SubjectId			,
	@Syllabusdetail			,
	@PreparedBy			,
	
	@EntryBY			,
	@EntryDate			
	
)














GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
















--- Exec BookDisretInformation 'Stu-00000000001','','00002','00001','N','25 Sep 2005','Lia'
CREATE           procedure TCInfo
(
	
	@StudentId			varchar(15),
	@Shift	 			Varchar(30),
	@ClassID			varchar(5),
	@SectionID			varchar(5),
	@classRoll			int,
	@TCTypeId			varchar(5),
	@TCDate				datetime,
	@TcNote				varchar(80),
	@EntryBY			varchar(10),
	@EntryDate			DateTime,
	@Approved			varchar(1)
	
)

as



if exists (select * from TcInformation where  StudentId= @StudentId ) 		


Update TcInformation set  	


			
	StudentId	=		@StudentId,
	Shift	 	=		@Shift,
	ClassID		=		@ClassID	,
	SectionID	=		@SectionID,
	classRoll	=		@classRoll,
	TCTypeId	=		@TCTypeId,
	TCDate		=		@TCDate,
	TcNote		=		@TcNote,
	EntryBY		=		@EntryBY,
	EntryDate	=		@EntryDate,
	Approved	=		@Approved
	
where StudentId= @StudentId
else


 insert into TcInformation

(

	StudentId	,
	Shift	 	,
	ClassID		,
	SectionID	,
	classRoll	,
	TCTypeId	,
	TCDate		,
	TcNote		,
	EntryBY		,
	EntryDate	,
	Approved	
	
	
	
			
		
	
)
values
(
	@StudentId	,
	@Shift	 	,
	@ClassID		,
	@SectionID	,
	@classRoll	,
	@TCTypeId	,
	@TCDate		,
	@TcNote		,
	@EntryBY		,
	@EntryDate	,
	@Approved	
	
	
)


















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO












--- Exec BookDisretInformation 'Stu-00000000001','','00002','00001','N','25 Sep 2005','Lia'
CREATE          procedure TCInfoApprove
(
	
	@StudentId			varchar(15),
	@TcNote				varchar(80),
	@Approved			varchar(1),
	@ApprovedBy			varchar(10),
	@Approveddate			datetime,
	@ActiveStu 			varchar(1)	
)

as



if exists (select * from TcInformation where  StudentId= @StudentId ) 		


Update TcInformation set  	

		
	Approved	=		@Approved,
	TcNote		=		@TcNote,
	ApprovedBy	=		@ApprovedBy,
	Approveddate	=		@Approveddate
	
	
where StudentId= @StudentId

Update StudentEvaluation Set
Active=@ActiveStu
Where StudentId=@StudentID
  if @Approved='Y' 
     begin 	
	update studentAdmission
	 set active_std=0
	where studentID=@studentID
     end
 else
    begin
      update studentAdmission
	 set active_std=1
	where studentID=@studentID
    
   end



















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO














--- Exec BookDisretInformation 'Stu-00000000001','','00002','00001','N','25 Sep 2005','Lia'
CREATE         procedure TCType
(
	
	@TCID				varchar(5),
	@TCName	 			Varchar(50),
	@Note				varchar(80),
	@EntryBY			varchar(10),
	@EntryDate			DateTime
	
)

as



if exists (select * from TCTypeSetUp where  TCID= @TCID ) 		


Update TCTypeSetUp set  	


			
	TCName	 		= 	@TCName,
	Note			=	@Note	,
	EntryBY			=	@EntryBY	,
	EntryDate		=	@EntryDate	


where TCID= @TCID 
else


 insert into TCTypeSetUp

(

	TCID				,
	TCName	 			,
	Note				,
	EntryBY			,
	EntryDate			
	
	
			
		
	
)
values
(
	@TCID				,
	@TCName	 			,
	@Note				,
	@EntryBY			,
	@EntryDate			
			
	
)
















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE procedure Tax_Setup_I_U_D
@Mode varchar(5),
@Type int,
@Desig_Code varchar (10),
@Range_From varchar (10),
@Range_To varchar (10),
@Percentage varchar (10),
@U_Id varchar(10),
@Tx_Code varchar(10)
AS
/*
default for all 			1	single record
default for designation  		2	multiple record
default for pay range 			3	multiple record
determine during perparation 		0	single record
*/
IF @Mode='I'
BEGIN
	delete from tax_setup where Type!=@Type  
	insert into tax_setup 
		(Type,Desig_Code,Range_From,Range_To,Percentage,U_Id)
	values(@Type,@Desig_Code,convert(money,@Range_From),convert(money,@Range_To)
		,convert(float,@Percentage),@U_Id)
END
IF @Mode='U'
BEGIN	
	Update Tax_Setup SET Type=@Type,Desig_Code=@Desig_Code,Range_From=convert(money,@Range_From)
	,Range_To=convert(money,@Range_To),Percentage=convert(float,@Percentage),U_Id=@U_Id 
	where Tx_Code=convert(varchar(10),@Tx_Code)
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






----------------------------------------------------------------------------

CREATE  procedure Tax_Slab_I_U_D

@Mode varchar(5)
,@Fin_Year varchar(9)
,@Slab_No varchar(10)
,@Slab_Amount money
,@Slab_Percent float
,@track_Id int
,@U_id varchar(10)

AS
set nocount on

declare @Message varchar(150)

If @Mode='I'

BEGIN
	
	if exists(Select * from Tax_Slab where Slab_No=@Slab_No and Fin_Year=@Fin_Year)
	BEGIN
		set @Message ='Data already exists !'
	end

	else
	begin
		insert into Tax_Slab(Fin_Year,Slab_No,Slab_Amount,Slab_Percent,U_id)
		values(@Fin_Year,@Slab_No,@Slab_Amount,@Slab_Percent,@U_id) 

		set @Message ='Data saved successfully !'

	end

END

If @Mode='U'

BEGIN
	update Tax_Slab 
	set Fin_Year=@Fin_Year,Slab_No=@Slab_No,Slab_Amount=@Slab_Amount
		,Slab_Percent=@Slab_Percent,U_id=@U_id
	where track_Id=@track_Id

	set @Message ='Update done successfully !'
	

END


If @Mode='D'

BEGIN
	delete from Tax_Slab where track_Id=@track_Id
	set @Message ='Data deleted successfully !'

END

Select Message=@Message

set nocount off


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE  procedure Taxable_Ceiling_I_U_D

@Mode varchar(5)
,@Head_Code varchar(10)
,@Ceiling_Amount money
,@track_Id int
,@U_id varchar(10)

AS

set nocount on

declare @Message varchar(150)

If @Mode='I'

BEGIN
	
	if exists(Select * from Taxable_Ceiling where Head_Code=@Head_Code)
	BEGIN
		set @Message ='Data already exists !'
	end

	else
	begin
		insert into Taxable_Ceiling(Head_Code,Ceiling_Amount,U_id)
		values(@Head_Code,@Ceiling_Amount,@U_id) 
		set @Message ='Data saved successfully !'

	end

END

If @Mode='U'

BEGIN
	update Taxable_Ceiling 
	set Head_Code=@Head_Code,Ceiling_Amount=@Ceiling_Amount,U_id=@U_id
	where track_Id=@track_Id
	set @Message ='Update done successfully !'
	

END


If @Mode='D'

BEGIN
	delete from Taxable_Ceiling where track_Id=@track_Id
	set @Message ='Data deleted successfully !'

END



select Message=@Message

set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




create  Procedure Tbl

AS
select Tables=Name from sysobjects where type='u'
and  name not like '%odit' and name not like '%Audit' 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Tier_setup_I_U_D    Script Date: 3/7/02 10:43:41 AM ******/
/****** Object:  Stored Procedure dbo.Tier_setup_I_U_D    Script Date: 10/28/01 4:34:35 PM ******/
/****** Object:  Stored Procedure dbo.Tier_setup_I_U_D    Script Date: 10/22/01 6:24:34 PM ******/
/****** Object:  Stored Procedure dbo.Tier_setup_I_U_D    Script Date: 9/17/00 4:34:07 PM ******/
/****** Object:  Stored Procedure dbo.Tier_setup_I_U_D    Script Date: 9/4/01 6:30:54 PM ******/
CREATE procedure Tier_setup_I_U_D		
@opr varchar(1),
@Code varchar (2),
@Policy_Code varchar (10),
@Tier_1 varchar (1),
@Tier_2 varchar (1),
@Tier_3 varchar(1),
@Tier_4 varchar (1),
@Final_Tier varchar(10),
@U_id varchar(10)
as
if @opr='I'
begin
	insert into Tier_setup (
		Code,Policy_Code,Tier_1,Tier_2,Tier_3,Tier_4,Final_Tier,U_id
	)
	values
	(
		@Code,@Policy_Code,@Tier_1,@Tier_2,@Tier_3,@Tier_4,@Final_Tier,@U_id
	)
end
---------------------------------------
if @opr='U'
begin
	Update Tier_setup set
		Code=@Code,Policy_Code=@Policy_Code,Tier_1=@Tier_1,Tier_2=@Tier_2,Tier_3=@Tier_3,
		Tier_4=@Tier_4,Final_Tier=@Final_Tier,U_id=@U_id
	Where Policy_Code=@Policy_Code
end
---------------------------------------
if @opr='D'
begin
	Delete from Tier_setup
	where  Policy_Code=@Policy_Code
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Time_Keeper_I_U    Script Date: 3/7/02 10:43:41 AM ******/
/****** Object:  Stored Procedure dbo.Time_Keeper_I_U    Script Date: 10/28/01 4:34:35 PM ******/
/****** Object:  Stored Procedure dbo.Time_Keeper_I_U    Script Date: 10/22/01 6:24:35 PM ******/
/****** Object:  Stored Procedure dbo.Time_Keeper_I_U    Script Date: 9/17/00 4:34:08 PM ******/
CREATE procedure Time_Keeper_I_U
@Mode varchar(11),
@latest_time varchar(11),
@U_id varchar(10)
as
if @Mode='Boot'
begin
	if (select status=count(a.Boot_Up) from time_keeper a  where 
		datepart(day,a.Dt)=datepart(day,getdate())and
		datepart(month,a.Dt)=datepart(month,getdate())and
		datepart(year,a.Dt)=datepart(year,getdate()))=0
	begin
		insert into time_keeper(Boot_Up,U_id,Dt )values(@latest_time,'Boot_Up',getdate())
	end
	else
	begin
		---delete from time_keeper
		update time_keeper set
		 Boot_Up=@latest_time,Dt=getdate()
	end
return
end
if @Mode='App_Start'
begin
	if (select status=count(*) from time_keeper a where 
		datepart(day,a.Dt)=datepart(day,getdate())and
		datepart(month,a.Dt)=datepart(month,getdate())and
		datepart(year,a.Dt)=datepart(year,getdate()))=0
	begin
		insert into time_keeper (latest_time,U_id)
		values (@latest_time,'Start_Up')
	end
	else
	begin
		update time_keeper set
		 latest_time=@latest_time,U_id=@U_id
	end
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Uptodate_Job_detail    Script Date: 3/7/02 10:43:51 AM ******/
/****** Object:  Stored Procedure dbo.Uptodate_Job_detail    Script Date: 10/28/01 4:34:48 PM ******/
/****** Object:  Stored Procedure dbo.Uptodate_Job_detail    Script Date: 10/22/01 6:24:46 PM ******/
/****** Object:  Stored Procedure dbo.Uptodate_Job_detail    Script Date: 9/17/00 4:34:19 PM ******/
/****** Object:  Stored Procedure dbo.Uptodate_Job_detail    Script Date: 9/4/01 6:31:01 PM ******/
create procedure Uptodate_Job_detail
@Mode varchar(5),
@emp_id varchar(10),
@Effect_date datetime
as
if @Mode='Ind'
begin
declare @var1 as varchar(20)
declare @var2 as varchar(20)
declare @var3 as varchar(20)
declare @var4 as varchar(20)
declare @var5 as varchar(20)
declare @var6 as varchar(20)
declare @var7 as varchar(20)
declare @Effective_date as datetime
declare @Emp_Rank as varchar(20)
declare @Emp_desig as varchar(20)
declare @Emp_branch as varchar(20)
declare @Emp_dept as varchar(20)
declare @Emp_section as varchar(20)
declare @Emp_job_type as varchar(20)
set @Effective_date =""
set @Emp_Rank = ""
set @Emp_desig = ""
set @Emp_branch = ""
set @Emp_dept = ""
set @Emp_section = ""
set @Emp_job_type = ""
declare @c as char(1)
DECLARE abc CURSOR FOR
select Effect_date,Emp_Rank,Emp_desig,Emp_branch,Emp_dept,Emp_section,Emp_job_type
from emp_job_hist_future where emp_id=@Emp_id and Effect_date <=@Effect_date
OPEN abc
set @c=1
WHILE @c = 1
begin
    FETCH NEXT FROM abc into @var1,@var2,@var3,@var4,@var5,@var6,@var7
    if @@FETCH_STATUS = 0 
	begin
		if @var1 <> ""
			set @Effective_date=@var1 
		if @var2 <> ""
			set @Emp_Rank=@var2
		if @var3 <> ""
			set @Emp_desig=@var3
		if @var4 <> ""
			set @Emp_branch=@var4
		if @var5 <> ""
			set @Emp_dept=@var5
		if @var6 <> ""
			set @Emp_section=@var6
		if @var7 <> ""
			set @Emp_job_type=@var7
	end
    else set @c = 0
	
end
select Effect_date=@Effect_date,Emp_Rank=@Emp_Rank,
Emp_desig=@Emp_desig,Emp_branch=@Emp_branch,Emp_dept=@Emp_dept,
Emp_section=@Emp_section,Emp_job_type=@Emp_job_type
CLOSE abc
DEALLOCATE abc
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



create Procedure Validate_PW 
@User varchar (45)
,@Given_PW varchar (40)
AS
set nocount on

Declare @Message varchar(50)
,@PW varchar(40)
,@Can varchar(1)


set @Can=(select Cancel from Soft_Pass where u_id=@User)

if @Can='' or @Can is null
set @Message='Invalid user or password!'

if @Can='1'
set @Message='Your entry has been restricted!'

if @Can='0'
begin
	set @PW=(select user_pass from Soft_Pass where u_id=@User)

	if @PW =@Given_PW	
		set @Message='Valid'
	else
		set @Message='Invalid password, please try typing it again!'

end

select @Message

set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE     procedure VaxinInformation
(
	
	@StudentID		varchar(15),
	@VaxinName	 	varchar(30),
	@VaxinDate		datetime
	--@NextVaxinDate		datetime
	
	
	
)

AS if exists (select * from VaxinInfo where  StudentID= @StudentID and VaxinName=@VaxinName) 		
Update VaxinInfo set

	
	StudentID			=	@StudentID		,
	VaxinName	 		=	@VaxinName	 	,
	VaxinDate			=	@VaxinDate		
	--NextVaxinDate			=	@NextVaxinDate
	
	where StudentID= @StudentID and VaxinName=@VaxinName	
else
 insert into VaxinInfo 

(
	StudentID			,
	VaxinName			,
	VaxinDate			
	--NextVaxinDate
		
		
	
)
values
(
	@StudentID			,
	@VaxinName	 		,
	@VaxinDate			
	--@NextVaxinDate
        		
)







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






CREATE       procedure Vou_I_Payroll_Main_U

@pay_month varchar(12),
@pay_year varchar(4),
@Emp_dept varchar(35),
@Operation_Mode varchar(10),
@A_Gen Varchar(10),
--------------------------------------------------
@Pay_Type varchar(2),
@vou_no [varchar] (10),
@vou_date [datetime],
@vou_slno [varchar] (4),
@cost_code [varchar] (10),
@cust_ord_no [varchar] (25),
@vou_narr [varchar] (200),
--@vou_desc[varchar] (200), 
@Vou_Dr [varchar] (10),
@Vou_Cr [varchar] (10),
--@acc_code [varchar] (10),
--@dollar money default (0), 
--@rate money default (0), 
--@dr_amt money default (0), 
--@cr_amt money default (0),	
@Vou_Amount money, 
@vou_type [char] (2),
@uid [varchar] (10)
AS

set nocount on

declare @Message varchar(200)
,@Acct_Send_Flag int
,@Num_Processed int
,@Num_Disbursed int
set @Acct_Send_Flag=(select dbo.GetParamFlag(77))

---------------------------------------------------------------------
If @Operation_Mode='Ind' 
begin
	if (select count(*) from Payroll_main Where A_Gen=@A_Gen)=0
		set @Message='Data not found!'	
	else

	begin
		Update Payroll_main set  Pay_Stat='1' Where A_Gen=@A_Gen
		set @Message='Disbursement done successfully !'
	end

end


If @Operation_Mode='Dept' 
	begin
		set @Num_Processed=(select count(*) from Payroll_main where Pay_Type=@Pay_Type
			and pay_month=@pay_month and pay_year=@pay_year and Pay_Stat=0
			and Emp_Id in (Select emp_id from Emp_Job_Hist_Current where Emp_dept=@Emp_dept))


		set @Num_Disbursed=(select count(*) from Payroll_main where Pay_Type=@Pay_Type
			and pay_month=@pay_month and pay_year=@pay_year and Pay_Stat=1
			and Emp_Id in (Select emp_id from Emp_Job_Hist_Current where Emp_dept=@Emp_dept))

		if @Num_Processed!=0

		begin
			Update Payroll_main set  Pay_Stat='1' Where Pay_Type=@Pay_Type
			and pay_month=@pay_month and pay_year=@pay_year
			and Emp_Id in (Select emp_id from Emp_Job_Hist_Current where Emp_dept=@Emp_dept)

			set @Vou_Amount=(select dbo.GetVouAmount('Dept',@Pay_Type,@Pay_Month,@Pay_Year,@Emp_dept))
			set @Message='Disbursement done successfully to '+ convert(varchar(6),@Num_Processed)+' employee(s) of the '+ @Emp_dept + ' department'
		end
			
		else if @Num_Disbursed=@Num_Processed

			set @Message='Already disbursed to "'+ @Emp_Dept+ '" department'
		
		else if @Num_Processed=0
			set @Message='Data have not yet been processed for " '+ @Emp_dept + ' " department'
	end

If @Operation_Mode='All' 

	begin

  
		set @Num_Processed=(select count(*)from Payroll_main Where Pay_Type=@Pay_Type
			and pay_month=@pay_month and pay_year=@pay_year
			and Pay_Stat='0')

		set @Num_Disbursed=(select count(*)from Payroll_main Where Pay_Type=@Pay_Type
				and pay_month=@pay_month and pay_year=@pay_year
					and Pay_Stat='1')
		

		if @Num_Processed!=0

		begin
			Update Payroll_main set  Pay_Stat='1' Where Pay_Type=@Pay_Type
				and pay_month=@pay_month and pay_year=@pay_year

			set @Vou_Amount=(select dbo.GetVouAmount('All',@Pay_Type,@Pay_Month,@Pay_Year,''))
			set @Message='Disbursement done successfully to '+ convert(varchar(6),@Num_Processed)+ ' employee(s) !'
		end

		if @Num_Processed=0
			set @Message='Data have not yet been processed for all employees !'	

		if @Num_Disbursed!=0

			set @Message='Already disbursed !'	


	End

--------------------------Debit Part---------------------------------
--Acc_Code=Vou_Dr



if @Acct_Send_Flag=1

Begin

	insert into Time_Keeper..vou 
		(vou_no,vou_date,vou_slno,cost_code ,cust_ord_no
		,acc_code,dr_amt,vou_type ,uid)

	values(@vou_no,@vou_date,@vou_slno,@cost_code ,@cust_ord_no
	,@Vou_Dr,@Vou_Amount, @vou_type, @uid)

--------------------------Credit Part---------------------------------
--Acc_Code=Vou_Cr

	insert into Time_Keeper..vou 
		(vou_no,vou_date,vou_slno,cost_code ,cust_ord_no
		,acc_code,cr_amt,vou_type,uid)
	
	values(@vou_no,@vou_date,@vou_slno,@cost_code ,@cust_ord_no  
		,@Vou_Cr,@Vou_Amount, @vou_type, @uid)

----------------------------------------------------------------------
---Export to Accounting part
----------------------------------------------------------------------

	insert into Ayman_acct..vou 
		(vou_no,vou_date,vou_slno,cost_code ,cust_ord_no
		,acc_code,dr_amt,vou_type ,uid)
	
	values(@vou_no,@vou_date,@vou_slno,@cost_code ,@cust_ord_no 
		,@Vou_Dr,@Vou_Amount, @vou_type,  @uid)

--------------------------Credit Part---------------------------------
--Acc_Code=Vou_Cr

	insert into Ayman_acct..vou 
		(vou_no,vou_date,vou_slno,cost_code ,cust_ord_no
		,acc_code,cr_amt,vou_type ,uid)
	
	values(@vou_no,@vou_date,@vou_slno,@cost_code ,@cust_ord_no 
		,@Vou_Cr,@Vou_Amount, @vou_type, @uid)


END


Select Message=@Message

set nocount off









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.Worker_Group_I_U_D    Script Date: 3/7/02 10:43:42 AM ******/
/****** Object:  Stored Procedure dbo.Worker_Group_I_U_D    Script Date: 10/28/01 4:34:35 PM ******/
/****** Object:  Stored Procedure dbo.Worker_Group_I_U_D    Script Date: 10/22/01 6:24:35 PM ******/
/****** Object:  Stored Procedure dbo.Worker_Group_I_U_D    Script Date: 9/17/00 4:34:08 PM ******/
---------------------------------------
CREATE procedure Worker_Group_I_U_D
@opr varchar (1),
@Gr_Code varchar(5),
@Emp_ID varchar(10),
@Stored_Emp_ID varchar(10),
@U_id    varchar(10)
as
if @opr='I'
begin
	insert into Worker_Group(Gr_Code,Emp_ID,U_id)
	values(@Gr_Code,@Emp_ID,@U_id)
end 
if @opr='U'
begin
	Update Worker_Group Set Emp_ID=@Emp_ID,U_id=@U_id
	Where Emp_ID=@Stored_Emp_ID
end 
if @opr='D'
begin
	Delete from Worker_Group
	Where Emp_ID=@Stored_Emp_ID
end 




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO





/*

Yearly_Holidays '2003','Friday'
*/

CREATE  PROCEDURE Yearly_Holidays
@Yr varchar(4)
,@Day varchar(9)

AS

set nocount on 

DECLARE  
@Date varchar (10)
,@Name varchar (15)
,@Chk_Dt DATETIME
,@Days int
,@Count int
,@Td int
,@Wd int
,@Wc int
,@Hol_Type varchar(25)



----set @Yr='2003'
----SET @Day ='FRIDAY'
SET @Count=1
set @Wc=1
set @Hol_Type='Government Holiday'

create table #Holidays
(Hol_name varchar(30)
,Hol_desc varchar(50)null
,Str_date datetime
,End_date datetime
,Category varchar(30)
,U_id varchar(10)null
,Dt datetime null
)



-----------------------------------------------------------------------
set @Date =@yr+'-02-21'
set @Chk_Dt = convert(datetime,@Date)
insert into #Holidays values(right(@Yr,2)+' Int. Mother Language Day',' ',@Chk_Dt,@Chk_Dt,@Hol_Type,'dsl',getdate())
-----------
set @Date =@yr+'-03-26'
set @Chk_Dt = convert(datetime,@Date)
insert into #Holidays values(right(@Yr,2)+' Independance Day',' ',@Chk_Dt,@Chk_Dt,@Hol_Type,'dsl',getdate())
-----------
set @Date =@yr+'-04-14'
set @Chk_Dt = convert(datetime,@Date)
insert into #Holidays values(right(@Yr,2)+' Poila Baishakh',' ',@Chk_Dt,@Chk_Dt,@Hol_Type,'dsl',getdate())
-----------
set @Date =@yr+'-05-01'
set @Chk_Dt = convert(datetime,@Date)
insert into #Holidays values(right(@Yr,2)+' May Day',' ',@Chk_Dt,@Chk_Dt,@Hol_Type,'dsl',getdate())
-----------
set @Date =@yr+'-11-07'
set @Chk_Dt = convert(datetime,@Date)
insert into #Holidays values(right(@Yr,2)+' Solidarity Day',' ',@Chk_Dt,@Chk_Dt,@Hol_Type,'dsl',getdate())
-----------
set @Date =@yr+'-12-16'
set @Chk_Dt = convert(datetime,@Date)
insert into #Holidays values(right(@Yr,2)+' Victory Day',' ',@Chk_Dt,@Chk_Dt,@Hol_Type,'dsl',getdate())
-----------
set @Date =@yr+'-12-25'
set @Chk_Dt = convert(datetime,@Date)
insert into #Holidays values(right(@Yr,2)+' Christmas Day',' ',@Chk_Dt,@Chk_Dt,@Hol_Type,'dsl',getdate())
--------------------------------------------------------------------------


-----------------------------------------------------

SELECT @Wd=CASE @Day 
	WHEN 'SUNDAY' THEN 1
	WHEN 'MONDAY' THEN 2
	WHEN 'TUESDAY' THEN 3
	WHEN 'WEDNESDAY' THEN 4
	WHEN 'THURSDAY' THEN 5
	WHEN 'FRIDAY' THEN 6
	WHEN 'SATURDAY' THEN 7     	
END
----------------------------------------Number of days 
if @Yr % 4 = 0 
	Begin
		if @Yr % 100 <> 0 or @Yr % 400 = 0 
			set @Days=366
		else
			set @Days=365
	End
else			
	set @Days=365
-----------------------------------------------------
set @Date =@yr+'-01-01'
set @Chk_Dt = convert(datetime,@Date)

while @Count<=@Days
BEGIN 		
	SET @Chk_Dt=DATEADD(DAY,1,@Chk_Dt)	
	SET @Td=DATEPART(dw,@Chk_Dt)

	if @Wd=@Td 
	begin	
		set  @Name=right(@Yr,2)+' '+@Day+convert(varchar(3),@Wc)	
		
		if not exists (select * from  #Holidays where [Str_Date]=@Chk_Dt)
		insert into #Holidays values(@Name,'',@Chk_Dt,@Chk_Dt,'Weekend','dsl',getdate())

		set @Wc	=@Wc+1
	end	

	set @Count=@Count+1
END

insert into  Hol_List
select * from #Holidays where Hol_Name not in (select Hol_Name from Hol_List)

drop table #Holidays



set nocount off

---delete from holiday

select * from Hol_List




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.add_scrn    Script Date: 3/7/02 10:43:23 AM ******/
/****** Object:  Stored Procedure dbo.add_scrn    Script Date: 10/28/01 4:34:20 PM ******/
/****** Object:  Stored Procedure dbo.add_scrn    Script Date: 10/22/01 6:24:21 PM ******/
/****** Object:  Stored Procedure dbo.add_scrn    Script Date: 9/17/00 4:33:55 PM ******/
/****** Object:  Stored Procedure dbo.add_scrn    Script Date: 9/4/01 6:30:39 PM ******/
/****** Object:  Stored Procedure dbo.add_scrn    Script Date: 6/20/01 12:00:32 PM ******/
/****** Object:  Stored Procedure dbo.add_scrn    Script Date: 5/14/01 9:49:53 AM ******/
/****** Object:  Stored Procedure dbo.add_scrn    Script Date: 1/1/99 3:01:38 AM ******/
/****** Object:  Stored Procedure dbo.add_scrn    Script Date: 5/9/01 8:13:06 AM ******/
/****** Object:  Stored Procedure dbo.add_scrn    Script Date: 5/8/01 4:09:44 AM ******/
CREATE PROCEDURE add_scrn
@software as varchar(50),
@scr_no as varchar(50),
@descript as text
 AS
if (select count(scr_no) from soft_bag where scr_no=@scr_no) = 0
begin
insert into soft_bag(software,scr_no,descript)values(@software,@scr_no,@descript)
end
else
begin
update soft_bag set software=@software,scr_no=@scr_no,descript=@descript where software=@software
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




Create Procedure emp_bank_account_I_U

@Emp_Id varchar(10)
,@Bank_Code int  
,@Acc_No varchar(30)                        
,@U_id varchar(10)      

AS

if exists (select * from emp_bank_account 
where Emp_Id=@Emp_Id)
begin 
	update emp_bank_account 
	set Bank_Code=@Bank_Code,Acc_No=@Acc_No,U_id=@U_id 
	where Emp_Id=@Emp_Id
end
else
begin
	insert into emp_bank_account(Emp_Id,Bank_Code,Acc_No,U_id)
	Values(@Emp_Id,@Bank_Code,@Acc_No,@U_id)
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




----exec generate_position 'a','00001','00002','M','01','01','2006'
CREATE      procedure generate_position
(
   @mode               varchar(1),
   @ClassID            varchar(12),
   @sectionID          varchar(12),
   @shift              varchar(12),
   @ExamType           varchar(6),
   @ExamID             varchar(6),
   @AcaYr              varchar(20)
)
as
set nocount on
create table #temp_table(
   stdid              varchar(15),
   ClassID            varchar(12),
   sectionID          varchar(12),
   shift              varchar(12),
   AcaYr              varchar(10),    
   ExamType           varchar(6),
   ExamID             varchar(6),
   std_marks  decimal)
                       
begin
        declare @m_serial_no as integer
        declare @s_serial_no as integer 
        declare @f_term_wd as integer  
        declare @f_term_p as integer  
        declare @s_term_wd as integer  
        declare @s_term_p as integer  
        declare @wd as integer  
        declare @p as integer  
	
	
	
	declare @subject  as varchar(50)
        declare @class_test_marks decimal
        declare @term_marks       decimal
        declare @M_srl_no as int
        declare @s_srt_no as int
        declare @total_marks as decimal
        declare @StdID as varchar(15)     
        declare @i  as integer
        declare @std_marks as decimal
        declare @fst_term_marks as decimal
        declare @snd_term_marks as decimal
        declare @fst_sec_trm_marks as decimal
if @mode='a' 
begin
   set @i=0
   declare mycursor cursor for
	   select distinct(a.StdID)
	       from  result_main b, result_sub a 
	      where a.M_Slr_no=b.M_Slr_no  and
	          b.ClassID=@classid
	        and b.SectionID=@sectionID 
                and b.shift=@shift and
	         b.AcaYr= @AcaYr and           
	          b.ExamType=@ExamType and        
	          b.ExamID=@ExamID ---- group by a.M_Slr_no,a.S_Slr_no
                  

 open mycursor
 fetch next from mycursor into  @stdid
     
 while @@fetch_status=0 
     begin

        set @f_term_wd =0 
        set @f_term_p =0
        set @s_term_wd =0 
        set @s_term_p =0 
        set @wd=0 
        set @p =0  
	set @class_test_marks=0
        set @term_marks =0
 

   if @ExamType='01'  and @examid='01' ---first term and pre term
       begin
          if exists(select * from position where  ClassID=@classid
	        and SectionID=@sectionID 
                and shift=@shift 
	         and AcaYr=@AcaYr           
	          and ExamType = @ExamType       
	          and ExamID= @ExamId)
           begin
                 delete from  position  where  ClassID=@classid
	       and  SectionID= @sectionID 
                and shift=@shift and
	         AcaYr= @AcaYr and           
	          ExamType =@ExamType and        
	          ExamID= @ExamId
          end     

           set @term_marks=(select  isnull(sum(a.obtainedMarks),0)
	         from   result_main b, result_sub a 
	      where  a.M_Slr_no=b.M_Slr_no 
	          and
	          b.ClassID=@ClassID 
	        and b.SectionID=@SectionID
	          and 
	          b.AcaYr= @AcaYr and           
	          b.ExamType ='01' and  ----@ExamType 
                  b.shift=@shift 
	          and a.StdID=@StdID and        
	          b.ExamID= '01' )
         


         insert into  #temp_table values(@StdID,@ClassID,
                     @sectionID,@shift,@AcaYr,@ExamType,@ExamID,   
                       @term_marks)
       fetch next from mycursor into   @stdid

       
    end  ---end of exam type='01' and exam_id='01'

  if @ExamType='01'  and @examid='02'  ---first term and term final
        begin

             if exists(select * from position where  ClassID=@classid
	        and SectionID=@sectionID 
                 and shift=@shift 
	         and AcaYr=@AcaYr           
	          and ExamType = @ExamType       
	          and ExamID= @ExamId)
           begin
                 delete from  position  where  ClassID=@classid
	        and SectionID= @sectionID 
                 and shift=@shift and
	         AcaYr= @AcaYr and           
	          ExamType =@ExamType and        
	          ExamID= @ExamId
          end     

           set @term_marks=(select  isnull(sum(a.obtainedMarks),0)
	         from   result_main b, result_sub a 
	      where  a.M_Slr_no=b.M_Slr_no 
	          and
	          b.ClassID=@ClassID 
	        and b.SectionID=@SectionID
	          and 
	          b.AcaYr= @AcaYr and           
	          b.ExamType ='01' and  ----@ExamType 
                  b.shift=@shift 
	          and a.StdID=@StdID and        
	          b.ExamID= '02' )



              select  @class_test_marks=isnull(sum(a.obtainedMarks),0)
		         from   result_main b, result_sub a 
		      where    a.M_Slr_no=b.M_Slr_no
		           and
		          b.ClassID=@ClassID 
		        and b.SectionID=@SectionID
		          and 
		          b.AcaYr= @AcaYr and           
		          b.ExamType ='01'   ----@ExamType 
		          and a.StdID=@StdID and        
		          b.ExamID NOT IN('01','02') 

             set @class_test_marks=(@class_test_marks*5)/100 ---class test marks
            
           set   @term_marks=@term_marks+@class_test_marks


     insert into  #temp_table values(@StdID,@ClassID,
                     @sectionID,@shift,@AcaYr,@ExamType,@ExamID,   
                       @term_marks)
       fetch next from mycursor into   @stdid

       
   end ----end of exam type='01' and exam_id='02'


      
  
if @ExamType='02'  and @examid='01' ---second term and pre term
       begin
          if exists(select * from position where  ClassID=@classid
	        and SectionID=@sectionID 
                and shift=@shift 
	         and AcaYr=@AcaYr           
	          and ExamType = @ExamType       
	          and ExamID= @ExamId)
           begin
                 delete from  position  where  ClassID=@classid
	       and  SectionID= @sectionID 
                and shift=@shift and
	         AcaYr= @AcaYr and           
	          ExamType =@ExamType and        
	          ExamID= @ExamId
          end     

           set @term_marks=(select  isnull(sum(a.obtainedMarks),0)
	         from   result_main b, result_sub a 
	      where  a.M_Slr_no=b.M_Slr_no 
	          and
	          b.ClassID=@ClassID 
	        and b.SectionID=@SectionID
	          and 
	          b.AcaYr= @AcaYr and           
	          b.ExamType ='02' and  ----@ExamType 
                  b.shift=@shift 
	          and a.StdID=@StdID and        
	          b.ExamID= '01' )
         


         insert into  #temp_table values(@StdID,@ClassID,
                     @sectionID,@shift,@AcaYr,@ExamType,@ExamID,   
                       @term_marks)
       fetch next from mycursor into   @stdid

       
      end  ---end of exam type='02' and exam_id='01'

    if @ExamType='02'  and @examid='02'  ---second term and term final
        begin

             if exists(select * from position where  ClassID=@classid
	        and SectionID=@sectionID 
                 and shift=@shift 
	         and AcaYr=@AcaYr           
	          and ExamType = @ExamType       
	          and ExamID= @ExamId)
           begin
                 delete from  position  where  ClassID=@classid
	        and SectionID= @sectionID 
                 and shift=@shift and
	         AcaYr= @AcaYr and           
	          ExamType =@ExamType and        
	          ExamID= @ExamId
          end     

           set @term_marks=(select  isnull(sum(a.obtainedMarks),0)
	         from   result_main b, result_sub a 
	      where  a.M_Slr_no=b.M_Slr_no 
	          and
	          b.ClassID=@ClassID 
	        and b.SectionID=@SectionID
	          and 
	          b.AcaYr= @AcaYr and           
	          b.ExamType ='02' and  ----@ExamType 
                  b.shift=@shift 
	          and a.StdID=@StdID and        
	          b.ExamID= '02' )



              select  @class_test_marks=isnull(sum(a.obtainedMarks),0)
		         from   result_main b, result_sub a 
		      where    a.M_Slr_no=b.M_Slr_no
		           and
		          b.ClassID=@ClassID 
		        and b.SectionID=@SectionID
		          and 
		          b.AcaYr= @AcaYr and           
		          b.ExamType ='02'   ----@ExamType 
		          and a.StdID=@StdID and        
		          b.ExamID NOT IN('01','02') 

             set @class_test_marks=(@class_test_marks*5)/100 ---class test marks
            
           set   @term_marks=@term_marks+@class_test_marks


    insert into  #temp_table values(@StdID,@ClassID,
                     @sectionID,@shift,@AcaYr,@ExamType,@ExamID,   
                       @term_marks)
       fetch next from mycursor into   @stdid

       
      end   ---end of exam type='01' and exam_id='02'

 if @ExamType='03'  and @examid='01'  ---Final
        begin
             if exists(select * from position where  ClassID=@classid
	        and SectionID=@sectionID 
                 and shift=@shift 
	         and AcaYr=@AcaYr           
	          and ExamType=@ExamType       
	          and ExamID=@ExamId)
           begin
                delete from  position  where  ClassID=@classid
	        and SectionID= @sectionID 
                 and shift=@shift and
	         AcaYr= @AcaYr and           
	          ExamType =@ExamType and        
	          ExamID= @ExamId
          end     

          set @term_marks=(select  isnull(sum(a.obtainedMarks),0)
	         from   result_main b, result_sub a 
	      where  a.M_Slr_no=b.M_Slr_no 
	          and
	          b.ClassID=@ClassID 
	        and b.SectionID=@SectionID
	          and 
	          b.AcaYr= @AcaYr and           
	          b.ExamType ='03' and  ----@ExamType 
                  b.shift=@shift 
	          and a.StdID=@StdID and        
	          b.ExamID= '01' )

           select @wd=isnull(count(distinct(attn_date)),0)
	          from StudentAttendanceLeaveInfo 
	  where aca_yr=@AcaYr  and ClassID=@ClassID
            


          select @p=isnull(count(distinct(attn_date)),0)
	         from StudentAttendanceLeaveInfo 
	  where aca_yr=@AcaYr and ClassID=@ClassID
	        and StudentID=@stdid and Present='P'

          if ((@wd*80)/100)>=@p 
              begin
                set @term_marks= @term_marks+(@term_marks*5)/100
             end

              select  @class_test_marks=isnull(sum(a.obtainedMarks),0)
		         from   result_main b, result_sub a 
		      where    a.M_Slr_no=b.M_Slr_no
		           and
		          b.ClassID=@ClassID 
		        and b.SectionID=@SectionID
		          and 
		          b.AcaYr= @AcaYr and           
		          b.ExamType ='03'   ----@ExamType 
		          and a.StdID=@StdID and        
		          b.ExamID NOT IN('01') 

             set @class_test_marks=(@class_test_marks*5)/100 ---class test marks
            
           set   @term_marks=@term_marks+@class_test_marks

            set @fst_term_marks=(select  isnull(sum(a.obtainedMarks),0)
	         from   result_main b, result_sub a 
	      where  a.M_Slr_no=b.M_Slr_no 
	          and
	          b.ClassID=@ClassID 
	        and b.SectionID=@SectionID
	          and 
	          b.AcaYr= @AcaYr and           
	          b.ExamType ='01' and  ----@ExamType 
                  b.shift=@shift 
	          and a.StdID=@StdID and        
	          b.ExamID= '02' )

           set @snd_term_marks=(select  isnull(sum(a.obtainedMarks),0)
	         from   result_main b, result_sub a 
	      where  a.M_Slr_no=b.M_Slr_no 
	          and
	          b.ClassID=@ClassID 
	        and b.SectionID=@SectionID
	          and 
	          b.AcaYr= @AcaYr and           
	          b.ExamType ='02' and  ----@ExamType 
                  b.shift=@shift 
	          and a.StdID=@StdID and        
	          b.ExamID= '02' )

         set  @fst_sec_trm_marks=((@fst_term_marks+@snd_term_marks)*5)/100

         set   @term_marks=@term_marks+@fst_sec_trm_marks


     insert into  #temp_table values(@StdID,@ClassID,
                     @sectionID,@shift,@AcaYr,@ExamType,@ExamID,   
                       @term_marks)
       fetch next from mycursor into   @stdid

       
      end 

 
      end       
 ---end of exam type='03' 

     declare mycursor1 cursor for
          select * from #temp_table  order by std_marks desc
 
    open mycursor1
    fetch next from mycursor1 into   @stdid,@ClassID,@sectionID,@shift,
                                     @AcaYr,@ExamType,@ExamID,@std_marks
   
   while @@fetch_status=0
       begin  
          set @i=@i+1       
          insert into position values(@stdid,@ClassID,@sectionID,@shift,
                                     @AcaYr,@ExamType,@ExamID,@std_marks,@i)
          fetch next from mycursor1 into   @stdid,@ClassID,@sectionID,@shift,
                                     @AcaYr,@ExamType,@ExamID,@std_marks
  

end

Close MyCursor
deallocate MyCursor

Close MyCursor1
deallocate MyCursor1

  end
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE    procedure getStudentInfo(
	@StudentID	VARCHAR(15))
AS
SET NOCOUNT ON

SELECT     StudentInfo.StudentName, StudentAdmission.StudentId, ClassInfo.ClassName, StudentAdmission.ClassId, SectionInfo.Sectiondsc, 
                      StudentAdmission.SectionId, Shift.Shift_Name, Shift.Shift_Code
FROM         StudentInfo INNER JOIN
                      StudentAdmission ON StudentInfo.StudentID = StudentAdmission.StudentId INNER JOIN
                      ClassInfo ON StudentAdmission.ClassId = ClassInfo.ClassID INNER JOIN
                      SectionInfo ON StudentAdmission.SectionId = SectionInfo.SectionID CROSS JOIN
                      Shift
where StudentAdmission.StudentId = @StudentID

SET NOCOUNT OFF

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.get_max_Eff_date    Script Date: 3/7/02 10:43:44 AM ******/
/****** Object:  Stored Procedure dbo.get_max_Eff_date    Script Date: 10/28/01 4:34:38 PM ******/
/****** Object:  Stored Procedure dbo.get_max_Eff_date    Script Date: 10/22/01 6:24:37 PM ******/
/****** Object:  Stored Procedure dbo.get_max_Eff_date    Script Date: 9/17/00 4:34:10 PM ******/
/****** Object:  Stored Procedure dbo.get_max_Eff_date    Script Date: 9/4/01 6:30:56 PM ******/
create procedure get_max_Eff_date 
@emp_ID varchar(10)
as 
select max_date=max(Effect_date)from emp_job_hist_future
where emp_ID=@emp_id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







CREATE PROCEDURE give_pmt
@Mode varchar(10),
@Portion varchar(5),
@code int,
@u_id varchar(10)

 AS

IF @Mode='Single'

BEGIN

	insert into permit(code,u_id) values (@code,@u_id)

END


IF @Mode='All'

BEGIN

declare @S_Code as int

	IF @Portion='All'
	
	BEGIN

		DECLARE Code_Cursor  CURSOR FOR
			select distinct Code from Soft_Bag

		OPEN Code_Cursor

		FETCH NEXT FROM Code_Cursor into @S_Code

		WHILE @@FETCH_STATUS = 0     
		
		BEGIN
			
			insert into permit(code,u_id) values (@S_Code,@u_id)
			
			FETCH NEXT FROM Code_Cursor into @S_Code
		END
		

		CLOSE Code_Cursor
		DEALLOCATE Code_Cursor
	END

	ELSE

	BEGIN

		DECLARE Code_Cursor_1  CURSOR FOR
			select distinct Code from Soft_Bag where portion=@portion
		OPEN Code_Cursor_1

		FETCH NEXT FROM Code_Cursor_1 into @S_Code

		WHILE @@FETCH_STATUS = 0     
		
		BEGIN
			
			insert into permit(code,u_id) values (@S_Code,@u_id)

			FETCH NEXT FROM Code_Cursor_1 into @S_Code
		END
		

		CLOSE Code_Cursor_1
		DEALLOCATE Code_Cursor_1
	END


END




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE    PROCEDURE [pro_pass_entry] 
AS
insert into Soft_pass (u_id,u_name,user_pass,uid,udt,cancel)
values('DSL','DBA','123123897112345','Default',getdate(),0)

----PW word 'pmis'

insert into permit (code,u_id)values(1,'DSL')
insert into permit (code,u_id)values(2,'DSL')
insert into permit (code,u_id)values(3,'DSL')
insert into permit (code,u_id)values(4,'DSL')
insert into permit (code,u_id)values(5,'DSL')
insert into permit (code,u_id)values(6,'DSL')
insert into permit (code,u_id)values(7,'DSL')
insert into permit (code,u_id)values(8,'DSL')
insert into permit (code,u_id)values(9,'DSL')
insert into permit (code,u_id)values(10,'DSL')
insert into permit (code,u_id)values(11,'DSL')
insert into permit (code,u_id)values(12,'DSL')







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








/* Object:  Stored Procedure dbo.pro_security    Script Date: 3/7/02 10:43:47 AM ******/

CREATE     PROCEDURE [pro_security]
@carry as varchar(50),
@cur_user as varchar(20)
 AS
declare @Result as char(1)
set @Result= (select Case @carry
    
	when 'Form2' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form2') and u_id=@cur_user)
	when 'Form3' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form3') and u_id=@cur_user)
	when 'Form4' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form4') and u_id=@cur_user)
	when 'Form5' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form5') and u_id=@cur_user)
	when 'Form6' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form6') and u_id=@cur_user)
	when 'Form7' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form7') and u_id=@cur_user)
	when 'Form8' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form8') and u_id=@cur_user)
	when 'Form9' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form9') and u_id=@cur_user)END)
	

if @Result = 1 goto stop
set @Result= (select Case @carry

	when 'Form10' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form10') and u_id=@cur_user)
	when 'Form11' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form11') and u_id=@cur_user)		
	when 'Form12' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form12') and u_id=@cur_user)
	when 'Form13' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form13') and u_id=@cur_user)
	when 'Form14' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form14') and u_id=@cur_user)
	when 'Form15' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form15') and u_id=@cur_user)
	when 'Form16' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form16') and u_id=@cur_user)
	when 'Form17' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form17') and u_id=@cur_user)
	when 'Form18' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form18') and u_id=@cur_user)
	when 'Form19' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form19') and u_id=@cur_user)
	when 'Form20' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form20') and u_id=@cur_user)END)

if @Result = 1 goto stop
set @Result= (select Case @carry

	when 'Form21' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form21') and u_id=@cur_user)
	when 'Form22' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form22') and u_id=@cur_user)
	when 'Form25' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form25') and u_id=@cur_user)
	when 'Form27' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form27') and u_id=@cur_user)
	when 'Form28' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form28') and u_id=@cur_user)
	when 'Form29' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form29') and u_id=@cur_user)
	when 'Form30' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form30') and u_id=@cur_user)
	when 'Form31' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form31') and u_id=@cur_user)END)
	
if @Result = 1 goto stop
set @Result= (select Case @carry

	when 'Form32' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form32') and u_id=@cur_user)
	when 'Form33' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form33') and u_id=@cur_user)
	when 'Form34' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form34') and u_id=@cur_user)
	when 'Form35' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form35') and u_id=@cur_user)
	when 'Form36' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form36') and u_id=@cur_user)
	when 'Form37' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form37') and u_id=@cur_user)
	when 'Form39' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form39') and u_id=@cur_user)
	when 'Form40' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form40') and u_id=@cur_user)
	when 'Form41' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form41') and u_id=@cur_user)
	when 'frmAttn_Back' then (select count(code)from permit
		where code=(select code from soft_bag where software ='Daffodil PMIS'and scr_no='frmAttn_Back')and u_id=@cur_user)End)

if @Result =1 goto stop
set @Result= (select Case @carry
    
    	when 'Report1' then (select count(code) from permit
    		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report1') and u_id=@cur_user)
	when 'Report2' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report2') and u_id=@cur_user)
	when 'Report3' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report3') and u_id=@cur_user)
	when 'Report4' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report4') and u_id=@cur_user)
	when 'Report5' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report5') and u_id=@cur_user)
	when 'Report6' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report6') and u_id=@cur_user)
	when 'Report7' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report7') and u_id=@cur_user)
	when 'Report8' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report8') and u_id=@cur_user)
	when 'Report9' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report9') and u_id=@cur_user)
	when 'Report10' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report10') and u_id=@cur_user)
	when 'Report11' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report11') and u_id=@cur_user)
	when 'Report12' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report12') and u_id=@cur_user)
	when 'Report13' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report13') and u_id=@cur_user)
	when 'Report14' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report14') and u_id=@cur_user)
	when 'Report15' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Report15') and u_id=@cur_user)END)
	
if @Result = 1 goto stop
set @Result= (select Case @carry

	when 'Form90' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form90') and u_id=@cur_user)
	when 'Form92' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form92') and u_id=@cur_user)
	when 'Form93' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form93') and u_id=@cur_user)	
	when 'Form94' then (select count(code) from permit
		where code=(select code from soft_bag where software='Daffodil PMIS' and scr_no='Form94') and u_id=@cur_user)END)

Stop:
select Result=@Result


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO






CREATE     PROCEDURE pro_security_entry
AS


insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form90','User Creation','S')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form92','Changing Password','S')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form93','Software Privilege','S')
--insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','','Software Privilege','S')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form21','Database Backup & Restore','S')
-------------------------------------------------------------------------------------------------------------------------------------
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form94','Administrative Setup','T')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form2','Company Details','T')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form3','Salary Fixing','P')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form4','Holiday Setup','T')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form5','Job Ending','E')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form6','Performance Evaluation','E')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form7','Leave Setup','T')
-------------------------------------------------------------------------------------------------------------------------------------
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form9','Employee Attendance','A')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form10','Employee Personal Information','E')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form11','Late Attendance Notes','A')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','frmAttn_Back','Attendance Backup','A')

insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form12','Leave Application','L')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form13','Company Setup','T')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form14','Pay Preparation','P')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form15','Job Detail','E')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form17','Office Time Setup','T')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form18','Promotion,Increment,Transfer ','E')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form30','Overtime Rate Fixing ','P')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form31','Overtime Preparation','P')
------------------------------------------------------------------------------------------------------------------------------------
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Report1','Employee Personal Information','R')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Report2','Duty Plan Reports','R')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Report3','Attendance Reports','R')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Report4','Leave Reports','R')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Report5','Movement Reports','R')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Report6','Loan/PF Reports','R')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Report7','Payroll Reports','R')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Report8','Overtime Reports','R')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Report9','Confidential Reports','R')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Report10','Holiday List','R')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Report11','Miscellaneous Reports','R')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Report12','Todays Movement at a Glance','R')

------------------------------------------------------------------------------------------------------------------------------------
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form28','Move Out','M')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form29','Movement Scheduler','M')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form31','Move Back','M')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form32','View Movement Schedule','M')
--insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form34','Approving various applications ','S')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form35','Employee Educational & Ref. Info.','E')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form37','Payscale Setup','T')
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Form39','Salary Disbursment','P')
------------------------------------------------------------------------------------------------------------------------------------
insert into soft_bag (software,scr_no,descript,portion) values ('Daffodil PMIS','Tax1','Tax Setup','T')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





--pro_soft_pass '9-001','','','','','I'

CREATE  PROCEDURE [pro_soft_pass]

@u_id varchar (15),
@u_name varchar (45),
@user_pass varchar (45),
@uid varchar (15),
@cancel bit,
@status char(1)
 
AS

set nocount on

Declare @Message varchar(200)

if @status='I'
begin

	if exists (select * from soft_pass where u_id=@U_id)
	begin
		set @Message='User id'+ ' (' +@U_id + ') ' + 'already exixts!'
	end
	else
	begin
		insert into soft_pass (u_id,u_name,uid,cancel) 
		values (@u_id,@u_name,@uid,@cancel)

		set @Message='User'+ ' (' +@U_id + ') ' +'created successfully!'
	end
end

if @status='U'
begin
	update soft_pass set u_name=@u_name,uid=@uid,cancel=@cancel 
	where u_id=@u_id
	
	set @Message='Update done successfully!'

end


if @status='P'
begin
	update soft_pass set user_pass=@user_pass 
	where u_id=@u_id
	set @Message='Password saved Successfully!'

end


if @status='C'
begin
	update soft_pass set user_pass=@user_pass 
	where u_id=@u_id
	set @Message='Password changed Successfully!'

end





if @status='D'

begin
	delete from soft_pass where u_id=@u_id

	set @Message='User'+ ' (' +@U_id + ') ' +'deleted successfully!'
end


select Message=@Message


set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO






CREATE   Procedure rptBrtalert
(
	@mode int,@date1 datetime,@date2 datetime

)

as
  if @mode=0 
     begin
        SELECT s.StudentID, s.StudentName, s.StuFatherName,
               s.StuMotherName,s.LegalGerdian,s.StuMoOrFaLet,
               s.StuCountryOfBirth,s.StuReligion,s.StuDateOfBirth,
               s.StuHight,s.StuBloodGroup,s.StuStreetPAddress, 
               s.StuCStreetAddress,s.StuPhone,s.StuEmail,
               (select a.classRoll From StudentAdmission a
                 where a.StudentId=s.StudentID
                 and a.serial_no=(select max(serial_no)  From 
                    StudentAdmission   
                 where StudentId=s.StudentID)),
                (select a.classid From StudentAdmission a
                 where a.StudentId=s.StudentID
                 and a.serial_no=(select max(serial_no)  From 
                    StudentAdmission   
                 where StudentId=s.StudentID)) as class_id,
                (select ClassName from classinfo 
                where classid=((select a.classid From StudentAdmission a
                 where a.StudentId=s.StudentID
                 and a.serial_no=(select max(serial_no)  From 
                    StudentAdmission   
                 where StudentId=s.StudentID)))) as class_name,
                s.ImmAddress,s.ImmPhone,s.ImmMob

           FROM         StudentInfo s

       WHERE     s.StuDateOfBirth between @date1 and @date2
    end






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE PROCEDURE rptEducation_Info
@emp_id varchar(10)
 AS

select emp_id,edu_count,exam_name,m_subject,pass_year,degree_from, result from emp_education
where emp_id=@emp_id


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO





CREATE Procedure rptStudentAdmisionInfo
(
	@ClassID varchar(5),
	@AdmissionDate1 datetime,
	@AdmissionDate2 datetime
)

as

SELECT     StudentAdmission.StudentId, StudentAdmission.AdmissionDate, StudentAdmission.ClassRoll, StudentAdmission.Shift, 
                      StudentAdmission.AdmitApproveBy, StudentAdmission.AdmitApproveDate, StudentInfo.StudentName, StudentInfo.StuFatherName, 
                      ClassInfo.ClassName, SectionInfo.Sectiondsc
FROM         StudentAdmission INNER JOIN
                      StudentInfo ON StudentAdmission.StudentId = StudentInfo.StudentID INNER JOIN
                      ClassInfo ON StudentAdmission.ClassId = ClassInfo.ClassID INNER JOIN
                      SectionInfo ON ClassInfo.ClassID = SectionInfo.ClassID
WHERE     (StudentAdmission.ClassId = @ClassID)

 

SELECT     StudentAdmission.StudentId, StudentAdmission.AdmissionDate, StudentAdmission.ClassRoll, StudentAdmission.Shift, 
                      StudentAdmission.AdmitApproveBy, StudentAdmission.AdmitApproveDate, StudentInfo.StudentName, StudentInfo.StuFatherName, 
                      ClassInfo.ClassName, SectionInfo.Sectiondsc
FROM         StudentAdmission INNER JOIN
                      StudentInfo ON StudentAdmission.StudentId = StudentInfo.StudentID INNER JOIN
                      ClassInfo ON StudentAdmission.ClassId = ClassInfo.ClassID INNER JOIN
                      SectionInfo ON ClassInfo.ClassID = SectionInfo.ClassID
WHERE     (StudentAdmission.AdmissionDate between @AdmissionDate1 and @AdmissionDate2 )



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO





CREATE Procedure rptStudentInfo
(
	@StudentID varchar(50)

)

as

SELECT     StudentID, StudentName, StuFatherName, StuMotherName, LegalGerdian, StuMoOrFaLet, StuCountryOfBirth, StuReligion, StuDateOfBirth, StuHight, 
                      StuBloodGroup, StuStreetPAddress, StuCStreetAddress, StuPhone, StuEmail, ImmAddress, ImmPhone, ImmMob
FROM         StudentInfo
WHERE     StudentID =@StudentID



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





/*
exec rpt_Bank_Statment '4-003','December','2004'

select * from Emp_job_Hist_Current
select * from Pay_Struc
*/


CREATE PROCEDURE rpt_Bank_Statment
	@Emp_ID varchar (10),
	@pay_Month varchar(9),
	@pay_Year varchar(4)
AS

set nocount on

Declare 
	 @Comp_Name varchar(75)
	,@Address varchar(150)

set @Comp_Name=(Select [Name] from Comp_Det_Info)
set @Address=(Select Address+ ', '+ City from Comp_Det_Info)


BEGIN
if @Emp_ID='' 
			SELECT a.emp_id,
			Emp_name=(select Emp_fna+" "+Emp_mna+" "+Emp_lna from Emp_Per_info where emp_id=a.emp_id),
			Dept=(select emp_dept from emp_Job_hist_current where emp_id=a.emp_id),
			Category=(select emp_rank from emp_Job_hist_current where emp_id=a.emp_id),
			Desig=(select emp_desig from emp_Job_hist_current where emp_id=a.emp_id),
			Comp_Name=@Comp_Name,
			Address=@Address,
			sum(CASE b.head_name WHEN 'Consolidated Salary' THEN c.amount ELSE 0 END) AS  Cons_Salary,
			sum(CASE b.head_name WHEN 'PF (Own Cont)' THEN c.amount ELSE 0 END) AS [PF_Own_Cont],
			sum(CASE b.head_name WHEN 'PF Loan' THEN c.amount ELSE 0 END) AS [PF_Loan],
			a.pay_Month,a.pay_Year 
			from Payroll_Main a,pay_struc b, payroll_sub c
			where a.pay_Month=@pay_Month AND a.pay_Year=@pay_Year 
			and a.a_gen=c.a_gen and b.head_code=c.head_code 
			group by a.emp_id,a.pay_Month,a.pay_Year
else
		SELECT a.emp_id,
			Emp_name=(select Emp_fna+" "+Emp_mna+" "+Emp_lna from Emp_Per_info where emp_id=a.emp_id),
			Dept=(select emp_dept from emp_Job_hist_current where emp_id=a.emp_id),
			Category=(select emp_rank from emp_Job_hist_current where emp_id=a.emp_id),
			Desig=(select emp_desig from emp_Job_hist_current where emp_id=a.emp_id),
			Comp_Name=@Comp_Name,
			Address=@Address,
			sum(CASE b.head_name WHEN 'Consolidated Salary' THEN c.amount ELSE 0 END) AS  Cons_Salary,
			sum(CASE b.head_name WHEN 'PF (Own Cont)' THEN c.amount ELSE 0 END) AS [PF_Own_Cont],
			sum(CASE b.head_name WHEN 'PF Loan' THEN c.amount ELSE 0 END) AS [PF_Loan],
			a.pay_Month,a.pay_Year 
			from Payroll_Main a,pay_struc b, payroll_sub c
			where a.pay_Month=@pay_Month AND a.pay_Year=@pay_Year 
			and a.a_gen=c.a_gen and b.head_code=c.head_code  
			group by a.emp_id,a.pay_Month,a.pay_Year
end


/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/







set nocount off








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREATE            procedure rpt_Emp_Att_Info  
@Mode varchar (10),
@Param varchar (35),
@pay_Month varchar(12),
@pay_Year varchar(4)
AS
declare @M_Days varchar(12)

set @M_days=(select dbo.GetMonthDays(@pay_Month,@pay_Year))

declare @Tot_Hol  int
, @Tot_Weekend int
, @Tot_Leave int


set @Tot_Hol = (select count(*) from hol_list where 
(datepart(year,Str_date)=@pay_Year and datename(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datename(month,End_date)=@Pay_month)
)

set @Tot_Weekend = (select count(*) from hol_list where Category ='Weekend'and
(datepart(year,Str_date)=@pay_Year and datename(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datename(month,End_date)=@Pay_month)
)


IF @mode='Ind'
BEGIN


set @Tot_Leave = (dbo.GetLeaveDuration(@Param,@Pay_month,@pay_Year))

select a.emp_id,Emp_fna=(a.Emp_fna +' '+ a.Emp_mna +' '+ a.Emp_lna),
	b.emp_desig,b.emp_dept,emp_login=convert(char(8),c.emp_login,14),
	Emp_logout=convert(char(8),c.Emp_logout,14),c.Entry_dt,
	In_Status = case c.In_status when '1' then '*' end,
	out_Status =case c.Out_status when '8' then '*' end
	,Notes=(select dbo.GetAttnNotes(Emp_LOgIn))

	----------------Calculates Office Staying Hours----------------------------------
	,Hr_Stay=(select dbo.GetHrsInOffice(@Param,Emp_LogIn))	
	----------------Calculates Movement Hours----------------------------------------
	,Hr_Movement=(select dbo.GetHrInMovement(@Param,Emp_LogIn))
	----------------Calculates Movement Number---------------------------------------
	,Num_Movement=(select dbo.GetMovementNum(@Param,Emp_LogIn))
	---------------------------------------------------------------------------------

	,Tot_Att=(dbo.GetTotAttn(@Param,@Pay_month,@pay_Year))

	,Late =(dbo.GetTotLateAttn(@Param,@Pay_month,@pay_Year))

	,Working_Days=(@M_Days-@Tot_Hol)
	,Weekend=@Tot_Weekend
	,Other_holidays=(@Tot_Hol-@Tot_Weekend)
	,Leave=@Tot_Leave
	from emp_Per_info a,Emp_Job_Hist_Current b,Emp_Att_info c where 
	a.Emp_ID=B.Emp_ID and a.Emp_ID=c.Emp_ID and a.Emp_ID=@Param
	and datepart(year,c.Entry_dt)=@pay_Year
	and datename(month,c.Entry_dt)=@Pay_month
order by c.Entry_dt 	 
	
end




IF @mode='All'	-------Returns individual  record for Attn Summery Report

BEGIN


declare @Present int
,@Late int
,@DailyWorkHr float

set @DailyWorkHr=(select dbo.GetDailyWorkingHr(@Pay_Month,@Pay_Year))

set @Param=ltrim(rtrim(@Param))

Set @Present =(dbo.GetTotAttn(@Param,@Pay_month,@pay_Year))
set @Late =(dbo.GetTotLateAttn(@Param,@Pay_month,@pay_Year))

set @Tot_Leave = (dbo.GetLeaveDuration(@Param,@Pay_month,@pay_Year))

SELECT DISTINCT a.emp_id,Emp_fna=(a.Emp_fna +'  '+ a.Emp_mna +'  '+ a.Emp_lna),
		b.emp_desig,Dept='Department: '+b.emp_dept,Month_year=@pay_Month+', '+@Pay_Year
		,Work_Day=(@M_Days-@Tot_Hol)
		,MonthlyWorkHr=@DailyWorkHr*(@M_Days-@Tot_Hol)
		,Tot_Hol=@Tot_Hol 
		,Weekend=@Tot_Weekend 
		,Present=@Present
		,Late=@Late	
		,Leave=@Tot_Leave	
		,Off_Stay=convert(Char(6),(select dbo.GetOffStayHr(@Param,@pay_Month,@pay_Year)))
		,Movement=(select dbo.GetMonthlyMovementHr(@Param,@pay_Month,@pay_Year))
		from emp_Per_info a,Emp_Job_Hist_Current b where 
		a.Emp_ID=B.Emp_ID and a.Emp_ID=@Param
		--and datename(year,c.Entry_dt)=@pay_Year
		--and datename(month,c.Entry_dt)=@pay_Month 
end


IF @mode='Dpt_AllInd'
BEGIN
declare @WorkingDs int
set @WorkingDs =(select dbo.GetWorkingDays(@pay_month,@pay_Year))
select distinct a.emp_id,Emp_fna=(a.Emp_fna +' ' + a.Emp_mna +' '+ a.Emp_lna),
	b.emp_desig,Department=@Param,
	emp_login=convert(char(8),c.emp_login,14),
	Emp_logout=convert(char(8),c.Emp_logout,14),c.Entry_dt,
	In_Status = case c.In_status when '1' then 'Late' end,
	out_Status =case c.Out_status when '8' then 'Early' end

	,WorkingDays=@WorkingDs,Week_End=@Tot_Weekend,Other_Holidays=(@Tot_Hol-@Tot_Weekend)
	,Present=(select dbo.GetPresence(a.emp_id,@pay_month,@Pay_year))
	,Late=(select dbo.GetLateCount(a.emp_id,@pay_month,@Pay_year))
	,Absent=(select dbo.GetAbsDays(a.emp_id,@pay_month,@Pay_year))
	,Leave=(select dbo.GetLeaveDays(a.emp_id,@pay_month,@Pay_year))
	
	----------------Calculates Office Staying Hour----------------------------------
	,Hr_Stay=(select dbo.GetHrsInOffice(a.Emp_Id,c.Emp_LogIn))	
	--------------------------------------------------------------------------------

	,Month_Year=@pay_Month+', '+@pay_Year

	from emp_Per_info a,Emp_Job_Hist_Current b,Emp_Att_info c where 
	a.Emp_ID=B.Emp_ID and a.Emp_ID=c.Emp_ID and b.emp_dept=@param
	and datepart(year,c.Entry_dt)=@pay_Year
	and datename(month,c.Entry_dt)=@Pay_month	 
	
end






















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







/****** Object:  Stored Procedure dbo.rpt_Emp_Att_Summery    Script Date: 3/7/02 10:43:52 AM ******/
--rpt_Emp_Att_Summery 'Software','May','2002'
--rpt_Emp_Att_Info 'All','9-015','May','2002'


CREATE          Procedure rpt_Emp_Att_Summery 
@emp_dept varchar(35),
@pay_Month varchar(12),
@pay_Year varchar(4)
AS
set nocount on
declare @Emp_ID varchar(10)
-----------------------------------------------
create table #Attendance(
		emp_id varchar(10)
		,Emp_fna varchar(45)
	        ,emp_desig varchar(25)
		,emp_dept varchar(50)   
		,Month_Year varchar(20)
		,Work_Day int
		,Monthly_Work_Hr char(6)
		,Tot_Hol int
		,Weekend int
		,Present int
		,Late int
		,Leave int
		,Office_Stay char(6)
		,Movement char(6))
-----------------------------------------------
DECLARE Emp_ID_cursor CURSOR FOR
Select Distinct Emp_Id from Emp_Job_Hist_Current where emp_dept=@emp_dept
OPEN Emp_ID_cursor

    FETCH NEXT FROM Emp_ID_cursor into @Emp_ID
	
	WHILE (@@FETCH_STATUS = 0)

	BEGIN
              
		INSERT INTO #Attendance(emp_id,Emp_fna,emp_desig,emp_dept,
			Month_Year,Work_Day,Monthly_Work_Hr,Tot_Hol,Weekend,
			Present,Late,Leave,Office_Stay,Movement)
		
		EXEC rpt_Emp_Att_Info 'All',@Emp_id,@pay_Month,@pay_Year
		
		FETCH NEXT FROM Emp_ID_cursor into @Emp_ID	
	END


CLOSE Emp_ID_cursor
DEALLOCATE Emp_ID_cursor
Select DISTINCT * from #Attendance



set nocount off









GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.rpt_Emp_Leave_Info    Script Date: 3/7/02 10:43:47 AM ******/
/****** Object:  Stored Procedure dbo.rpt_Emp_Leave_Info    Script Date: 10/28/01 4:34:42 PM ******/
/****** Object:  Stored Procedure dbo.rpt_Emp_Leave_Info    Script Date: 10/22/01 6:24:41 PM ******/
/****** Object:  Stored Procedure dbo.rpt_Emp_Leave_Info    Script Date: 9/17/00 4:34:14 PM ******/
/****** Object:  Stored Procedure dbo.rpt_Emp_Leave_Info    Script Date: 9/4/01 6:30:48 PM ******/
/****** Object:  Stored Procedure dbo.rpt_Emp_Leave_Info    Script Date: 6/20/01 12:00:41 PM ******/
CREATE procedure rpt_Emp_Leave_Info
@Mode varchar (5),
@Emp_id varchar(10),
@Start_dt datetime,                   
@End_dt datetime
as 
if @mode='Ind'				---@Emp_ID, ***** Regardless date ***********
begin
	select 
		A.Emp_Id,emp_Name=(A.Emp_fna+" "+A.emp_mna+" "+A.emp_lna),
		B.leave_name,B.Start_dt,B.End_dt,B.Des,
		C.Emp_desig,C.Emp_section 
		from Emp_per_info A, Emp_Leave_info B,Emp_Job_hist_current c
		where a.emp_Id= b.emp_id and b.emp_Id= c.emp_id and a.emp_id=@Emp_ID
		
end
if @mode='Ind'				---@Emp_ID,@Start_dt,@End_dt,**** date specific***
begin
	select 
		A.Emp_Id,emp_Name=(A.Emp_fna+" "+A.emp_mna+" "+A.emp_lna),
		B.leave_name,B.Start_dt,B.End_dt,B.Des,
		C.Emp_desig,C.Emp_section 
		from Emp_per_info A, Emp_Leave_info B,Emp_Job_hist_current c
		where a.emp_Id= b.emp_id and b.emp_Id= c.emp_id and a.emp_id=@Emp_ID
		and B.Start_dt BETWEEN @Start_dt AND @End_dt
end
if @Mode='All'					------@Start_dt,@End_dt
begin
	select 
		A.Emp_Id,emp_Name=(A.Emp_fna+" "+A.emp_mna+" "+A.emp_lna),
		B.leave_name,B.Start_dt,B.End_dt,B.Des,
		C.Emp_desig,C.Emp_section
		from Emp_per_info A, Emp_Leave_info B,Emp_Job_hist_current c
		where a.emp_Id= b.emp_id and b.emp_Id= c.emp_id
		and B.Start_dt BETWEEN @Start_dt AND @End_dt
		order by a.emp_id
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE procedure rpt_Emp_Performance --'9-023','4','2002'
@Emp_id varchar (10),
@pay_Month varchar(2),
@pay_Year varchar(4)
AS
declare @M_Days varchar(2)
,@Month_NM varchar(12)
Set @M_Days=(CASE @pay_Month
		when '1' THEN '31'
		when '2' THEN '28'
		when '3' THEN '31'
		when '4' THEN '30'
		when '5' THEN '31'
		when '6' THEN '30'
		when '7' THEN '31'
		when '8' THEN '31'
		when '9' THEN '30'
		when '10' THEN '31'
		When '11' THEN '30'
		When '12' THEN '31'
	end)

Set @Month_NM=(CASE @pay_Month
		when '1' THEN 'January'
		when '2' THEN 'February'
		when '3' THEN 'March'
		when '4' THEN 'April'
		when '5' THEN 'May'
		when '6' THEN 'June'
		when '7' THEN 'July'
		when '8' THEN 'August'
		when '9' THEN 'September'
		when '10' THEN 'October'
		When '11' THEN 'November'
		When '12' THEN 'December'
	end)

declare @Tot_Hol  int
, @Tot_Weekend int
, @Tot_Leave int


set @Tot_Hol = (select count(*) from hol_list where 
(datepart(year,Str_date)=@pay_Year and datepart(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datepart(month,End_date)=@Pay_month)
)

set @Tot_Weekend = (select count(*) from hol_list where Category ='Weekend'and
(datepart(year,Str_date)=@pay_Year and datepart(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datepart(month,End_date)=@Pay_month)
)


declare @Present int
declare @Late int

Set @Present =(select count(*)from emp_att_info where emp_id=@Emp_id
		and datename(year,Entry_dt)=@pay_Year and datepart(month,Entry_dt)=@pay_Month)

set @Late = (Select count(*) from emp_Att_info where In_status='1'and emp_id=@Emp_id
		and datename(year,Entry_dt)=@pay_Year
		and datepart(month,Entry_dt)=@pay_Month)

set @Tot_Leave = (select count(*) from Emp_Leave_Info where emp_Id=@Emp_Id and 
(datepart(year,Start_dt)=@pay_Year and datepart(month,Start_dt)=@Pay_month)and
(datepart(year,End_dt)=@pay_Year and datepart(month,End_dt)=@Pay_month)
)		

SELECT DISTINCT a.emp_id,Emp_fna=(a.Emp_fna + '' + a.Emp_mna +''+ a.Emp_lna),
		b.emp_desig,b.emp_dept
		,b.Emp_join_date,b.Emp_job_type
		,Pay_type = case b.Pay_type when '1' then 'Salary based'else 'Wage based' end
		,Duty_type =case b.Duty_type when '1' then 'Fixed time'else 'Shifting' end
		,Month_NM=@Month_NM,Mn_Num=@pay_Month
		,[Year]=@Pay_Year
		,Work_Day=(@M_Days-@Tot_Hol)
		,Tot_Hol=@Tot_Hol 
		,Weekend=@Tot_Weekend 
		,Present=@Present
		,Late=@Late	
		,Leave=@Tot_Leave	

		from emp_Per_info a,Emp_Job_Hist_Current b where 
		a.Emp_ID=B.Emp_ID and a.Emp_ID=@Emp_ID		






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.rpt_Employee_list    Script Date: 3/7/02 10:43:48 AM ******/
/****** Object:  Stored Procedure dbo.rpt_Employee_list    Script Date: 10/28/01 4:34:43 PM ******/
/****** Object:  Stored Procedure dbo.rpt_Employee_list    Script Date: 10/22/01 6:24:42 PM ******/
/****** Object:  Stored Procedure dbo.rpt_Employee_list    Script Date: 9/17/00 4:34:14 PM ******/
/****** Object:  Stored Procedure dbo.rpt_Employee_list    Script Date: 9/4/01 6:30:49 PM ******/
/****** Object:  Stored Procedure dbo.rpt_Employee_list    Script Date: 6/20/01 12:00:41 PM ******/
CREATE procedure rpt_Employee_list
@Mode varchar(5),
@Param_1 varchar(20),
@Param_2 varchar (15)
as
if @Mode='Desig'
begin
	select
	a.Emp_id,
	Emp_Name=(a.Emp_fna+" "+a.Emp_mna+" "+a.Emp_lna),	
	a.Emp_perm_add ,
	a.Emp_perm_town ,
	a.Emp_tel1,
	b.Emp_join_date,
	b.Emp_desig,
	b.Emp_Dept
	from emp_per_info a,emp_job_hist_current b
	where a.emp_id=
		(select emp_id from emp_job_hist_current	
			where Emp_desig=@Param_1)
ORDER BY a. emp_id 
end
if @Mode='Dept'
begin
	select
	a.Emp_id,
	Emp_Name=(a.Emp_fna+" "+a.Emp_mna+" "+a.Emp_lna),	
	a.Emp_perm_add ,
	a.Emp_perm_town ,
	a.Emp_tel1,
	b.Emp_join_date,
	b.Emp_desig,
	b.Emp_Dept
	from emp_per_info a,emp_job_hist_current b
	where a.emp_id=
		(select emp_id from emp_job_hist_current	
			where Emp_dept=@Param_1)
ORDER BY a. emp_id 
end
if @Mode='BldGr'
begin
	select
	a.Emp_id,
	Emp_Name=(a.Emp_fna+" "+a.Emp_mna+" "+a.Emp_lna),	
	a.Emp_perm_add ,
	a.Emp_perm_town ,
	a.Emp_tel1 ,
	a.Blood_group,
	b.Emp_desig,
	b.Emp_Dept
	from emp_per_info a,emp_job_hist_current b
	where a.emp_id=
		(select emp_id from emp_job_hist_current
			where Blood_group=@Param_1)
ORDER BY a. emp_id 
end
if @Mode='JbTp'
begin
	select
	a.Emp_id,
	Emp_Name=(a.Emp_fna+" "+a.Emp_mna+" "+a.Emp_lna),	
	a.Emp_perm_add ,
	a.Emp_perm_town ,
	a.Emp_tel1 ,
	b.Emp_desig,
	b.Emp_Dept,
	b.Emp_join_date,
	b.Emp_job_type
	from emp_per_info a,emp_job_hist_current b
	where a.emp_id=
		(select emp_id from emp_job_hist_current
			where Emp_job_type=@Param_1)
ORDER BY a. emp_id 
end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.rpt_Hol_list    Script Date: 3/7/02 10:43:37 AM ******/
/****** Object:  Stored Procedure dbo.rpt_Hol_list    Script Date: 10/28/01 4:34:31 PM ******/
/****** Object:  Stored Procedure dbo.rpt_Hol_list    Script Date: 1/29/01 10:24:54 AM ******/
create procedure rpt_Hol_list
as
select Hol_name,Hol_desc,Str_date,End_date,Category from Hol_list




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/*
rpt_Payroll_Ind_All 'All','','December','2004'
select * from Emp_job_Hist_Current
select * from Pay_Struc
*/


CREATE      PROCEDURE rpt_Payroll_Ind_All
	@Mode varchar (5),
	@Param_1 varchar (40),			/*  Param_1---------->Emp_Id / Department   */
	@pay_Month varchar(12),
	@pay_Year varchar(4)
AS

set nocount on

Declare @Comp_Name varchar(75)
	,@Address varchar(150)

set @Comp_Name=(Select [Name] from Comp_Det_Info)
set @Address=(Select Address+ ', '+ City from Comp_Det_Info)

IF @mode='All'
BEGIN
SELECT a.emp_id,
Emp_name=(select Emp_fna+" "+Emp_mna+" "+Emp_lna from Emp_Per_info where emp_id=a.emp_id),
Dept=(select emp_dept from emp_Job_hist_current where emp_id=a.emp_id),
Category=(select emp_rank from emp_Job_hist_current where emp_id=a.emp_id),
Desig=(select emp_desig from emp_Job_hist_current where emp_id=a.emp_id),
Comp_Name=@Comp_Name,
Address=@Address,
sum(CASE b.head_name WHEN 'Consolidated Salary' THEN c.amount ELSE 0 END) AS  Cons_Salary,
sum(CASE b.head_name WHEN 'House Rent' THEN c.amount ELSE 0 END) AS House_Rent,
sum(CASE b.head_name WHEN 'Washing Allowance' THEN c.amount ELSE 0 END) AS  Washing_Allowance,
sum(CASE b.head_name WHEN 'Medical Allowance' THEN c.amount  ELSE 0 END) AS Medical_Allowance,
sum(CASE b.head_name WHEN 'T.A (ADD)' THEN c.amount ELSE 0 END) AS  [Travel_All(+)],
sum(CASE b.head_name WHEN 'Tiffin' THEN c.amount ELSE 0 END) AS Tiffin,
sum(CASE b.head_name WHEN 'Responsibility Allowance' THEN c.amount ELSE 0 END) AS [Responsibility_Allowance],
sum(CASE b.head_name WHEN 'Overtime' THEN c.amount ELSE 0 END) AS [Overtime],
sum(CASE b.head_name WHEN 'T.A' THEN c.amount ELSE 0 END) AS [T.A(-)],
sum(CASE b.head_name WHEN 'PF (Own Cont)' THEN c.amount ELSE 0 END) AS [PF_Own_Cont],
sum(CASE b.head_name WHEN 'PF Loan' THEN c.amount ELSE 0 END) AS [PF_Loan],
sum(CASE b.head_name WHEN 'Overtime' THEN c.amount ELSE 0 END) AS [Overtime],
sum(CASE b.head_name WHEN 'Others (-)' THEN c.amount ELSE 0 END) AS [Others_Deduct],
sum(CASE b.head_name WHEN 'Others (+)' THEN c.amount ELSE 0 END) AS [Others_Add],
a.pay_Month,a.pay_Year 
from Payroll_Main a,pay_struc b, payroll_sub c
where a.pay_Month=@pay_Month AND a.pay_Year=@pay_Year 
and a.a_gen=c.a_gen and b.head_code=c.head_code
group by a.emp_id,a.pay_Month,a.pay_Year
end


/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/


IF @mode='Dept'
BEGIN
SELECT a.emp_id,
Emp_name=(select Emp_fna+" "+Emp_mna+" "+Emp_lna from Emp_Per_info where emp_id=a.emp_id),
Dept=(select emp_dept from emp_Job_hist_current where emp_id=a.emp_id),
Category=(select emp_rank from emp_Job_hist_current where emp_id=a.emp_id),
Desig=(select emp_desig from emp_Job_hist_current where emp_id=a.emp_id),
sum(CASE b.head_name WHEN 'Consolidated Salary' THEN c.amount ELSE 0 END) AS  Cons_Salary,
sum(CASE b.head_name WHEN 'House Rent' THEN c.amount ELSE 0 END) AS House_Rent,
---sum(CASE b.head_name WHEN 'Washing Allowance' THEN c.amount ELSE 0 END) AS  [Others(+)],
sum(CASE b.head_name WHEN 'Washing Allowance' THEN c.amount ELSE 0 END) AS  Washing_Allowance,
sum(CASE b.head_name WHEN 'Medical Allowance' THEN c.amount  ELSE 0 END) AS Medical_Allowance,
sum(CASE b.head_name WHEN 'T.A (ADD)' THEN c.amount ELSE 0 END) AS  [Travel_All(+)],
sum(CASE b.head_name WHEN 'Tiffin' THEN c.amount ELSE 0 END) AS Tiffin,
sum(CASE b.head_name WHEN 'Responsibility Allowance' THEN c.amount ELSE 0 END) AS [Responsibility_Allowance],
sum(CASE b.head_name WHEN 'Overtime' THEN c.amount ELSE 0 END) AS [Overtime],
sum(CASE b.head_name WHEN 'T.A' THEN c.amount ELSE 0 END) AS [T.A(-)],
sum(CASE b.head_name WHEN 'PF (Own Cont)' THEN c.amount ELSE 0 END) AS [PF_Own_Cont],
sum(CASE b.head_name WHEN 'PF Loan' THEN c.amount ELSE 0 END) AS [PF_Loan],
sum(CASE b.head_name WHEN 'Overtime' THEN c.amount ELSE 0 END) AS [Overtime],
sum(CASE b.head_name WHEN 'Others (-)' THEN c.amount ELSE 0 END) AS [Others_Deduct],
sum(CASE b.head_name WHEN 'Others (+)' THEN c.amount ELSE 0 END) AS [Others_Add],


pay_Month,a.pay_Year,q.emp_dept
from Payroll_Main a,pay_struc b, payroll_sub c,  emp_Job_hist_current q
where a.pay_Month=@pay_Month AND a.pay_Year=@pay_Year
and a.a_gen=c.a_gen and b.head_code=c.head_code and q.emp_dept=@Param_1
and a.emp_id=q.emp_id
group by a.emp_id,a.pay_Month,a.pay_Year,q.emp_dept
order by a.emp_id
end

/*
		**************SALARY HEAD SETUP FOR STATE MEDICAL FACULTY**************

01		Consolidated Salary		+	F	dsl	7/19/2005
02		House Rent			+	F	dsl	7/19/2005
03		Washing Allowance		+	V	dsl	7/19/2005
04		Medical Allowance		+	F	dsl	7/19/2005
07		T.A				+	V	dsl	7/19/2005
09		Tiffin				+	F	dsl	7/19/2005
05		Responsibility Allowance	+	V	dsl	7/19/2005
06		Overtime			+	V	dsl	7/19/2005 
08		T.A				-	V	dsl	7/19/2005
10		PF (Own Cont)			-	V	dsl	7/19/2005
11		PF Loan				-	V	dsl	7/19/2005
12		Others (-)			-	V	dsl	7/19/2005
13		Others (+)			+	V	dsl	7/19/2005
*/

/*+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/

IF @mode='Acc'
BEGIN
SELECT a.emp_id,
Emp_name=(select Emp_fna+" "+Emp_mna+" "+Emp_lna from Emp_Per_info where emp_id=a.emp_id),
Dept=(select emp_dept from emp_Job_hist_current where emp_id=a.emp_id),
Category=(select emp_rank from emp_Job_hist_current where emp_id=a.emp_id),
Desig=(select emp_desig from emp_Job_hist_current where emp_id=a.emp_id),
Account_No=(select Acc_No from emp_Bank_Account where emp_id=a.emp_id),
sum(CASE b.head_name WHEN 'Consolidated Salary' THEN c.amount ELSE 0 END) AS  Cons_Salary,
sum(CASE b.head_name WHEN 'House Rent' THEN c.amount ELSE 0 END) AS House_Rent,
---sum(CASE b.head_name WHEN 'Washing Allowance' THEN c.amount ELSE 0 END) AS  [Others(+)],
sum(CASE b.head_name WHEN 'Washing Allowance' THEN c.amount ELSE 0 END) AS  Washing_Allowance,
sum(CASE b.head_name WHEN 'Medical Allowance' THEN c.amount  ELSE 0 END) AS Medical_Allowance,
sum(CASE b.head_name WHEN 'T.A' THEN c.amount ELSE 0 END) AS  [Travel_All(+)],
sum(CASE b.head_name WHEN 'Tiffin' THEN c.amount ELSE 0 END) AS Tiffin,
sum(CASE b.head_name WHEN 'Responsibility Allowance' THEN c.amount ELSE 0 END) AS [Responsibility_Allowance],
sum(CASE b.head_name WHEN 'Overtime' THEN c.amount ELSE 0 END) AS [Overtime],
sum(CASE b.head_name WHEN 'T.A' THEN c.amount ELSE 0 END) AS [T.A(-)],
sum(CASE b.head_name WHEN 'PF (Own Cont)' THEN c.amount ELSE 0 END) AS [PF_Own_Cont],
sum(CASE b.head_name WHEN 'PF Loan' THEN c.amount ELSE 0 END) AS [PF_Loan],
sum(CASE b.head_name WHEN 'Overtime' THEN c.amount ELSE 0 END) AS [Overtime],
sum(CASE b.head_name WHEN 'Others (-)' THEN c.amount ELSE 0 END) AS [Others_Deduct],
sum(CASE b.head_name WHEN 'Others (+)' THEN c.amount ELSE 0 END) AS [Others_Add],

pay_Month,a.pay_Year,q.emp_dept
from Payroll_Main a,pay_struc b, payroll_sub c,  emp_Job_hist_current q

where a.pay_Month=@pay_Month AND a.pay_Year=@pay_Year
and a.a_gen=c.a_gen and b.head_code=c.head_code and q.emp_dept=@Param_1
and a.emp_id=q.emp_id
group by a.emp_id,a.pay_Month,a.pay_Year,q.emp_dept
order by a.emp_id
end



set nocount off







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



/*
rpt_Payroll_Payslip 'All','','August','2003'

select * from Emp_job_Hist_Current

select * from Pay_Struc

dbo.GetTotAttn
dbo.GetTot_Holday
dbo.GetTotLateAttn
dbo.GetWorkingDays

*/

CREATE    PROCEDURE rpt_Payroll_Payslip
@Mode varchar (5),
@Param_1 varchar (40),			/*  Param_1---------->Emp_Id / Department   */
@pay_Month varchar(12),
@pay_Year varchar(4)
AS

set nocount on

Declare @Comp_Name varchar(75)
,@Address varchar(150)

set @Comp_Name=(Select [Name] from Comp_Det_Info)
set @Address=(Select Address+ ', '+ City from Comp_Det_Info)

IF @mode='All'
BEGIN
SELECT a.emp_id,
Emp_name=(select Emp_fna+" "+Emp_mna+" "+Emp_lna from Emp_Per_info where emp_id=a.emp_id),
Dept=(select emp_dept from emp_Job_hist_current where emp_id=a.emp_id),
Category=(select emp_rank from emp_Job_hist_current where emp_id=a.emp_id),
Desig=(select emp_desig from emp_Job_hist_current where emp_id=a.emp_id),
Comp_Name=@Comp_Name,
Address=@Address,
---------------------------------------------------------------------
Late=dbo.GetTotLateAttn(a.Emp_Id,@Pay_Month,@Pay_Year),
Wrk_Days=dbo.GetWorkingDays(@pay_Month,@pay_Year),
Attn=dbo.GetTotAttn(a.Emp_Id,@Pay_Month,@Pay_Year),
Leave= dbo.GetLeaveDuration(a.Emp_Id,@Pay_Month,@Pay_Year),
Absent=dbo.GetAbsentDays(a.Emp_Id,@Pay_Month,@Pay_Year),
---Modified on January 20, 2004 By Shameem Ferdous
Total_Weekends=dbo.GetTotalWeekEnds(@Pay_Month,@Pay_Year),
Total_Other_Holidays=dbo.GetTot_OtherHolday(@Pay_Month,@Pay_Year),
Total_Hartal=dbo.GetTotalHartals(@Pay_Month,@Pay_Year),
---------------------------------------------------------------------

sum(CASE b.head_name WHEN 'Consolidated Salary' THEN c.amount ELSE 0 END) AS  Cons_Salary,
sum(CASE b.head_name WHEN 'Telephone/Mobile' THEN c.amount ELSE 0 END) AS Telephone_Mobile,
sum(CASE b.head_name WHEN 'Others(+)' THEN c.amount ELSE 0 END) AS  [Others(+)],
sum(CASE b.head_name WHEN 'HR Deduction' THEN c.amount  ELSE 0 END) AS HR_Deduction,
sum(CASE b.head_name WHEN 'Loan Deduction' THEN c.amount ELSE 0 END) AS  Loan_Deduc,
sum(CASE b.head_name WHEN 'Advance Deduction' THEN c.amount ELSE 0 END) AS [Adv_Deduc],
sum(CASE b.head_name WHEN 'Absent Deduction' THEN c.amount ELSE 0 END) AS [Abs_Deduc],
sum(CASE b.head_name WHEN 'Others(-)' THEN c.amount ELSE 0 END) AS [Others(-)]

,a.pay_Month,a.pay_Year
 

from Payroll_Main a,pay_struc b, payroll_sub c
where a.pay_Month=@pay_Month AND a.pay_Year=@pay_Year 
and a.a_gen=c.a_gen and b.head_code=c.head_code
group by a.emp_id,a.pay_Month,a.pay_Year
end

set nocount off







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



---exec rpt_Provident_Fund_Statment 'September','2005'

CREATE  PROCEDURE rpt_Provident_Fund_Statment
	@pay_Month varchar(9),
	@pay_Year varchar(4)
AS

set nocount on

BEGIN

SELECT  sum(Amount) AS PF_Contribution,isnull((SELECT isnull(sum(b.amount),0) 
FROM Payroll_sub b,Payroll_main d
WHERE b.head_code = '11' and b.a_gen=d.a_gen and d.pay_month=@pay_Month and 
d.pay_year=@pay_Year group by b.head_code),0) AS PF_loan 
FROM Payroll_sub a,Payroll_main c
WHERE Head_Code = '10' and c.pay_month=@pay_Month and c.pay_year=@pay_Year
and a.a_gen=c.a_gen group by a.head_code 


end

set nocount off





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



---exec rpt_SalStat_SendtoBank 'December','2004'



CREATE PROCEDURE rpt_SalStat_SendtoBank
	@pay_Month varchar(9),
	@pay_Year varchar(4)
AS

set nocount on

Declare 
		@Comp_Name varchar(75),
		@Address varchar(150)

set @Comp_Name=(Select [Name] from Comp_Det_Info)
set @Address=(Select Address+ ', '+ City from Comp_Det_Info)

BEGIN
		SELECT a.emp_id,
		Emp_name=(select Emp_fna+" "+Emp_mna+" "+Emp_lna from Emp_Per_info where emp_id=a.emp_id),
		Dept=(select emp_dept from emp_Job_hist_current where emp_id=a.emp_id),
		Category=(select emp_rank from emp_Job_hist_current where emp_id=a.emp_id),
		Desig=(select emp_desig from emp_Job_hist_current where emp_id=a.emp_id),
		AccountNo=(SELECT AccountNo FROM Fixed_Pay WHERE  (Emp_id =a.emp_id)),
		Comp_Name=@Comp_Name,
		Address=@Address,
		sum(CASE b.head_name WHEN 'Consolidated Salary' THEN c.amount ELSE 0 END) AS  Cons_Salary,
		sum(CASE b.head_name WHEN 'PF (Own Cont)' THEN c.amount ELSE 0 END) AS [PF_Own_Cont],
		sum(CASE b.head_name WHEN 'PF Loan' THEN c.amount ELSE 0 END) AS [PF_Loan],
		a.pay_Month,a.pay_Year 
		from Payroll_Main a,pay_struc b, payroll_sub c
		where a.pay_Month=@pay_Month AND a.pay_Year=@pay_Year 
		and a.a_gen=c.a_gen and b.head_code=c.head_code
		group by a.emp_id,a.pay_Month,a.pay_Year
end

set nocount off




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE procedure s_u_d_leave_info
(
 @mode int,
 @Stdid varchar(20),
 @acayr  varchar(20),
 @stdate datetime,
 @eddate datetime,
 @cause  varchar(200),
 @serial  integer
)
as
declare @Srl as int
  begin
     if @mode=1 
        begin
          select @srl=isnull(max(serial),0)+1 from student_leave_info
          insert into Student_leave_info values(@Stdid,@acayr,@stdate, @eddate,@cause,@srl)
        end 
     if @mode=1 
        begin
           update Student_leave_info set L_st_dt=@stdate, L_ed_dt=@eddate,remarks=@cause
           where serial=@serial
        end 
     if @mode=3 
        begin
          delete from  Student_leave_info where serial=@serial
        end 
     
end 

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







CREATE PROCEDURE size_pmt
@Mode varchar(10),
@Portion varchar(5),
@code int,
@U_Id varchar(10)

 AS

IF @Mode='Single'

BEGIN

	delete from permit where code=@code and U_Id=@U_Id

END


IF @Mode='All'

BEGIN
	IF @Portion ='All'

	BEGIN

		delete from permit where U_Id=@U_Id

	END

	ELSE

	BEGIN

		delete from permit where U_Id=@U_Id
		and Code in (SELECT Code from soft_bag where U_Id=@U_Id and Portion=@Portion)


	END


END








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.sp_leaves_only_for_u    Script Date: 3/7/02 10:43:39 AM ******/
/****** Object:  Stored Procedure dbo.sp_leaves_only_for_u    Script Date: 10/28/01 4:34:32 PM ******/
/****** Object:  Stored Procedure dbo.sp_leaves_only_for_u    Script Date: 10/22/01 6:24:32 PM ******/
/****** Object:  Stored Procedure dbo.sp_leaves_only_for_u    Script Date: 9/17/00 4:34:05 PM ******/
/****** Object:  Stored Procedure dbo.sp_leaves_only_for_u    Script Date: 9/4/01 6:30:51 PM ******/
CREATE PROCEDURE sp_leaves_only_for_u
@u_id varchar(10)
AS
select * from approval where Tier_1=@u_id 
union
select * from approval where Tier_2=@u_id
union
select * from approval where Tier_3=@u_id
union
select * from approval where Tier_4=@u_id
union
select * from approval where Final_Tier=@u_id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




/****** Object:  Stored Procedure dbo.sp_set_position    Script Date: 3/7/02 10:43:41 AM ******/
/****** Object:  Stored Procedure dbo.sp_set_position    Script Date: 10/28/01 4:34:34 PM ******/
/****** Object:  Stored Procedure dbo.sp_set_position    Script Date: 10/22/01 6:24:34 PM ******/
/****** Object:  Stored Procedure dbo.sp_set_position    Script Date: 9/17/00 4:34:07 PM ******/
/****** Object:  Stored Procedure dbo.sp_set_position    Script Date: 9/4/01 6:30:53 PM ******/
CREATE PROCEDURE sp_set_position  
@app_id varchar(10),
@pos varchar(1)
AS
update approval set move_pos=@pos where app_id=@app_id




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO








CREATE   procedure temp_Collec_save
(
		    @mode	   varchar(1),
            @Seq_no    integer,
            @Srl_no    integer,
            @user_id    varchar(50),
            @std_id    varchar(15),  
			@class_code  varchar(5),
			@Fee_code  varchar(2),
            @Fee_title  varchar(100), 
            @Act_amount  decimal,
            @fine  decimal,
            @Discount  decimal
             )         
	AS
         declare @loc_seq as   int
         declare @loc_srl  as int
          

         if @mode='s'
               begin
		    			if not exists (select Seq_no from temp_collect where Seq_no=@Seq_no)
                            begin
                                set @Seq_no= (select isnull(max(Seq_no),0)+1 from temp_collect )
                                set @loc_srl=1
				                   
				                insert into temp_collect(Seq_no,Srl_no,u_id,class_code,Fee_code,Fee_title,Act_amount,Fine,Discount,std_id) 
				                            values(@Seq_no,@loc_srl,@user_id,@class_code,@Fee_code,@Fee_title,@Act_amount,@Fine,@Discount,@std_id)	


                           end 
                       else
                          if not exists (select Seq_no from temp_collect where Seq_no=@Seq_no and Fee_code=@Fee_code)
                           begin
                          
                                set @loc_srl=(select isnull(max(srl_no),0)+1 from temp_collect where Seq_no=@Seq_no)
				                   
				                insert into temp_collect(Seq_no,Srl_no,u_id,class_code,Fee_code,Fee_title,Act_amount,Fine,Discount,std_id) 
				                            values(@Seq_no,@loc_srl,@user_id,@class_code,@Fee_code,@Fee_title,@Act_amount,@Fine,@Discount,@std_id)	 
               
                          end
               end

     if @mode='p' ----popup delete
        delete from  temp_collect where  Seq_no=@Seq_no and Srl_no=@Srl_no
    
		                          





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

