if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Log_sub1_Access_Log]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Log_sub1] DROP CONSTRAINT FK_Log_sub1_Access_Log
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Job_Hist_Current_Branch_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Job_Hist_Current] DROP CONSTRAINT FK_Emp_Job_Hist_Current_Branch_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Job_Hist_Current_Department_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Job_Hist_Current] DROP CONSTRAINT FK_Emp_Job_Hist_Current_Department_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Att_Notes_Emp_Att_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Att_Notes] DROP CONSTRAINT FK_Emp_Att_Notes_Emp_Att_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Conveyance_Emp_Movement]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Conveyance] DROP CONSTRAINT FK_Conveyance_Emp_Movement
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Att_Info_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Att_Info] DROP CONSTRAINT FK_Emp_Att_Info_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Bank_Account_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Bank_Account] DROP CONSTRAINT FK_Emp_Bank_Account_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Discipline_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Discipline] DROP CONSTRAINT FK_Emp_Discipline_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Education_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Education] DROP CONSTRAINT FK_Emp_Education_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Job_Hist_Current_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Job_Hist_Current] DROP CONSTRAINT FK_Emp_Job_Hist_Current_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Job_Hist_Future_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Job_Hist_Future] DROP CONSTRAINT FK_Emp_Job_Hist_Future_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Leave_Info_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Leave_Info] DROP CONSTRAINT FK_Emp_Leave_Info_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Movement_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Movement] DROP CONSTRAINT FK_Emp_Movement_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Payscale_Hist_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Payscale_Hist] DROP CONSTRAINT FK_Emp_Payscale_Hist_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_performance_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_performance] DROP CONSTRAINT FK_Emp_performance_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Reference_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Reference] DROP CONSTRAINT FK_Emp_Reference_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Salary_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Salary] DROP CONSTRAINT FK_Emp_Salary_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Fixed_Pay_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Fixed_Pay] DROP CONSTRAINT FK_Fixed_Pay_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Increment_History_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Increment_History] DROP CONSTRAINT FK_Increment_History_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Move_schedule_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Move_schedule] DROP CONSTRAINT FK_Move_schedule_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Payroll_main_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Payroll_main] DROP CONSTRAINT FK_Payroll_main_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Worker_Group_Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Worker_Group] DROP CONSTRAINT FK_Worker_Group_Emp_Per_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Group_Shift_Group_NM]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Group_Shift] DROP CONSTRAINT FK_Group_Shift_Group_NM
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Worker_Group_Group_NM]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Worker_Group] DROP CONSTRAINT FK_Worker_Group_Group_NM
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Job_Hist_Current_Job_Title]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Job_Hist_Current] DROP CONSTRAINT FK_Emp_Job_Hist_Current_Job_Title
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Job_Hist_Current_Job_Type]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Job_Hist_Current] DROP CONSTRAINT FK_Emp_Job_Hist_Current_Job_Type
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Leave_Info_Leave_List]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Leave_Info] DROP CONSTRAINT FK_Emp_Leave_Info_Leave_List
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Fixed_Pay_Pay_Struc]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Fixed_Pay] DROP CONSTRAINT FK_Fixed_Pay_Pay_Struc
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Payroll_sub_Pay_Struc]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Payroll_sub] DROP CONSTRAINT FK_Payroll_sub_Pay_Struc
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Payscale_Sub_Pay_Struc]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Payscale_Sub] DROP CONSTRAINT FK_Payscale_Sub_Pay_Struc
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Taxable_Ceiling_Pay_Struc]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Taxable_Ceiling] DROP CONSTRAINT FK_Taxable_Ceiling_Pay_Struc
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Payroll_sub_Payroll_main]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Payroll_sub] DROP CONSTRAINT FK_Payroll_sub_Payroll_main
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Payscale_Hist_Payscale_Main]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Payscale_Hist] DROP CONSTRAINT FK_Emp_Payscale_Hist_Payscale_Main
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Increment_History_Payscale_Main]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Increment_History] DROP CONSTRAINT FK_Increment_History_Payscale_Main
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Payscale_Sub_Payscale_Main]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Payscale_Sub] DROP CONSTRAINT FK_Payscale_Sub_Payscale_Main
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Job_Hist_Current_Rank]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Job_Hist_Current] DROP CONSTRAINT FK_Emp_Job_Hist_Current_Rank
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Payscale_Main_Rank]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Payscale_Main] DROP CONSTRAINT FK_Payscale_Main_Rank
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Emp_Job_Hist_Current_Sec_Info]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Emp_Job_Hist_Current] DROP CONSTRAINT FK_Emp_Job_Hist_Current_Sec_Info
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Group_Shift_Shift]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Group_Shift] DROP CONSTRAINT FK_Group_Shift_Shift
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Log_sub1_soft_bag]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Log_sub1] DROP CONSTRAINT FK_Log_sub1_soft_bag
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Access_Log_soft_pass]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Access_Log] DROP CONSTRAINT FK_Access_Log_soft_pass
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_permit_soft_pass]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[permit] DROP CONSTRAINT FK_permit_soft_pass
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Access_Log]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Access_Log]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Approval]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Approval]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AuthorInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AuthorInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BonusPreparation]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BonusPreparation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookDistributionandReturnInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BookDistributionandReturnInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookIssueRefund]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BookIssueRefund]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookIssueRefundSub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BookIssueRefundSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookIssueRule]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BookIssueRule]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookList]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BookList]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BookRecievedInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BookRecievedInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Branch_Info]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Branch_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CafeItem]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CafeItem]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CafeSupplyer]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CafeSupplyer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClassInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClassInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClassRoutine]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClassRoutine]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Collec_details]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Collec_details]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Collec_master]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Collec_master]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Comp_Det_Info]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Comp_Det_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CompanyInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CompanyInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Company_Policy]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Company_Policy]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Conveyance]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Conveyance]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Country]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Country]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Department_Info]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Department_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[District]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[District]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DriverSchedule]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DriverSchedule]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[EmpIncrementInformationInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[EmpIncrementInformationInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Att_Info]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Att_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Att_Notes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Att_Notes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Bank_Account]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Bank_Account]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Discipline]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Discipline]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Education]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Education]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Job_End]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Job_End]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Job_Hist_Current]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Job_Hist_Current]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Job_Hist_Future]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Job_Hist_Future]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Leave_Info]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Leave_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Movement]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Movement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Payscale_Hist]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Payscale_Hist]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Per_Info]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Per_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Reference]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Reference]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_Salary]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_Salary]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Emp_performance]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Emp_performance]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExamGuardPlan]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ExamGuardPlan]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExamRoutine]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ExamRoutine]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExamSchedule]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ExamSchedule]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExamSitPlan]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ExamSitPlan]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ExamTypeInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ExamTypeInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Exam_setup]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Exam_setup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fee_info]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Fee_info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fee_setup]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Fee_setup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fixed_Pay]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Fixed_Pay]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Group_NM]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Group_NM]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Group_Shift]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Group_Shift]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Hol_List]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Hol_List]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Increment_History]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Increment_History]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Job_Title]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Job_Title]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Job_Type]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Job_Type]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Leave_List]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Leave_List]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LectureInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LectureInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LibraryBook]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LibraryBook]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LibraryBookAuthor]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LibraryBookAuthor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LibraryBookInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LibraryBookInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LibraryBookList]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LibraryBookList]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LibraryBookSub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LibraryBookSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LibrarySubjectInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LibrarySubjectInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LoanRefund]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LoanRefund]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LoanSanction_Info]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LoanSanction_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Loan_Type]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Loan_Type]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Log_sub1]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Log_sub1]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Login_Fail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Login_Fail]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Ls_plan_Master]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Ls_plan_Master]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Ls_plan_details]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Ls_plan_details]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Ls_plan_topic]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Ls_plan_topic]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MarksCategory]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MarksCategory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Move_Out_In]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Move_Out_In]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Move_schedule]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Move_schedule]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OT_Fix]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[OT_Fix]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Office_time]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Office_time]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Out_of_Office]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Out_of_Office]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Overtime]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Overtime]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Overtime_Preparation]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Overtime_Preparation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Param_Tbl]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Param_Tbl]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pay_Struc]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Pay_Struc]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Payroll_main]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Payroll_main]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Payroll_sub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Payroll_sub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Payscale_Main]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Payscale_Main]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Payscale_Sub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Payscale_Sub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Promotion_Record]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Promotion_Record]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PublisherInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PublisherInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rank]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Rank]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RequisetionSub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RequisetionSub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RequisitionInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RequisitionInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Results]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Results]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ScholerShipinfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ScholerShipinfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ScholershipType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ScholershipType]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sec_Info]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Sec_Info]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SectionInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SectionInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Shift]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Shift]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Std_Study_performance]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Std_Study_performance]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StudentAdmission]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StudentAdmission]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StudentAttendanceLeaveInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StudentAttendanceLeaveInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StudentEvaluation]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StudentEvaluation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StudentInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[StudentInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SubjectInfoMain]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SubjectInfoMain]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SubjectMarksDistribution]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SubjectMarksDistribution]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Subject_info_sub]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Subject_info_sub]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SupplierInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SupplierInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SyllabusPreperation]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SyllabusPreperation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TCInformation]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TCInformation]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TCTypeSetUp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TCTypeSetUp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Tax_Slab]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Tax_Slab]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Taxable_Ceiling]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Taxable_Ceiling]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TeacherInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TeacherInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Tier_setup]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Tier_setup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Time_Keeper]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Time_Keeper]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UserInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UserInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[VaxinInfo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[VaxinInfo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Vou]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Vou]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Worker_Group]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Worker_Group]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[permit]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[permit]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[scholerShipNameSetup]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[scholerShipNameSetup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[soft_bag]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[soft_bag]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[soft_pass]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[soft_pass]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[temp_collect]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[temp_collect]
GO

CREATE TABLE [dbo].[Access_Log] (
	[U_Id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Access_Id] [int] NOT NULL ,
	[LogIn] [datetime] NULL ,
	[Logout] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Approval] (
	[App_id] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Policy_Code] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Tier_1] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tier_1_Chk] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tier_1_Remarks] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tier_2] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tier_2_Chk] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tier_2_Remarks] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tier_3] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tier_3_Chk] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tier_3_Remarks] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tier_4] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tier_4_Chk] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tier_4_Remarks] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Final_Tier] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Final_Tier_Chk] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Final_Tier_Remarks] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Move_pos] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AuthorInfo] (
	[AuthorCode] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AuthorName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AuthorNote] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AuthorEntryBy] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AuthorEntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BonusPreparation] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PayMonth] [char] (9) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PayYear] [char] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Amount] [money] NOT NULL ,
	[EntryBy] [char] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BookDistributionandReturnInfo] (
	[TrackNo] [bigint] IDENTITY (1, 1) NOT NULL ,
	[StudentId] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Shift] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ClassId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SubjectId] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[BookRecievedBack] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryDate] [datetime] NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DeliveryApproved] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DeliveryApprovedBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DeliveryApprovedDate] [datetime] NULL ,
	[ReturnApproved] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ReturnApprovedBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ReturnApprovedDate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BookIssueRefund] (
	[StudentId] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BookId] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BookCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IssueDate] [datetime] NOT NULL ,
	[ExpReturnDate] [datetime] NOT NULL ,
	[ReqNo] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ReqDate] [datetime] NULL ,
	[ActualReturnDate] [datetime] NULL ,
	[DelayDay] [int] NOT NULL ,
	[FineAmt] [money] NOT NULL ,
	[Remarks] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BookIssueRefundSub] (
	[IssueId] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BookId] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IssueDate] [datetime] NOT NULL ,
	[ExpReturnDate] [datetime] NOT NULL ,
	[ActRetuenDate] [datetime] NULL ,
	[DelayDay] [int] NOT NULL ,
	[FineAmt] [money] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BookIssueRule] (
	[ClassCode] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MaxNumberOfBook] [int] NOT NULL ,
	[MaxDayOfUse] [int] NOT NULL ,
	[FineAmount] [money] NOT NULL ,
	[Remarks] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[IsRuEntryDate] [datetime] NULL ,
	[ISRuEntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BookList] (
	[ClassId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EYear] [int] NOT NULL ,
	[SubjectId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Book] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Writter] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entryby] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[entrydate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BookRecievedInfo] (
	[TrackId] [bigint] IDENTITY (1, 1) NOT NULL ,
	[RecieveNo] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SuppId] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[RecDate] [datetime] NOT NULL ,
	[ClassId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SubjectId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Qty] [money] NOT NULL ,
	[Notes] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryDate] [datetime] NOT NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Branch_Info] (
	[Title] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Address] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Telephone] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CafeItem] (
	[ItemCode] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ItemName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ItemType] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ReorderLevel] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CafeSupplyer] (
	[SupCode] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SupSupplyerName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SupplyerAdress] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SupDistrictCode] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SupCountryCode] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Remarks] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SupEntryDate] [datetime] NULL ,
	[PubEntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClassInfo] (
	[ClassID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ShIftname] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[StartTime] [datetime] NULL ,
	[EndTime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClassRoutine] (
	[ClassId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Shift] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ListOfday] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[subjectid] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Starttime] [datetime] NULL ,
	[EndTime] [datetime] NULL ,
	[TeacherId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryBy] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entrydate] [datetime] NOT NULL ,
	[academic_yr] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Collec_details] (
	[C_Srl] [int] NOT NULL ,
	[serial_no] [int] NOT NULL ,
	[Fee_code] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Act_Amount] [decimal](18, 2) NOT NULL ,
	[Discount] [decimal](18, 0) NOT NULL ,
	[Fine] [decimal](18, 0) NOT NULL ,
	[Entry_by] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entry_date] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Collec_master] (
	[C_Srl] [int] NOT NULL ,
	[Std_id] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Class_id] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Mon] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Yr] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Remark] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_date] [datetime] NOT NULL ,
	[Entry_by] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Collec_date] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Comp_Det_Info] (
	[Name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Type] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Address] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[City] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Country] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Phone1] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Phone2] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Fax] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Email] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Notes] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_Dt] [datetime] NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CompanyInfo] (
	[CompanyName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CompanyAddress] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Company_Policy] (
	[Policy_code] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Policy_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Policy_detail] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Conveyance] (
	[Move_Id] [int] NOT NULL ,
	[Transport] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Amount] [money] NULL ,
	[Dt] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Country] (
	[CntCoutryCode] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CntCountryName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CntCountryShortName] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CntEntryDate] [datetime] NULL ,
	[CntEntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Department_Info] (
	[Title] [varchar] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[District] (
	[DisDistrictCode] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DisDistrictName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DisDistrictShortName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DisCountryCode] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DisEntryDate] [datetime] NULL ,
	[DisEntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DriverSchedule] (
	[Emp_Id] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ScheduleMonth] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ScheduleYear] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ScheduleStTime] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ScheduleEndTime] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DayoftheSchedule] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entry_Dt] [datetime] NULL ,
	[Entry_By] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Track_Id] [int] IDENTITY (1, 1) NOT NULL ,
	[AbsentTime] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[EmpIncrementInformationInfo] (
	[Emp_ID] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[IncrementAmount] [int] NULL ,
	[LastIncrementDt] [datetime] NULL ,
	[NextIncremntDt] [datetime] NOT NULL ,
	[EffectiveDt] [datetime] NULL ,
	[Remarks] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_Dt] [datetime] NOT NULL ,
	[Entry_By] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Track_Id] [int] IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Att_Info] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Attn_Id] [int] IDENTITY (1, 1) NOT NULL ,
	[Emp_login] [datetime] NOT NULL ,
	[Hol_Flag] [int] NOT NULL ,
	[Emp_logout] [datetime] NULL ,
	[In_Status] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[In_Notes] [int] NOT NULL ,
	[Out_Status] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Out_Notes] [int] NOT NULL ,
	[Entry_Dt] [datetime] NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Att_Notes] (
	[Attn_Id] [int] NOT NULL ,
	[Note_Type] [int] NULL ,
	[Reason] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Description] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Track_Id] [int] IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Bank_Account] (
	[Emp_Id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Bank_Code] [int] NULL ,
	[Acc_No] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Discipline] (
	[Emp_Id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Track_Id] [int] IDENTITY (1, 1) NOT NULL ,
	[Ref_No] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Reason] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Penalty] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_Id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Education] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Edu_Count] [int] IDENTITY (1, 1) NOT NULL ,
	[Exam_Name] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[M_Subject] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Pass_Year] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Degree_From] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Result] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Job_End] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Job_end_dt] [datetime] NULL ,
	[Type] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Des] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_Dt] [datetime] NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Job_Hist_Current] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Emp_join_date] [datetime] NULL ,
	[Emp_rank] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_desig] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_branch] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_dept] [varchar] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_section] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_job_type] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Responsibility] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Pay_type] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Duty_type] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Super_ID] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL ,
	[Permdate] [datetime] NULL ,
	[CBasic] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Job_Hist_Future] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Effect_date] [datetime] NOT NULL ,
	[Emp_Rank] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_desig] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_branch] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_dept] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_section] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_job_type] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Leave_Info] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Leave_name] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Start_dt] [datetime] NULL ,
	[End_dt] [datetime] NULL ,
	[Des] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Address] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Tel1] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Super_ID] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NOT NULL ,
	[App_Id] [int] NOT NULL ,
	[App_Type] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[track_id] [int] IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Movement] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Move_Id] [int] IDENTITY (1, 1) NOT NULL ,
	[Mode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Place] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Move_Out_Dt] [datetime] NOT NULL ,
	[Exp_Rtn_Dt] [datetime] NULL ,
	[Rtn_Dt] [datetime] NULL ,
	[Cont_Tel] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Move_des] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Move_Status] [int] NULL ,
	[Entry_Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Payscale_Hist] (
	[Emp_Id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Track_Id] [int] IDENTITY (1, 1) NOT NULL ,
	[Scale_Code] [int] NULL ,
	[Effect_Dt] [datetime] NULL ,
	[U_Id] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Per_Info] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Emp_fna] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Emp_mna] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Emp_lna] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Emp_fa_na] [varchar] (45) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Emp_ma_na] [varchar] (45) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_d_of_b] [datetime] NOT NULL ,
	[Emp_age] [tinyint] NOT NULL ,
	[Blood_group] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Emp_marital_st] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Emp_gender] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[S_Security] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Voter_ID] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_nat] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Emp_sp_qta] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_religion] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_perm_add] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Emp_perm_town] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Emp_post1] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Emp_country1] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_tel1] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_fax1] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_email1] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Contact_person] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Contact_add] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Contact_town] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Contact_post] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Contact_tel] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Contact_fax] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Contact_email] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_eye] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_height] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_weight] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Emp_disable] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[pick_pic_path] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Job_Stat] [int] NOT NULL ,
	[Entry_Dt] [datetime] NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Reference] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Ref_Count] [int] IDENTITY (1, 1) NOT NULL ,
	[Ref_Name] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Ref_Occup] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Ref_Add] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Ref_town] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Ref_post] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Ref_Country] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Ref_tel] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Ref_fax] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Ref_email] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Ref_Relation] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_Salary] (
	[Emp_ID] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Basic_Amount] [int] NULL ,
	[Dt] [datetime] NULL ,
	[U_Id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Emp_performance] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Ev_dt_from] [datetime] NOT NULL ,
	[Ev_dt_to] [datetime] NOT NULL ,
	[Job] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Skill] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Comn] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Sincere] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Discip] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Cooper] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Laeader] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Motiv] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Plann] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Initiate] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Att] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[App] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ExamGuardPlan] (
	[Examdate] [datetime] NOT NULL ,
	[RoomNo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TeacherID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Responsibility] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[StartTime] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ExamRoutine] (
	[ExamYear] [int] NOT NULL ,
	[ExamID] [int] NOT NULL ,
	[SubjectID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Startdate] [datetime] NULL ,
	[MarksCataApplied] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CategoryID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ExamDate] [datetime] NOT NULL ,
	[ExamStartTime] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TotalMarks] [int] NULL ,
	[Note] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ExamSchedule] (
	[serial_no] [int] NOT NULL ,
	[ClassId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ExamYear] [int] NOT NULL ,
	[ExamId] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ExamTypeID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ExamStartDate] [datetime] NOT NULL ,
	[MarksCataApplied] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Note] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ExamSitPlan] (
	[ExamDate] [datetime] NOT NULL ,
	[RoomNo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[StartTime] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Shift] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[StartRoll] [int] NOT NULL ,
	[EndRoll] [int] NOT NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ExamTypeInfo] (
	[ETypeID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ETypeName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Note] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Exam_setup] (
	[Group_code] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Exam_code] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Exam_title] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Remarks] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_date] [datetime] NOT NULL ,
	[Entry_by] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Fee_info] (
	[Fee_code] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Fee_title] [varchar] (75) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entry_by] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entry_date] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Fee_setup] (
	[Srl_No] [int] NOT NULL ,
	[Class_id] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Fee_Code] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Acc_code] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Fee_amt] [decimal](18, 2) NOT NULL ,
	[Entry_date] [datetime] NOT NULL ,
	[Entry_by] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Fixed_Pay] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Head_code] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Amount] [money] NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL ,
	[AccountNo] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Group_NM] (
	[Gr_Code] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Gr_Name] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Group_Shift] (
	[Gr_Code] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Shift_Code] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Start_Dt] [datetime] NOT NULL ,
	[End_Dt] [datetime] NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Hol_List] (
	[Hol_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Hol_desc] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Str_date] [datetime] NOT NULL ,
	[End_date] [datetime] NOT NULL ,
	[Category] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Increment_History] (
	[Emp_Id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Scale_Code] [int] NULL ,
	[Num_Incr] [int] NULL ,
	[Effect_Dt] [datetime] NULL ,
	[Done] [int] NULL ,
	[Track_Id] [int] IDENTITY (1, 1) NOT NULL ,
	[Dt] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Job_Title] (
	[Title] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Job_Type] (
	[Title] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Leave_List] (
	[Leave_name] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Duration] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Des] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LectureInfo] (
	[LectureId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LectureDsc] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SubjectId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LectureDetail] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LecLessonPrepareBy] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LIOpenForStu] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NULL ,
	[EntryBy] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[LibraryBook] (
	[PurchaseId] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PurchaseDate] [datetime] NOT NULL ,
	[Remarks] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LibraryBookAuthor] (
	[BookCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AuthorId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Trackid] [int] IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LibraryBookInfo] (
	[BookCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BookName] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SubCode] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PubCode] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Remarks] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryDate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LibraryBookList] (
	[PurchaseId] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BookCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BookId] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[UnitRate] [money] NOT NULL ,
	[FineAmt] [money] NOT NULL ,
	[Demage] [bit] NOT NULL ,
	[TrackId] [int] IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LibraryBookSub] (
	[PurchaseId] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BookCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[UnitRate] [money] NOT NULL ,
	[Qty] [int] NOT NULL ,
	[FineAmt] [money] NOT NULL ,
	[TrackId] [int] IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LibrarySubjectInfo] (
	[SubSubjectCode] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SubSubjectName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SubNote] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SubEntryDate] [datetime] NULL ,
	[SubEntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LoanRefund] (
	[Emp_Id] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Loan_Type] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PaidInstament] [int] NULL ,
	[PaidAmount] [decimal](18, 0) NULL ,
	[Paid_Dt] [datetime] NOT NULL ,
	[TrackID] [int] IDENTITY (1, 1) NOT NULL ,
	[Entry_By] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entr_Dt] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LoanSanction_Info] (
	[Emp_id] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Loan_Type] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[InstalmentNo] [int] NOT NULL ,
	[GracePeriod] [int] NULL ,
	[SanctionAmount] [decimal](18, 0) NULL ,
	[InstallmentAmount] [decimal](18, 0) NULL ,
	[SanctionDate] [datetime] NULL ,
	[Entry_Dt] [datetime] NULL ,
	[Entry_By] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Track_Id] [int] IDENTITY (1, 1) NOT NULL ,
	[IstInstallmentDatePrin] [datetime] NOT NULL ,
	[IstInstallmentDateInt] [datetime] NOT NULL ,
	[InstallmentAmountInt] [decimal](18, 0) NULL ,
	[InterestInstallmetNo] [int] NULL ,
	[PaidInstallment] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Loan_Type] (
	[Loan_Type] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Loan_Description] [char] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_By] [char] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_Dt] [datetime] NULL ,
	[IntRate] [decimal](18, 2) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Log_sub1] (
	[Access_Id] [int] NULL ,
	[Track_Id] [int] IDENTITY (1, 1) NOT NULL ,
	[Access_Area] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_Time] [datetime] NULL ,
	[Exit_Time] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Login_Fail] (
	[u_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[u_name] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[user_pass] [varchar] (45) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[udt] [datetime] NULL ,
	[psl] [numeric](18, 0) IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Ls_plan_Master] (
	[Srl_no] [decimal](10, 0) NOT NULL ,
	[Class_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Section_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Term_id] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Exam_id] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Sub_id] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Ls_plan_details] (
	[Srl_no] [int] NOT NULL ,
	[Topic_srl] [int] NOT NULL ,
	[Details_srl] [int] NOT NULL ,
	[Ls_date] [datetime] NOT NULL ,
	[LS_Week] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[HW_CW] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Oral] [varchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Written] [varchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_date] [datetime] NOT NULL ,
	[Entry_by] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Ls_plan_topic] (
	[Srl_no] [int] NOT NULL ,
	[Topic_srl] [int] NOT NULL ,
	[Topic_title] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LS_Week] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_by] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entry_date] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MarksCategory] (
	[MCategoryID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MCategoryDsc] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Note] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Move_Out_In] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Mode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Place] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Move_Out_Dt] [datetime] NOT NULL ,
	[Exp_Rtn_Dt] [datetime] NULL ,
	[Rtn_Dt] [datetime] NULL ,
	[Cont_Tel] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Move_des] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Move_schedule] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Mode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Place] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Move_Out_Dt] [datetime] NOT NULL ,
	[Exp_Rtn_Dt] [datetime] NOT NULL ,
	[Cont_Tel] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Move_des] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entry_Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[OT_Fix] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Gen_Rate] [money] NOT NULL ,
	[Out_Rate] [money] NOT NULL ,
	[Hol_Rate] [money] NOT NULL ,
	[U_Id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Office_time] (
	[Start_time] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[End_time] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Relaxed] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Abs_time] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Sp_Start_time] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Sp_Start_Day] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Sp_End_time] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Sp_End_Day] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Effect_date] [datetime] NOT NULL ,
	[U_id] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL ,
	[Ref_Key] [numeric](18, 0) IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Out_of_Office] (
	[Emp_Id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Str_Date] [datetime] NOT NULL ,
	[End_Date] [datetime] NOT NULL ,
	[Description] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Track_Id] [int] IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Overtime] (
	[Emp_ID] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Pay_Month] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[pay_Year] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Gen_OT_Duration] [int] NOT NULL ,
	[Gen_OT_Pay] [money] NOT NULL ,
	[Out_OT_Duration] [int] NOT NULL ,
	[Out_OT_Pay] [money] NOT NULL ,
	[Hol_OT_Duration] [int] NOT NULL ,
	[Hol_OT_Pay] [money] NOT NULL ,
	[P_Status] [int] NOT NULL ,
	[Dt] [datetime] NOT NULL ,
	[U_Id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Overtime_Preparation] (
	[Emp_ID] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Date_of_OT] [datetime] NULL ,
	[Ot_St_Time] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OT_End_Time] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[OT_Amount] [decimal](18, 2) NULL ,
	[P_Status] [int] NOT NULL ,
	[NOOfHr] [decimal](18, 0) NULL ,
	[TraceID] [int] IDENTITY (1, 1) NOT NULL ,
	[Entry_Dt] [datetime] NULL ,
	[Entry_By] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Remarks] [varchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Param_Tbl] (
	[Module] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Policy_No] [int] NULL ,
	[Policy] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Flag] [int] NULL ,
	[value] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[description] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Pay_Struc] (
	[Head_code] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Head_name] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Operation] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Mode] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Payroll_main] (
	[Emp_Id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[pay_Month] [varchar] (9) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[pay_Year] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[A_Gen] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Pay_Type] [int] NULL ,
	[Pay_Stat] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Payroll_sub] (
	[A_Gen] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Head_Code] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Amount] [money] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Payscale_Main] (
	[Scale_Code] [int] NOT NULL ,
	[Desig_Code] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SB] [int] NULL ,
	[Incr] [int] NULL ,
	[EB] [int] NULL ,
	[EBIncr] [int] NULL ,
	[EBMax] [int] NULL ,
	[Dt] [datetime] NULL ,
	[U_Id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Payscale_Sub] (
	[Scale_Code] [int] NULL ,
	[Head_Code] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Calc_Type] [int] NULL ,
	[Prcnt_of_Basic] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Promotion_Record] (
	[Emp_Id] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Pre_Desig] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[C_Desig] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Pre_Dept] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[C_Dept] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Pre_Scale] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[C_Scale] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Last_Pro_Dt] [datetime] NULL ,
	[Pro_Effe_Dt] [datetime] NULL ,
	[Entry_Dt] [datetime] NOT NULL ,
	[Trace_ID] [int] IDENTITY (1, 1) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PublisherInfo] (
	[PubCode] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PubPublisherName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PublisherAdress] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PubCountryCode] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Remarks] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PubEntryDate] [datetime] NULL ,
	[PubEntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Rank] (
	[Title] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RequisetionSub] (
	[RequisitionId] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SubCode] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[BookCode] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TrackId] [int] IDENTITY (1, 1) NOT NULL ,
	[Status] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[RequisitionInfo] (
	[RequisitionId] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[RequisitionDate] [datetime] NOT NULL ,
	[StudentId] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Remarks] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Results] (
	[Emp_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Attn_Id] [int] NOT NULL ,
	[Emp_login] [datetime] NOT NULL ,
	[Hol_Flag] [int] NOT NULL ,
	[Emp_logout] [datetime] NULL ,
	[In_Status] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[In_Notes] [int] NOT NULL ,
	[Out_Status] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Out_Notes] [int] NOT NULL ,
	[Entry_Dt] [datetime] NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ScholerShipinfo] (
	[StudentId] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ScholerYear] [int] NOT NULL ,
	[ClassId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Type] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Scname] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[grade] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Amount] [float] NOT NULL ,
	[NoOfvalidYear] [int] NOT NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[entrydate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ScholershipType] (
	[SchTypeId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ScTypeName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Notes] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Sec_Info] (
	[Title] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Code] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Description] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SectionInfo] (
	[SectionID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Sectiondsc] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionRoomNo] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassTeacher] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SecMonitor1] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[secmonitor2] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryDate] [datetime] NULL ,
	[EntryBy] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Shift] (
	[Shift_Code] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Shift_Name] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Shift_Start] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Shift_End] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Delay] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Abs_time] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Std_Study_performance] (
	[Student_id] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Classid] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Sectionid] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Class_roll] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Srl_no] [int] NOT NULL ,
	[Topic_srl] [int] NOT NULL ,
	[Details_srl] [int] NOT NULL ,
	[Prfm] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Remarks] [char] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_by] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entry_date] [datetime] NOT NULL ,
	[Academic_yr] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[StudentAdmission] (
	[StudentId] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AdmissionDate] [datetime] NOT NULL ,
	[Shift] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassRoll] [int] NOT NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL ,
	[AdmitApproveBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[AdmitApproveDate] [datetime] NULL ,
	[Approval] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[AdmissionCancel] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[serial_no] [int] NULL ,
	[active_std] [tinyint] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[StudentAttendanceLeaveInfo] (
	[StudentID] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Shift] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassRoll] [int] NOT NULL ,
	[Present] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryTime] [datetime] NOT NULL ,
	[attn_date] [datetime] NOT NULL ,
	[PresentCancel] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CancelDate] [datetime] NULL ,
	[CancelBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CancelNote] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[StudentEvaluation] (
	[StudentID] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EvaluationDate] [datetime] NOT NULL ,
	[Shift] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassRoll] [int] NOT NULL ,
	[EntryBy] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entrydate] [datetime] NOT NULL ,
	[Active] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ApproveBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ApproveDate] [datetime] NULL ,
	[ActiveClass] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[StudentInfo] (
	[StudentID] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[StudentName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[StuFatherName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[StuMotherName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LegalGerdian] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StuMarraigeDate] [datetime] NULL ,
	[StuMoOrFaLet] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[StuBroNo] [tinyint] NULL ,
	[StuSisNo] [tinyint] NULL ,
	[StuCountryOfBirth] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StuReligion] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StuDateOfBirth] [datetime] NULL ,
	[Computer] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Internet] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StuHight] [float] NULL ,
	[StuWeight] [float] NULL ,
	[StuBloodGroup] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StuStreetPAddress] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StuPDistrict] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StuPCountry] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StuCStreetAddress] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StuCDistrict] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StuCCountry] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StuPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[StuEmail] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ImmAddress] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ImmPhone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ImmMob] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[NextvaccineDate] [datetime] NULL ,
	[StuEntryDate] [datetime] NOT NULL ,
	[StuEntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SubjectInfoMain] (
	[M_code] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[M_title] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SubjectUnit] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SubjectType] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SubjectMarksDistribution] (
	[ClassID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SubjectID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[term_code] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Exam_code] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CategoryID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SubMarks] [float] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Subject_info_sub] (
	[M_code] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Sub_code] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Class_code] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Sub_title] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Teacher_id] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entry_by] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entry_date] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SupplierInfo] (
	[SuppId] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Suppname] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SuppAddr] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Phone] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryDate] [datetime] NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[SyllabusPreperation] (
	[ClassId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Eyear] [int] NOT NULL ,
	[SubjectId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Syllabusdetail] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PreparedBy] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entryby] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entrydate] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[TCInformation] (
	[StudentID] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Shift] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SectionID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ClassRoll] [int] NOT NULL ,
	[TCTypeID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TCDate] [datetime] NOT NULL ,
	[TCNote] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL ,
	[Approved] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ApprovedBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ApprovedDate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TCTypeSetUp] (
	[TCID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TcName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Note] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entrydate] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Tax_Slab] (
	[Fin_year] [varchar] (9) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Slab_No] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Slab_Amount] [money] NULL ,
	[Slab_Percent] [float] NULL ,
	[Track_id] [int] IDENTITY (1, 1) NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Taxable_Ceiling] (
	[Head_Code] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Ceiling_Amount] [money] NULL ,
	[Track_id] [int] IDENTITY (1, 1) NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TeacherInfo] (
	[TeacherId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[TeacherName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Tier_setup] (
	[Code] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Policy_Code] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Tier_1] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Tier_2] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Tier_3] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Tier_4] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Final_Tier] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Time_Keeper] (
	[Boot_Up] [varchar] (11) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Latest_time] [varchar] (11) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[U_id] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Dt] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[UserInfo] (
	[UserID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[UserName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Address] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[UserCategory] [smallint] NULL ,
	[Phone] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Fax] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EMail] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[UserPass] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[UserStatus] [bit] NOT NULL ,
	[Remarks] [varchar] (75) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[VaxinInfo] (
	[StudentId] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[VaxinName] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vaxinDate] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Vou] (
	[vou_no] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vou_date] [datetime] NOT NULL ,
	[vou_slno] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[cost_code] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[cust_ord_no] [varchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vou_narr] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[vou_desc] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[acc_code] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[dollar] [money] NULL ,
	[rate] [money] NULL ,
	[dr_amt] [money] NULL ,
	[cr_amt] [money] NULL ,
	[vou_type] [char] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[vou_chq] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[track_id] [int] IDENTITY (1, 1) NOT NULL ,
	[uid] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Worker_Group] (
	[Gr_Code] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[emp_ID] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[U_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Dt] [datetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[permit] (
	[cnt] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[code] [numeric](18, 0) NOT NULL ,
	[u_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[scholerShipNameSetup] (
	[ScholerShipNameId] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SchType] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ScName] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[SchBy] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Address] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[DesOfSch] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entryby] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[soft_bag] (
	[code] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[software] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[scr_no] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[descript] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[portion] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[soft_pass] (
	[psl] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[u_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[u_name] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[user_pass] [varchar] (45) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[uid] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[udt] [datetime] NULL ,
	[cancel] [bit] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[temp_collect] (
	[Seq_no] [int] NOT NULL ,
	[Srl_no] [int] NOT NULL ,
	[u_id] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[class_code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Fee_code] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Fee_title] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Act_amount] [int] NOT NULL ,
	[Fine] [int] NOT NULL ,
	[Discount] [int] NOT NULL ,
	[std_id] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

