-----------------------To update
CREATE TABLE [StudentAdmission] (
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


CREATE TABLE [Ls_plan_details] (
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


CREATE TABLE [Ls_plan_Master] (
	[Srl_no] [decimal](10, 0) NOT NULL ,
	[Class_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Section_id] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Term_id] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Exam_id] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Sub_id] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO


CREATE TABLE [Ls_plan_topic] (
	[Srl_no] [int] NOT NULL ,
	[Topic_srl] [int] NOT NULL ,
	[Topic_title] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[LS_Week] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Entry_by] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Entry_date] [datetime] NOT NULL 
) ON [PRIMARY]
GO


CREATE TABLE [MarksCategory] (
	[MCategoryID] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MCategoryDsc] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Note] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EntryBy] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EntryDate] [datetime] NOT NULL 
) ON [PRIMARY]
GO


CREATE TABLE [Std_Study_performance] (
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


CREATE TABLE [StudentAttendanceLeaveInfo] (
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


CREATE TABLE [ClassRoutine] (
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


CREATE TABLE [ExamRoutine] (
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


CREATE TABLE [ExamSchedule] (
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


