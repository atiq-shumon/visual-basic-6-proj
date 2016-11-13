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
	[active_std] [tinyint] NULL ,
	[aca_yr] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Student_status] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,


    to be added:

	[EffectDateForFee] [datetime] NULL 


) ON [PRIMARY]
GO










SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


ALTER    procedure StuAdmissionEvaluationInformation
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
        @Aca_yr				        varchar(6),
        @student_status				varchar(2),
	@EffectDateForFee 			datetime  
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
	active_std,
	student_status,
	
	
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
        @aca_yr,
        1,
	@student_status)
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
	ActiveClass		=@ActiveClass,
	EffectDateForFee	=@EffectDateForFee
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
	ActiveClass		,
	EffectDateForFee	
	
)
values
(
	@StudentID		,
	@AdmissionEvaluationDate,
	@Shift		,
	@ClassID	,
	@SectionID	,
	@ClassRoll	,
	@EntryBy	,
	@Entrydate	,
	@Active		,
	@ActiveClass	,
	@EffectDateForFee		    
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

