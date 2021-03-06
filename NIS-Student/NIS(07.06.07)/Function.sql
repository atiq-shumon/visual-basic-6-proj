if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AttendenceREportDeptWise]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[AttendenceREportDeptWise]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AttendenceREportEmployeeWise]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[AttendenceREportEmployeeWise]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetAbsDays]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetAbsDays]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetAbsentDays]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetAbsentDays]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetAbsent_Deduction]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetAbsent_Deduction]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetAttnNotes]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetAttnNotes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetAutoGen_No]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetAutoGen_No]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetAuto_SlNo]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetAuto_SlNo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetBonusAmount]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetBonusAmount]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetDailyTimeDiff]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetDailyTimeDiff]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetDailyWorkingHr]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetDailyWorkingHr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetDateMatched]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetDateMatched]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetDateOnly]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetDateOnly]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetDateRangeMatched]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetDateRangeMatched]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetDayName]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetDayName]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetDefaultTime]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetDefaultTime]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetFebruary_Days]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetFebruary_Days]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetHolDayFlag]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetHolDayFlag]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetHrInMovement]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetHrInMovement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetHrInMovementWLM]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetHrInMovementWLM]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetHrsInOffice]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetHrsInOffice]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetIf_In_Movement]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetIf_In_Movement]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetIncomeTax]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetIncomeTax]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetInitIncrAmount]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetInitIncrAmount]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetInitIncrNum]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetInitIncrNum]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetLateCount]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetLateCount]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetLatestIncrAmount]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetLatestIncrAmount]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetLatestIncrDate]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetLatestIncrDate]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetLatestIncrNum]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetLatestIncrNum]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetLeaveDays]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetLeaveDays]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetLeaveDuration]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetLeaveDuration]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetLeaveRangeMatched]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetLeaveRangeMatched]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetLeaveValidation_Mt]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetLeaveValidation_Mt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetLogOut]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetLogOut]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetLogOutOrNot]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetLogOutOrNot]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetMonthDays]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetMonthDays]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetMonthlyMovementHr]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetMonthlyMovementHr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetMonthlySalaryBrkUp]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetMonthlySalaryBrkUp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetMovementNum]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetMovementNum]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetNetBonusPay]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetNetBonusPay]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetNetPay]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetNetPay]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetNewID_DCL]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetNewID_DCL]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetNoMOnthDiff]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetNoMOnthDiff]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetNoMinutesOT]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetNoMinutesOT]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetOTAmount]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetOTAmount]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetOT_Hr]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetOT_Hr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetOffStayHr]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetOffStayHr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetOfficeDuration]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetOfficeDuration]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetOfficeEnd]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetOfficeEnd]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetOfficeStart]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetOfficeStart]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetOfficeStartWD]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetOfficeStartWD]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetOutRangeMatched]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetOutRangeMatched]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetParamFlag]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetParamFlag]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetParamValue]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetParamValue]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetPayType_PlcNo]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetPayType_PlcNo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetPresence]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetPresence]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetPresenceBonus]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetPresenceBonus]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetPresentBasic]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetPresentBasic]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetPresentOrNot]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetPresentOrNot]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetPresentSalary]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetPresentSalary]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetSalFixErrMsg]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetSalFixErrMsg]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetSalaryBrkUp]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetSalaryBrkUp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetSalaryHead]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetSalaryHead]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetStartingBasic]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetStartingBasic]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetStartingSalary]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetStartingSalary]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetStayInOffice]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetStayInOffice]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetTimeAdded]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetTimeAdded]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetTimeOnly]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetTimeOnly]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetTotAttn]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetTotAttn]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetTotLateAttn]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetTotLateAttn]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetTot_Holday]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetTot_Holday]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetTot_OtherHolday]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetTot_OtherHolday]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetTotalHartals]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetTotalHartals]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetTotalWeekEnds]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetTotalWeekEnds]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetVar_PreMonth]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetVar_PreMonth]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetVouAmount]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetVouAmount]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetWorkingDays]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetWorkingDays]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_If_In_Leave]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Get_If_In_Leave]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Get_pkeys]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Get_pkeys]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Month_Name]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Month_Name]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Month_No]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Month_No]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PadString]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[PadString]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RPT_MR]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[RPT_MR]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_Department_Wise_Attendance]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Rpt_Department_Wise_Attendance]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_marks_Sheet]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Rpt_marks_Sheet]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_marks_Sheet_all]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Rpt_marks_Sheet_all]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_marks_distribution]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Rpt_marks_distribution]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Rpt_statement_of_prog]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Rpt_statement_of_prog]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ShowMaxiamumEmpID]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[ShowMaxiamumEmpID]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

-----------select * from AttendenceREportDeptWise('Daffodil Software Ltd.','July',2005,'2005 jan 01')
CREATE   function AttendenceREportDeptWise(
	@Dept varchar (30),
	@PayMonth varchar(12),
	@PayYear int,
	@Today datetime)

RETURNS @TmpTable  Table (
	emp_id varchar(20),
	Emp_fna varchar(45),
        emp_desig varchar(25),
	emp_dept varchar(20),   
	Month_NM varchar(20),           
	Work_Day int,
	Tot_Hol int,
	Weekend int,
	Present int,
	Late int,
	Leave int,
	Absent int)
as
begin
	DECLARE 
	@emp_id varchar(20),
	@Emp_fna varchar(45),
        @emp_desig varchar(25),
	@emp_dept varchar(20),   
	@Month_NM varchar(20),           
	@Work_Day int,
	@Tot_Hol int,
	@Weekend int,
	@Present int,
	@Late int,
	@Leave int,
	@Absent int,
	@StartDate datetime,
	@EndingDate datetime,
	@MonthDays int

	DECLARE MyCursor CURSOR FOR
		SELECT Emp_id FROM Emp_Job_Hist_Current WHERE (Emp_dept = @Dept)
							
		Open MyCursor
				FETCH NEXT FROM MyCursor into @emp_id
				WHILE @@FETCH_STATUS = 0 
				Begin
					select @Emp_fna = Emp_Nm, @emp_desig = Emp_Desig, @emp_dept = Emp_Dept, @Month_NM = Month_Year, @Leave = Leave, @Tot_Hol = Hol_Day, @Weekend = Weekend, @Work_Day = Work_Day, @Late = Late, @Present = Attn, @Absent = Absent   from AttendenceREportEmployeeWise (@PayMonth, @PayYear, @emp_id, @Today) 
					INSERT INTO @TmpTable VALUES (@emp_id, @Emp_fna, @emp_desig, @emp_dept, @Month_NM, @Work_Day, @Tot_Hol, @Weekend, @Present, @Late, @Leave, @Absent)

				FETCH NEXT FROM MyCursor into @emp_id
			    End
	CLOSE MyCursor
	DEALLOCATE MyCursor

return
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

Create   function AttendenceREportEmployeeWise(
	@PayMonth varchar(12),
	@PayYear int,
	@Emp_id varchar(10),
	@Today datetime)

RETURNS @TmpTable  Table (
	Emp_id varchar(10)null,
	Emp_Nm varchar(50),
	Emp_Desig varchar(50),
	Emp_Dept varchar(50),
	[Date]varchar(20),
	Emp_login varchar(8)null,
	In_Status varchar(4)null,
	Notes  varchar(50)null,
	Emp_logout varchar(8)null,
	Leave int,
	Hol_Day int,
	Weekend int,
	Work_Day int,
	Late int,
	Attn int,
	Absent int,
	Out_Office int,
	Month_Year varchar(20),
	StartDate datetime,
	EndingDate datetime)
as
begin
	DECLARE 
	@Emp_Nm varchar(50),
	@Emp_Desig varchar(50),
	@Emp_Dept varchar(50),
	@AttenDate varchar(20),
	@Emp_login varchar(8),
	@In_Status varchar(4),
	@Notes  varchar(50),
	@Emp_logout varchar(8),
	@Leave int,
	@Hol_Day int,
	@Weekend int,
	@Work_Day int,
	@Late int,
	@Attn int,
	@Absent int,
	@Out_Office int,
	@Month_Year varchar(20),
	@StartDate datetime,
	@EndingDate datetime,
	@M_Days int,
	@LogonDate varchar (20)

select @StartDate = StartDate, @EndingDate = EndingDate  from getStartEndingPeriod (@PayMonth, @PayYear) 
set @M_Days = datediff(day, @StartDate, @EndingDate)

SELECT @Emp_Nm = Emp_Per_Info.Emp_fna + ' ' + Emp_Per_Info.Emp_mna + ' ' + Emp_Per_Info.Emp_lna, 
@Emp_Desig = Emp_Job_Hist_Current.Emp_desig, @Emp_Dept = Emp_Job_Hist_Current.Emp_dept
FROM Emp_Per_Info INNER JOIN Emp_Job_Hist_Current ON Emp_Per_Info.Emp_id = Emp_Job_Hist_Current.Emp_id WHERE     (Emp_Per_Info.Emp_id = @emp_id)

SET @Leave = dbo.TotalLeave(@Emp_ID,@PayMonth,@PayYear)
SET @Hol_Day = dbo.TotalHolday(@PayMonth,@PayYear)
SELECT @Weekend = count(*) from hol_list where Category ='Weekend' and Str_Date between @StartDate and @EndingDate
SET @Work_Day = @M_Days - @Hol_Day

WHILE @StartDate <= @EndingDate
	BEGIN
		SET @Emp_login = NULL
		SET @In_Status = NULL
		SET @Emp_logout = NULL
		SET @Notes = NULL
		SET @In_Status = NULL

		select @Emp_login=(convert(varchar(8),Emp_login,8))
		,@Emp_logout=convert(varchar(8),Emp_logout,8) 
		from emp_att_info where emp_id=@Emp_id --and emp_login = @StartDate
		AND SUBSTRING(CONVERT(VarChar(20), emp_login, 9), 1, 11) = @StartDate

		select @In_Status = Status from getLogonStatus (@Emp_ID, @StartDate) 
		set @Notes=(select Reason  from emp_Att_notes where Attn_ID=
			(select Attn_ID from emp_Att_Info where emp_Id = @Emp_ID AND SUBSTRING(CONVERT(VarChar(20), emp_login, 9), 1, 11) = @StartDate))
		IF @Notes IS NULL
			set @Notes= (select Category= 
				case Category when 'Weekend'then Category
			      	when 'Hartal'then Category
			      	else Hol_Name end
				from hol_list where Str_date = @StartDate)

		IF @Notes IS NULL 
			BEGIN
				IF EXISTS (SELECT * FROM Emp_Leave_Info
				WHERE (Emp_id = @Emp_ID) AND @StartDate BETWEEN Start_dt AND End_dt)
				SET @Notes = 'Leave'
			END
		ELSE
			SET @Notes = 'Absent'
		SET @LogonDate = SUBSTRING(CONVERT(VarChar(20), @StartDate, 3), 1, 11)
		INSERT INTO @TmpTable VALUES (@Emp_id, @Emp_Nm, @Emp_Desig, @Emp_Dept, @LogonDate, @Emp_login, @In_Status, @Notes, @Emp_logout, 0, 0, 0, 0, 0, 0, 0, 0, @Paymonth + ' ' + cast(@PayYear as varchar (4)), @StartDate, @EndingDate)
		
		SET @StartDate = DATEADD(DAY,1,@StartDate)
	END

	update @TmpTable set Notes='' 
	where rtrim(left([Date],2))>datepart(day,@Today)
	and datename(month,@Today)= @Paymonth
	and datename(year,@Today)=@PayYear
	----------------------------------------------------------------------
	update @TmpTable set Leave=@Leave,Hol_Day=@Hol_Day
	,Weekend=@Weekend,Work_Day=@Work_Day
	,Late=(select count(In_Status)from @TmpTable where In_Status='Late')
	,Attn =(select count(Emp_Login)from @TmpTable where Emp_Login !='')
	,Absent=(select count(Notes)from @TmpTable where Notes='Absent')

return
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






---select dbo.GetAbsDays('9001','may','2002')

create Function GetAbsDays(@Emp_id varchar (10),
					@pay_Month varchar(12),
						@pay_Year varchar(4))
Returns int

AS

Begin

declare @M_Days varchar(12)
, @Febru_days varchar(2)
, @Tot_Hol  int
, @Tot_Weekend int
, @Tot_Leave int
, @Tot_Abs int
, @Late int
, @Tot_Att int
, @Working_Days int



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

-------------------------Holidays---------------------------------------------
set @Tot_Hol = (select count(*) from hol_list where 
(datepart(year,Str_date)=@pay_Year and datename(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datename(month,End_date)=@Pay_month)
)
-------------------------Working Days-----------------------------------------
set @Working_Days=(@M_Days-@Tot_Hol)

-------------------------Leave------------------------------------------------
set @Tot_Leave = (select count(*) from Emp_Leave_Info where emp_Id=@Emp_Id and 
(datepart(year,Start_dt)=@pay_Year and datename(month,Start_dt)=@Pay_month)and
(datepart(year,End_dt)=@pay_Year and datename(month,End_dt)=@Pay_month)
)		

-------------------------Presence---------------------------------------------
Set @Tot_Att=(Select count (Emp_Id)from emp_att_info where Emp_ID=@Emp_id 
		and datepart(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@Pay_month)

--------------------------------------------------------------------------
/*
Late = (Select count(*) from emp_Att_info where In_status='1'and emp_id=@param 
		and datename(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@pay_Month)

*/


---- Absent days = Working days - (present days + leave) 
/*
dbo.GetAbsentDays includes late attendance ie. 1 day  absent 
for 3 late attendance
----------------------
dbo.GetAbsDays (This Fx.)does not consider late attendance.

*/
Set @Tot_Abs=(@Working_Days-(@Tot_Att+@Tot_Leave))

return @Tot_Abs

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





---select dbo.GetAbsentDays('9001','may','2002')

CREATE  Function GetAbsentDays(@Emp_id varchar (10),
					@pay_Month varchar(12),
						@pay_Year varchar(4))
Returns int

AS

Begin

declare @M_Days varchar(12)
, @Febru_days varchar(2)
, @Tot_Hol  int
, @Tot_Weekend int
, @Tot_Leave int
, @Tot_Abs int
, @Late int
, @Tot_Att int
, @Working_Days int



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

-------------------------Holidays---------------------------------------------
set @Tot_Hol = (select count(*) from hol_list where 
(datepart(year,Str_date)=@pay_Year and datename(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datename(month,End_date)=@Pay_month)
)
-------------------------Working Days-----------------------------------------
set @Working_Days=(@M_Days-@Tot_Hol)

-------------------------Leave------------------------------------------------
set @Tot_Leave = (select count(*) from Emp_Leave_Info where emp_Id=@Emp_Id and 
(datepart(year,Start_dt)=@pay_Year and datename(month,Start_dt)=@Pay_month)and
(datepart(year,End_dt)=@pay_Year and datename(month,End_dt)=@Pay_month)
)		

-------------------------Presence---------------------------------------------
Set @Tot_Att=(Select count (Emp_Id)from emp_att_info where Emp_ID=@Emp_id 
		and datepart(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@Pay_month)

-------------------------Late Presence-------------------------------------------
Set @Late = (Select count(*) from emp_Att_info where In_status='1'and emp_id=@Emp_id
		and datename(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@pay_Month)
--------------------------------------------------------------------------
---- Absent days = Working days - (present days + leave) 
---- Absent due to late (1 absent  for 3 late present)=late present/3
---- Total absent = Absent days + Absent due to late

Set @Tot_Abs=(@Working_Days-(@Tot_Att+@Tot_Leave))+ (@Late/3)

return @Tot_Abs

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




CREATE    Function GetAbsent_Deduction(@Emp_id varchar (10),
					@pay_Month varchar(12),
						@pay_Year varchar(4))
Returns money

AS

Begin


declare @M_Days varchar(12)
, @Febru_days varchar(2)
, @Tot_Hol  int
, @Tot_Weekend int
, @Tot_Leave int
, @Tot_Abs int
, @Late int
, @LAbs int
, @Tot_Att int
, @Working_Days int
, @Total_Fixed_Pay money
, @Abs_Deduct money
---------------------------------------------------------
,@Late_Flag int
,@Late_Value int
set @Late_Flag =(select dbo.GetParamFlag(62))
set @Late_Value =(select dbo.GetParamValue(62))
---------------------------------------------------------

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

-------------------------Holidays---------------------------------------------
set @Tot_Hol = (select count(*) from hol_list where 
(datepart(year,Str_date)=@pay_Year and datename(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datename(month,End_date)=@Pay_month)
)
-------------------------Weekend----------------------------------------------
/*
set @Tot_Weekend = (select count(*) from hol_list where Category ='Weekend'and
(datepart(year,Str_date)=@pay_Year and datename(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datename(month,End_date)=@Pay_month)
)
*/
-------------------------Working Days-----------------------------------------
set @Working_Days=(@M_Days-@Tot_Hol)

-------------------------Leave------------------------------------------------
set @Tot_Leave = (select count(*) from Emp_Leave_Info where emp_Id=@Emp_Id and 
(datepart(year,Start_dt)=@pay_Year and datename(month,Start_dt)=@Pay_month)and
(datepart(year,End_dt)=@pay_Year and datename(month,End_dt)=@Pay_month)
)		

-------------------------Presence---------------------------------------------
Set @Tot_Att=(Select count (Emp_Id)from emp_att_info where Emp_ID=@Emp_id 
		and datepart(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@Pay_month)


-------------------------Late Presence-------------------------------------------
Set @Late = (Select count(*) from emp_Att_info where In_status='1'and emp_id=@Emp_id
		and datename(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@pay_Month)

-------------------------Fixed Pay-------------------------------------------

set @Total_Fixed_Pay=(select sum(Amount)from fixed_pay where emp_Id=@emp_Id)

---------------------------------------------------------------------------
---- Absent days = Working days - (present days + leave) 
---- Absent due to late (1 absent  for 3 late present)=late present/3
---- Total absent = Absent days + Absent due to late

if @Late_Flag=1
	begin
		Set @Tot_Abs=(@Working_Days-(@Tot_Att+@Tot_Leave))+ (@Late/@Late_Value)
	end

if @Late_Flag=0
	begin
		Set @Tot_Abs=(@Working_Days-(@Tot_Att+@Tot_Leave))
	end


set @Abs_Deduct =(@Total_Fixed_Pay/convert(int,@M_Days))* @Tot_Abs

set @Abs_Deduct=ceiling(@Abs_Deduct)
return @Abs_Deduct

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


/*
select * from Emp_Att_Notes
select * from emp_att_info where emp_id='9-015'and in_status=1
insert into Emp_Att_Notes(Attn_Id,Note_Type,Reason,Description)
values(362,0,'Traffic Jam','')

2002-05-18 09:05:05.000
*/

/*

select dbo.GetAttnNotes('2002-05-06 10:47:47.180')
*/

CREATE  Function GetAttnNotes (@LoginDt datetime)

returns Varchar(150)

AS

BEGIN
	Declare @Notes varchar(150)

	Set @Notes=(select Reason+':'+ [Description] from Emp_Att_Notes where Attn_Id=
		(Select Attn_Id from Emp_Att_Info where Emp_LogIn=@LoginDt))

Return @Notes

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





--select * from test
--update test set A_Gen='SL002062'
/*
update Pay_type_ind set FB_No='FB000001'
select * from Pay_type_ind

select dbo.GetAutoGen_No(21)

*/

CREATE    Function GetAutoGen_No(@Category int)
returns varchar(8)
AS
Begin

declare @Pay_Type char(2)
,@Latest_Num varchar(10)
,@Last_Num int
,@Num_Len int	

if @Category is null set  @Pay_Type='SL'
if @Category =0 set  @Pay_Type='SL'

/*
select 
Head_code, Head_name from pay_struc where mode='B'
------------------------------------
15        Eid Bonus
16        AGM
17        BATEXPO
18        Election Bonus
19        Travel/tour Allowance
20        Incentive 
*/
if @Category =15 set  @Pay_Type='ED'
if @Category =16 set  @Pay_Type='GM'
if @Category =17 set  @Pay_Type='BX'
if @Category =18 set  @Pay_Type='EL'
if @Category =19 set  @Pay_Type='TR'
if @Category =20 set  @Pay_Type='IN'
if @Category =21 set  @Pay_Type='OT'

		
	set @Last_Num =(select distinct max(convert(int,right(A_Gen,6)))
				from Payroll_Main where A_Gen like @Pay_Type + '%')


		if @Last_Num is null or @Last_Num=''
			begin
				set @Last_Num=1
				set @Latest_Num = @Pay_Type+'00000'+ convert(char(1),@Last_Num)
			end
		else 	
			Begin		
				set @Last_Num=@Last_Num+1
				
				set @Num_Len=len(@Last_Num)

				if @Num_Len=1
					set @Latest_Num = @Pay_Type+'00000'+ convert(char(1),@Last_Num)
				
				if @Num_Len=2
					set @Latest_Num = @Pay_Type+'0000'+ convert(char(2),@Last_Num)

				if @Num_Len=3
					set @Latest_Num = @Pay_Type+'000'+ convert(char(3),@Last_Num)

				if @Num_Len=4
					set @Latest_Num = @Pay_Type+'00'+ convert(char(4),@Last_Num)

				if @Num_Len=5
					set @Latest_Num = @Pay_Type+'0'+ convert(char(5),@Last_Num)	

				if @Num_Len=6
					set @Latest_Num = @Pay_Type+ convert(char(6),@Last_Num)

			end


			set @Latest_Num=ltrim(@Latest_Num)

Return @Latest_Num

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



/*
select dbo.GetAuto_SlNo('Payscale')
select dbo.GetAuto_SlNo('Payroll')
select dbo.GetAuto_SlNo('Acs_Log')
select dbo.GetAuto_SlNo('Leave')
*/
CREATE     Function GetAuto_SlNo(@Tbl varchar(10))

returns int

AS
begin

declare @Last_Num int
--------------------------------------------------------------------------------
	if @Tbl='Payroll'
	
	begin	

		set @Last_Num = (select max(A_Gen)from payroll_main)
	
		if @Last_Num is null 		----There is no entry at all
			set @Last_Num=1
		else 	
			set @Last_Num=@Last_Num+1
	end
--------------------------------------------------------------------------------
	
	if @Tbl='Payscale'
	
	begin	

		set @Last_Num = (select max(Scale_Code)from payscale_main)
	
		if @Last_Num is null 		----There is no entry at all
			set @Last_Num=1
		else 	
			set @Last_Num=@Last_Num+1
	end
----------------------------------------------------------------------------
	if @Tbl='Acs_Log'
	
	begin	

		set @Last_Num = (select max(Access_Id)from Access_Log)
	
		if @Last_Num is null 		----There is no entry at all
			set @Last_Num=1
		else 	
			set @Last_Num=@Last_Num+1
	end

----------------------------------------------------------------------------
	if @Tbl='Leave'
	
	begin	

		set @Last_Num = (select max(App_Id)from Emp_Leave_Info)
	
		if @Last_Num is null 		----There is no entry at all
			set @Last_Num=1
		else 	
			set @Last_Num=@Last_Num+1
	end

Return @Last_Num
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

select dbo.GetPresentBasic ('9-027')

Select dbo.GetBonusAmount('9-027',16)

*/

create   function GetBonusAmount(@Emp_Id varchar(10),@Pay_Type int)

returns Money

AS

BEGIN
Declare 
@Base_Flag int
,@Prcnt float
,@Bonus_Amount money

set @Base_Flag =(select dbo.GetParamFlag (dbo.GetPayType_PlcNo(@Pay_Type,0)))
set @Prcnt = convert(float,(select dbo.GetParamValue (dbo.GetPayType_PlcNo(@Pay_Type,0))))

if @Base_Flag=0		--Percent of current basic salary
	set @Bonus_Amount=(select dbo.GetPresentBasic (@Emp_Id))*(@Prcnt/100)

if @Base_Flag=1		--Percent of current gross salary
	set @Bonus_Amount=(select dbo.GetPresentSalary(@Emp_Id))*(@Prcnt/100)


Return @Bonus_Amount
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

/*
select dbo.GetDailyTimeDiff('December','2005','BS-008')
*/

CREATE Function GetDailyTimeDiff(
		@Pay_Month varchar(12),
		@Pay_Year varchar(4),
		@EmpId varchar(10)
)

Returns Float

AS

BEGIN

Declare @DailyWorkHr float
declare 

	 @Start_time varchar(8)
	,@End_time varchar(8)
	,@Sp_Start_time varchar(10)
	,@Sp_Start_Day varchar(10)
	,@Sp_End_time varchar(10)
	,@Sp_End_Day varchar(10)
	,@Min varchar(2)
	,@EMin varchar(4)
--------------------------

	SELECT  @Start_time=Emp_login, @End_time=Emp_logout FROM Emp_Att_Info
				where emp_id=@EmpId



set @Start_time=rtrim(ltrim(@Start_time))
set @Min=(select substring(@Start_time,4,2))			----Cut minutes

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

Set @End_time=rtrim(ltrim(@End_time))
set @EMin=(select substring(@End_time,3,3))			----Cut minutes
								----Cut and add 12 with hour
if (select substring(@End_time,7,1))='P'and (select substring(@End_time,1,2))!='12' 			 
	Set @End_time=convert(char,(convert(int,substring(Rtrim(@End_time),1,2))+12))
else
	Set @End_time=convert(char,substring(@End_time,1,2))
set @End_time=rtrim(@End_time)+ +@EMin+':00'			----Add seconds
set @End_time=ltrim(rtrim(@End_time))


---set @DailyWorkHr=round(datediff(n,@Start_time,@End_time)*0.01667,1)

RETURN @DailyWorkHr 

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

/*

select AA=round(datediff(n,'09:00:00','18:30:00')*0.01667,1)

select dbo.GetDailyWorkingHr('november','2002')

*/

CREATE Function GetDailyWorkingHr (@Pay_Month varchar(12),@Pay_Year varchar(4))

Returns Float

AS

BEGIN

Declare @DailyWorkHr float
declare 

@Start_time varchar(8)
,@End_time varchar(8)

,@Sp_Start_time varchar(10)
,@Sp_Start_Day varchar(10)
,@Sp_End_time varchar(10)
,@Sp_End_Day varchar(10)

--------------------------
,@Min varchar(2)
,@EMin varchar(4)
--------------------------

select @Start_time=Start_time,@End_time=End_time

--		,@Sp_Start_time=Sp_Start_time,@Sp_Start_day=Sp_Start_day
	--	,@Sp_end_time=Sp_end_time,@Sp_End_day=Sp_End_day

from office_time where Effect_date=(
	select Mdt=(select max(Effect_date)from office_time 
		where datepart(Month,Effect_date)<=(select dbo.Month_No(@Pay_Month))
		and datepart(year,Effect_date)<=@pay_year))


set @Start_time=rtrim(ltrim(@Start_time))
set @Min=(select substring(@Start_time,4,2))			----Cut minutes

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

Set @End_time=rtrim(ltrim(@End_time))
set @EMin=(select substring(@End_time,3,3))			----Cut minutes
								----Cut and add 12 with hour
if (select substring(@End_time,7,1))='P'and (select substring(@End_time,1,2))!='12' 			 
	Set @End_time=convert(char,(convert(int,substring(Rtrim(@End_time),1,2))+12))
else
	Set @End_time=convert(char,substring(@End_time,1,2))
set @End_time=rtrim(@End_time)+ +@EMin+':00'			----Add seconds
set @End_time=ltrim(rtrim(@End_time))


set @DailyWorkHr=round(datediff(n,@Start_time,@End_time)*0.01667,1)

RETURN @DailyWorkHr 

END

---select * from office_time

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

this function gets two date values to compare them.
if dates are identical then returns 1 else 0

select dbo.GetDateMatched('2002-09-22','2002-09-20')
select dbo.GetDateMatched('2002-09-22','2002-09-22')

*/

create Function GetDateMatched(@Date1 Datetime,@Date2 Datetime)

Returns Int

AS	

BEGIN

Declare @Return_Value int

if datepart(day,@Date1)=datepart(day,@Date2)
 and datepart(month,@Date1)=datepart(month,@Date2)
 and datepart(Year,@Date1)=datepart(year,@Date2)

	set @Return_Value=1

else 

	set @Return_Value=0


return @Return_Value

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



CREATE  Function GetDateOnly(@Date datetime)

returns datetime

as
begin
declare @Rtn_Dt datetime
,@Conv_dt char(6)
,@Modi_dt char(10)

	set @Conv_dt=(select convert(char(6),@Date,12))
	set @Modi_dt='20'+substring(@Conv_dt,1,2)+'-'
			+substring(@Conv_dt,3,2)+'-'
			+substring(@Conv_dt,5,2)

	set @Rtn_Dt=convert(datetime,@Modi_dt)

return @Rtn_Dt

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

this function gets three date values to find
if the first date lies between the second and the third one.
if positive then returns 1 else 0

GetDateRangeMatched(Seach date',Date from,Date to)

select dbo.GetDateRangeMatched('2002-09-04','2002-09-05','2002-09-06')
select dbo.GetDateRangeMatched(getdate(),'2002-09-04','2002-09-05')

*/

create  Function GetDateRangeMatched(@Date0 Datetime,@Date1 Datetime,@Date2 Datetime)

Returns Int

AS	

BEGIN

Declare @Return_Value int

if @Date0 between @Date1 and @Date2

	set @Return_Value=1

else 

	set @Return_Value=0


return @Return_Value

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

/*
select dbo.GetDayName ('2002-07-25')
*/

CREATE Function GetDayName(@Dt Datetime)
Returns Varchar(9)
AS
BEGIN

DECLARE @Td varchar(1)
,@Today varchar(10)

SET @Td=DATEPART(dw,@Dt)

Set @Today=(SELECT [Day Name] =
	CASE @Td
	      	WHEN '1' THEN 'Sunday'
     	 	WHEN '2' THEN 'Monday'
		WHEN '3' THEN 'Tuesday'
      		WHEN '4' THEN 'Wednesday'
		WHEN '5' THEN 'Thursday'
		WHEN '6' THEN 'Friday'
		WHEN '7' THEN 'Saturday' END)

set @Today=ltrim(rtrim(@Today))

Return @Today

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

/*
select dbo.GetDefaultTime('2002-05-13 09:02:06.000')

*/

CREATE Function GetDefaultTime(@Dt datetime)
returns datetime

AS

BEGIN
declare @strDefault_Dt_Time varchar(26)
,@Default_Dt_Time datetime
declare @MM int,@YY int,@DD int

declare @strMM char(2),@strDD char(2)

SET @YY=(select datepart(Year,@Dt))
SET @MM=(select datepart(month,@Dt))
SET @DD=(select datepart(Day,@Dt))

	if len(@MM)=1
		set @strMM='0'+convert(char(1),@MM)
	else	
		set @strMM=convert(char(2),@MM)

	if len(@DD)=1
		set @strDD='0'+convert(char(1),@DD)
	else
		set @strDD=convert(char(2),@DD)

								
	set @strDefault_Dt_Time=convert(char(4),@YY)+'-'+ @strMM+'-'+ @strDD+space(1)+'13:00:00'

	set @Default_Dt_Time=convert(datetime,@strDefault_Dt_Time)

Return @Default_Dt_Time


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



CREATE    Function GetFebruary_Days(@Year int)
Returns varchar(2) 
AS
Begin


Declare	@Days varchar(2)
	

	if @Year % 4 = 0 
		Begin
			if @Year % 100 <> 0 or @Year % 400 = 0 
				set @Days = '29'
			else
				set @Days = '28'
		End
	
	else
				
	set @Days = '28'

return @Days 

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

/*

select * from emp_att_info
select * from hol_list


select dbo.GetHolDayFlag(getdate()) 
select dbo.GetHolDayFlag('2002-08-31') 

when some one present at office on a 
holiday a flag will automatically set as following:

Weekend ----> 1
Government Holiday----> 2
Public Holiday----> 3
Other Holiday----> 4

---------------------------------
Working day----->0

*/


create function GetHolDayFlag (@Today datetime)

returns int

AS

Begin

declare @Flag int

if exists (select * from hol_list where 
convert(char(8),Str_Date,1)=convert(char(8),@Today,1))

Begin 

	select @Flag = 
	case Category
		when 'Weekend'then 1
		when 'Government Holiday'then 2
		when 'Public Holiday'then 3
		else 4
	end
from hol_list where 
convert(char(8),Str_Date,1)=convert(char(8),@Today,1)

End

else
	set @Flag=0

return @Flag

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



/*
	select Move_Out_Dt,Entry_Dt from Emp_Movement where emp_id='9-015'
	and datepart(month,Move_Out_Dt)=5 and datepart(year,Move_Out_Dt)=2002
	and datepart(day,Move_Out_Dt)=13

select dbo.GetHrInMovement('9-015','2002-05-7')

*/

CREATE    Function GetHrInMovement(@Emp_Id varchar(10),@Dt datetime)  
returns char(3)
AS

BEGIN


Declare @Hr Float
,@Move_Out datetime
,@Move_Back datetime
,@Office_End varchar(8) 
,@MDt varchar(24)

set @Hr=0

---------------------------------------------------------------------------------

declare Movement_Cursor Cursor for

select Move_Out_Dt,Rtn_Dt from Emp_Movement where emp_id=@emp_id
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
			
			Set @Hr=@Hr+(Select convert(char(4),(ROUND(convert(float,DATEDIFF(n,@Move_Out,@Move_Back)*0.01666667),1))))

			--Set @Hr=@Hr+(select DATEDIFF(n,@Move_Out,@Move_Back)* 0.01667)
			
			fetch next from Movement_Cursor into @Move_Out,@Move_Back

	END


CLOSE Movement_Cursor
DEALLOCATE Movement_Cursor

Return convert(char(3),@Hr)

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



/*
	select Move_Out_Dt,Entry_Dt from Emp_Movement where emp_id='9-015'
	and datepart(month,Move_Out_Dt)=5 and datepart(year,Move_Out_Dt)=2002
	and datepart(day,Move_Out_Dt)=13

select dbo.GetHrInMovement('9-015','2002-05-7')
select dbo.GetHrInMovementWLM('9-015','2002-05-7')

*/

Create  Function GetHrInMovementWLM(@Emp_Id varchar(10),@Dt datetime)  
returns char(3)
AS

BEGIN


Declare @Hr Float
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

					and Move_Id !=(Select max(Move_Id)from Emp_Movement where emp_id=@emp_id
						and datepart(Year,Move_Out_Dt)=datepart(Year,@Dt)
							and datepart(month,Move_Out_Dt)=datepart(month,@Dt)
								and datepart(Day,Move_Out_Dt)=datepart(Day,@Dt))

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
			
			Set @Hr=@Hr+(Select convert(char(4),(ROUND(convert(float,DATEDIFF(n,@Move_Out,@Move_Back)*0.01666667),1))))

			--Set @Hr=@Hr+(select DATEDIFF(n,@Move_Out,@Move_Back)* 0.01667)
			
			fetch next from Movement_Cursor into @Move_Out,@Move_Back

	END


CLOSE Movement_Cursor
DEALLOCATE Movement_Cursor

Return convert(char(3),@Hr)

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



/*

select dbo.GetHrInMovement('9-015','2002-05-07 07:37:33.113')
select dbo.GetHrInMovementWLM('9-015','2002-05-07 07:37:33.113')

select dbo.GetHrsInOffice('9-015','2002-05-23 08:49:23.603')

select dbo.GetLogOut('9-015','2002-05-07 07:37:33.113')


select * from emp_att_info where emp_id='9-015'
and datepart(month,emp_login)=5 

*/
CREATE   Function GetHrsInOffice(@Emp_Id varchar(10),@Emp_LogIn datetime)

Returns char(4)
AS

BEGIN

declare @Hr_Stay_Before_Last_Movement float
,@Hr_on_Movement_Without_Last_Movement float
,@Hr_Stay char(4)



set @Hr_Stay_Before_Last_Movement=(Select convert(char(4),(ROUND(convert(float,DATEDIFF(n,@Emp_LogIn,(select dbo.GetLogOut(@Emp_Id,@Emp_LogIn)))*0.01666667),1))))

set @Hr_on_Movement_Without_Last_Movement=(select dbo.GetHrInMovementWLM(@Emp_Id,@Emp_LogIn))


Set @Hr_Stay=(@Hr_Stay_Before_Last_Movement-@Hr_on_Movement_Without_Last_Movement)

Return @Hr_Stay

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

/*
select dbo.GetIf_In_Movement('9-004','2002-08-10') 

*/

CREATE Function GetIf_In_Movement (@Emp_Id varchar(10),@Dt datetime)

Returns int
Begin
Declare @Flag int

	if exists (select * from emp_movement where emp_id=@emp_id 
		and Move_Status=1
		and convert(char(8),Move_Out_Dt,1)= convert(char(8),@Dt,1))

			set @Flag=1
		else
			set @Flag=0

return @Flag
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
select dbo.GetIncomeTax (100000,50000,150000,10,18,25,123596)

*/

create Function GetIncomeTax(@Slab1 money,@Slab2 money,@Slab3 money
	,@TxP2 float,@TxP3 float,@TxP4 float,@Sal money)

Returns money

AS


BEGIN
--------------------------------------------
declare @Tax money
--------------------------------------------

if @Sal>@Slab1
	BEGIN
		set @Sal=@Sal-@Slab1
	
		if @Sal>0
			BEGIN
				if @Sal<@Slab2
					begin
						set @Tax=@Sal*(@TxP2/100)
					end
				else
					begin
						set @Tax=@Slab2*(@TxP2/100)
					end
		
				set @Sal=@Sal-@Slab2
			END
	END
----------------------------------------------------------------------------------------
 
if @Sal>0
	BEGIN
		if @Sal<@Slab3
			begin
				set @Tax=@Tax+(@Sal*(@TxP3/100))
			end
		else
			begin
				set @Tax=@Tax+(@Slab3*(@TxP3/100))
			end
		set @Sal=@Sal-@Slab3
	END
----------------------------------------------------------------------------------------

if @Sal>0
	BEGIN	
		set @Tax=@Tax+(@Sal*(@TxP4/100))
	END

----------------------------------------------------------------------------------------
--set @Tax=round(@Tax,0)
set @Tax=floor(@Tax)
Return @Tax

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



--select * from increment_history
--select * from payscale_main
--  select dbo.GetInitIncrAmount ('9-027')

Create Function GetInitIncrAmount (@Emp_id varchar(10))

Returns money

AS

begin


declare @SB money
,@Incr int
,@Num_Incr int
,@Increment_Amount money

------------------------------------------------------------------------
select @SB=SB,@Incr=Incr from payscale_main where Scale_Code=
	(select Scale_Code from increment_history where Emp_Id=@Emp_id and Track_Id=
		(select min(Track_Id) from increment_history where Emp_Id=@Emp_Id))
------------------------------------------------------------------------
set @Num_Incr=(select Num_Incr from increment_history where Emp_Id=@Emp_id and Track_Id=
		(select min(Track_Id) from increment_history where Emp_Id=@Emp_Id))
------------------------------------------------------------------------

set @Increment_Amount=(@Incr*@Num_Incr)

return @Increment_Amount

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



/*

select dbo.GetInitIncrNum ('9-005')
select * from increment_history where emp_id='9-005'

*/

create   Function GetInitIncrNum(@Emp_Id varchar(10))

Returns int

AS

begin


declare @Num_Incr int


-------------------------------------------------------------------------------
set @Num_Incr=(select Num_Incr from increment_history where Emp_Id=@Emp_id and Track_Id=
		(select min(Track_Id) from increment_history where Emp_Id=@Emp_Id))
-------------------------------------------------------------------------------


return @Num_Incr

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






---select dbo.GetAbsentDays('9001','may','2002')

Create   Function GetLateCount(@Emp_id varchar (10),
					@pay_Month varchar(12),
						@pay_Year varchar(4))
Returns int

AS

Begin

declare @Late int
-------------------------Late Presence-------------------------------------------
Set @Late = (Select count(*) from emp_Att_info where In_status='1'and emp_id=@Emp_id
		and datename(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@pay_Month)
--------------------------------------------------------------------------

return @Late

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



--select * from increment_history
--select * from payscale_main
--  select dbo.GetLatestIncrAmount ('9-001')

Create Function GetLatestIncrAmount (@Emp_id varchar(10))

Returns money

AS

begin


declare @SB money
,@Incr int
,@Num_Incr int
,@Increment_Amount money

------------------------------------------------------------------------
select @SB=SB,@Incr=Incr from payscale_main where Scale_Code=
	(select Scale_Code from increment_history where Emp_Id=@Emp_id and Track_Id=
		(select min(Track_Id) from increment_history where Emp_Id=@Emp_Id))
------------------------------------------------------------------------
set @Num_Incr=(select Num_Incr from increment_history where Emp_Id=@Emp_id and Track_Id=
		(select max(Track_Id) from increment_history where Emp_Id=@Emp_Id))
------------------------------------------------------------------------

set @Increment_Amount=(@Incr*@Num_Incr)

return @Increment_Amount

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



--select * from increment_history
--select * from payscale_main

--  select dbo.GetLatestIncrDate ('9-001')

Create Function GetLatestIncrDate(@Emp_id varchar(10))

Returns datetime

AS

begin


declare @Last_Incr_Dt datetime

------------------------------------------------------------------------
set @Last_Incr_Dt=(select Effect_Dt from increment_history where Emp_Id=@Emp_id and Track_Id=
		(select max(Track_Id) from increment_history where Emp_Id=@Emp_Id))
------------------------------------------------------------------------

return @Last_Incr_Dt

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



/*

select dbo.GetLatestIncrNum ('9-005')
select * from increment_history where emp_id='9-005'

*/

create   Function GetLatestIncrNum(@Emp_Id varchar(10))

Returns int

AS

begin


declare @Num_Incr int


-------------------------------------------------------------------------------
set @Num_Incr=(select Num_Incr from increment_history where Emp_Id=@Emp_id and Track_Id=
		(select max(Track_Id) from increment_history where Emp_Id=@Emp_Id))
-------------------------------------------------------------------------------


return @Num_Incr

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




---select dbo.GetAbsentDays('9001','may','2002')

create  Function GetLeaveDays(@Emp_id varchar (10),
					@pay_Month varchar(12),
						@pay_Year varchar(4))
Returns int

AS

Begin

declare @Tot_Leave int


set @Tot_Leave = (select count(*) from Emp_Leave_Info where emp_Id=@Emp_Id and 
(datepart(year,Start_dt)=@pay_Year and datename(month,Start_dt)=@Pay_month)and
(datepart(year,End_dt)=@pay_Year and datename(month,End_dt)=@Pay_month)
)		
return @Tot_Leave

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




/*
select * from emp_leave_info

select dbo.GetLeaveDuration('9-001','August','2002')

*/

CReate  Function GetLeaveDuration(@Emp_id varchar(10)
		,@Pay_Month varchar(12),@Pay_Year varchar(4))

Returns Int

AS	

BEGIN

Declare @Total_Leave int
,@Duration int
,@Start_Dt datetime
,@End_Dt datetime

set @Total_Leave=0


Declare Leave_Cur cursor for

select Start_Dt,End_Dt from emp_leave_info where emp_id=@emp_id
and datename(month,Start_Dt)=@Pay_Month 
and datename(year,Start_Dt)=@Pay_Year 
--------------------------------------------------------------------
open Leave_Cur 

	
	fetch next from Leave_Cur into @Start_Dt,@End_Dt

	while @@fetch_status=0
	begin


		set @Duration=datediff(day,@Start_Dt,@End_Dt)+1
			if @Duration is null set @Duration=0
		set @Total_Leave=@Total_Leave+@Duration
		
		fetch next from Leave_Cur into @Start_Dt,@End_Dt	
	end	

close Leave_Cur 
deallocate Leave_Cur 


return @Total_Leave

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


/*
this function gets emp_id and two date values to find
if leave range already exixts or not
if positive then returns 1 else 0

select dbo.GetLeaveRangeMatched('9-001','2002-08-09','2002-08-13')

*/

create  Function GetLeaveRangeMatched(@Emp_id varchar(10)
		,@Date_From Datetime,@Date_To Datetime)

Returns Int

AS	

BEGIN

Declare @Return_Value int
,@Start_Dt datetime
,@End_Dt datetime


Declare Leave_Cur cursor for

select Start_Dt,End_Dt from emp_leave_info where emp_id=@emp_id
--------------------------------------------------------------------

open Leave_Cur 

	
	fetch next from Leave_Cur into @Start_Dt,@End_Dt

	while @@fetch_status=0
	begin

		if (@Date_From between @Start_Dt and @End_Dt)or 
			(@Date_To between @Start_Dt and @End_Dt)or
				(@Start_Dt between @Date_From and @Date_To)or
					(@End_Dt between @Date_From and @Date_To)
					
			set @Return_Value=1
		else 
			set @Return_Value=0

	
		if @Return_Value=1 	

			break
		else
			fetch next from Leave_Cur into @Start_Dt,@End_Dt	
	
	end	


	if @Return_Value is null set @Return_Value=0

close Leave_Cur 
deallocate Leave_Cur 


return @Return_Value

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


/*

select dbo.GetLeaveValidation_Mt('9-017')

*/
create function GetLeaveValidation_Mt(@Emp_ID varchar(10))
Returns int
as

BEGIN
declare @Result int

if (select Emp_gender from emp_per_Info where emp_ID=@Emp_ID)='Female'and (select Emp_marital_st
from emp_per_Info  where emp_ID=@Emp_ID)='Married'
	
		set @Result=1
	else
		set @Result=0

if @Result is null set @Result=0

return @Result
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


/*
select dbo.GetLogOut ('9-015','2002-05-13 09:02:06.000')

*/


CREATE  function GetLogOut(@Emp_Id varchar(10),@Login_Dt datetime)

returns Datetime

AS
BEGIN


declare @MM int
,@YY int
,@DD int
,@emp_logout datetime


SET @YY=(select datepart(Year,@Login_Dt))
SET @MM=(select datepart(month,@Login_Dt))
SET @DD=(select datepart(Day,@Login_Dt))

set @emp_logout=(select emp_logout from emp_att_info where Emp_Id=@Emp_Id
	and Emp_LogIn=@Login_Dt)


	if @emp_logout is null or @emp_logout=''

	BEGIN
		declare @Move_Id int,@Move_Out_Dt datetime

		Set @Move_Id=(select max(Move_Id) from emp_movement where emp_id=@Emp_Id
				and datepart(month,Move_Out_Dt)=@MM
					and datepart(Year,Move_Out_Dt)=@YY
				 		and datepart(Day,Move_Out_Dt)=@DD)
				
		set @Move_Out_Dt=(select Move_Out_Dt from emp_movement 
				where emp_id=@Emp_Id
			
					and datepart(month,Move_Out_Dt)=@MM
						and datepart(Year,Move_Out_Dt)=@YY
							and datepart(Day,Move_Out_Dt)=@DD	
								and Move_Id=@Move_Id)	
				/*	
				and  datepart(Year,Entry_Dt)=@YY
						and datepart(month,Entry_Dt)=@MM
							and datepart(Day,Entry_Dt)>@DD)
				*/

				if @Move_Out_Dt is null or @Move_Out_Dt =''
	
					begin

						set @emp_logout=(select dbo.GetDefaultTime(@Login_Dt))
						
			
					end

	
					else

					BEGIN
						set @emp_logout=@Move_Out_Dt
					END

	END

	ELSE
	
	BEGIN
		set @emp_logout=@emp_logout
	END
	

return @emp_logout

END


/*

2002-05-13 09:02:06.000                                NULL                                                   NULL
2002-05-15 09:03:40.053                                NULL                                                   NULL
2002-05-16 09:03:01.827                                2002-05-16 17:02:25.040   

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
select dbo.GetLogOutOrNot('9-015','2002-05-22')
*/


CREATE Function GetLogOutOrNot(@Emp_Id varchar(10),@Dt datetime)
returns int
AS

BEGIN
Declare @Result int

	if exists(select * from emp_Att_Info where emp_Id=@Emp_id and
		convert(char(12),Emp_Login,1 )=convert(char(12),@Dt,1)and 
 				emp_Logout is null )

		set @Result=0	---Not yet logged out
	else
		set @Result=1	---Already Logged out

return @Result
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

/*
select dbo.GetMonthDays('February','2002')
*/

CREATE function GetMonthDays(@pay_Month varchar(12),@pay_Year varchar(4))
returns int

AS

BEGIN
declare @M_Days varchar(12)
,@Days varchar(2)
,@Year int 	

set @Year=convert(int,@Pay_Year)

	if @Year % 4 = 0 
		Begin
			if @Year % 100 <> 0 or @Year % 400 = 0 
				set @Days = '29'
			else
				set @Days = '28'
		End
	
	else

	set @Days = '28'

------------------------------------------------------------------

Set @M_Days=(CASE @pay_Month 
		when 'January' THEN '31'
		when 'February' THEN @Days
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

return @M_Days

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






/*

select dbo.GetMonthlyMovementHr('9-015','May','2002')

*/



CREATE  Function GetMonthlyMovementHr (@Emp_Id varchar(10)
				,@pay_Month varchar(12)
					,@pay_year varchar(4))

returns char(4)

AS

BEGIN

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
		

		
		set @Hr=(select dbo.GetHrInMovement(@Emp_Id,@Emp_LogIn))	
				
		SET @Total_Hr=@Total_Hr+@Hr

		fetch next from LogInDt_Cursor into @Emp_LogIn

	END

	

close LogInDt_Cursor
deallocate LogInDt_Cursor

SET @Total_Hr=(select convert(char(4),@Total_Hr))

return @Total_Hr

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



--select * from increment_history
--select * from payscale_main
--select * from payscale_sub

--  select dbo.GetMonthlySalaryBrkUp ('9-027',8,'2002')

CREATE  Function GetMonthlySalaryBrkUp(@Emp_id varchar(10),
			@Pay_Month int,@Pay_year varchar(4))

Returns money

AS

begin

declare @SB money
,@Incr int
,@Num_Incr int
,@Current_Basic money

------------------------------------------------------------------------
select @SB=SB,@Incr=Incr from payscale_main where Scale_Code=
	(select Scale_code from Increment_History where Emp_Id=@Emp_Id and 
	datepart(month,Effect_Dt)<=@Pay_Month and datepart(year,Effect_Dt)=@Pay_Year)

------------------------------------------------------------------------
set @Num_Incr=(select sum(Num_Incr) from increment_history where Emp_Id=@Emp_id and 
scale_code =(select Scale_code from Increment_History where Emp_Id=@Emp_Id and 
	datepart(month,Effect_Dt)<=@Pay_Month and datepart(year,Effect_Dt)=@Pay_Year))

------------------------------------------------------------------------

set @Current_Basic=@SB+(@Incr*@Num_Incr)

return @Current_Basic

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

/*

select dbo.GetMovementNum('9-015','2002-05-13')

*/


CREATE Function GetMovementNum(@Emp_Id varchar(10),@Dt datetime)
returns int

AS

BEGIN
DECLARE @Movement_Number int

SET @Movement_Number=(

		SELECT count(*) FROM Emp_Movement WHERE emp_id=@emp_id
		and datepart(Year,Move_Out_Dt)=datepart(Year,@Dt)
		and datepart(month,Move_Out_Dt)=datepart(month,@Dt)
		and datepart(Day,Move_Out_Dt)=datepart(Day,@Dt)

			)		
RETURN @Movement_Number

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





/*

select dbo.GetNetBonusPay ('9-001','18','August','2002')

*/

create     Function GetNetBonusPay(@Emp_Id varchar(10)
			,@Pay_Type varchar(2)
			,@Pay_Month varchar(12)
			,@Pay_year varchar(4))
Returns Money

AS

BEGIN

Declare @Amount money

set @Amount=(select Amount from Payroll_sub where A_Gen=
		(Select A_Gen from payroll_main where emp_id=@Emp_Id and 
			Pay_month=@Pay_month and Pay_Year=@Pay_Year and Pay_Type=@Pay_Type))


if @Amount is null set @Amount=0





Return @Amount

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






/*

select dbo.GetNetPay ('9-027','May','2002')

*/

CREATE   Function GetNetPay(@Emp_Id varchar(10)
			,@Pay_Month varchar(10)
			,@Pay_year varchar(4))

Returns Money

AS

BEGIN

Declare @Amount money
,@Brk_Up money
,@Head_code varchar(10)
,@Operation varchar(1)

set @Amount=0
----------------------------------------------------------
declare Salary_Head_Cursor Cursor for
Select Head_code,Operation from Pay_Struc where Head_Code >0 and Head_Code <15
----------------------------------------------------------

	Open Salary_Head_Cursor
	fetch next from  Salary_Head_Cursor into @Head_code,@Operation

	while @@Fetch_Status=0
		BEGIN

			set @Brk_Up =(select dbo.GetSalaryBrkUp (@Head_Code,@Emp_Id,@Pay_Month,@Pay_Year))

				if @Operation='+'
					begin
						set @Amount=(@Amount+@Brk_Up)
					end
				if @Operation='-'
					begin			
					    set @Amount=(@Amount-@Brk_Up)
					end

			fetch next from  Salary_Head_Cursor into @Head_code,@Operation
		END
		
	Close Salary_Head_Cursor
	Deallocate Salary_Head_Cursor


if @Amount is null set @Amount=0

Return @Amount

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






--select * from test
--update test set A_Gen='SL002062'
/*
update Pay_type_ind set FB_No='FB000001'
select * from Pay_type_ind

select dbo.GetNewID_DCL('9')

*/

CREATE  Function GetNewID_DCL(@Prefix varchar(2))
returns varchar(6)
AS
Begin

declare @Pay_Type varchar(3)
,@Latest_Num varchar(6)
,@Last_Num int
,@Num_Len int	

if len(@Prefix)=2 and right(@Prefix,1)='-'
	set  @Pay_Type=@Prefix

if len (@Prefix)=1
	set  @Pay_Type=@Prefix+'-'

if len (@Prefix)=2 and right(@Prefix,1)!='-'
	set  @Pay_Type=@Prefix+'-'

		
	set @Last_Num =(select distinct max(convert(int,right(Emp_Id,3)))
				from Emp_per_Info where Emp_Id like @Pay_Type + '%')


		if @Last_Num is null or @Last_Num=''
			begin
				set @Last_Num=1
				set @Latest_Num = @Pay_Type+'00'+ convert(char(1),@Last_Num)
			end
		else 	
			Begin		
				set @Last_Num=@Last_Num+1
				
				set @Num_Len=len(@Last_Num)

				if @Num_Len=1
					set @Latest_Num = @Pay_Type+'00'+ convert(char(1),@Last_Num)
				
				if @Num_Len=2
					set @Latest_Num = @Pay_Type+'0'+ convert(char(2),@Last_Num)

				if @Num_Len=3
					set @Latest_Num = @Pay_Type+ convert(char(3),@Last_Num)

			end


			set @Latest_Num=ltrim(@Latest_Num)

Return @Latest_Num

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

----select dbo.GetNoMOnthDiff ('2005-04-05','2005/06/14') as DiffMonthofTwoDates


CREATE  Function GetNoMOnthDiff(
		 @Start_Time datetime,
		 @End_Time datetime
)
Returns int

AS	

BEGIN

Declare 
	@GetTotalNoofMonth int
	


Set @GetTotalNoofMonth=(select datediff(month,@Start_Time,@End_Time))


return @GetTotalNoofMonth

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

---select  dbo.GetNoMinutesOT('09:30:20 AM','10:10:10 AM')


CREATE   Function GetNoMinutesOT(
		 @OT_St_Time datetime,
		 @OT_End_Time datetime
)
Returns int

AS	

BEGIN

Declare @GetTotalHours int


Set @GetTotalHours=(select datediff(minute,@OT_St_Time,@OT_End_Time))

return @GetTotalHours

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



/*
select * from param_tbl
select dbo.GetPresentBasic ('9-027')
select dbo.GetOT_Hr('9-015','July','2002')
Select dbo.GetOTAmount('9-027','July','2002')

*/

cReate  function GetOTAmount(@Emp_Id varchar(10)
		,@Pay_Month varchar(12)
		,@Pay_Year varchar(4))
returns Money

AS

BEGIN
Declare 

@OT_Hr real
,@Off_Hr real
,@Prcnt float
,@Base_Flag int
,@Amount money
,@Days int
,@Day_Flag int

set @Off_Hr=(select dbo.GetOfficeDuration(@Pay_Month,@Pay_year))
set @OT_Hr=(select dbo.GetOT_Hr(@Emp_Id,@Pay_Month,@Pay_Year))

set @Day_Flag=(select dbo.GetParamFlag(76))

set @Base_Flag=(select dbo.GetParamFlag(75))
set @Prcnt=(select dbo.GetParamValue(75))

if @Base_Flag=0 	---Percent of Basic or gross salary
	set @Amount=(select dbo.GetPresentBasic(@Emp_Id))
if @Base_Flag=1
	set @Amount=(select dbo.GetPresentSalary(@Emp_Id))


if @Day_Flag=0 		---Days of the month or working days
	Set @Days=(select dbo.GetMonthDays(@Pay_Month,@Pay_year))
if @Day_Flag=1
	Set @Days=(select dbo.GetWorkingDays(@Pay_Month,@Pay_year))


	set @Amount=((@Amount)/(@Days*@Off_Hr))*(@Prcnt/100) 
	set @Amount= @Amount * @OT_Hr

Return @Amount
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




/*

select dbo.GetOT_Hr('9-015','July','2002')


*/

CREATE    Function GetOT_Hr (@Emp_Id varchar(10)
				,@pay_Month varchar(12)
					,@pay_year varchar(4))
returns real

AS

BEGIN

Declare 
@Emp_LogOut  datetime
,@LogOut_Tm char(8)
,@Office_End  char(8)
,@Total_Hr float
,@Hr float
,@Flag int
,@Time_Brkup varchar(2)

set @Total_Hr=0
set @Flag=(select dbo.GetParamFlag(74))

Declare LogInDt_Cursor Cursor for

	select Emp_LogOut from emp_Att_info where Emp_Id=@Emp_Id
		and DATENAME(MONTH,Entry_Dt)=@pay_Month 
		and DATEPART(YEAR,Entry_Dt)=@pay_year
		and Emp_LogOut is not null

-------------------------------------------------------------------
open LogInDt_Cursor

	fetch next from LogInDt_Cursor into @Emp_LogOut

	while @@fetch_Status=0

	BEGIN

		
			set @Office_End=(select dbo.GetOfficeEnd (@Emp_Id,@Emp_LogOut))
		
		

		if @Flag=1		--- has OT break up

		begin
			
			set @Time_Brkup=(select dbo.GetParamValue(74))
			set @Office_End=(select dbo.GetTimeAdded(@Office_End,@Time_Brkup))
		end


		set @LogOut_Tm=(Select dbo.GetTimeOnly (@Emp_LogOut))


		if @LogOut_Tm>@Office_End

			begin
				Set @Hr=ROUND(SUM(DATEDIFF(n,@Office_End,@LogOut_Tm)* 0.01667),1) 
				
				SET @Total_Hr=@Total_Hr+@Hr
			end

		fetch next from LogInDt_Cursor into @Emp_LogOut

	END

	

close LogInDt_Cursor
deallocate LogInDt_Cursor


return @Total_Hr

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





/*

select dbo.GetOffStayHr('9-015','May','2002')

*/



CREATE   Function GetOffStayHr (@Emp_Id varchar(10)
				,@pay_Month varchar(12)
					,@pay_year varchar(4))

returns float

AS

BEGIN

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
		

		
		set @Hr=(select dbo.GetHrsInOffice(@Emp_Id,@Emp_LogIn))	
				
		SET @Total_Hr=@Total_Hr+@Hr

		fetch next from LogInDt_Cursor into @Emp_LogIn

	END

	

close LogInDt_Cursor
deallocate LogInDt_Cursor


return @Total_Hr

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



/*
select dbo.GetOfficeStart('9-015','2002-07-25')
select dbo.GetOfficeEnd ('9-015','2002-07-25')
select dbo.GetOfficeDuration('July','2002')
*/


create  Function GetOfficeDuration(@Pay_Month varchar(12),@Pay_Year varchar(8))
Returns real

AS

BEGIN

Declare 
@Start_time varchar(8)
,@End_time varchar(8)
,@Min varchar(2)
,@Duration real





	select @Start_time=Start_time
		
		from office_time where Effect_date=(
		select Mdt=(select max(Effect_date)from office_time 
		where datepart(Month,Effect_date)<=(select dbo.Month_No(@Pay_Month))
		and datepart(year,Effect_date)<=@pay_year))

set @Start_time=(select rtrim(ltrim(@Start_time)))
set @Min=(select substring(@Start_time,4,2))			----Cut minutes

if len(@Min)=1
set @Min='0'+convert(char,@Min)
if len(@Min)=1
set @Min='0'+convert(char,@Min)
								----Cut and add 12 with hour
if (select substring(@Start_time,7,1))='P' and (select substring(@Start_time,1,2))!='12'			
	Set @Start_time=convert(char,(convert(int,substring(@Start_time,1,2))+12))
else
	Set @Start_time=convert(char,substring(@Start_time,1,2))
set @Start_time=rtrim(@Start_time)+ ':'+@Min+':00'		----Add seconds	

set @Start_time=rtrim(ltrim(@Start_time))
----------------------------------------------------------------------------


select @End_time=End_time from office_time where Effect_date=(
		select Mdt=(select max(Effect_date)from office_time 
		where datepart(Month,Effect_date)<=(select dbo.Month_No(@Pay_Month))
		and datepart(year,Effect_date)<=@pay_year))


set @End_time=rtrim(ltrim(@End_time))
set @Min=(select substring(@End_time,4,2))			----Cut minutes

if len(@Min)=1
set @Min='0'+convert(char,@Min)
						----Cut and add 12 with hour
if (select substring(@End_time,7,1))='P' and (select substring(@End_time,1,2))!='12'			
	Set @End_time=convert(char,(convert(int,substring(@End_time,1,2))+12))
else
	Set @End_time=convert(char,substring(@End_time,1,2))
set @End_time=rtrim(@End_time)+ ':'+@Min+':00'		----Add seconds	

set @End_time=rtrim(ltrim(@End_time))



set @Duration=(select ROUND(SUM(DATEDIFF(n,@Start_time,@End_time)* 0.01667),1)) 

if @Duration is null set @Duration=8

return @Duration

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




/*

select dbo.GetOfficeStart('9-015','2002-07-25')
select dbo.GetOfficeEnd ('9-015','2002-07-25')

*/

CREATE  Function GetOfficeEnd(@Emp_Id varchar(10),@Dt datetime)

Returns varchar(8)

AS

BEGIN

Declare @Duty_Type varchar(1)
,@Pay_Month varchar(12),@Pay_Year varchar(4)
,@End_time varchar(8),@Relaxed varchar(2)
,@Sp_End_time varchar(10),@Sp_End_Day  varchar(10)
,@Min varchar(2)
------------------------------------------------------------------------------
Set @Pay_Month=(Select datename(month,@Dt))
Set @Pay_Year=(Select datepart(Year,@Dt))

--------Duty Type (Fixed/Shifting)----------------------------------------------------------------------
set @Duty_Type=(Select Duty_type from emp_job_hist_current where Emp_Id=@Emp_Id)

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
	--select @S_Dt
	--select @E_Dt
	--------------------------------------------------------
	select @End_time=Shift_End from Shift where Shift_Code=(
		select Shift_Code from Group_Shift where Gr_Code=(
			select Gr_Code from Worker_Group where emp_ID=@ID))and 
				datepart(day,@Dt)>=datepart(day,@S_Dt) and
				datepart(month,@Dt)=datepart(month,@S_Dt)and
				datepart(year,@Dt)=datepart(year,@S_Dt)and 
				datepart(day,@Dt)<=datepart(day,@E_Dt) and
				datepart(month,@Dt)=datepart(month,@E_Dt)and
				datepart(year,@Dt)=datepart(year,@E_Dt)
END


if @Duty_Type='1'	--Fixed duty time

Begin

	select @End_time=End_time,@Relaxed=Relaxed
		,@Sp_End_time=Sp_End_time,@Sp_End_day=Sp_End_day
		from office_time where Effect_date=(
		select Mdt=(select max(Effect_date)from office_time 
		where datepart(Month,Effect_date)<=(select dbo.Month_No(@Pay_Month))
		and datepart(year,Effect_date)<=@pay_year))

END
--------What day is it?-----------------------------------------------------------------------

Declare @Today varchar(9)
set @Today=(select dbo.GetDayName (@Dt))

IF @Today=@Sp_End_Day 

	BEGIN
		IF @Sp_End_time IS NOT NULL or @Sp_End_time=''

			begin

				SET @End_time=@Sp_End_time
			end
	
	END

-------------------------------------------------------------------------------
set @End_time=rtrim(ltrim(@End_time))
set @Min=(select substring(@End_time,4,2))			----Cut minutes

if len(@Min)=1
set @Min='0'+convert(char,@Min)
						----Cut and add 12 with hour
if (select substring(@End_time,7,1))='P' and (select substring(@End_time,1,2))!='12'			
	Set @End_time=convert(char,(convert(int,substring(@End_time,1,2))+12))
else
	Set @End_time=convert(char,substring(@End_time,1,2))
set @End_time=rtrim(@End_time)+ ':'+@Min+':00'		----Add seconds	

set @End_time=rtrim(ltrim(@End_time))


RETURN @End_time

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

/*

select dbo.GetOfficeStart('9-015','2002-07-25')

*/

CREATE Function GetOfficeStart(@Emp_Id varchar(10),@Dt datetime)

Returns varchar(8)

AS

BEGIN

Declare @Duty_Type varchar(1)
,@Pay_Month varchar(12),@Pay_Year varchar(4)
,@Start_time varchar(8),@Relaxed varchar(2)
,@Sp_Start_time varchar(10),@Sp_Start_Day varchar(10)
,@Min varchar(2)
------------------------------------------------------------------------------
Set @Pay_Month=(Select datename(month,@Dt))
Set @Pay_Year=(Select datepart(Year,@Dt))

--------Duty Type (Fixed/Shifting)----------------------------------------------------------------------
set @Duty_Type=(Select Duty_type from emp_job_hist_current where Emp_Id=@Emp_Id)

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
	--select @S_Dt
	--select @E_Dt
	--------------------------------------------------------
	select @Start_time=Shift_Start,@Relaxed=[Delay] 
	from Shift where Shift_Code=(
		select Shift_Code from Group_Shift where Gr_Code=(
			select Gr_Code from Worker_Group where emp_ID=@ID))and 
				datepart(day,@Dt)>=datepart(day,@S_Dt) and
				datepart(month,@Dt)=datepart(month,@S_Dt)and
				datepart(year,@Dt)=datepart(year,@S_Dt)and 
				datepart(day,@Dt)<=datepart(day,@E_Dt) and
				datepart(month,@Dt)=datepart(month,@E_Dt)and
				datepart(year,@Dt)=datepart(year,@E_Dt)
END


if @Duty_Type='1'	--Fixed duty time

Begin

	select @Start_time=Start_time,@Relaxed=Relaxed
		,@Sp_Start_time=Sp_Start_time,@Sp_Start_day=Sp_Start_day
		from office_time where Effect_date=(
		select Mdt=(select max(Effect_date)from office_time 
		where datepart(Month,Effect_date)<=(select dbo.Month_No(@Pay_Month))
		and datepart(year,Effect_date)<=@pay_year))

END
--------What day is it?-----------------------------------------------------------------------

Declare @Today varchar(9)
set @Today=(select dbo.GetDayName (@Dt))

IF @Today=@Sp_Start_Day 

	BEGIN
		IF @Sp_Start_time IS NOT NULL or @Sp_Start_time=''

			begin

				SET @Start_time=@Sp_Start_time
			end
	
	END

-------------------------------------------------------------------------------

set @Start_time=rtrim(ltrim(@Start_time))
set @Min=(select substring(@Start_time,4,2))			----Cut minutes

if len(@Min)=1
set @Min='0'+convert(char,@Min)
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


RETURN @Start_time

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


/*

select dbo.GetOfficeStartWD('9-015','2002-07-25')

*/

Create  Function GetOfficeStartWD(@Emp_Id varchar(10),@Dt datetime)

Returns varchar(8)

AS

BEGIN

Declare @Duty_Type varchar(1)
,@Pay_Month varchar(12),@Pay_Year varchar(4)
,@Start_time varchar(8),@Relaxed varchar(2)
,@Sp_Start_time varchar(10),@Sp_Start_Day varchar(10)
,@Min varchar(2)
------------------------------------------------------------------------------
Set @Pay_Month=(Select datename(month,@Dt))
Set @Pay_Year=(Select datepart(Year,@Dt))

--------Duty Type (Fixed/Shifting)----------------------------------------------------------------------
set @Duty_Type=(Select Duty_type from emp_job_hist_current where Emp_Id=@Emp_Id)

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
	--select @S_Dt
	--select @E_Dt
	--------------------------------------------------------
	select @Start_time=Shift_Start,@Relaxed=[Delay] 
	from Shift where Shift_Code=(
		select Shift_Code from Group_Shift where Gr_Code=(
			select Gr_Code from Worker_Group where emp_ID=@ID))and 
				datepart(day,@Dt)>=datepart(day,@S_Dt) and
				datepart(month,@Dt)=datepart(month,@S_Dt)and
				datepart(year,@Dt)=datepart(year,@S_Dt)and 
				datepart(day,@Dt)<=datepart(day,@E_Dt) and
				datepart(month,@Dt)=datepart(month,@E_Dt)and
				datepart(year,@Dt)=datepart(year,@E_Dt)
END


if @Duty_Type='1'	--Fixed duty time

Begin

	select @Start_time=Start_time,@Relaxed=Relaxed
		,@Sp_Start_time=Sp_Start_time,@Sp_Start_day=Sp_Start_day
		from office_time where Effect_date=(
		select Mdt=(select max(Effect_date)from office_time 
		where datepart(Month,Effect_date)<=(select dbo.Month_No(@Pay_Month))
		and datepart(year,Effect_date)<=@pay_year))

END
--------What day is it?-----------------------------------------------------------------------

Declare @Today varchar(9)
set @Today=(select dbo.GetDayName (@Dt))

IF @Today=@Sp_Start_Day 

	BEGIN
		IF @Sp_Start_time IS NOT NULL or @Sp_Start_time=''

			begin

				SET @Start_time=@Sp_Start_time
			end
	
	END

-------------------------------------------------------------------------------

set @Start_time=rtrim(ltrim(@Start_time))
set @Min=(select substring(@Start_time,4,2))			----Cut minutes

if len(@Min)=1
set @Min='0'+convert(char,@Min)
--set @Min=convert(int,@Min)+ convert(int,@Relaxed)
--if len(@Min)=1
--set @Min='0'+convert(char,@Min)
								----Cut and add 12 with hour
if (select substring(@Start_time,7,1))='P' and (select substring(@Start_time,1,2))!='12'			
	Set @Start_time=convert(char,(convert(int,substring(@Start_time,1,2))+12))
else
	Set @Start_time=convert(char,substring(@Start_time,1,2))
set @Start_time=rtrim(@Start_time)+ ':'+@Min+':00'		----Add seconds	

set @Start_time=rtrim(ltrim(@Start_time))


RETURN @Start_time

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



/*
this function gets emp_id and two date values to find
if leave range already exixts or not
if positive then returns 1 else 0

select dbo.GetLeaveRangeMatched('9-001','2002-08-09','2002-08-13')

*/

CREATE   Function GetOutRangeMatched(@Emp_id varchar(10)
		,@Date_From Datetime,@Date_To Datetime)

Returns Int

AS	

BEGIN

Declare @Return_Value int
,@Str_Date datetime
,@End_Date datetime


Declare Out_Cur cursor for

select Str_Date,End_Date from Out_of_Out where emp_id=@emp_id
--------------------------------------------------------------------

open Out_Cur 

	
	fetch next from Out_Cur into @Str_Date,@End_Date

	while @@fetch_status=0
	begin

		if (@Date_From between @Str_Date and @End_Date)or 
			(@Date_To between @Str_Date and @End_Date)or
				(@Str_Date between @Date_From and @Date_To)or
					(@End_Date between @Date_From and @Date_To)
					
			set @Return_Value=1
		else 
			set @Return_Value=0

	
		if @Return_Value=1 	

			break
		else
			fetch next from Out_Cur into @Str_Date,@End_Date	
	
	end	


	if @Return_Value is null set @Return_Value=0

close Out_Cur 
deallocate Out_Cur 


return @Return_Value

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

/*
select dbo.GetParamFlag(21)
select * from param_tbl

*/
---------------------------------------------------------------

CREATE  Function GetParamFlag(@Policy_No int)
returns int

AS

BEGIN
declare @Flag int

	select 	@Flag=Flag from Param_Tbl where Policy_No=@Policy_No

return @Flag
	
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

---------------------------------------------------------------

Create Function GetParamValue(@Policy_No int)
returns varchar(50)

AS

BEGIN
declare @Value varchar(50)

	select 	@Value=Value from Param_Tbl where Policy_No=@Policy_No

return @Value
	
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



/*
select dbo.GetPayType_PlcNo (16,0)

*/


create   Function GetPayType_PlcNo(@Pay_Type int,@No int)
Returns Int

As

Begin

Declare @Policy_No int

/*
Pay_Type	Policy No and description  @No
----------------------------------------------
15 Eid Bonus	63  Eid_Bonus_Prcnt         0 
		64  Eid_Bonus_Deservs       1 
----------------------------------------------
16 AGM		65  AGM_Prcnt               0 
		66  AGM_Deservs             1 
----------------------------------------------
17 BATEXPO	67  BTEXPO_Prcnt            0 
		68  BTEXPO_Deservs          1 
----------------------------------------------
21 Overtime	69  OT_Deservs              0 
----------------------------------------------
18 Election	70  Election_Alwnc_Prcnt    0 
   Bonus	71  Election_Alwnc_Deserves 1
----------------------------------------------
19 Travel/
tour Allowance	72  Travel_tour_Alwnc       0	
----------------------------------------------
20 Incentive 	73  Incentive               0 
----------------------------------------------
*/

IF @No=0
BEGIN
	if @Pay_Type=15 Set @Policy_No=63
	if @Pay_Type=16 Set @Policy_No=65
	if @Pay_Type=17 Set @Policy_No=67
	if @Pay_Type=18 Set @Policy_No=70
	if @Pay_Type=19 Set @Policy_No=72
	if @Pay_Type=20 Set @Policy_No=73
	if @Pay_Type=21 Set @Policy_No=69
END

IF @No=1
BEGIN
	if @Pay_Type=15 Set @Policy_No=64
	if @Pay_Type=16 Set @Policy_No=66
	if @Pay_Type=17 Set @Policy_No=68
	if @Pay_Type=18 Set @Policy_No=71
END



return @Policy_No
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


create    Function GetPresence(@Emp_id varchar (10),
				@pay_Month varchar(12),
				@pay_Year varchar(4))
Returns Int

AS

Begin

declare 
@Tot_Att int
Set @Tot_Att=(Select count (Emp_Id)from emp_att_info where Emp_ID=@Emp_id 
		and datepart(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@Pay_month)
return @Tot_Att
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




--select dbo.GetPresenceBonus('9005','June','2002')

--rpt_Payroll_Ind_All 'Sec','Banking','June','2002'

create  Function GetPresenceBonus(@Emp_id varchar (10),
				@pay_Month varchar(12),
						@pay_Year varchar(4))
Returns money

AS

Begin

declare @Desig varchar(25)
,@M_Days varchar(12)
, @Febru_days varchar(2)
, @Tot_Hol  int
, @Tot_Att int
, @Working_Days int
, @Pres_Bonus varchar(10)

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

-------------------------Holidays---------------------------------------------
set @Tot_Hol = (select count(*) from hol_list where 
(datepart(year,Str_date)=@pay_Year and datename(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datename(month,End_date)=@Pay_month)
)
-------------------------Working Days-----------------------------------------
set @Working_Days=(@M_Days-@Tot_Hol)

-------------------------Presence---------------------------------------------
Set @Tot_Att=(Select count (Emp_Id)from emp_att_info where Emp_ID=@Emp_id 
		and datepart(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@Pay_month)

-------------------------Presence Bonus---------------------------------------------

if @Tot_Att=23 --@M_Days

Begin
	set @Desig=(Select Emp_Desig from emp_job_hist_current where Emp_Id=@Emp_Id)
	set @Pres_Bonus =(select case @Desig	

	WHEN 'Analyst Programmer'THEN '400'
	WHEN 'Junior Programmer'THEN '100'
	WHEN 'Office Assistant'THEN '100'
	WHEN 'Programmer'THEN '200'
	WHEN 'Senior Programmer'THEN '250'
	WHEN 'Software Specialist'THEN '500'
	WHEN 'Systems Analyst'THEN '500'
	WHEN 'Trainee Programmer'THEN '0'	
	
	END)

END
ELSE

BEGIN
	set @Pres_Bonus='0'
END

------------------------------------------------------------------------
declare @PBonus money
set @PBonus=convert(money,@Pres_Bonus)

return @PBonus

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



--select * from increment_history
--select * from payscale_main
--select * from payscale_sub

--  select dbo.GetPresentBasic ('9-027')

CREATE  Function GetPresentBasic (@Emp_id varchar(10))

Returns money

AS

begin

declare @SB money
,@Incr int
,@Num_Incr int
,@Present_Basic money

------------------------------------------------------------------------
select @SB=SB,@Incr=Incr from payscale_main where Scale_Code=
	(select Scale_Code from increment_history where Emp_Id=@Emp_id and Track_Id=
		(select min(Track_Id) from increment_history where Emp_Id=@Emp_Id))
------------------------------------------------------------------------
set @Num_Incr=(select sum(Num_Incr) from increment_history where Emp_Id=@Emp_id and  Scale_Code=
		(select Scale_Code from increment_history where Emp_Id=@Emp_id and Track_Id=
				(select min(Track_Id) from increment_history where Emp_Id=@Emp_Id)))

------------------------------------------------------------------------

set @Present_Basic=@SB+(@Incr*@Num_Incr)

return @Present_Basic

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


create Function GetPresentOrNot(@Emp_Id varchar(10),@Dt datetime)
returns int
AS

BEGIN
Declare @Result int

	if exists(select * from emp_Att_Info where Emp_Id=@Emp_Id and
		convert(char(12),Emp_Login,1 )=convert(char(12),@Dt,1))

		set @Result=1	
	else
		set @Result=0

return @Result
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



/*
select dbo.GetPresentSalary('9-001')
select sum(Amount) from fixed_pay where emp_id='9-001'
select * from fixed_Pay where emp_Id='9-001'

*/

CREATE Function GetPresentSalary(@Emp_Id varchar(10))
Returns Money

AS

Begin

	declare @Head_Code varchar(2)
	,@Prcnt float
	,@Calc_Type int
	,@Amount money
	,@Scale_Code int
	,@Present_Basic money
	,@Present_Salary money

--------------------------------------------------------------------------------
set @Present_Basic=(select dbo.GetPresentBasic(@Emp_Id))
Set @Present_Salary=@Present_Basic

set @Scale_Code= (select Scale_Code from increment_history where Emp_Id=@Emp_id and Track_Id=
		(select min(Track_Id) from increment_history where Emp_Id=@Emp_Id))

--------------------------------------------------------------------------------
DECLARE Payscale_Sub_Cursor CURSOR FOR
select  @Head_Code,Calc_Type,Prcnt_of_Basic from payscale_sub 
where Scale_Code=@Scale_Code
--------------------------------------------------------------------------------

OPEN Payscale_Sub_Cursor

    FETCH NEXT FROM Payscale_Sub_Cursor into @Head_Code,@Calc_Type,@Prcnt

	WHILE (@@FETCH_STATUS = 0)
	BEGIN
	
		if @Calc_Type= 1 			
		begin
			set @Amount= (@Present_Basic* @Prcnt)/100
		end
		if @Calc_Type= 0				
		begin
			set @Amount= @Prcnt
		end

		set @Present_Salary=@Present_Salary+@Amount		
	
	    FETCH NEXT FROM Payscale_Sub_Cursor into @Head_Code,@Calc_Type,@Prcnt

	END



close Payscale_Sub_Cursor
deallocate Payscale_Sub_Cursor


return @Present_Salary


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


create  function GetSalFixErrMsg (@Scale_Code int,@Basic_Amount int)

returns Varchar(250)
AS

begin

declare @Message varchar(250)
,@SB int
,@Incr int
,@EB int

	set @SB=(select SB from payscale_main where scale_code=@scale_code)
	set @Incr=(select Incr from payscale_main where scale_code=@scale_code)
	set @EB=(select EB from payscale_main where scale_code=@scale_code)
		
	if @Basic_Amount>@EB
		set @Message='Amount overflows, please follow the payscale !'

	else if (@Basic_Amount % @Incr)!=0
		set @Message='Amount does not conform payscale definition !'

	else if @Basic_Amount < @SB
		set @Message='Please follow the payscale !'
	else 
		set @Message='Valid'	

return @Message
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
select dbo.GetSalaryBrkUp ('01','9-001','June','2002')

*/

CREATE Function GetSalaryBrkUp (@Head_Code varchar(10)
				,@Emp_Id varchar(10)
				,@Pay_Month varchar(10)
				,@Pay_year varchar(4))
Returns Money

AS

BEGIN

Declare @Amount money

select @Amount=Amount from Payroll_sub where Head_Code=@Head_Code and A_Gen=(
	select A_Gen from Payroll_Main where Emp_Id=@Emp_Id 
		and Pay_Month=@Pay_Month and Pay_year=@Pay_year)

Return @Amount

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


/*
select dbo.GetSalaryHead ('01')

*/

CREATE function GetSalaryHead (@Head_Code varchar(10))
returns Varchar(40)

AS

BEGIN

declare @Head_Name varchar(40)

	set @Head_Name=(select Head_Name from Pay_Struc
	where Head_Code=@Head_Code)
	

return @Head_Name

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


--select * from increment_history
--select * from payscale_main
--  select dbo.GetStartingBasic ('9-027')

CREATE Function GetStartingBasic (@Emp_id varchar(10))

Returns money

AS

begin


declare @SB money
,@Incr int
,@Num_Incr int
,@Starting_Salary money

------------------------------------------------------------------------
select @SB=SB,@Incr=Incr from payscale_main where Scale_Code=
	(select Scale_Code from increment_history where Emp_Id=@Emp_id and Track_Id=
		(select min(Track_Id) from increment_history where Emp_Id=@Emp_Id))
------------------------------------------------------------------------
set @Num_Incr=(select Num_Incr from increment_history where Emp_Id=@Emp_id and Track_Id=
		(select min(Track_Id) from increment_history where Emp_Id=@Emp_Id))
------------------------------------------------------------------------

set @Starting_Salary=@SB+(@Incr*@Num_Incr)

return @Starting_Salary

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




/*
select dbo.GetStartingSalary('9-001')
select sum(Amount) from fixed_pay where emp_id='9-001'
select * from fixed_Pay where emp_Id='9-001'

*/

CREATE   Function GetStartingSalary(@Emp_Id varchar(10))
Returns Money

AS

Begin

	declare @Head_Code varchar(2)
	,@Prcnt float
	,@Calc_Type int
	,@Amount money
	,@Scale_Code int
	,@Starting_Basic money
	,@Starting_Salary money

--------------------------------------------------------------------------------
set @Starting_Basic=(select dbo.GetStartingBasic(@Emp_Id))
Set @Starting_Salary=@Starting_Basic

set @Scale_Code= (select Scale_Code from increment_history where Emp_Id=@Emp_id and Track_Id=
		(select min(Track_Id) from increment_history where Emp_Id=@Emp_Id))

--------------------------------------------------------------------------------
DECLARE Payscale_Sub_Cursor CURSOR FOR
select  @Head_Code,Calc_Type,Prcnt_of_Basic from payscale_sub 
where Scale_Code=@Scale_Code
--------------------------------------------------------------------------------

OPEN Payscale_Sub_Cursor

    FETCH NEXT FROM Payscale_Sub_Cursor into @Head_Code,@Calc_Type,@Prcnt

	WHILE (@@FETCH_STATUS = 0)
	BEGIN
	
		if @Calc_Type= 1 			
		begin
			set @Amount= (@Starting_Basic* @Prcnt)/100
		end
		if @Calc_Type= 0				
		begin
			set @Amount= @Prcnt
		end

		set @Starting_Salary=@Starting_Salary+@Amount		
	
	    FETCH NEXT FROM Payscale_Sub_Cursor into @Head_Code,@Calc_Type,@Prcnt

	END


close Payscale_Sub_Cursor
deallocate Payscale_Sub_Cursor

return @Starting_Salary


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


/*

select dbo.GetHrInMovement('9-015','2002-05-7')

select dbo.GetStayInOffice('9-015','2002-05-7')

*/
CREATE Function GetStayInOffice(@Emp_Id varchar(10),@Emp_LogIn datetime)

Returns char(4)

AS

BEGIN

declare @Hr_Stay float
,@Office_Hour float

set @Office_Hour=datediff(n,(select dbo.GetOfficeStartWD(@Emp_Id,@Emp_LogIn)),(select dbo.GetOfficeEnd (@Emp_Id,@Emp_LogIn)))* 0.01666667

set @Office_Hour=round(@Office_Hour,1)

set @Hr_Stay=@Office_Hour-(select dbo.GetHrInMovement(@Emp_Id,@Emp_LogIn))

return convert(char(4),@Hr_Stay)

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



/*
select dbo.GetTimeAdded('18:30:00','30')

*/

create  Function GetTimeAdded (@Time char(8),@Add char(2))
returns char(8)

AS

Begin
declare @minutes int
,@Hour int
,@Return_Time Char(8)
,@chrHour char(2)
,@chrMin char(2)

set @Hour=convert(int,substring(@Time,1,2))
set @minutes= convert(int,substring(@Time,4,2))
set @minutes=@minutes+convert(int,@Add)

if @minutes >=60
begin
	set @Hour=@Hour+1
	set @minutes = @minutes-60
end

if len(@Hour)=1 
	set @chrHour='0'+convert(char(1),@Hour)
else 
	set @chrHour=convert(char(2),@Hour)

if len(@minutes)=1 
	set @chrMin='0'+convert(char(1),@minutes)
else
	set @chrMin=convert(char(2),@minutes)

set @Return_Time=@chrHour+':'+@chrMin+substring(@Time,6,3)

return @Return_Time

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

select dbo.GetTimeOnly (getdate())

*/

create  Function GetTimeOnly(@Date datetime)

returns char(8)

as
begin

declare @Conv_dt char(8)

	set @Conv_dt=(select convert(char(8),@Date,8))
	

return @Conv_dt

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
select * from emp_leave_info

select dbo.GetTotAttn('9-001','August','2002')


Late = (Select count(*) from emp_Att_info where In_status='1'and emp_id=@Emp_id
		and datename(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@pay_Month)

*/

CREATE  Function GetTotAttn(@Emp_id varchar(10),@Pay_Month varchar(12),@Pay_Year varchar(4))
Returns Int

AS	

BEGIN

Declare @Total_Attn int

	set @Total_Attn =(select count (*)from emp_att_info where Emp_ID=@Emp_id 
				and datename(month,Entry_dt)=@Pay_month
				and datepart(year,Entry_dt)=@pay_Year)		

	if @Total_Attn is null set @Total_Attn=0

	return @Total_Attn

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


/*
select * from emp_leave_info

select dbo.GetTotLateAttn('9-001','August','2002')


Late = (Select count(*) from emp_Att_info where In_status='1'and emp_id=@Emp_id
		and datename(year,Entry_dt)=@pay_Year
		and datename(month,Entry_dt)=@pay_Month)

*/

create  Function GetTotLateAttn(@Emp_id varchar(10)
			,@Pay_Month varchar(12),@Pay_Year varchar(4))
Returns Int

AS	

BEGIN

Declare @Total_Attn int

	set @Total_Attn =(select count (*)from emp_att_info where Emp_ID=@Emp_id 
				and datename(month,Entry_dt)=@Pay_month
				and datepart(year,Entry_dt)=@pay_Year
				and In_status='1')		

	if @Total_Attn is null set @Total_Attn=0

	return @Total_Attn

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






/*
select * from hol_list

select dbo.GetTot_Holday('december','2002')

*/

create   Function GetTot_Holday(@Pay_Month varchar(12)
		,@Pay_Year varchar(4))

Returns Int

AS	

BEGIN

Declare @Total_Hol int
,@Duration int
,@Str_Date datetime
,@End_Date datetime

set @Total_Hol=0


Declare Hol_Cur cursor for

select Str_Date,End_Date from hol_list where 
datename(month,Str_Date)=@Pay_Month 
and datename(year,Str_Date)=@Pay_Year and  category !='Hartal'
--------------------------------------------------------------------
open Hol_Cur 

	
	fetch next from Hol_Cur into @Str_Date,@End_Date

	while @@fetch_status=0
	begin


		set @Duration=datediff(day,@Str_Date,@End_Date)+1
			if @Duration is null set @Duration=0
		set @Total_Hol=@Total_Hol+@Duration
		
		fetch next from Hol_Cur into @Str_Date,@End_Date	
	end	

close Hol_Cur 
deallocate Hol_Cur 


return @Total_Hol

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







/*
select * from hol_list

select dbo.GetTot_Holday('december','2002')

*/

Create Function GetTot_OtherHolday(@Pay_Month varchar(12)
		,@Pay_Year varchar(4))

Returns Int

AS	

BEGIN

Declare @Total_Hol int
,@Duration int
,@Str_Date datetime
,@End_Date datetime

,@Total_WeekEnds int
,@Total_Other_HolDays int


set @Total_Hol=0


Declare Hol_Cur cursor for

select Str_Date,End_Date from hol_list where 
datename(month,Str_Date)=@Pay_Month 
and datename(year,Str_Date)=@Pay_Year and  category !='Hartal'
--------------------------------------------------------------------
open Hol_Cur 

	
	fetch next from Hol_Cur into @Str_Date,@End_Date

	while @@fetch_status=0
	begin


		set @Duration=datediff(day,@Str_Date,@End_Date)+1
			if @Duration is null set @Duration=0
		set @Total_Hol=@Total_Hol+@Duration
		
		fetch next from Hol_Cur into @Str_Date,@End_Date	
	end	


close Hol_Cur 
deallocate Hol_Cur 


set @Total_WeekEnds =(select count(*) from hol_list where 
	datename(month,Str_Date)=@Pay_Month 
	and datename(year,Str_Date)=@Pay_Year 
	and  rtrim(ltrim(category))='Weekend')

set @Total_Other_HolDays=(@Total_Hol-@Total_WeekEnds)

return @Total_Other_HolDays

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



CREATE  Function GetTotalHartals(@Pay_Month varchar(12)
		,@Pay_Year varchar(4))

Returns Int

AS	

BEGIN

Declare @GetTotalHartals int



set @GetTotalHartals=(select count(*) from hol_list where 
		datename(month,Str_Date)=@Pay_Month 
		and datename(year,Str_Date)=@Pay_Year
		and  rtrim(ltrim(category))='Hartal')

return @GetTotalHartals

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







/*
select * from hol_list

select dbo.GetTotalWeekEnds('december','2002')

*/

CREATE  Function GetTotalWeekEnds(@Pay_Month varchar(12)
		,@Pay_Year varchar(4))

Returns Int

AS	

BEGIN

Declare @Total_WeekEnds int



set @Total_WeekEnds =(select count(*) from hol_list where 
	datename(month,Str_Date)=@Pay_Month 
	and datename(year,Str_Date)=@Pay_Year 
	and  rtrim(ltrim(category))='Weekend')

return @Total_WeekEnds

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



/*
select dbo.GetVar_PreMonth ('9-026','July','2002','Telephone Allowance') 

*/

CREATE   function GetVar_PreMonth(
				@Emp_id varchar (10),
				@pay_Month varchar(12),
				@pay_Year varchar(4),
				@Head_Name varchar(35)
					)
Returns money

AS

Begin

declare @Preceding_Month varchar(12)
,@Amount money

set @Preceding_Month=(select dbo.Month_Name (dbo.Month_No (@Pay_month)-1))


Set @Amount=(select Amount from payroll_sub where A_Gen=(select A_Gen from payroll_main 
					where emp_id=@Emp_Id and Pay_month=@Preceding_Month 
						and Pay_year=@Pay_year and Pay_Type=0)
		and Head_Code=(Select Head_Code from Pay_Struc 
			where Head_Name=@Head_Name)
)

----if there is no salary prepared for the preceding month

if @Amount is null	
set @Amount=0

return @Amount

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



/*
select dbo.GetVouAmount('Dept','0','August','2002','Software')
*/


create  function GetVouAmount(@Mode varchar(5)
			,@Pay_Type varchar(35)
			,@Pay_Month varchar(12)
			,@Pay_Year varchar(4)
			,@Emp_Dept varchar(35)
			)
returns Money

AS
Begin
declare @Amount money


if @Mode='Dept'

set @Amount=(select sum(amount) from payroll_sub where A_Gen in 
	(select A_Gen from payroll_main where pay_month=@pay_month
  		and pay_year=pay_year and pay_type=@pay_type and emp_id in 
			(Select Emp_Id from emp_job_hist_Current 
				where @Emp_dept=Emp_dept)))
if @Mode='All'

set @Amount=(select sum(amount) from payroll_sub where A_Gen in 
	(select A_Gen from payroll_main where pay_month=@pay_month
  		and pay_year=pay_year and pay_type=@pay_type))


if @Amount is null set @Amount =0

return @Amount 
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








---select dbo.GetWorkingDays('June','2002')

CREATE     Function GetWorkingDays(@pay_Month varchar(12),@pay_Year varchar(4))

Returns Int

AS

Begin

declare @M_Days varchar(12)
, @Febru_days varchar(2)
, @Tot_Hol  int
, @Working_Days int

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

-------------------------Holidays---------------------------------------------
set @Tot_Hol = (select count(*) from hol_list where 
(datepart(year,Str_date)=@pay_Year and datename(month,Str_date)=@Pay_month)and
(datepart(year,End_date)=@pay_Year and datename(month,End_date)=@Pay_month)
)

-------------------------Working Days-----------------------------------------
set @Working_Days=(@M_Days-@Tot_Hol)

return @Working_Days

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





CREATE      Function Get_If_In_Leave (@Emp_Id varchar(10),@Dt datetime)

returns Varchar(50)

AS

Begin


Declare @Result varchar(50)

	if Exists (select * from emp_leave_info 
		where Emp_Id=@Emp_Id 
		and @Dt between Start_Dt and End_Dt) 
	
		set @Result='Leave'

	else 

		if (select count(*) from Out_of_Office
			where Emp_Id=@Emp_Id 
			and @Dt between Str_Date and End_Date)!=0
	
		set @Result='Out of Office: '  + 
			(select [Description] from Out_of_Office
				where Emp_Id=@Emp_Id 
				and @Dt between Str_Date and End_Date)

	else

		set @Result =''

Return @Result
	
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

CREATE Function Get_pkeys(@table_name	sysname)

returns sysname
			
as
begin
	DECLARE @table_id		int
	DECLARE @full_table_name	nvarchar(255)
	DECLARE @Pk_Column varchar(50)

		SET @full_table_name = 'dbo.' + quotename(@table_name)


	/*	Get Object ID */
	SELECT @table_id = object_id(@full_table_name)

     select  @Pk_Column =convert(sysname,c.name)

	from
		sysindexes i, syscolumns c, sysobjects o --, syscolumns c1
	where
		o.id = @table_id
		and o.id = c.id
		and o.id = i.id
		and (i.status & 0x800) = 0x800
		--and c.name = index_col (@full_table_name, i.indid, c1.colid)
		and (c.name = index_col (@full_table_name, i.indid,  1) or
		     c.name = index_col (@full_table_name, i.indid,  2) or
		     c.name = index_col (@full_table_name, i.indid,  3) or
		     c.name = index_col (@full_table_name, i.indid,  4) or
		     c.name = index_col (@full_table_name, i.indid,  5) or
		     c.name = index_col (@full_table_name, i.indid,  6) or
		     c.name = index_col (@full_table_name, i.indid,  7) or
		     c.name = index_col (@full_table_name, i.indid,  8) or
		     c.name = index_col (@full_table_name, i.indid,  9) or
		     c.name = index_col (@full_table_name, i.indid, 10) or
		     c.name = index_col (@full_table_name, i.indid, 11) or
		     c.name = index_col (@full_table_name, i.indid, 12) or
		     c.name = index_col (@full_table_name, i.indid, 13) or
		     c.name = index_col (@full_table_name, i.indid, 14) or
		     c.name = index_col (@full_table_name, i.indid, 15) or
		     c.name = index_col (@full_table_name, i.indid, 16)
		    )
		--and c1.colid <= i.keycnt	/* create rows from 1 to keycnt */
		--and c1.id = @table_id

Return @Pk_Column 

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


--select dbo.Month_Name(1)

CREATE Function Month_Name(@Month_No int)

returns Varchar(12)

AS

BEGIN

declare @Month_Name varchar(12)
,@Month_Num varchar(2)

set @Month_Num =convert(varchar(2),@Month_No)



set @Month_Name=(CASE @Month_No
	when  '1' THEN 'January'
	when  '2'THEN  'February'
	when  '3'THEN  'March'
	when  '4'THEN  'April'
	when  '5'THEN  'May' 
	when  '6'THEN  'June'
	when  '7'THEN  'July'
	when  '8'THEN  'August'
	when  '9'THEN  'September'
	when  '10'THEN 'October'
	When  '11'THEN 'November' 
	When  '12'THEN 'December'
	end
)




return @Month_Name

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



--select dbo.Month_No('March')

Create  Function Month_No(@Month_Name varchar(12))

returns int

AS

BEGIN

declare @Month_No int


set @Month_No=(CASE @Month_Name
	when  'January' THEN '1'
	when  'February'THEN  '2'
	when  'March'THEN  '3'
	when  'April'THEN  '4'
	when  'May'THEN  '5' 
	when  'June'THEN  '6'
	when  'July'THEN  '7'
	when  'August'THEN  '8'
	when  'September'THEN  '9'
	when  'October'THEN '10'
	When  'November'THEN '11' 
	When  'December'THEN '12'
	end
)




return @Month_No

END






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




--select dbo.Padstring ('001',5,'0','L')
---select dbo.Padstring('010',5,'0','R')
Create    FUNCTION PadString
-- Input dimensions in centimeters
(
@Values VarChar(50), 
@ValueLength Int,
@PadChar Char(1),
@PadType Char(1)
)
RETURNS VarChar(50) 

--<<WE>> -- This is security symbol, Please do not remove it.
AS
Begin
DECLARE @ValuesLen int
DECLARE @DiffLen int
DECLARE @ReturnValues varchar(50)
DECLARE @MidPoint int
select @valuesLen=len(@Values)
select @DiffLen=(@ValueLength-@valuesLen)
select @ReturnValues =@Values
select @MidPoint=(@DiffLen/2)
while  @DiffLen>=1
Begin
   if @PadType = 'R'
	begin
		select @ReturnValues=@ReturnValues+@PadChar
	end
   if @PadType = 'C'
	begin
		if @DiffLen>= (@MidPoint+1)
			select @ReturnValues =@PadChar + @ReturnValues
		else
			select @ReturnValues =@ReturnValues + @PadChar
	end
   if @PadType = 'L'
	begin
		select @ReturnValues =@PadChar+@ReturnValues
	end
	select @DiffLen = @DiffLen -1
	continue
End
RETURN (@ReturnValues)
End



























GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO




--drop procedure RPT_MR

--SELECT * FROM RPT_MR ('S','1') ORDER BY fee_code

CREATE   function RPT_MR
(           @mode	   varchar(2),
            @Mr_no    integer
            )
     returns @TmpTable table(std_id  varchar(15),
                    std_name varchar(100),
                    class_id varchar(15),
                    class_title varchar(50),
                    Mon      varchar(5),
                    Yr       varchar(5),
                    collect_date datetime,
                    fee_code  varchar(5),
                    fee_title varchar(60),
                    act_amount decimal,   
                    discount  decimal,
                    fine      decimal,
                    userid   varchar(10),
                    username varchar(60)  
                    )
as
begin

DECLARE @std_id  varchar(15),
        @std_name varchar(100),
        @class_id varchar(15),
        @class_title varchar(50),
        @Mon      varchar(5),
        @Yr       varchar(5),
        @collect_date datetime,
        @fee_code  varchar(5),
        @fee_title  varchar(60),
        @act_amount decimal,
        @discount  decimal,
        @fine      decimal,
        @userid   varchar(10),
        @username   varchar(60)
       

DECLARE MyCursor CURSOR FOR
	SELECT Collec_master.Std_id, Collec_master.Class_id, Collec_master.Collec_date, Collec_master.Mon, Collec_master.Yr, Collec_details.Fee_code, 
	Collec_details.Act_Amount, Collec_details.Discount, Collec_details.Fine,Collec_master.Entry_by FROM Collec_master INNER JOIN
	Collec_details ON Collec_master.C_Srl = Collec_details.C_Srl WHERE     (Collec_master.C_Srl = @Mr_no)

	OPEN MyCursor
		FETCH NEXT FROM MyCursor INTO @std_id, @class_id, @collect_date, @Mon, @Yr, @fee_code, @act_amount, @discount, @fine,@userid
		WHILE @@FETCH_STATUS = 0 
		BEGIN
		   SELECT @fee_title = Fee_title FROM         Fee_info WHERE     (Fee_code = @fee_code)
           select @std_name=studentname from studentinfo where studentid=@Std_id
           select @class_title=ClassName from classinfo where ClassID=@Class_id 
           select @username=userName from userinfo where userid=@userid  
          
		   INSERT INTO @TmpTable VALUES (@std_id, @std_name,@class_id, @class_title,@Mon, @Yr, @collect_date, @fee_code, @fee_title,@act_amount, @discount, @fine,@userid,@username)
		FETCH NEXT FROM MyCursor INTO @std_id, @class_id, @collect_date, @Mon, @Yr, @fee_code, @act_amount, @discount, @fine,@userid
		END
	CLOSE MyCursor
	DEALLOCATE MyCursor


  DECLARE MyCursor1 CURSOR FOR
	SELECT Fee_setup.Fee_Code, Fee_info.Fee_title FROM Fee_setup INNER JOIN
	Fee_info ON Fee_setup.Fee_Code = Fee_info.Fee_code WHERE     (Fee_setup.Class_id = @class_id) AND (Fee_setup.Fee_Code NOT IN (SELECT fee_code FROM @TmpTable))
OPEN MyCursor1
		FETCH NEXT FROM MyCursor1 INTO @fee_code, @fee_title
		WHILE @@FETCH_STATUS = 0 
		BEGIN

			
			SELECT @fee_title = Fee_title FROM         Fee_info WHERE     (Fee_code = @fee_code)
			INSERT INTO @TmpTable VALUES (@std_id, @std_name,@class_id, @class_title,@Mon, @Yr, @collect_date, @fee_code, @fee_title,0, 0, 0,@userid,@username)
		FETCH NEXT FROM MyCursor1 INTO @fee_code, @fee_title
		END
	CLOSE MyCursor1
	DEALLOCATE MyCursor1


return
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

-- select * from Rpt_Department_Wise_Attendance ('Daffodil Software Ltd.','November',2004, '17 sep 2005') 

CREATE  function Rpt_Department_Wise_Attendance(
	@Dept varchar (30),
	@PayMonth varchar(12),
	@PayYear int,
	@Today datetime)

RETURNS @TmpTable  Table (
	emp_id varchar(20),
	Emp_fna varchar(45),
        emp_desig varchar(25),
	emp_dept varchar(20),   
	Month_NM varchar(20),           
	Work_Day int,
	Tot_Hol int,
	Weekend int,
	Present int,
	Late int,
	Leave int,
	Absent int)
as
begin
	DECLARE 
	@emp_id varchar(20),
	@Emp_fna varchar(45),
        @emp_desig varchar(25),
	@emp_dept varchar(20),   
	@Month_NM varchar(20),           
	@Work_Day int,
	@Tot_Hol int,
	@Weekend int,
	@Present int,
	@Late int,
	@Leave int,
	@Absent int,
	@StartDate datetime,
	@EndingDate datetime,
	@MonthDays int

	DECLARE MyCursor CURSOR FOR
		SELECT Emp_id FROM Emp_Job_Hist_Current WHERE (Emp_dept = @Dept)
							
		Open MyCursor
				FETCH NEXT FROM MyCursor into @emp_id
				WHILE @@FETCH_STATUS = 0 
				Begin
					select @Emp_fna = Emp_Nm, @emp_desig = Emp_Desig, @emp_dept = Emp_Dept, @Month_NM = Month_Year, @Leave = Leave, @Tot_Hol = Hol_Day, @Weekend = Weekend, @Work_Day = Work_Day, @Late = Late, @Present = Attn, @Absent = Absent   from dbo.Monthly_Attn_Summery(@PayMonth, @PayYear, @emp_id) 
					INSERT INTO @TmpTable VALUES (@emp_id, @Emp_fna, @emp_desig, @emp_dept, @Month_NM, @Work_Day, @Tot_Hol, @Weekend, @Present, @Late, @Leave, @Absent)

				FETCH NEXT FROM MyCursor into @emp_id
			    End
	CLOSE MyCursor
	DEALLOCATE MyCursor

return
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


---exec Rpt_marks_Sheet 'a','STI-000001','2006','00001','01','01'
--SELECT * FROM Rpt_marks_Sheet('a','STI-000001','2006','00001','01','01')
CREATE   function Rpt_marks_Sheet
(
    @mode           varchar(1),
	@std_id         varchar(10),
    @AcaYr          varchar(5),
    @ClassID        varchar(10),
    @ExamType       varchar(5),
    @ExamID         varchar(5)
)
returns @marksheettemp table
    (
     subid    varchar(10),
     subjecttitle varchar(40),
     category   varchar(10),
     categorytitle varchar(50),
     obtainedmarks int,
     passmarks     int,
     fullmarks     int
)
as
  begin
declare @serial as  int
declare @p_marks as  int
declare @f_marks as  int
declare @subjectid as varchar(10)
declare @markscate as varchar(10) 
declare @sub_title as varchar(60)
declare @cat_title as varchar(60)
 
if @mode='a' 

declare mycursor1 cursor for 
   select  M_Slr_no,SubID,categoryid from result_main 
       where ClassID=@ClassID and AcaYr=@AcaYr
             and ExamType= @ExamType
             and ExamID=@ExamID 
open mycursor1
fetch next from mycursor1 into @serial,@subjectid,@markscate
while @@fetch_status=0
   begin
   insert into @marksheettemp  select  @subjectid,(select Sub_title 
                            from subject_info_sub
                         where Sub_code=@subjectid
                               and Class_code=@ClassID) as sub_title,
     @markscate,(select MCategoryDsc from markscategory 
         where MCategoryID=@markscate) as category,
    ObtainedMarks,PassMarks,Fullmarks from result_sub 
       where M_Slr_no=@serial
             and StdID=@std_id


   fetch next from mycursor1 into @serial,@subjectid,@markscate
  end 
close mycursor1
deallocate mycursor1




declare mycursor2 cursor for
  select Sub_code from   subject_info_sub
   where Sub_code not in (select subid from @marksheettemp) 
                               and Class_code=@ClassID

open mycursor2
fetch next from  mycursor2 into @subjectid
while @@fetch_status =0 
   begin
      select @markscate=CategoryID,@p_marks=passmarks,@f_marks=fullmarks
       from subjectmarksdistribution
    where ClassID=@ClassID and SubjectID=@subjectid 
         and term_code=@ExamType and Exam_code= @ExamID 

     select @sub_title=Sub_title 
           from subject_info_sub
         where Sub_code=@subjectid
        and Class_code=@ClassID

      select @cat_title=MCategoryDsc 
         from markscategory where MCategoryID=@markscate

     insert into @marksheettemp(subid,subjecttitle,category,
                  categorytitle,obtainedmarks,passmarks,
                  fullmarks ) values (@subjectid,@sub_title,@markscate,
                  @cat_title,0,@p_marks,@f_marks)
    
   
 
fetch next from mycursor2 into @subjectid
   end
close mycursor2
deallocate mycursor2



return

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

/****** Encrypted object is not transferable, and script can not be generated. ******/

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

/****** Encrypted object is not transferable, and script can not be generated. ******/

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
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

---select dbo.ShowMaxiamumEmpID()
---select isnull(max(cast(substring(Emp_Id,4,3) as int)),0)FROM Emp_Per_Info 



CREATE  Function ShowMaxiamumEmpID()

returns char(3)

as
begin

declare 
	@Conv_Value char(8)

	set @Conv_Value=(select isnull(max(cast(substring(Emp_Id,4,3) as int)),0)FROM Emp_Per_Info )
	

return @Conv_Value

end




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

