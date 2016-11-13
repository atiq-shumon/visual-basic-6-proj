SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


ALTER        procedure Save_VaccineInfo
(	
	@VaccineID			varchar(3),
	@VaccineName	 		Varchar(70),
	@Entry_BY			varchar(10),
	@Entry_Date			DateTime	
)

as
if exists (select * from VaccineInfo where  VaccineID= @VaccineID ) 		
Update VaccineInfo set  				
	VaccineID	=	@VaccineID,
	VaccineName	=	@VaccineName,
	Entry_BY	=	@Entry_BY,
	Entry_Date	=	@Entry_Date			
	
where VaccineID= @VaccineID
else
insert into VaccineInfo
(
	VaccineID,
	VaccineName,
	Entry_BY,
	Entry_Date	
)
values
(
	@VaccineID,
	@VaccineName,
	@Entry_BY,
	@Entry_Date		
)


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

