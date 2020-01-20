 
	DECLARE 
		@workdays SMALLINT,
		@workdaysprev SMALLINT,
		@workdaysweek SMALLINT,
		@week1 SMALLINT,
		@prevmos SMALLINT,
		@currmos SMALLINT;

	SET @week1			= DATEPART(WK,DATEFROMPARTS(YEAR(GETDATE()),MONTH(GETDATE()),1)); 
	SET @prevmos		= MONTH(DATEADD(M,-1, GETDATE()));
	SET @currmos		= MONTH(GETDATE());
	SET @workdays		= Convert(decimal(10,2),(Select NULLIF(Count(Distinct Date),0) as DaysWorked from dbo.eq_prod where SAP_ID = 51763970 and Month(Date) = @currmos and Year(Date) = Year(GetDate())));
	SET @workdaysprev	= Convert(decimal(10,2),(Select NULLIF(Count(Distinct Date),0) as DaysWorked from dbo.eq_prod where SAP_ID = 51763970 and Month(Date) = @prevmos and Year(Date) = Year(DATEADD(MONTH, DATEDIFF(MONTH, -1, GETDATE())-1, -1))));
	SET @workdaysweek	= 7.0;

	Select AveProdPrevRank,CompletePercentPrevRank,OTCPrevRank from(
	SELECT 
	RANK() OVER(ORDER BY AveProdPrev DESC) as AveProdPrevRank,
	RANK() OVER(ORDER BY IIF(CompletePercentPrev IS NULL, 1, 0),CompletePercentPrev DESC) as CompletePercentPrevRank, 
	RANK() OVER(ORDER BY IIF(OTCPrev IS NULL, 1, 0),OTCPrev desc) as OTCPrevRank, sapno
	
	FROM
	(SELECT 
	-- Previous Month 
		(SELECT COUNT(*)/@workdaysprev FROM eq_prod WHERE SAP_ID = tb_user.sapno and MONTH(Date) = @prevmos) as AveProdPrev,
		(Select Convert(DEcimal(18,2),COUNT(*)) from eq_prod where Task_Resolution = 'COMPLETE' and SAP_ID = tb_user.sapno  and MONTH(Date) = @prevmos) /
		(Select nullif(COUNT(*),0) from eq_prod where Task_Status = 'RESOLVED' and SAP_ID = tb_user.sapno  and MONTH(Date) = @prevmos ) as CompletePercentPrev,

		-- Instead of IIF sum of both returns may be more accurate returning only one condition
		-- WithinSLA /
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Convert(DEcimal(18,2),COUNT(*)) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID = tb_user.sapno and MONTH(Date) = @prevmos),
			(Select Convert(DEcimal(18,2),COUNT(*)) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID = tb_user.sapno  and MONTH(Date) = @prevmos)) from eq_prod
		where SAP_ID = tb_user.sapno and MONTH(Date) = @prevmos)/(
		-- WithinSLA + NotWithinSLA
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Convert(DEcimal(18,2),COUNT(*)) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID = tb_user.sapno and MONTH(Date) = @prevmos),
			(Select Convert(DEcimal(18,2),COUNT(*)) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID = tb_user.sapno  and MONTH(Date) = @prevmos)) from eq_prod
		where SAP_ID = tb_user.sapno and MONTH(Date) = @prevmos)
		+
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Convert(DEcimal(18,2),COUNT(*)) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (10*24)) and
			SAP_ID = tb_user.sapno  and MONTH(Date) = @prevmos),
			(Select Convert(DEcimal(18,2),COUNT(*)) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (5*24)) and
			SAP_ID = tb_user.sapno  and MONTH(Date) = @prevmos)) from eq_prod
		where SAP_ID = tb_user.sapno and MONTH(Date) = @prevmos)) as OTCPrev,
		 
		sapno 
		from 
		(select sap_id as sapno from users) tb_user) tb_data ) tb where sapno = 51763970