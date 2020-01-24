 
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
	SET @workdaysweek	= 5.0;

	Select AveProdPrevRank,CompletePercentPrevRank,OTCPrevRank,AveProdRank,CompletePercentRank,OTCRank,AveProdWk1Rank,CompletePercentWk1Rank,OTCWk1Rank,AveProdWk2Rank,CompletePercentWk2Rank,OTCWk2Rank,AveProdWk3Rank,CompletePercentWk3Rank,OTCWk3Rank,AveProdWk4Rank,CompletePercentWk4Rank,OTCWk4Rank,AveProdWk5Rank,CompletePercentWk5Rank,OTCWk5Rank  from(
	SELECT
	RANK() OVER(ORDER BY AveProdPrev DESC) as AveProdPrevRank,
	RANK() OVER(ORDER BY IIF(CompletePercentPrev IS NULL, 1, 0),CompletePercentPrev DESC) as CompletePercentPrevRank, 
	RANK() OVER(ORDER BY IIF(OTCPrev IS NULL, 1, 0),OTCPrev desc) as OTCPrevRank,  
	RANK() OVER(ORDER BY IIF(AveProd IS NULL, 1, 0),AveProd desc) as AveProdRank, 
	RANK() OVER(ORDER BY IIF(CompletePercent IS NULL, 1, 0),CompletePercent desc) as CompletePercentRank,
	RANK() OVER(ORDER BY IIF(OTC IS NULL, 1, 0),OTC desc) as OTCRank,   
	RANK() OVER(ORDER BY IIF(AveProdWk1 IS NULL, 1, 0),AveProdWk1 desc) as AveProdWk1Rank, 
	RANK() OVER(ORDER BY IIF(CompletePercentWk1 IS NULL, 1, 0),CompletePercentWk1 desc) as CompletePercentWk1Rank,
	RANK() OVER(ORDER BY IIF(OTCWk1 IS NULL, 1, 0),OTCWk1 desc) as OTCWk1Rank,  
	RANK() OVER(ORDER BY IIF(AveProdWk2 IS NULL, 1, 0),AveProdWk2 desc) as AveProdWk2Rank, 
	RANK() OVER(ORDER BY IIF(CompletePercentWk2 IS NULL, 1, 0),CompletePercentWk2 desc) as CompletePercentWk2Rank,
	RANK() OVER(ORDER BY IIF(OTCWk2 IS NULL, 1, 0),OTCWk2 desc) as OTCWk2Rank,  
	RANK() OVER(ORDER BY IIF(AveProdWk3 IS NULL, 1, 0),AveProdWk3 desc) as AveProdWk3Rank, 
	RANK() OVER(ORDER BY IIF(CompletePercentWk3 IS NULL, 1, 0),CompletePercentWk3 desc) as CompletePercentWk3Rank,
	RANK() OVER(ORDER BY IIF(OTCWk3 IS NULL, 1, 0),OTCWk3 desc) as OTCWk3Rank,  
	RANK() OVER(ORDER BY IIF(AveProdWk4 IS NULL, 1, 0),AveProdWk4 desc) as AveProdWk4Rank, 
	RANK() OVER(ORDER BY IIF(CompletePercentWk4 IS NULL, 1, 0),CompletePercentWk4 desc) as CompletePercentWk4Rank,
	RANK() OVER(ORDER BY IIF(OTCWk4 IS NULL, 1, 0),OTCWk4 desc) as OTCWk4Rank,  
	RANK() OVER(ORDER BY IIF(AveProdWk5 IS NULL, 1, 0),AveProdWk5 desc) as AveProdWk5Rank, 
	RANK() OVER(ORDER BY IIF(CompletePercentWk5 IS NULL, 1, 0),CompletePercentWk5 desc) as CompletePercentWk5Rank,
	RANK() OVER(ORDER BY IIF(OTCWk5 IS NULL, 1, 0),OTCWk5 desc) as OTCWk5Rank,  
	sapno
	
	FROM
	(SELECT 
	-- Previous Month 
		(SELECT Count(*)/@workdaysprev FROM eq_prod WHERE SAP_ID =sapno and MONTH(Date) = @prevmos) as AveProdPrev,
		(Select Count(*) from eq_prod where Task_Resolution = 'COMPLETE' and SAP_ID =sapno  and MONTH(Date) = @prevmos)/
		Nullif((Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and SAP_ID =sapno  and MONTH(Date) = @prevmos),0) as CompletePercentPrev,

		-- Instead of IIF sum of both returns may be more accurate returning only one condition
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID =sapno and MONTH(Date) = @prevmos),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID =sapno  and MONTH(Date) = @prevmos)) from eq_prod
		where SAP_ID =sapno and MONTH(Date) = @prevmos)/NULLIF((SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (10*24)) and
			SAP_ID =sapno  and MONTH(Date) = @prevmos),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (5*24)) and
			SAP_ID =sapno  and MONTH(Date) = @prevmos)) from eq_prod
		where SAP_ID =sapno and MONTH(Date) = @prevmos),0) as OTCPrev,

		(SELECT Count(*)/@workdays FROM eq_prod WHERE SAP_ID =sapno and MONTH(Date) = @currmos) as AveProd,
		(Select Count(*) from eq_prod where Task_Resolution = 'COMPLETE' and SAP_ID =sapno  and MONTH(Date) = @currmos)/
		Nullif((Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and SAP_ID =sapno  and MONTH(Date) = @currmos),0) as CompletePercent,
		-- Instead of IIF sum of both returns may be more accurate returning only one condition
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID =sapno and MONTH(Date) = @currmos),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID =sapno  and MONTH(Date) = @currmos)) from eq_prod
		where SAP_ID =sapno and MONTH(Date) = @currmos)/Nullif((SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (10*24)) and
			SAP_ID =sapno  and MONTH(Date) = @currmos),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (5*24)) and
			SAP_ID =sapno  and MONTH(Date) = @currmos)) from eq_prod
		where SAP_ID =sapno and MONTH(Date) = @currmos),0) as OTC,
		-- Week 1
		(SELECT Count(*)/@workdaysweek FROM eq_prod WHERE SAP_ID =sapno and DATEPART(wk,Date) = @week1) as AveProdWk1,
		(Select Count(*) from eq_prod where Task_Resolution = 'COMPLETE' and SAP_ID =sapno  and DATEPART(wk,Date) = @week1)/
		Nullif((Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and SAP_ID =sapno  and DATEPART(wk,Date) = @week1),0) as CompletePercentWk1,

		-- Instead of IIF sum of both returns may be more accurate returning only one condition
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID =sapno and DATEPART(wk,Date) = @week1),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1)) from eq_prod
		where SAP_ID =sapno and DATEPART(wk,Date) = @week1)/Nullif(
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (10*24)) and
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (5*24)) and
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1)) from eq_prod
		where SAP_ID =sapno and DATEPART(wk,Date) = @week1),0) as OTCWk1,

		-- Week 2
		(SELECT Count(*)/@workdaysweek FROM eq_prod WHERE SAP_ID =sapno and DATEPART(wk,Date) = (@week1 + 1)) as AveProdWk2,
		(Select Count(*) from eq_prod where Task_Resolution = 'COMPLETE' and SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 1)/Nullif(
		(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 1),0) as CompletePercentWk2,

		-- Instead of IIF sum of both returns may be more accurate returning only one condition
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 1),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 1)) from eq_prod
		where SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 1)/Nullif(
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (10*24)) and
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 1),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (5*24)) and
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 1)) from eq_prod
		where SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 1),0) as OTCWk2,

		-- Week 3
		(SELECT Count(*)/@workdaysweek FROM eq_prod WHERE SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 2) as AveProdWk3,
		(Select Count(*) from eq_prod where Task_Resolution = 'COMPLETE' and SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 2)/Nullif(
		(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 2),0) as CompletePercentWk3,

		-- Instead of IIF sum of both returns may be more accurate returning only one condition
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 2),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 2)) from eq_prod
		where SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 2)/Nullif(
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (10*24)) and
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 2),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (5*24)) and
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 2)) from eq_prod
		where SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 2),0) as OTCWk3,

		-- Week 4
		(SELECT Count(*)/@workdaysweek FROM eq_prod WHERE SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 3) as AveProdWk4,
		(Select Count(*) from eq_prod where Task_Resolution = 'COMPLETE' and SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 3)/Nullif(
		(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 3),0) as CompletePercentWk4,

		-- Instead of IIF sum of both returns may be more accurate returning only one condition
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 3),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 3)) from eq_prod
		where SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 3)/Nullif(
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (10*24)) and
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 3),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (5*24)) and
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 3)) from eq_prod
		where SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 3),0) as OTCWk4,

		-- Week 5
		(SELECT Count(*)/@workdaysweek FROM eq_prod WHERE SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 4) as AveProdWk5,
		(Select Count(*) from eq_prod where Task_Resolution = 'COMPLETE' and SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 4)/Nullif(
		(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 4),0) as CompletePercentWk5,

		-- Instead of IIF sum of both returns may be more accurate returning only one condition
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 4),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 4)) from eq_prod
		where SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 4)/Nullif(
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (10*24)) and
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 4),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (5*24)) and
			SAP_ID =sapno  and DATEPART(wk,Date) = @week1 + 4)) from eq_prod
		where SAP_ID =sapno and DATEPART(wk,Date) = @week1 + 4),0) as OTCWk5,
		sapno 
		from 
		(select sap_id as sapno from users) tb_user) tb_data ) tb where sapno = 51763970