	DECLARE  
		@currmos SMALLINT; 

	SET @currmos		= MONTH(GETDATE()); 
	Select top 5 dbo.fn_EqProdScore(AveProd,CompletePercent,OTC,cc,euc,bc,Absenteeism,wpu, lms) as ProdScore, sapno from 
	(Select(
SELECT Count(*)/workdays FROM eq_prod WHERE SAP_ID =sapno and MONTH(Date) = @currmos) as AveProd,
 
		Convert(Decimal(18,2),(Select Count(*) from eq_prod where Task_Resolution = 'COMPLETE' and SAP_ID =sapno  and MONTH(Date) = @currmos))/
		Nullif((Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and SAP_ID =sapno  and MONTH(Date) = @currmos),0) as CompletePercent,
		-- Instead of IIF sum of both returns may be more accurate returning only one condition
		 (SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID = sapno and MONTH(Date) = @currmos),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID = sapno  and MONTH(Date) = @currmos)) from eq_prod
		where SAP_ID = sapno and MONTH(Date) = @currmos)/
		 Convert(Decimal(18,2),(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID = sapno and MONTH(Date) = @currmos),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID = sapno  and MONTH(Date) = @currmos)) from eq_prod
		where SAP_ID = sapno and MONTH(Date) = @currmos)+
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (10*24)) and
			SAP_ID = sapno  and MONTH(Date) = @currmos),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (5*24)) and
			SAP_ID = sapno  and MONTH(Date) = @currmos)) from eq_prod
		where SAP_ID = sapno and MONTH(Date) = @currmos)) as OTC,
		1.0 - ((select count(*) from audit where agent_sap = sapno and month(audit_date) = @currmos and euc = 1)/
		nullif((select count(*) from audit where agent_sap = sapno and month(audit_date) = @currmos),0)) as EUC,
		1.0 - ((select count(*) from audit where agent_sap = sapno and month(audit_date) = @currmos and bc = 1)/
		nullif((select count(*) from audit where agent_sap = sapno and month(audit_date) = @currmos),0)) as BC,
		1.0 - ((select count(*) from audit where agent_sap = sapno and month(audit_date) = @currmos and cc = 1)/
		nullif((select count(*) from audit where agent_sap = sapno and month(audit_date) = @currmos),0)) as CC,
		 
		(Select CONVERT(Decimal(18,2),(Select Count(*) from 
		(Select 
			day, 
			shift,
			(select top 1 finesse_detail_event_datetime from finesse where CAST(FLOOR(CAST(finesse_detail_event_datetime as FLOAT)) as DateTime) = day and sap_id = sapno) as login from
		(select distinct(day) as day, sap_id as sapno, shift from schedule where sap_id = sapno and  month(day) = @currmos) tb) tb2 where login IS NULL and shift != 'OFF'))
		/nullif((Select Count(*) from schedule where sap_id = sapno and month(day) = @currmos and shift != 'OFF'),0)) as Absenteeism,
		(select  top 1
			(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week1,10)) +
			(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week2,10)) +
			(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week3,10)) +
			(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week4,10)) +
			(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week5,10)) as wpu
		from
		(select week1, week2, week3, week4, week5, sapno from monthweek where month = MONTH( GETDATE()) and year = Year(GETDATE())) tbl1) as wpu
		 ,
		 (Select top 1 IIF(status='Completed' And completed_date<due_date,5,IIF(status='Completed', 3, 1)) as LmsScore from lms where user_id=sapno) as lms
		 ,sapno 
		from (Select sap_id as sapno,Convert(decimal(10,2),(Select NULLIF(Count(Distinct Date),0) as DaysWorked from dbo.eq_prod where SAP_ID = users.sap_id and Month(Date) = @currmos and Year(Date) = Year(GetDate()))) as workdays, name from users)tb) pio order by prodscore desc