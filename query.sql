DECLARE  
		@currmos SMALLINT,
		@score decimal(18,2), 
		@firstday date,
		@count int,
		@maxRows int,
		@sap int,@aveprod decimal(18,2), 
		@name nvarchar(max);
	Set @count		= 0;
	Set @maxRows		= (SELECT Count(*) from users where sub_department ='Sleep EQ')-1;
	SET @currmos	= MONTH(GETDATE()); 
	SET @firstday		= GETDATE()-DAY(GETDATE())+1;

Select dbo.fn_EqProdScore(AveProd,CompletePercent,OTC,cc,euc,bc,Absenteeism,isnull(wpu,0), isnull(lms,1)) as ProdScore,AveProd, sapno, nm from(
Select(
SELECT Count(*)/workdays FROM eq_prod WHERE SAP_ID =sapno and Date >=  @firstday) as AveProd,
 
		Convert(Decimal(18,2),(Select Count(*) from eq_prod where Task_Resolution = 'COMPLETE' and SAP_ID =sapno  and Date >=  @firstday))/
		Nullif((Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and SAP_ID =sapno  and Date >=  @firstday),0) as CompletePercent,
		-- Instead of IIF sum of both returns may be more accurate returning only one condition
		 (SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID = sapno and Date >=  @firstday),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID = sapno  and Date >=  @firstday)) from eq_prod
		where SAP_ID = sapno and Date >=  @firstday)/
		 nullif(Convert(Decimal(18,2),(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (10*24)) and 
			SAP_ID = sapno and Date >=  @firstday),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) < (5*24)) and 
			SAP_ID = sapno  and Date >=  @firstday)) from eq_prod
		where SAP_ID = sapno and Date >=  @firstday)+
		(SELECT top 1 
		IIF(queue_name = 'SLP CEN WEST AUTH EQ',
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (10*24)) and
			SAP_ID = sapno  and Date >=  @firstday),
			(Select Count(*) from eq_prod where Task_Status = 'RESOLVED' and (DATEDIFF(hh,Task_Create_Date_Time,ISNULL(Task_Resolved_Timestamp,GETDATE())) > (5*24)) and
			SAP_ID = sapno  and Date >=  @firstday)) from eq_prod
		where SAP_ID = sapno and Date >=  @firstday)),0) as OTC,
			isnull(1.0 - ((select count(*) from audit where agent_sap = sapno and audit_date >=  @firstday and euc = 1)/
			nullif(Convert(Decimal(18,2),(select count(*) from audit where agent_sap = sapno and  audit_date >=  @firstday)),0)),0) as EUC,
			isnull(1.0 - ((select count(*) from audit where agent_sap = sapno and  audit_date >=  @firstday and bc = 1)/
			nullif(Convert(Decimal(18,2),(select count(*) from audit where agent_sap = sapno and  audit_date >=  @firstday)),0)),0) as BC,
			isnull(1.0 - ((select count(*) from audit where agent_sap = sapno and  audit_date >=  @firstday and cc = 1)/
			nullif(Convert(Decimal(18,2),(select count(*) from audit where agent_sap = sapno and  audit_date >=  @firstday)),0)),0) as CC,
		 
		(Select isnull(CONVERT(Decimal(18,2),(Select Count(*) from 
		(Select 
			day, 
			shift,
			(select top 1 finesse_detail_event_datetime from finesse where CAST(FLOOR(CAST(finesse_detail_event_datetime as FLOAT)) as DateTime) = day and sap_id = sapno) as login from
		(select distinct(day) as day, sap_id as sapno, shift from schedule where sap_id = sapno and day >=  @firstday) tb) tb2 where login IS NULL and shift != 'OFF'  and day < getdate()-1))
		/nullif((Select Count(*) from schedule where sap_id = sapno and day >=  @firstday and shift != 'OFF'),0),0)) as Absenteeism,
		(select (isnull(week1marks,0) + isnull(week2marks,0) + isnull(week3marks,0) + isnull(week4marks,0) + isnull(week5marks,0))/Nullif(isnull(week1marks/nullif(week1marks,0),0)+isnull(week2marks/nullif(week2marks,0),0)+isnull(week3marks/nullif(week3marks,0),0)+isnull(week4marks/nullif(week4marks,0),0)+isnull(week5marks/nullif(week5marks,0),0),0) as monthmarks from
		(select 
		(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week1,10)) as week1marks,
		(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week2,10)) as week2marks,
		(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week3,10)) as week3marks,
		(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week4,10)) as week4marks,
		(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week5,10)) as week5marks
		from
		(select week1, week2, week3, week4, week5, sapno from monthweek where month = MONTH( GETDATE()) and year = Year(GETDATE())) tb1) tb2) as wpu,
		 (Select top 1 IIF(status='Completed' And completed_date<due_date,5,IIF(status='Completed', 3, 1)) as LmsScore from lms where user_id=sapno) as lms
		 ,sapno , nm
		from (Select sap_id as sapno,Convert(decimal(10,2),(Select NULLIF(Count(Distinct Date),0) as DaysWorked from dbo.eq_prod where SAP_ID = users.sap_id and Date >=  @firstday and Year(Date) >= Year(GetDate()-30))) as workdays, name as nm from users)tb)pio where AveProd is not null order by prodscore desc 