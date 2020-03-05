DECLARE  
		@firstday date,
		@rank SMALLINT, 
		@count SMALLINT, 
		@maxrows SMALLINT, 
		@sap int, 
		@eom_score decimal(18,2),
		@prevmos SMALLINT,
		@currmos SMALLINT;
		SET @firstday		= GETDATE()-DAY(GETDATE())+1;
		SET @currmos		= MONTH(GETDATE());
		SET @count		= 0;
		SET @maxrows 		= (Select count(*) from users where sub_department = 'Kaiser SMC Resupply');
	Declare FirstCursor cursor READ_ONLY for 
	SELECT RANK() OVER(ORDER BY eom_score DESC) as rank, sapno, eom_score  from(
		Select (aveprodscore *0.45) + (QualityScore *0.3) + (AbsenteeismScore *0.15)+ (LmsScore *0.1) as eom_score, sapno from(
		SELECT  
		(select sum(ave_prod_score)/nullif(count(ave_prod_score),0) from kaiser_smc_overall where date > @firstday and sap_id = sapno) as aveprodscore,
		dbo.fn_EqQual((
		(select count(*) from audit where agent_sap = sapno and  audit_date >=  @firstday and bc = 1)/nullif((select count(*) from audit where agent_sap = sapno and  audit_date >=  @firstday ),0)),
		((select count(*) from audit where agent_sap = sapno and  audit_date >=  @firstday and euc = 1)/nullif((select count(*) from audit where agent_sap = sapno and  audit_date >=  @firstday ),0)),
		((select count(*) from audit where agent_sap = sapno and  audit_date >=  @firstday and cc = 1)/nullif((select count(*) from audit where agent_sap = sapno and  audit_date >=  @firstday ),0))) as QualityScore,
		dbo.fn_EqAbs(
		(select isnull((Select CONVERT(Decimal(18,2),
			(Select Count(*) from (Select day,shift,
			(select top 1 finesse_detail_event_datetime from finesse where CAST(FLOOR(CAST(finesse_detail_event_datetime as FLOAT)) as DateTime) = day and sap_id = sapno) as login 
			from (select distinct(day) as day, sap_id as sapno, shift from schedule where sap_id = sapno and month(day) = @currmos) tb) tb2 where login IS NULL and shift != 'OFF'  and shift != 'Leave-Vacation' and shift != 'Leave-Planned' and shift != 'Leave-Bereavement' and shift != 'Leave-LOA' and shift != 'Leave-Maternity' and shift != 'Leave-Emergency' and day < getdate()-1)))
		/nullif(
		(Select Count(*) from schedule where sap_id = sapno and month(day) = @currmos and shift != 'OFF'  and shift != 'Leave-Vacation' and shift != 'Leave-Planned' and shift != 'Leave-Bereavement' and shift != 'Leave-LOA' and shift != 'Leave-Maternity' and shift != 'Leave-Emergency' and day < getdate()-1)
		,0),0))) as AbsenteeismScore,
		 isnull(dbo.fn_EqComp(
		 isnull((select (isnull(week1marks,0) + isnull(week2marks,0) + isnull(week3marks,0) + isnull(week4marks,0) + isnull(week5marks,0))/Nullif(isnull(week1marks/nullif(week1marks,0),0)+isnull(week2marks/nullif(week2marks,0),0)+isnull(week3marks/nullif(week3marks,0),0)+isnull(week4marks/nullif(week4marks,0),0)+isnull(week5marks/nullif(week5marks,0),0),0) as monthmarks from
			(select 
				(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week1,10)) as week1marks,
				(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week2,10)) as week2marks,
				(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week3,10)) as week3marks,
				(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week4,10)) as week4marks,
				(select top 1 marks_obtained from wpu where wpu.sap_id = sapno and left(wpu.week,10)= left(week5,10)) as week5marks
			from
			(select week1, week2, week3, week4, week5 from monthweek where month = MONTH( GETDATE()) and year = Year(GETDATE())) tb) tbl),0)
		 ,(Select top 1 IIF(status='Completed' And completed_date<due_date,5,IIF(status='Completed', 3, 1)) as LmsScore from lms where user_id=sapno)),1) as LmsScore
		 ,sapno , nm from (Select sap_id as sapno,Convert(decimal(10,2),(Select NULLIF(Count(Distinct Date),0) as DaysWorked from dbo.eq_prod where SAP_ID = users.sap_id and Date >=  @firstday and Year(Date) >= Year(GetDate()-30))) as workdays, name as nm from users where users.sub_department='Kaiser SMC Resupply')tb )pio) poi

		 open FirstCursor
		while @count < @maxRows
			begin
				fetch FirstCursor into @rank, @sap, @eom_score; 
				if(@eom_score is not null)
				begin 
					update TOP (1) kaiser_smc_overall set rank = @rank, eom_score = @eom_score where 
					kaiser_smc_overall_id = (select top 1 kaiser_smc_overall_id from kaiser_smc_overall where sap_id= @sap order by date desc)
				end
				set @count = @count + 1;
			end   
	close FirstCursor 
	deallocate FirstCursor 