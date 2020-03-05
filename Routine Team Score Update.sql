Declare @sapid int = 0, @maxRows int = 0, @count int = 0;
Set @maxRows	= (Select Count(sap_id) from users where user_role = 'Team Leader' and (sub_department = 'Kaiser Closet'));
		
Declare newCursor cursor FAST_FORWARD for Select sap_id from users where user_role = 'Team Leader' and (sub_department = 'Kaiser Closet');
		open newCursor
		while @count < @maxRows
			BEGIN
				fetch newCursor into @sapid;   
				Exec dbo.update_TeamScoresKaiserCloset @sap_id = @sapid;
			Set @count = @count + 1;
			END 
		close newCursor 
		deallocate newCursor 
Go