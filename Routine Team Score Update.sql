Declare @sapid int = 0, @maxRows int = 0, @count int = 0;
Set @maxRows	= (Select Count(sap_id) from users where user_role = 'Team Leader' and sub_department = 'Sleep EQ');
		
Declare newCursor cursor FAST_FORWARD for Select sap_id from users where user_role = 'Team Leader' and sub_department = 'Sleep EQ';
		open newCursor
		while @count < @maxRows
			BEGIN
				fetch newCursor into @sapid;   
				Exec dbo.update_TeamScores @sap_id = @sapid;
			Set @count = @count + 1;
			END 
		close newCursor 
		deallocate newCursor 
Go