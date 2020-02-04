

Select * from users where group_id = (Select top 1 group_id from [group] where group_leader = (Select top 1 user_id from users where sap_id=51691175))