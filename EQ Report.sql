Select 
(SELECT Count(eq_prod_id)
FROM eq_prod where 
(Task_Resolution = 'REQUEUE' OR Task_Resolution = 'COMPLETE' OR Task_Resolution = 'CANCEL') and (queue_name = 'SLP CEN WEST NG FAX EQ' ) and Month(eq_prod.Date)=1 and 
SAP_ID = sapid) as TotalProd, 
(SELECT Count(eq_prod_id)
FROM eq_prod where 
(Task_Resolution = 'REQUEUE') and (queue_name = 'SLP CEN WEST NG FAX EQ' )  and Month(eq_prod.Date)=1 and 
SAP_ID = sapid) as TotalRequeue, 
(SELECT Count(eq_prod_id)
FROM eq_prod where 
(Task_Resolution = 'COMPLETE') and (queue_name = 'SLP CEN WEST NG FAX EQ' ) and Month(eq_prod.Date)=1 and 
SAP_ID = sapid) as TotalCompletes, 
(SELECT Count(eq_prod_id)
FROM eq_prod where 
(Task_Resolution = 'CANCEL') and (queue_name = 'SLP CEN WEST NG FAX EQ' ) and Month(eq_prod.Date)=1 and  
SAP_ID = sapid) as TotalCancels, 

(Select Count(Distinct(Date)) from eq_prod where SAP_ID = sapid and Month(eq_prod.Date)=1) as Days, 
(Select top 1 Name from eq_prod where SAP_ID = sapid) as Name,
sapid
from 
(Select Distinct(eq_prod.SAP_ID) as sapid from eq_prod) tb