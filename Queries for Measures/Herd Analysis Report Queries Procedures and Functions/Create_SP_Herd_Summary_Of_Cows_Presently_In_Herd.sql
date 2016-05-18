
DELIMITER //
DROP PROCEDURE Summary_Of_Cows_Present_In_Herd //
Create PROCEDURE Summary_Of_Cows_Present_In_Herd(Birth_Start_Date date,Birth_End_Date date,Bull_TurnOut_Date date)
BEGIN
DECLARE est_turn_date DATE;
set est_turn_date=Estimated_Bull_Turnout_Date(Birth_Start_Date, Birth_End_Date, Bull_Turnout_Date);
SELECT DISTINCT 
CASE b.Measures
	WHEN 'Total_Cows_Exposed' THEN 'Total_Cows_Exposed'
	WHEN 'Total_Cows_Kept_For_Calving' THEN 'Total_Cows_Kept_For_Calving'
	WHEN 'Number_Of_Cows_Aborted' THEN 'Number_Of_Cows_Aborted'
	WHEN 'Number_Of_Cows_Open' THEN 'Number_Of_Cows_Open'
	WHEN 'Number_Of_Cows_Calving' THEN 'Number_Of_Cows_Calving'
	WHEN 'Number_Of_Cows_Losing_Calf' THEN 'Number_Of_Cows_Losing_Calf'
	WHEN 'Number_Of_Cows_Weaning_Calf' THEN 'Number_Of_Cows_Weaning_Calf'
END as Measures,
CASE b. Measures
	WHEN 'Total_Cows_Exposed' THEN Total_Cows_Exposed
	WHEN 'Total_Cows_Kept_For_Calving' THEN Total_Cows_Kept_For_Calving
	WHEN 'Number_Of_Cows_Aborted' THEN Number_Of_Cows_Aborted
	WHEN 'Number_Of_Cows_Open' THEN Number_Of_Cows_Open
	WHEN 'Number_Of_Cows_Calving' THEN Number_Of_Cows_Calving
	WHEN 'Number_Of_Cows_Losing_Calf' THEN Number_Of_Cows_Losing_Calf
	WHEN 'Number_Of_Cows_Weaning_Calf' THEN Number_Of_Cows_Weaning_Calf
END as Count
FROM(
SELECT
(SELECT (select SUM(CASE WHEN weaning_tbl.manage_code='T' Then 0.5 ELSE 1 END)
from cattle_info_tbl INNER JOIN weaning_tbl ON cattle_info_tbl.chaps_id=weaning_tbl.chaps_id  where cattle_info_tbl.birth_date >= Birth_Start_Date AND  cattle_info_tbl.birth_date <=Birth_End_Date)
+
(select count(*) from owners_tbl where exit_date between est_turn_date AND DATE_ADD(est_turn_date, INTERVAL 365 DAY ))) as Total_Cows_Exposed,
@kept_for_calving:=(select SUM(CASE WHEN weaning_tbl.manage_code='T' Then 0.5 ELSE 1 END)
from cattle_info_tbl INNER JOIN weaning_tbl ON cattle_info_tbl.chaps_id=weaning_tbl.chaps_id  where cattle_info_tbl.birth_date >= Birth_Start_Date AND  cattle_info_tbl.birth_date <=Birth_End_Date) as Total_Cows_Kept_For_Calving,
@cows_aborted:=(select count(*) from weaning_tbl where entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date and manage_code ='B') as Number_Of_Cows_Aborted,
@cows_open:=(select count(*) from weaning_tbl where entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date and manage_code ='A') as Number_Of_Cows_Open,
@cows_calving:=@kept_for_calving-@cows_aborted-@cows_open as Number_Of_Cows_Calving,
@cows_losing_calves:=(select count(*) from weaning_tbl inner join cattle_info_tbl on weaning_tbl.chaps_id=cattle_info_tbl.chaps_id where cattle_info_tbl.birth_date >= Birth_Start_Date and  cattle_info_tbl.birth_date<=Birth_End_Date  and weaning_tbl.manage_code in ('C','D','F','K')) as Number_Of_Cows_Losing_Calf,
@cows_weaning_calves:=@cows_calving-@cows_losing_calves as Number_Of_Cows_Weaning_Calf
)a
cross join
(
	select 'Total_Cows_Exposed' as Measures
	union all select 'Total_Cows_Kept_For_Calving'
	union all select 'Number_Of_Cows_Aborted'
	union all select 'Number_Of_Cows_Open'
	union all select 'Number_Of_Cows_Calving'	
	union all select 'Number_Of_Cows_Losing_Calf'
	union all select 'Number_Of_Cows_Weaning_Calf'

)b;

END //

CALL Summary_Of_Cows_Present_In_Herd('2014-1-1','2014-12-31','2013-08-01')
