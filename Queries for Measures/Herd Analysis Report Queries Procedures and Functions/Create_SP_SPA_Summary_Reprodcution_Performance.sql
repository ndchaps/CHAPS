DELIMITER //
DROP PROCEDURE IF EXISTS SPA_Summary_Reprodcution_Performance //
CREATE PROCEDURE SPA_Summary_Reprodcution_Performance(Birth_Start_Date date,Birth_End_Date date,Bull_Turn_Out_Date date)
BEGIN 
SELECT 
CASE  b.Critical_Succes_Factors 	
	WHEN "Pregnancy_Percentage" THEN 'Pregnancy Percentage'	
	WHEN "Pregnancy_Loss_Percentage" THEN 'Pregnancy Loss Percentage'
	WHEN "Calving_Percentage" THEN 'Calving Percentage'
	WHEN "Calf_Death_Loss" THEN 'Calf Death Loss'
	WHEN "Weaning_Percentage" THEN 'Calf Crop or Weaning Percentage'
	WHEN "Female_Replacement_Rate_Percentage" THEN 'Female Replacement Rate Percentage'
	WHEN "Calf_Death_Loss_Based_On_Calves_Born" THEN 'Calf Death Loss Based on # of Calves Born'
	WHEN "Pounds_Weaned_Per_Exposed_Female" THEN 'Pounds_Weaned_Per_Exposed_Female'
END as Critical_Success_Factors,
CASE  b.Critical_Succes_Factors 
	WHEN "Pregnancy_Percentage" THEN CONCAT(FORMAT(Pregnancy_Percentage* 100 ,2),' %')	
	WHEN "Pregnancy_Loss_Percentage" THEN CONCAT(FORMAT(Pregnancy_Loss_Percentage* 100 ,2),' %')
	WHEN "Calving_Percentage" THEN CONCAT(FORMAT(Calving_Percentage* 100 ,2),' %')
	WHEN "Calf_Death_Loss" THEN CONCAT(FORMAT(Calf_Death_Loss* 100 ,2),' %')
	WHEN "Weaning_Percentage" THEN CONCAT(FORMAT(Weaning_Percentage* 100 ,2),' %')
	WHEN "Female_Replacement_Rate_Percentage" THEN CONCAT(FORMAT(Female_Replacement_Rate_Percentage*100,2),'%')
	WHEN "Calf_Death_Loss_Based_On_Calves_Born" THEN CONCAT(FORMAT(Calf_Death_Loss_Based_On_Calves_Born*100,2),'%')
	WHEN "Pounds_Weaned_Per_Exposed_Female" THEN CONCAT(Pounds_Weaned_Per_Exposed_Female, ' lbs')
END as Your_Herd_Performance
FROM(
	SELECT 
	(@total_cows_kept_for_calving-@cows_open)/@denom as Pregnancy_Percentage,
	(@cows_aborted/(@total_cows_kept_for_calving+@cows_aborted)) as Pregnancy_Loss_Percentage,	
	(@total_cows_kept_for_calving-@cows_aborted-@cows_open)/@denom as Calving_Percentage,
	@cows_losing_calves/@denom as Calf_Death_Loss,
	(@cows_weaning_calves+@x-@f)/@denom as Weaning_Percentage,
	(@cows_losing_calves/@total_cows_calving) as Calf_Death_Loss_Based_On_Calves_Born,
	@total_wt/@denom as Pounds_Weaned_Per_Exposed_Female,
	@rep_calv as Female_Replacement_Rate_Percentage
	FROM(
		SELECT @total_cows_kept_for_calving:= (select SUM(CASE WHEN weaning_tbl.manage_code='T' Then 0.5 ELSE 1 END)
from cattle_info_tbl INNER JOIN weaning_tbl ON cattle_info_tbl.chaps_id=weaning_tbl.chaps_id  where cattle_info_tbl.birth_date >= Birth_Start_Date AND  cattle_info_tbl.birth_date <=Birth_End_Date),
		@total_cows_exposed:=@total_cows_kept_for_calving+(select count(*) from owners_tbl where exit_date between Bull_Turn_Out_Date AND DATE_ADD(Bull_Turn_Out_Date, INTERVAL 365 DAY ) ),
		@cows_aborted:=(select count(*) as 'Cows aborted' from weaning_tbl where entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date and manage_code ='B'),
		@cows_open:=(select count(*) as 'Cows open' from weaning_tbl where entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date and manage_code ='A'),
		@total_cows_calving:=@total_cows_kept_for_calving-@cows_aborted,
		@cows_losing_calves:=(select count(*) as 'Cows losing calves' from weaning_tbl inner join cattle_info_tbl on weaning_tbl.chaps_id=cattle_info_tbl.chaps_id where cattle_info_tbl.birth_date >= Birth_Start_Date and  cattle_info_tbl.birth_date<=Birth_End_Date  and weaning_tbl.manage_code in ('C','D','F','K')),
		@h:=(select count(*) as 'Cows_Sold_Because_Of_Age' from owners_tbl where exit_date between Bull_Turn_Out_Date AND DATE_ADD(Bull_Turn_Out_Date, INTERVAL 365 DAY ) and exit_code ='H'),
		@j:=(select count(*) as 'Cows_Sold_Because_Of_Physical_Defects' from owners_tbl where  exit_date between Bull_Turn_Out_Date AND DATE_ADD(Bull_Turn_Out_Date, INTERVAL 366 DAY ) and exit_code ='J'),
		@l:=(select count(*) as 'Cows_Sold_Because_Of_Inferior_Calves' from owners_tbl where exit_date between Bull_Turn_Out_Date AND DATE_ADD(Bull_Turn_Out_Date, INTERVAL 365 DAY ) and exit_code ='L'),
		@r:=(select count(*) as 'Cows_Sold_For_Replacement_Stock' from owners_tbl where exit_date between Bull_Turn_Out_Date AND DATE_ADD(Bull_Turn_Out_Date, INTERVAL 365 DAY ) and exit_code ='R'),
		@y:=(select count(*) as 'Cows_Sold_For_Unknown_Reason' from owners_tbl where exit_date between Bull_Turn_Out_Date AND DATE_ADD(Bull_Turn_Out_Date, INTERVAL 365 DAY ) and exit_code ='Y'),
		@cows_weaning_calves:=(select count(*) from measurement_tbl where weight>0 AND entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date),
		@f:=(select count(*) as 'foster or purchased calves'  from weaning_tbl where entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date and manage_code ='F'),
		@x:= (select count(*) as 'incomplete record in weaning tbl'  from weaning_tbl where entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date and manage_code ='X'),
		@denom:=@total_cows_exposed-@h-@j-@l-@r-@y,
		@total_wt:= (select sum(weight) from measurement_tbl where weight>0 AND entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date),
		#@rep_calv:=(select count(*) from cattle_info_tbl where cow_age<3 and birth_date >= Birth_Start_Date AND birth_date <=Birth_End_Date)/(select count(*) from cattle_info_tbl WHERE birth_date >= Birth_Start_Date AND birth_date <=Birth_End_Date)
		@rep_calv:=(select count(*) from cattle_info_tbl where cow_age<3 and birth_date >= Birth_Start_Date AND birth_date <=Birth_End_Date)/@denom
		
	)x
)a
cross join
(
	select "Pregnancy_Percentage" as Critical_Succes_Factors
	union all select "Pregnancy_Loss_Percentage"
	union all select "Calving_Percentage"
	union all select "Calf_Death_Loss"
	union all select "Weaning_Percentage"	
	union all select "Female_Replacement_Rate_Percentage"
	union all select  "Calf_Death_Loss_Based_On_Calves_Born"
	union all select "Pounds_Weaned_Per_Exposed_Female"
)b;

END //


CALL SPA_Summary_Reprodcution_Performance('2014-1-1','2014-12-31','2013-8-1')