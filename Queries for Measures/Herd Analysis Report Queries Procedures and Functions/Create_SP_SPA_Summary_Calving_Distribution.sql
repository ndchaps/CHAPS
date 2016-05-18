DELIMITER \\
DROP PROCEDURE IF EXISTS SPA_Summary_Calving_Distribution \\
CREATE PROCEDURE SPA_Summary_Calving_Distribution(Birth_Start_Date date,Birth_End_Date date,Bull_Turnout_Date date) 
BEGIN
DECLARE est_turn_date DATE;
set est_turn_date=Estimated_Bull_Turnout_Date(Birth_Start_Date, Birth_End_Date, Bull_Turnout_Date);
SELECT DISTINCT
CASE  b.Critical_Succes_Factors 	
	WHEN 'Calves_Born_During_First_21_Days' THEN 'Calves_Born_During_First_21_Days'	
	WHEN 'Calves_Born_During_First_42_Days'	THEN 'Calves_Born_During_First_42_Days'
	WHEN 'Calves_Born_During_First_63_Days' THEN 'Calves_Born_During_First_63_Days'
	WHEN 'Calves_Born_After_First_63_Days'	THEN 'Calves_Born_After_First_63_Days'
	WHEN 'Avg_Age_at_Weaning'	THEN 'Avg_Age_at_Weaning'
	WHEN 'Actual_Weaning_Wts_Steers' THEN 'Actual_Weaning_Wts_Steers'
	WHEN 'Actual_Weaning_Wts_Heifers'  THEN 'Actual_Weaning_Wts_Heifers'
	WHEN 'Actual_Weaning_Wts_Bulls' THEN 'Actual_Weaning_Wts_Bulls'
	WHEN 'Avg_Weaning_Wts' THEN 'Avg_Weaning_Wts'
	WHEN 'Pounds_Weaned_Per_Exposed_Female' THEN 'Pounds_Weaned_Per_Exposed_Female'
END as Critical_Success_Factors,
CASE  b.Critical_Succes_Factors 
		WHEN 'Calves_Born_During_First_21_Days' THEN CONCAT(FORMAT(Calves_Born_During_First_21_Days * 100 ,2),' %')
		WHEN 'Calves_Born_During_First_42_Days' THEN CONCAT(FORMAT(Calves_Born_During_First_42_Days * 100 ,2),' %')
		WHEN 'Calves_Born_During_First_63_Days' THEN CONCAT(FORMAT(Calves_Born_During_First_63_Days * 100 ,2),' %')
		WHEN 'Calves_Born_After_First_63_Days' THEN CONCAT(FORMAT(Calves_Born_After_First_63_Days * 100 ,2),' %')
		WHEN 'Avg_Age_at_Weaning' THEN Round(Avg_Age_at_Weaning,1)
		WHEN 'Actual_Weaning_Wts_Heifers' THEN Round(Actual_Weaning_Wts_Heifers,1)
		WHEN 'Actual_Weaning_Wts_Steers' THEN Round(Actual_Weaning_Wts_Steers,1)
		WHEN 'Actual_Weaning_Wts_Bulls' THEN Round(Actual_Weaning_Wts_Bulls,1)
		WHEN 'Avg_Weaning_Wts' THEN Round(Avg_Weaning_Wts,1)
		WHEN 'Pounds_Weaned_Per_Exposed_Female' THEN 'Pounds_Weaned_Per_Exposed_Female'
END as Your_Herd_Performance
FROM(
	SELECT  @within_21days/(@within_63days+@after_63days) AS Calves_Born_During_First_21_Days,
			@within_42days/(@within_63days+@after_63days) AS Calves_Born_During_First_42_Days,
			@within_63days/(@within_63days+@after_63days) AS Calves_Born_During_First_63_Days,
			@after_63days/(@within_63days+@after_63days) AS Calves_Born_After_First_63_Days,
			Avg_Age_at_Weaning, Actual_Weaning_Wts_Steers, Actual_Weaning_Wts_Heifers, Actual_Weaning_Wts_Bulls, Avg_Weaning_Wts
	FROM(
		SELECT
		@early_calves:=Early_Calves(Birth_Start_Date, Birth_End_Date,est_turn_date,4,null)AS Early_Calves,
		@within_21days:=Calving_Within_21Days(Birth_Start_Date, Birth_End_Date,est_turn_date,4,null),
		@within_42days:=Calving_Within_42Days(Birth_Start_Date, Birth_End_Date,est_turn_date,4,null) ,
		@within_63days:=Calving_Within_63Days(Birth_Start_Date, Birth_End_Date,est_turn_date,4,null) ,
		@after_63days:=Calving_After_63Days(Birth_Start_Date, Birth_End_Date,est_turn_date,4,null) ,
		Average_Calf_Age(Birth_Start_Date, Birth_End_Date) as Avg_Age_at_Weaning,
		COALESCE(Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,3),0) as Actual_Weaning_Wts_Steers,
		COALESCE(Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,2),0) as Actual_Weaning_Wts_Heifers,
		COALESCE(Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,1),0) as Actual_Weaning_Wts_Bulls,
		COALESCE(Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,4),0) as Avg_Weaning_Wts
		FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
			INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
			INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
			where cattle_info_tbl.birth_date >=Birth_Start_Date
			AND cattle_info_tbl.birth_date <=     Birth_End_Date 
			AND measurement_tbl.entry_date<>'0000-00-00'
	)x
)a
cross join
(
	select 'Calves_Born_During_First_21_Days' as Critical_Succes_Factors
	union all select 'Calves_Born_During_First_42_Days'
	union all select 'Calves_Born_During_First_63_Days'
	union all select 'Calves_Born_After_First_63_Days'
	union all select 'Avg_Age_at_Weaning'	
	union all select 'Actual_Weaning_Wts_Steers'
	union all select 'Actual_Weaning_Wts_Heifers'
	union all select 'Actual_Weaning_Wts_Bulls'
	union all select  'Avg_Weaning_Wts'
	#union all select 'Pounds_Weaned_Per_Exposed_Female'
)b;
END \\
Call SPA_Summary_Calving_Distribution('2014-1-1','2014-12-31','2013-08-01')