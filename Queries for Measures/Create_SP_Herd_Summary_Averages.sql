DELIMITER //
DROP PROCEDURE IF EXISTS Herd_Summary_Averages //
CREATE PROCEDURE Herd_Summary_Averages(Birth_Start_Date date,Birth_End_Date date)
BEGIN
	SELECT avg_age as Average_Age_At_Weaning,
			wt_2_day_gain as Weight_Per_Day_To_Weaning,
			sum(CASE WHEN birth_weight>0  THEN birth_weight ELSE 0 END)/SUM(CASE WHEN birth_weight>0 THEN 1 ELSE 0 END)as Average_Birth_Weight,
			#avg(birth_weight) as 'Average Birth Weight',
			avg_adj_205_day_wt as Average_Sex_Adjusted_Weight_205
	from (
		SELECT DISTINCTROW
		cattle_info_tbl.chaps_id,
		cattle_info_tbl.birth_weight,
		@avg_age:=Average_Calf_Age(Birth_Start_Date,Birth_End_Date ) as avg_age,
		Avg_Wt_Per_Day_Of_Age_To_Weaning(Birth_Start_Date,Birth_End_Date) as wt_2_day_gain,
		Adjusted_205_Day_Wt(Birth_Start_Date,Birth_End_Date ) as avg_adj_205_day_wt

	FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
	INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
	INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
	where cattle_info_tbl.birth_date >= Birth_Start_Date 
	AND cattle_info_tbl.birth_date <= Birth_End_Date 
	)a
;
END //
call Herd_Summary_Averages('2014-01-01','2014-12-31')