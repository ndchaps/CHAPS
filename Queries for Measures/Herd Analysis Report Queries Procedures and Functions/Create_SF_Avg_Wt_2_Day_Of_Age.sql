DELIMITER \\
DROP FUNCTION Avg_Wt_Per_Day_Of_Age_To_Weaning \\
CREATE FUNCTION  Avg_Wt_Per_Day_Of_Age_To_Weaning(Birth_Start_Date date,Birth_End_Date date) RETURNS DOUBLE
BEGIN
DECLARE Avg_Wt_Per_day double;
SELECT sum(wt2daygain)/sum(wt2daygaindenom) INTO  Avg_Wt_Per_day FROM(
SELECT 
		@age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS age_in_days,
		@avg_age:= Average_Calf_Age(Birth_Start_Date,Birth_End_Date),
		@irr_calf:=
			CASE 
				WHEN @age_in_days>@avg_age+45 or @age_in_days<@avg_age-45 THEN 'T'
				ELSE 'F'
			END AS irr_calf,

		@wt_2_day_gain:=
			CASE 
				WHEN  @age_in_days>0 AND weight>0 and @irr_calf='F' THEN weight/@age_in_days 
				ELSE 0
			END AS wt2daygain,

		@wt_2_day_gain_denom:=
			CASE 
				WHEN  @age_in_days>0 AND weight>0 and @irr_calf='F' THEN 1
				ELSE 0
			END AS wt2daygaindenom
		FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
		INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
		INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
		where cattle_info_tbl.birth_date >= Birth_Start_Date 
		AND cattle_info_tbl.birth_date <= Birth_End_Date 
		AND measurement_tbl.entry_date<>'0000-00-00'
)a;
return Avg_Wt_Per_day;
END\\
select Avg_Wt_Per_Day_Of_Age_To_Weaning('2014-1-1','2014-12-31')