DELIMITER \\
DROP FUNCTION Average_Daily_Gain \\
CREATE FUNCTION Average_Daily_Gain(Birth_Start_Date date,Birth_End_Date date) RETURNS double
BEGIN 
DECLARE avg_daily_gain  double;
SELECT SUM(CASE WHEN ADG>0 THEN ADG ELSE 0 END)/SUM(CASE WHEN ADG>0 THEN 1 ELSE 0 END) INTO avg_daily_gain FROM(
	SELECT 
		@age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS age_in_days,	
		CASE 
			WHEN @age_in_days > 0 and weight >0 THEN ROUND(@w2_day_gain:=(weight-birth_weight) /@age_in_days,1)
		ELSE 0
		END AS ADG
	FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
	INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
	INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
	where cattle_info_tbl.birth_date > Birth_Start_Date 
	AND cattle_info_tbl.birth_date < Birth_End_Date 
	AND measurement_tbl.entry_date<>'0000-00-00'
)a;
RETURN avg_daily_gain;
END\\

Select Average_Daily_Gain('2014-1-1','2014-12-31')