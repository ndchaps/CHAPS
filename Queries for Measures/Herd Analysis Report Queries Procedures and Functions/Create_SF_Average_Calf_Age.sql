Delimiter //
drop function Average_Calf_Age//
CREATE FUNCTION Average_Calf_Age(Birth_Start_Date date,Birth_End_Date date) RETURNS DOUBLE
BEGIN
	DECLARE Avg_Age double;
	SELECT  ROUND(SUM(age)/SUM(calf_count),2) INTO Avg_Age from
					(SELECT DISTINCTROW
						cattle_info_tbl.chaps_id,
						cattle_info_tbl.birth_date,
						measurement_tbl.entry_date,
						weaning_tbl.manage_code,
						@age:= CASE
								WHEN manage_code not in ('A','B','C','D','F','N','K','S','T','X') THEN DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) 
								ELSE 0
						END as age,						
						@calf_count:=CASE
								WHEN manage_code not in ('A','B','C','D','F','N','K','S','T','X') THEN  1
								ELSE 0
						END AS calf_count
					FROM cattle_info_tbl INNER JOIN measurement_tbl ON cattle_info_tbl.chaps_id = measurement_tbl.chaps_id
					INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
					INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
					where cattle_info_tbl.birth_date > Birth_Start_Date 
					AND cattle_info_tbl.birth_date < Birth_End_Date 
					AND measurement_tbl.entry_date<>'0000-00-00'
					)a;
return Avg_Age;
END //
select Average_Calf_Age('2013-1-1','2013-12-31')//