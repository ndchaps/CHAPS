Delimiter //
DROP FUNCTION Actual_Wean_Weight //
CREATE FUNCTION Actual_Wean_Weight(Birth_Start_Date date,Birth_End_Date date,calf_sex int) RETURNS DOUBLE
BEGIN
DECLARE act_wean_wt  double;
#SELECT measurement_tbl.frame_score as Frame_Score
SELECT 
#ROUND(SUM(CASE WHEN measurement_tbl.frame_score>0 THEN measurement_tbl.frame_score ELSE 0 END)/SUM(CASE WHEN measurement_tbl.frame_score>0 THEN 1 ELSE 0 END),1)
 SUM(CASE WHEN measurement_tbl.weight>0 THEN  measurement_tbl.weight ELSE 0 END)/ SUM(CASE WHEN measurement_tbl.weight>0 THEN 1 ELSE 0 END)into act_wean_wt 	
	FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
	INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
	INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
	where cattle_info_tbl.birth_date >= Birth_Start_Date 
	AND cattle_info_tbl.birth_date <=Birth_End_Date
	AND measurement_tbl.entry_date<>'0000-00-00'
	AND 
		CASE calf_sex
			WHEN 4 THEN cattle_info_tbl.sex IN (0,1,2,3)
			ELSE cattle_info_tbl.sex=calf_sex
	END;		
return act_wean_wt;
END //

select Actual_Wean_Weight('2013-1-1','2013-12-31',4)