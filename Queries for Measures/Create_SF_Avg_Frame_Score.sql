delimiter //
CREATE FUNCTION Avg_Frame_Score() RETURNS DOUBLE
BEGIN
	DECLARE avg_frame_score  double;
#SELECT measurement_tbl.frame_score as Frame_Score
	SELECT ROUND(SUM(CASE WHEN measurement_tbl.frame_score>0 THEN measurement_tbl.frame_score ELSE 0 END)/SUM(CASE WHEN measurement_tbl.frame_score>0 THEN 1 ELSE 0 END),1) into avg_frame_score 	
	FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
	INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
	INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
	where cattle_info_tbl.birth_date > '2014-01-01' 
	AND cattle_info_tbl.birth_date < '2014-12-31' 
	AND measurement_tbl.entry_date<>'0000-00-00';
RETURN avg_frame_score;

END //
