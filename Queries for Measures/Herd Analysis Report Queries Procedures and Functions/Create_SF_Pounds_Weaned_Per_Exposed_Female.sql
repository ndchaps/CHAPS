Delimiter //
DROP FUNCTION Pounds_Weaned_Per_Exposed_Female //
CREATE FUNCTION Pounds_Weaned_Per_Exposed_Female() RETURNS DOUBLE
BEGIN
DECLARE pounds_weaned  double;
SELECT 
 SUM(CASE WHEN measurement_tbl.weight>0 AND weaning_tbl.manage_code NOT IN('A','B') THEN  measurement_tbl.weight ELSE 0 END)/ Count(DISTINCT cattle_info_tbl.dam_id) into pounds_weaned 	
	FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
	INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
	INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
	where cattle_info_tbl.birth_date > '2014-01-01' 
	AND cattle_info_tbl.birth_date < '2014-12-31' 
	AND measurement_tbl.entry_date<>'0000-00-00';	
return pounds_weaned;
END //

select Pounds_Weaned_Per_Exposed_Female()