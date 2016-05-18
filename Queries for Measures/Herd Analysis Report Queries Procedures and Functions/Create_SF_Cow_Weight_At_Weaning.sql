DELIMITER \\
DROP FUNCTION Cow_Weight_At_Weaning \\
Create FUNCTION Cow_Weight_At_Weaning() RETURNS DOUBLE
BEGIN 
	DECLARE Avg_Wean_Weight double;
	SELECT @cow_id:=cattle_info_tbl.dam_ID as Cow_ID,
	@cow_weight:= (select weight from measurement_tbl where chaps_id=@cow_id)
	FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
	INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
	INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
	where cattle_info_tbl.birth_date > '2014-01-01' 
	AND cattle_info_tbl.birth_date < '2014-12-31' 
	AND measurement_tbl.entry_date<>'0000-00-00';
END \\