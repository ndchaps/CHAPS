SELECT DISTINCTROW
	cattle_info_tbl.sex,
	owners_tbl.herd_id,
	cattle_info_tbl.chaps_id,
	#FORMAT(cattle_info_tbl.birth_date,'DD-MM-YYYY') as birth_date,
	cattle_info_tbl.birth_date,
	measurement_tbl.entry_date,
	cattle_info_tbl.birth_weight,
	cattle_info_tbl.calving_ease,
	measurement_tbl.weight,
	weaning_tbl.manage_code,
	cattle_info_tbl.dam_ID,
	cattle_info_tbl.cow_age,
	cattle_info_tbl.sire_ID,
	measurement_tbl.frame_score as cframe,
	@age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS age_in_days

FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
where cattle_info_tbl.birth_date > '2014-01-01' 
AND cattle_info_tbl.birth_date < '2014-12-31' 
AND measurement_tbl.entry_date<>'0000-00-00'