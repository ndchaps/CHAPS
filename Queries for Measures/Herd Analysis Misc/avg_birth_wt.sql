
SELECT AVG(birth_weight) 
FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
where cattle_info_tbl.birth_date > '2014-01-01' 
AND cattle_info_tbl.birth_date < '2014-12-31' 
AND measurement_tbl.entry_date<>'0000-00-00'
