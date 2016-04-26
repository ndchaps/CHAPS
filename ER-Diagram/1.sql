SELECT AVG(measurement_tbl.weight) FROM measurement_tbl
INNER JOIN cattle_info_tbl ON cattle_info_tbl.chaps_id = measurement_tbl.chaps_id
INNER JOIN weaning_tbl ON cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE cattle_info_tbl.sex in ('0','1','2','3')
AND weaning_tbl.manage_code <>'A'
AND weaning_tbl.manage_code <> 'B'
AND cattle_info_tbl.cow_age=10
AND cattle_info_tbl.birth_date >= '2014-1-1'
AND cattle_info_tbl.birth_date <= '2014-12-31';