SELECT AVG(measurement_tbl.entry_date-cattle_info_tbl.birth_date) from measurement_tbl  
INNER JOIN cattle_info_tbl ON cattle_info_tbl.chaps_id=measurement_tbl.chaps_id  
INNER JOIN weaning_tbl ON cattle_info_tbl.chaps_id=weaning_tbl.chaps_id  
INNER JOIN owners_tbl ON cattle_info_tbl.chaps_id=owners_tbl.chaps_id
WHERE measurement_tbl.entry_date IS NOT null 
AND measurement_tbl.entry_date = '2014-13-11' 
AND cattle_info_tbl.birth_date > '2014-01-01' 
AND cattle_info_tbl.birth_date < '2014-31-12' 
AND owners_tbl.herd_id='H38'
Group by cattle_info_tbl.chaps_id
HAVING measurement_tbl.entry_date <> '0000-00-00' 
AND cattle_info_tbl.birth_date > '1900-01-01' 
AND weaning_tbl.manage_code NOT IN ('A','B','C','D')
 ;

