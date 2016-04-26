SELECT cow_age, COUNT(*) AS 'Cows open-aborted' FROM cattle_info_tbl 
INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE cattle_info_tbl.sex in ('0','1','2','3')
AND weaning_tbl.manage_code IN ('A','B')
AND cattle_info_tbl.birth_date >= '2014-1-1'
AND cattle_info_tbl.birth_date <= '2014-12-31'
Group by cow_age