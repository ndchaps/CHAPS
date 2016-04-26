SELECT COUNT(*) AS '1st 21 Calves' FROM cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE cattle_info_tbl.sex in ('0','1','2','3')
AND cattle_info_tbl.birth_date >= '2014-1-1'
AND  cattle_info_tbl.birth_date <= '2014-12-31'
AND weaning_tbl.manage_code NOT IN ('A','B')
AND cattle_info_tbl.cow_age=4
AND cattle_info_tbl.birth_date <= DATE_ADD('2013-07-19', INTERVAL 285+20 DAY );
