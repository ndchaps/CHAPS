SELECT cattle_info_tbl.cow_age as Dam_Age,
#attle_info_tbl.birth_date as birth_date,
CAST(from_unixtime(AVG(UNIX_TIMESTAMP(cattle_info_tbl.birth_date))) AS DATE)
#AVG(CONVERT(cattle_info_tbl.birth_date,SIGNED)) as birth_date_number
FROM cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
INNER JOIN measurement_tbl On cattle_info_tbl.chaps_id=measurement_tbl.chaps_id
WHERE cattle_info_tbl.sex in ('0','1','2','3')
AND cattle_info_tbl.birth_date >= '2014-1-1'
AND  cattle_info_tbl.birth_date <= '2014-12-31'
AND weaning_tbl.manage_code <>'A'
AND weaning_tbl.manage_code <>'B'
GROUP BY cattle_info_tbl.cow_age
order by Dam_Age