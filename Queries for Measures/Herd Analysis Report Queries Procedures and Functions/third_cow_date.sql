select birth_date as ThirdCowDate from (
select distinct cattle_info_tbl.dam_id,cattle_info_tbl.cow_age, owners_tbl.start_date AS enter_herd_date,cattle_info_tbl.birth_date as birth_date
from cattle_info_tbl INNER JOIN owners_tbl on cattle_info_tbl.chaps_id=owners_tbl.chaps_id
INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
where cattle_info_tbl.birth_date > '2013-01-01' 
AND cattle_info_tbl.birth_date < '2013-12-31'
and weaning_tbl.manage_code NOT IN( 'A','B','P' )
And cattle_info_tbl.cow_age > 2
order by cattle_info_tbl.birth_date
) A
order by birth_date LIMIT 2,1




