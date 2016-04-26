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
	cattle_info_tbl.sire_ID,
	measurement_tbl.frame_score as cframe,
	@age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS age_in_days,
	@AD205:=CASE
		WHEN manage_code not in ('A','B','C','D','F','N','K','S','T','X') THEN  @age_in_days
		ELSE 0
	END AS AD205,
	@calf_count:=CASE
		WHEN manage_code not in ('A','B','C','D','F','N','K','S','T','X') THEN  1
		ELSE 0
	END AS calf_count,
	@avg_age:=(select  SUM(AD205)/SUM(calf_count) as avg_age from
				(SELECT DISTINCTROW
					cattle_info_tbl.chaps_id,
					cattle_info_tbl.birth_date,
					measurement_tbl.entry_date,
					weaning_tbl.manage_code,
					@age:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS age,
					@AD205:=CASE
							WHEN manage_code not in ('A','B','C','D','F','N','K','S','T','X') THEN  @age
							ELSE 0
					END AS AD205,
					@calf_count:=CASE
							WHEN manage_code not in ('A','B','C','D','F','N','K','S','T','X') THEN  1
							ELSE 0
					END AS calf_count
				FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
				INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
				INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
				where cattle_info_tbl.birth_date > '2014-01-01' 
				AND cattle_info_tbl.birth_date < '2014-12-31' 
				AND measurement_tbl.entry_date<>'0000-00-00'
				)a) as avg_age,
	CASE 
		WHEN @age_in_days>@avg_age+45 or @age_in_days<@avg_age-45 THEN 'T'
		ELSE 'F'
	END AS irr_calf

	 
FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
 INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
where cattle_info_tbl.birth_date > '2014-01-01' 
AND cattle_info_tbl.birth_date < '2014-12-31' 
AND measurement_tbl.entry_date<>'0000-00-00'

