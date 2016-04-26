select sum(wt_2_day_gain)/sum(wt_2_day_gain_denom) as Weight_Per_Day_TO_Weaning from
(
SELECT DISTINCTROW
	cattle_info_tbl.sex,
	cattle_info_tbl.birth_date,
	measurement_tbl.entry_date,
	cattle_info_tbl.birth_weight,
	measurement_tbl.weight,
	weaning_tbl.manage_code,
	@age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS age_in_days,
	@wt_2_day_gain:=
	CASE 
		WHEN  @age_in_days>0 AND weight>0 THEN weight/@age_in_days 
		ELSE 0
	END AS wt_2_day_gain,
	@wt_2_day_gain_denom:=
	CASE 
		WHEN  @age_in_days>0 AND weight>0 THEN 1
		ELSE 0
	END AS wt_2_day_gain_denom

FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
where cattle_info_tbl.birth_date > '2014-01-01' 
AND cattle_info_tbl.birth_date < '2014-12-31' 
AND measurement_tbl.entry_date<>'0000-00-00'
)a
