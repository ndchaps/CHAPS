Delimiter //
drop function Calf_Count//
CREATE FUNCTION Calf_Count(calf_sex int,Birth_Start_Date date,Birth_End_Date date) RETURNS DOUBLE
BEGIN
	declare calf_count_205 double;
	select sum(ad_wt_205_count) into calf_count_205 from (
		select cattle_info_tbl.chaps_id,
		@age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS age_in_days,
		@avg_age:=Average_Calf_Age(Birth_Start_Date,Birth_End_Date),
	@irr_calf:=
	CASE 
		WHEN @age_in_days>@avg_age+45 or @age_in_days<@avg_age-45 THEN 'T'
		ELSE 'F'
	END AS irr_calf,
	@ad_wt_205_count:=
	CASE 
		WHEN @irr_calf='F' THEN 1
		ELSE 0
	END AS ad_wt_205_count  	
FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
where cattle_info_tbl.birth_date > Birth_Start_Date 
AND cattle_info_tbl.birth_date < Birth_End_Date 
AND measurement_tbl.entry_date<>'0000-00-00'
AND sex=calf_sex
)a;
return calf_count_205;
END //
select Calf_Count(2,'2014-01-01','2014-12-31')//