	SELECT avg_age as Average_Age_At_Weaning,
			sum(wt_2_day_gain)/sum(wt_2_day_gain_denom) as Weight_Per_Day_TO_Weaning,
			AVG(birth_weight) as Average_Birth_Weight,
			sum(adj_wt_205)/sum(ad_wt_205_count) as Average_Sex_Adjusted_Weight_205
	from (
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
						where cattle_info_tbl.birth_date > Birth_Start_Date 
						AND cattle_info_tbl.birth_date < Birth_End_Date 
						AND measurement_tbl.entry_date<>'0000-00-00'
						)a) as avg_age,
			@irr_calf:=
			CASE 
				WHEN @age_in_days>@avg_age+45 or @age_in_days<@avg_age-45 THEN 'T'
				ELSE 'F'
			END AS irr_calf,
			@wt_2_day_gain:=
			CASE 
				WHEN  @age_in_days>0 AND weight>0 and @irr_calf='F' THEN weight/@age_in_days 
				ELSE 0
			END AS wt_2_day_gain,
			@wt_2_day_gain_denom:=
			CASE 
				WHEN  @age_in_days>0 AND weight>0 and @irr_calf='F' THEN 1
				ELSE 0
			END AS wt_2_day_gain_denom,
			@adj_birth_wt:= 
			CASE
				WHEN cow_age <= 2 and birth_weight > 0 THEN birth_weight + 8
				WHEN cow_age =3 and birth_weight > 0 THEN birth_weight + 5
				WHEN cow_age =4 and birth_weight > 0 THEN birth_weight + 2
				WHEN cow_age >= 5 and cow_age <= 10 AND  birth_weight > 0 THEN birth_weight 
				WHEN cow_age >=11 and birth_weight > 0 THEN birth_weight + 3
				WHEN birth_weight = 0 and sex = 1 THEN 75
				WHEN birth_weight = 0 and sex = 2 THEN 70
				WHEN birth_weight = 0 and sex = 3 THEN 75
			END as adj_birth_wt,
			@dam:=
			CASE
				WHEN  sex = 2 and cow_age = 2 THEN 54
				WHEN  sex = 2 and cow_age = 3 THEN 36
				WHEN  sex = 2 and cow_age = 4 THEN 18
				WHEN  sex = 2 and cow_age >= 11 THEN 18
				WHEN  sex = 1 and cow_age = 2 THEN 60
				WHEN  sex = 1 and cow_age = 3 THEN 40
				WHEN  sex = 1 and cow_age = 4 THEN 20
				WHEN  sex = 1 and cow_age >=11 THEN 20
				WHEN  sex = 3 and cow_age = 2 THEN 60
				WHEN  sex = 3 and cow_age = 3 THEN 40
				WHEN  sex = 3 and cow_age = 4 THEN 20
				WHEN  sex = 3 and cow_age >=11 THEN 20
				ELSE 0
			END as dam,
			 CASE 
				WHEN @age_in_days>0 and weight>0 and sex=2 and @irr_calf='F' THEN ((((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam ) * 1.05
				WHEN @age_in_days>0 and weight>0 and sex=1 and @irr_calf='F' THEN ((((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam ) * 0.95
				WHEN @age_in_days>0 and weight>0 and sex=3 and @irr_calf='F' THEN ((((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam ) * 1.00
				WHEN @age_in_days>0 and weight>0 and sex=0 and @irr_calf='F' THEN ((((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam ) * 1.00
				ELSE 0
			 END as adj_wt_205,
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
	)a;