DELIMITER \\
DROP FUNCTION Adjusted_205_Day_Wt \\
Create FUNCTION Adjusted_205_Day_Wt(Birth_Start_Date date,Birth_End_Date date) RETURNS DOUBLE
BEGIN 
	DECLARE ADJ_205_WT double;
	SELECT sum(adj_wt_205)/sum(adj_wt_205_count) INTO  ADJ_205_WT FROM(
		SELECT  @age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS age_in_days,
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
				@adj_wt_205:=
				CASE 
					WHEN @age_in_days>0 and weight>0 and sex=2 THEN ((((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam ) * 1.05
					WHEN @age_in_days>0 and weight>0 and sex=1 THEN ((((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam ) * 0.95
					WHEN @age_in_days>0 and weight>0 and sex=3 THEN ((((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam ) * 1.00
					WHEN @age_in_days>0 and weight>0 and sex=0 THEN ((((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam ) * 1.00
					ELSE 0
				 END as adj_wt_205,
				@ad_wt_205_count:=
				CASE 
					WHEN @adj_wt_205 <> 0 THEN 1
					ELSE 0
				END AS adj_wt_205_count 
		FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
		INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
		INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
		where cattle_info_tbl.birth_date >=Birth_Start_Date 
		AND cattle_info_tbl.birth_date <=Birth_End_Date 
		#AND weaning_tbl.manage_code NOT IN ('A','B','C','D','F','K','N','P','S','T','X')
		AND measurement_tbl.entry_date<>'0000-00-00'
) A;
	RETURN ADJ_205_WT;
	
END \\

select Adjusted_205_Day_Wt('2012-1-1','2012-12-31')
