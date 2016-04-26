 DELIMITER //
drop function Heifer_SUM_ADJ_WEIGHT_205//
CREATE FUNCTION Heifer_SUM_ADJ_WEIGHT_205(calf_sex int,Birth_Start_Date date,Birth_End_Date date) RETURNS DOUBLE
BEGIN
	
		DECLARE Sum_Adj_205_wt DOUBLE;
		SELECT sum(Case when Adj_205_wt>0 then Adj_205_wt else 0 end ) into Sum_Adj_205_wt from(
		SELECT 	distinct
			@age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date),
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
			END ,
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
			END ,
			@adj205wt:=ROUND((((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam,1) as Adj_205_wt

		FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
		INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
		INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
		WHERE cattle_info_tbl.birth_date > Birth_Start_Date
		AND cattle_info_tbl.birth_date < Birth_End_Date 
		AND measurement_tbl.entry_date<>'0000-00-00'
		AND cattle_info_tbl.sex=calf_sex
	)a;

RETURN Sum_Adj_205_wt;
END //
select Heifer_SUM_ADJ_WEIGHT_205(3,'2013-01-01','2013-12-31')//