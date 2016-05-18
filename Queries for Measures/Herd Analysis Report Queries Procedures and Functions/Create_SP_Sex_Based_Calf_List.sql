 DELIMITER //
drop PROCEDURE IF EXISTS Sex_Based_Calf_List;//
CREATE PROCEDURE Sex_Based_Calf_List(IN calf_sex int,Birth_Start_Date date,Birth_End_Date date) 

BEGIN
	DECLARE sum_adj_205 DOUBLE;
	DECLARE adj_205_calf_count DOUBLE;
	SET sum_adj_205 = Heifer_SUM_ADJ_WEIGHT_205(calf_sex,Birth_Start_Date,Birth_End_Date);	
	SET adj_205_calf_count=Calf_Count(calf_sex,Birth_Start_Date,Birth_End_Date);
 SELECT Calf_ID, Birth_Date,Birth_Weight,Calving_Ease, Act_Wean_Weight, Age_In_Days, Adj_205_Wt,Adj_205_wt_ratio, MGT, Frame_Score,
ADG,WDA, Cow_ID, Cow_Breed, Cow_Age,Sire_ID, Sire_Breed FROM
 (SELECT DISTINCTROW
		cattle_info_tbl.chaps_id as Calf_ID,
		cattle_info_tbl.birth_date as Birth_Date ,
		cattle_info_tbl.birth_weight as Birth_weight,
		cattle_info_tbl.calving_ease as Calving_Ease,
		measurement_tbl.weight as Act_Wean_Weight,
		@age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS Age_in_Days,
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
		@adj205wt:=(((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam as Adj_205_Wt,

		((@adj205wt/sum_adj_205)*adj_205_calf_count)*100 as Adj_205_wt_ratio,

		weaning_tbl.manage_code AS MGT,

		measurement_tbl.frame_score as Frame_Score,

		CASE 
			WHEN @age_in_days > 0 and weight >0 THEN ROUND((weight-birth_weight) /@age_in_days,1)
		ELSE 0
		END AS ADG,

		CASE 
			WHEN @age_in_days > 0 and weight >0 THEN ROUND(weight/@age_in_days,1)
		ELSE 0
		END AS WDA,
	
		@cow_id:=cattle_info_tbl.dam_ID as Cow_ID,
		@cow_breed:=(select breed from cattle_info_tbl where chaps_id=@cow_id) as Cow_Breed,
		cattle_info_tbl.cow_age as Cow_Age,
		@sire_id:=cattle_info_tbl.sire_ID as Sire_ID,
		@sire_breed:=(select breed from cattle_info_tbl where chaps_id=@sire_id) as Sire_Breed

	FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
	INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
	INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
	where cattle_info_tbl.birth_date > Birth_Start_Date 
	AND cattle_info_tbl.birth_date < Birth_End_Date 
	#AND measurement_tbl.entry_date<>'0000-00-00'
	AND cattle_info_tbl.sex=calf_sex
)a;
END //

CALL Sex_Based_Calf_List(2,'2014-01-01','2014-12-31')//