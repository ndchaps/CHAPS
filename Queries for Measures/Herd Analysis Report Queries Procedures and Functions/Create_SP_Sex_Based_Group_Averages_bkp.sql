DELIMITER //
DROP PROCEDURE IF EXISTS Sex_Based_Group_Averages //
CREATE PROCEDURE Sex_Based_Group_Averages(calf_sex int,Birth_Start_Date date,Birth_End_Date date) 
BEGIN
	DECLARE calf_count,avg_age DOUBLE;
	SET calf_count=Calf_Count(calf_sex,Birth_Start_Date,Birth_End_Date);
	SET avg_age=Average_Calf_Age(Birth_Start_Date,Birth_End_Date);
	
SELECT
SUM(CASE WHEN Adj_205_Wt>0 AND irr_calf='F' THEN Adj_205_Wt else 0 end )/SUM(CASE WHEN Adj_205_Wt>0 AND irr_calf='F' THEN 1 ELSE 0 END) AS Avg_Adj_205_Wt,
SUM(CASE WHEN Birth_Weight>0 AND mgt_code not in ('A','B','C','D','F','K','N','P','S','T','X') THEN Birth_Weight ELSE 0 END)/SUM(CASE WHEN mgt_code not in('A','B','C','D','F','K','N','P','S','T','X' ) THEN 1 ELSE 0 END) AS Avg_Birth_Wt,
SUM(CASE WHEN  irr_calf='F' THEN Calving_Ease ELSE 0 END)/SUM(CASE WHEN irr_calf='F' THEN 1 ELSE 0 END) as Avg_Calving_Ease,
SUM(CASE WHEN Act_Wean_Weight>0 AND mgt_code not in ('A','B','C','D','F','K','N','P','S','T','X') THEN Act_Wean_Weight ELSE 0 END)/SUM(CASE WHEN Act_Wean_Weight>0 AND mgt_code not in ('A','B','C','D','F','K','N','P','S','T','X') THEN 1 ELSE 0 END) AS Avg_Act_Wean_Wt,
SUM(CASE WHEN Age_in_Days>0 AND mgt_code NOT IN ('A','B','C','D','F','K','N','P','S','T','X') THEN Age_in_Days ELSE 0 END )/SUM(CASE WHEN Age_in_Days>0 AND mgt_code NOT IN ('A','B','C','D','F','K','N','P','S','T','X') THEN 1 ELSE 0 END ) As Avg_Age_In_Days,
SUM(CASE WHEN Frame_Score>0 THEN Frame_Score ELSE 0 END)/SUM(CASE WHEN Frame_Score>0 THEN 1 ELSE 0 END) AS Avg_Frame_Score,
SUM(CASE WHEN ADG>0 THEN ADG ELSE 0 end)/SUM(CASE WHEN ADG>0 THEN 1 ELSE 0 end) as Avg_ADG,
SUM(CASE WHEN WDA>0 then WDA ELSE 0 end)/SUM(CASE WHEN WDA>0 then 1 ELSE 0 end) as Avg_WDA
FROM (
 SELECT DISTINCTROW
		cattle_info_tbl.chaps_id as Calf_ID,
		cattle_info_tbl.birth_date as Birth_Date ,
		cattle_info_tbl.birth_weight as Birth_weight,
		cattle_info_tbl.calving_ease as Calving_Ease,
		measurement_tbl.weight as Act_Wean_Weight,
		weaning_tbl.manage_code as mgt_code,
		@age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS Age_in_Days,
		@irr_calf:=
			CASE 
				WHEN @age_in_days>avg_age+45 or @age_in_days<avg_age-45 THEN 'T'
				ELSE 'F'
		END AS irr_calf,
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
			WHEN  calf_sex = 2 and cow_age = 2 THEN 54
			WHEN  calf_sex = 2 and cow_age = 3 THEN 36
			WHEN  calf_sex = 2 and cow_age = 4 THEN 18
			WHEN  calf_sex = 2 and cow_age >= 11 THEN 18
			WHEN  calf_sex = 1 and cow_age = 2 THEN 60
			WHEN  calf_sex = 1 and cow_age = 3 THEN 40
			WHEN  calf_sex = 1 and cow_age = 4 THEN 20
			WHEN  calf_sex = 1 and cow_age >=11 THEN 20
			WHEN  calf_sex = 3 and cow_age = 2 THEN 60
			WHEN  calf_sex = 3 and cow_age = 3 THEN 40
			WHEN  calf_sex = 3 and cow_age = 4 THEN 20
			WHEN  calf_sex = 3 and cow_age >=11 THEN 20
			ELSE 0
		END as dam,
		@adj205wt:=CASE WHEN @age_in_days>0 AND weight>0 THEN(((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam 
						ELSE 0 END as Adj_205_Wt,
		measurement_tbl.frame_score as Frame_Score,
		CASE 
			WHEN @age_in_days > 0 and weight >0 THEN ROUND(@w2_day_gain:=(weight-birth_weight) /@age_in_days,1)
		ELSE 0
		END AS ADG,
		CASE 
			WHEN @age_in_days > 0 and weight >0 THEN ROUND(@avg_daily_gain:=weight/@age_in_days,1)
		ELSE 0
		END AS WDA
		
	FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
	INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
	INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
	where cattle_info_tbl.birth_date >= Birth_Start_Date 
	AND cattle_info_tbl.birth_date <= Birth_End_Date 
	#AND measurement_tbl.entry_date<>'0000-00-00'
	AND cattle_info_tbl.sex=calf_sex
)a;
END // 

CALL Sex_Based_Group_Averages(2,'2013-01-01','2013-12-31')
