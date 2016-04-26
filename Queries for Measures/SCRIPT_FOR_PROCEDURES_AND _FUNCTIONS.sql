Delimiter //
/*-----------Heifer_SUM_ADJ_WEIGHT_205-----------------*/
drop function if exists Heifer_SUM_ADJ_WEIGHT_205//
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
/*-------Calf Count--------------------*/

drop function IF EXISTS Calf_Count//
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

/*----------------Average_Calf_Age-----------*/
drop function IF EXISTS  Average_Calf_Age//
CREATE FUNCTION Average_Calf_Age(Birth_Start_Date date,Birth_End_Date date) RETURNS DOUBLE
BEGIN
	DECLARE Avg_Age double;
	SELECT  ROUND(SUM(age)/SUM(calf_count),2) INTO Avg_Age from
					(SELECT DISTINCTROW
						cattle_info_tbl.chaps_id,
						cattle_info_tbl.birth_date,
						measurement_tbl.entry_date,
						weaning_tbl.manage_code,
						@age:= CASE
								WHEN manage_code not in ('A','B','C','D','F','N','K','S','T','X') THEN DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) 
								ELSE 0
						END as age,						
						@calf_count:=CASE
								WHEN manage_code not in ('A','B','C','D','F','N','K','S','T','X') THEN  1
								ELSE 0
						END AS calf_count
					FROM cattle_info_tbl INNER JOIN measurement_tbl ON cattle_info_tbl.chaps_id = measurement_tbl.chaps_id
					INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
					INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
					where cattle_info_tbl.birth_date > Birth_Start_Date 
					AND cattle_info_tbl.birth_date < Birth_End_Date 
					AND measurement_tbl.entry_date<>'0000-00-00'
					)a;
return Avg_Age;
END //

/*--------Avg_Wt_Per_Day_Of_Age_To_Weaning--*/
DROP FUNCTION IF EXISTS Avg_Wt_Per_Day_Of_Age_To_Weaning //
CREATE FUNCTION  Avg_Wt_Per_Day_Of_Age_To_Weaning(Birth_Start_Date date,Birth_End_Date date) RETURNS DOUBLE
BEGIN
DECLARE Avg_Wt_Per_day double;
SELECT sum(wt2daygain)/sum(wt2daygaindenom) INTO  Avg_Wt_Per_day FROM(
SELECT 
		@age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS age_in_days,
		@avg_age:= Average_Calf_Age(Birth_Start_Date,Birth_End_Date),
		@irr_calf:=
			CASE 
				WHEN @age_in_days>@avg_age+45 or @age_in_days<@avg_age-45 THEN 'T'
				ELSE 'F'
			END AS irr_calf,

		@wt_2_day_gain:=
			CASE 
				WHEN  @age_in_days>0 AND weight>0 and @irr_calf='F' THEN weight/@age_in_days 
				ELSE 0
			END AS wt2daygain,

		@wt_2_day_gain_denom:=
			CASE 
				WHEN  @age_in_days>0 AND weight>0 and @irr_calf='F' THEN 1
				ELSE 0
			END AS wt2daygaindenom
		FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
		INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
		INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
		where cattle_info_tbl.birth_date >= Birth_Start_Date 
		AND cattle_info_tbl.birth_date <= Birth_End_Date 
		AND measurement_tbl.entry_date<>'0000-00-00'
)a;
return Avg_Wt_Per_day;
END//

/*--------Adjusted_205_Day_Wt------------*/
DROP FUNCTION IF EXISTS Adjusted_205_Day_Wt //
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
	
END //

/*--------Actual_Wean_Weight------------*/
DROP FUNCTION IF EXISTS Actual_Wean_Weight //
CREATE FUNCTION Actual_Wean_Weight(Birth_Start_Date date,Birth_End_Date date,calf_sex int) RETURNS DOUBLE
BEGIN
DECLARE act_wean_wt  double;
SELECT 
#ROUND(SUM(CASE WHEN measurement_tbl.frame_score>0 THEN measurement_tbl.frame_score ELSE 0 END)/SUM(CASE WHEN measurement_tbl.frame_score>0 THEN 1 ELSE 0 END),1)
 SUM(CASE WHEN measurement_tbl.weight>0 THEN  measurement_tbl.weight ELSE 0 END)/ SUM(CASE WHEN measurement_tbl.weight>0 THEN 1 ELSE 0 END)into act_wean_wt 	
	FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
	INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
	INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
	where cattle_info_tbl.birth_date >= Birth_Start_Date 
	AND cattle_info_tbl.birth_date <=Birth_End_Date
	AND measurement_tbl.entry_date<>'0000-00-00'
	AND 
		CASE calf_sex
			WHEN 4 THEN cattle_info_tbl.sex IN (0,1,2,3)
			ELSE cattle_info_tbl.sex=calf_sex
	END;		
return act_wean_wt;
END //


/*--------Estimated_Bull_Turnout_Date------------*/
DROP FUNCTION IF EXISTS Estimated_Bull_Turnout_Date //
CREATE FUNCTION Estimated_Bull_Turnout_Date(Birth_Start_Date date, Birth_End_Date date, Bull_Turnout_Date date) returns date
BEGIN
DECLARE est_turn_date date;
SELECT CASE 
	WHEN DATEDIFF(@thrird_cow_date,Bull_Turnout_Date)<275 OR DATEDIFF(@thrird_cow_date,Bull_Turnout_Date)>295 THEN DATE_SUB(@thrird_cow_date,INTERVAL 285 DAY)
	ELSE Bull_Turnout_Date 
	END into est_turn_date from(
		select @thrird_cow_date:=birth_date as ThirdCowDate from (
			select distinct cattle_info_tbl.dam_id,cattle_info_tbl.cow_age, owners_tbl.start_date AS enter_herd_date,cattle_info_tbl.birth_date as birth_date
			from cattle_info_tbl INNER JOIN owners_tbl on cattle_info_tbl.chaps_id=owners_tbl.chaps_id
			INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
			where cattle_info_tbl.birth_date >=Birth_Start_Date 
			AND cattle_info_tbl.birth_date <= Birth_End_Date
			and weaning_tbl.manage_code NOT IN( 'A','B','P' )
			And cattle_info_tbl.cow_age > 2
			order by cattle_info_tbl.birth_date
		) A order by birth_date LIMIT 2,1
)B;
return est_turn_date;
END //

/*--------Average_Daily_Gain------------*/
DROP FUNCTION IF EXISTS Average_Daily_Gain //
CREATE FUNCTION Average_Daily_Gain(Birth_Start_Date date,Birth_End_Date date) RETURNS double
BEGIN 
DECLARE avg_daily_gain  double;
SELECT SUM(CASE WHEN ADG>0 THEN ADG ELSE 0 END)/SUM(CASE WHEN ADG>0 THEN 1 ELSE 0 END) INTO avg_daily_gain FROM(
	SELECT 
		@age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS age_in_days,	
		CASE 
			WHEN @age_in_days > 0 and weight >0 THEN ROUND(@w2_day_gain:=(weight-birth_weight) /@age_in_days,1)
		ELSE 0
		END AS ADG
	FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
	INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
	INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
	where cattle_info_tbl.birth_date > Birth_Start_Date 
	AND cattle_info_tbl.birth_date < Birth_End_Date 
	AND measurement_tbl.entry_date<>'0000-00-00'
)a;
RETURN avg_daily_gain;
END//

/*--------------Early_Calves-----------*/
DROP FUNCTION IF EXISTS Early_Calves //
CREATE FUNCTION  Early_Calves(Birth_Start_Date date,Birth_End_Date date,Est_Bull_Turnout_Date date,calf_sex int,mature_cow_or_heifer_cow bool) RETURNS DOUBLE
BEGIN
	DECLARE early_calves double;
	SELECT COALESCE(COUNT(*),0) into early_calves FROM cattle_info_tbl
	LEFT JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
	WHERE cattle_info_tbl.sex in ('0','1','2','3')
	AND cattle_info_tbl.birth_date >= Birth_Start_Date
	AND  cattle_info_tbl.birth_date <=Birth_End_Date
	AND weaning_tbl.manage_code <>'A'
	AND weaning_tbl.manage_code <>'B'
	AND cattle_info_tbl.birth_date < DATE_ADD(Est_Bull_Turnout_Date, INTERVAL 285 DAY )
	AND 
			CASE mature_cow_or_heifer_cow
				WHEN  true THEN cattle_info_tbl.cow_age>2 
				WHEN false THEN cattle_info_tbl.cow_age=2
				ELSE cattle_info_tbl.cow_age>=2
			END
	AND 
		CASE calf_sex
			WHEN 4 THEN cattle_info_tbl.sex IN (0,1,2,3)
			ELSE cattle_info_tbl.sex=calf_sex
		END;
	return early_calves;

END //
/*-------------Calving_Within_21Days-----------*/
DROP FUNCTION IF EXISTS Calving_Within_21Days //
CREATE FUNCTION  Calving_Within_21Days(Birth_Start_Date date,Birth_End_Date date,Est_Bull_Turnout_Date date,calf_sex int,mature_cow_or_heifer_cow bool) RETURNS DOUBLE
BEGIN
	DECLARE Calving_In21Days double;
	SELECT COUNT(*) INTO Calving_In21Days FROM cattle_info_tbl 
		INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
		AND cattle_info_tbl.birth_date >= Birth_Start_Date
		AND  cattle_info_tbl.birth_date <= Birth_End_Date
		AND weaning_tbl.manage_code <>'A'
		AND weaning_tbl.manage_code <>'B'
		#AND cattle_info_tbl.birth_date >=DATE_ADD('2013-07-19', INTERVAL 285 DAY )
		AND cattle_info_tbl.birth_date <=DATE_ADD(Est_Bull_Turnout_Date, INTERVAL 285+20 DAY )
		AND 
			CASE mature_cow_or_heifer_cow
				WHEN  true THEN cattle_info_tbl.cow_age>2 
				WHEN false THEN cattle_info_tbl.cow_age=2
				ELSE cattle_info_tbl.cow_age>=2
			END
		AND 
			CASE calf_sex
				WHEN 4 THEN cattle_info_tbl.sex IN (0,1,2,3)
				ELSE cattle_info_tbl.sex=calf_sex
			END;
	RETURN Calving_In21Days;
END //
/*-------------Calving_Within_42Days-----------*/
DROP FUNCTION IF EXISTS Calving_Within_42Days//
CREATE FUNCTION  Calving_Within_42Days(Birth_Start_Date date,Birth_End_Date date,Est_Bull_Turnout_Date date,calf_sex int,mature_cow_or_heifer_cow bool) RETURNS DOUBLE
BEGIN
	DECLARE Calving_In42Days double;
	SELECT COUNT(*) INTO Calving_In42Days FROM cattle_info_tbl 
		INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
		AND cattle_info_tbl.birth_date >= Birth_Start_Date
		AND  cattle_info_tbl.birth_date <= Birth_End_Date
		AND weaning_tbl.manage_code <>'A'
		AND weaning_tbl.manage_code <>'B'
		#AND cattle_info_tbl.birth_date >=DATE_ADD('2013-07-19', INTERVAL 285+42 DAY )
		AND cattle_info_tbl.birth_date <=DATE_ADD(Est_Bull_Turnout_Date, INTERVAL 285+41 DAY )
		AND 
			CASE mature_cow_or_heifer_cow
				WHEN  true THEN cattle_info_tbl.cow_age>2 
				WHEN false THEN cattle_info_tbl.cow_age=2
				ELSE cattle_info_tbl.cow_age>=2
			END
		AND 
		CASE calf_sex
			WHEN 4 THEN cattle_info_tbl.sex IN (0,1,2,3)
			ELSE cattle_info_tbl.sex=calf_sex
		END;
	RETURN Calving_In42Days;
END //
/*-------------Calving_Within_63Days-----------*/
DROP FUNCTION IF EXISTS Calving_Within_63Days //
CREATE FUNCTION  Calving_Within_63Days(Birth_Start_Date date,Birth_End_Date date,Est_Bull_Turnout_Date date,calf_sex int,mature_cow_or_heifer_cow bool) RETURNS DOUBLE
BEGIN
	DECLARE Calving_In63Days double;
	SELECT COUNT(*) INTO Calving_In63Days FROM cattle_info_tbl 
		INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
		AND cattle_info_tbl.birth_date >= Birth_Start_Date
		AND  cattle_info_tbl.birth_date <= Birth_End_Date
		AND weaning_tbl.manage_code <>'A'
		AND weaning_tbl.manage_code <>'B'
		#AND cattle_info_tbl.birth_date >=DATE_ADD('2013-07-19', INTERVAL 285 DAY )
		AND cattle_info_tbl.birth_date <=DATE_ADD(Est_Bull_Turnout_Date, INTERVAL 285+62 DAY )
		AND 
			CASE mature_cow_or_heifer_cow
				WHEN  true THEN cattle_info_tbl.cow_age>2 
				WHEN false THEN cattle_info_tbl.cow_age=2
				ELSE cattle_info_tbl.cow_age>=2
			END
		AND 
		CASE calf_sex
			WHEN 4 THEN cattle_info_tbl.sex IN (0,1,2,3)
			ELSE cattle_info_tbl.sex=calf_sex
		END;
	RETURN Calving_In63Days;
END //
/*-------------Calving_After_63Days-----------*/
DROP FUNCTION IF EXISTS Calving_After_63Days //
CREATE FUNCTION  Calving_After_63Days(Birth_Start_Date date,Birth_End_Date date,Est_Bull_Turnout_Date date,calf_sex int,mature_cow_or_heifer_cow bool) RETURNS DOUBLE
BEGIN
	DECLARE Calving_After63Days double;
	SELECT COUNT(*) INTO Calving_After63Days FROM cattle_info_tbl 
		INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
		AND cattle_info_tbl.birth_date >= Birth_Start_Date
		AND  cattle_info_tbl.birth_date <= Birth_End_Date
		AND weaning_tbl.manage_code <>'A'
		AND weaning_tbl.manage_code <>'B'
		AND cattle_info_tbl.birth_date >=DATE_ADD(Est_Bull_Turnout_Date, INTERVAL 285+63 DAY )
		AND cattle_info_tbl.birth_date <=DATE_ADD(Est_Bull_Turnout_Date, INTERVAL 285+84 DAY )
		AND 
			CASE mature_cow_or_heifer_cow
				WHEN  true THEN cattle_info_tbl.cow_age>2 
				WHEN false THEN cattle_info_tbl.cow_age=2
				ELSE cattle_info_tbl.cow_age>=2
			END
		AND 
		CASE calf_sex
			WHEN 4 THEN cattle_info_tbl.sex IN (0,1,2,3)
			ELSE cattle_info_tbl.sex=calf_sex
		END;
	RETURN Calving_After63Days;
END //
/*--------------Sex_Based_Calf_List-----------*/
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
/*--------------Sex_Based_Group_Averages-----------*/
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
/*--------------Sex_Based_Cow_Breed_Averages-----------*/
drop PROCEDURE IF EXISTS  Sex_Based_Cow_Breed_Averages;//
CREATE PROCEDURE Sex_Based_Cow_Breed_Averages(IN calf_sex int,Birth_Start_Date date,Birth_End_Date date) 

BEGIN
	DECLARE no_calves int;
SELECT Cow_Breed, No_Of_Calves,Avg_Adj_205_Wt,Avg_Birth_Wt,Avg_Calving_Ease, Avg_Act_Wean_Wt,Avg_Age_In_Days,Avg_Frame_Score,Avg_ADG,Avg_WDA from(
SELECT Cow_ID, Cow_Breed, 
SUM(Case WHEN mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X') Then 1 ELSE 0 END) AS No_Of_Calves, 
SUM(CASE WHEN Adj_205_Wt>0 AND mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X')THEN Adj_205_Wt else 0 end )/SUM(CASE WHEN Adj_205_Wt>0 AND mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X') THEN 1 ELSE 0 END) AS Avg_Adj_205_Wt,
SUM(CASE WHEN Birth_Weight>0 AND mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X')THEN Birth_Weight ELSE 0 END)/SUM(CASE WHEN Birth_Weight>0 AND mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X')THEN 1 ELSE 0 END) AS Avg_Birth_Wt,
SUM(CASE WHEN Calving_Ease>=0 AND Calving_Ease<=4 THEN Calving_Ease ELSE 0 END)/COUNT(CASE WHEN Calving_Ease>=0 AND Calving_Ease<=4 THEN 1 ELSE 0 END) as Avg_Calving_Ease,
SUM(CASE WHEN Act_Wean_Weight>0 THEN Act_Wean_Weight ELSE 0 END)/SUM(CASE WHEN Act_Wean_Weight>0 THEN 1 ELSE 0 END) AS Avg_Act_Wean_Wt,
SUM(Age_in_Days) as sum_age_in_days,
SUM(CASE WHEN Age_in_Days>0 AND mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X')THEN Age_in_Days ELSE 0 END )/SUM(CASE WHEN mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X') Then 1 ELSE 0 END) As Avg_Age_In_Days,
SUM(CASE WHEN Frame_Score>0 THEN Frame_Score ELSE 0 END)/SUM(CASE WHEN Frame_Score>0 THEN 1 ELSE 0 END) AS Avg_Frame_Score,
SUM(CASE WHEN ADG>0 THEN ADG ELSE 0 end)/SUM(CASE WHEN ADG>0 THEN 1 ELSE 0 end) as Avg_ADG,
SUM(CASE WHEN WDA>0 then WDA ELSE 0 end)/SUM(CASE WHEN WDA>0 then 1 ELSE 0 end) as Avg_WDA
FROM (
 SELECT DISTINCTROW
		#cattle_info_tbl.chaps_id as Calf_ID,
		weaning_tbl.manage_code as mgt_code,
		@cow_id:=cattle_info_tbl.dam_ID as Cow_ID,
		@cow_breed:=(select breed from cattle_info_tbl where chaps_id=@cow_id) as Cow_Breed,
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
		@adj205wt:=ROUND((((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam,1) as Adj_205_Wt,
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
	where cattle_info_tbl.birth_date > Birth_Start_Date 
	AND cattle_info_tbl.birth_date < Birth_End_Date
	AND measurement_tbl.entry_date<>'0000-00-00'
	AND cattle_info_tbl.sex=calf_sex
)a
Group by Cow_Breed
)b where No_Of_Calves>0;
END //

/*--------------Sex_Based_Sire_Averages-----------*/
drop PROCEDURE IF EXISTS Sex_Based_Sire_Averages;//
CREATE PROCEDURE Sex_Based_Sire_Averages(IN calf_sex int,Birth_Start_Date date,Birth_End_Date date) 

BEGIN
	DECLARE avg_age DOUBLE;
	SET avg_age=Average_Calf_Age(Birth_Start_Date,Birth_End_Date);
	
SELECT Sire_ID, Sire_Breed,
SUM(Case WHEN mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X') Then 1 ELSE 0 END) AS No_Of_Calves, 
SUM(CASE WHEN Adj_205_Wt>0 AND mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X')THEN Adj_205_Wt else 0 end )/SUM(CASE WHEN Adj_205_Wt>0 AND mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X') THEN 1 ELSE 0 END) AS Avg_Adj_205_Wt,
SUM(CASE WHEN Birth_Weight>0 AND mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X')THEN Birth_Weight ELSE 0 END)/SUM(CASE WHEN Birth_Weight>0 AND mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X')THEN 1 ELSE 0 END) AS Avg_Birth_Wt,
#SUM(CASE WHEN Calving_Ease>=0 AND Calving_Ease<=4 THEN Calving_Ease ELSE 0 END)/COUNT(CASE WHEN Calving_Ease>=0 AND Calving_Ease<=4 THEN 1 ELSE 0 END) as Avg_Calving_Ease,
SUM(Calving_Ease)/SUM(denom) as Average_Calving_Ease,
SUM(CASE WHEN Act_Wean_Weight>0 THEN Act_Wean_Weight ELSE 0 END)/SUM(CASE WHEN Act_Wean_Weight>0 THEN 1 ELSE 0 END) AS Avg_Act_Wean_Wt,
#SUM(Age_in_Days) as sum_age_in_days,
SUM(CASE WHEN Age_in_Days>0 AND mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X')THEN Age_in_Days ELSE 0 END )/SUM(CASE WHEN mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X') Then 1 ELSE 0 END) As Avg_Age_In_Days,
SUM(CASE WHEN Frame_Score>0 THEN Frame_Score ELSE 0 END)/SUM(CASE WHEN Frame_Score>0 THEN 1 ELSE 0 END) AS Avg_Frame_Score,
SUM(CASE WHEN ADG>0 THEN ADG ELSE 0 end)/SUM(CASE WHEN ADG>0 THEN 1 ELSE 0 end) as Avg_ADG,
SUM(CASE WHEN WDA>0 then WDA ELSE 0 end)/SUM(CASE WHEN WDA>0 then 1 ELSE 0 end) as Avg_WDA
FROM (
 SELECT DISTINCTROW
		cattle_info_tbl.chaps_id as Calf_ID,
		weaning_tbl.manage_code as mgt_code,
		@sire_id:=cattle_info_tbl.sire_ID as Sire_ID,
		@sire_breed:=(select breed from cattle_info_tbl where chaps_id=@sire_id) as Sire_Breed,
		cattle_info_tbl.birth_weight as Birth_weight,
		@age_in_days:= DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) AS Age_in_Days,
				@irr_calf:=
			CASE 
				WHEN @age_in_days>avg_age+45 or @age_in_days<avg_age-45 THEN 'T'
				ELSE 'F'
		END AS irr_calf,
		CASE WHEN @irr_calf='F' THEN 1 ELSE 0 END as denom,
		CASE WHEN @irr_calf='F' THEN measurement_tbl.weight ELSE 0 END as Act_Wean_Weight,
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
		@adj205wt:=ROUND((((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam,1) as Adj_205_Wt,
		measurement_tbl.frame_score as Frame_Score,
		CASE 
			WHEN @age_in_days > 0 and weight >0 THEN ROUND(@w2_day_gain:=(weight-birth_weight) /@age_in_days,1)
		ELSE 0
		END AS ADG,
		CASE 
			WHEN @age_in_days > 0 and weight >0 THEN ROUND(@avg_daily_gain:=weight/@age_in_days,1)
		ELSE 0
		END AS WDA,
		CASE WHEN @irr_calf='F' THEN cattle_info_tbl.calving_ease ELSE 0 END as Calving_Ease
	
		
	FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
	INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
	INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
	where cattle_info_tbl.birth_date > Birth_Start_Date 
	AND cattle_info_tbl.birth_date < Birth_End_Date 
	AND measurement_tbl.entry_date<>'0000-00-00'
	AND cattle_info_tbl.sex=calf_sex
)a
Group by sire_id;
END //
/*--------------Sire_Summary-----------*/
DROP PROCEDURE IF EXISTS Sire_Summary //
CREATE PROCEDURE Sire_Summary(Birth_Start_Date date,Birth_End_Date date)
BEGIN 
SELECT Sire_ID, Sire_Breed,
SUM(Case WHEN Irr_Calf='F' Then 1 ELSE 0 END) AS No_Of_Calves, 
SUM(CASE WHEN Birth_Weight>0 AND Irr_Calf='F' THEN Birth_Weight ELSE 0 END)/SUM(CASE WHEN Birth_Weight>0 AND Irr_Calf='F' THEN 1 ELSE 0 END) AS Avg_Birth_Wt,
SUM(CASE WHEN Act_Wean_Weight>0 AND  Irr_Calf='F' THEN Act_Wean_Weight ELSE 0 END)/SUM(CASE WHEN Act_Wean_Weight>0 AND  Irr_Calf='F' THEN 1 ELSE 0 END) AS Avg_Act_Wean_Wt,
SUM(CASE WHEN Irr_Calf='F'  THEN rev_adj205wt else 0 end )/SUM(CASE WHEN Irr_Calf='F' THEN 1 ELSE 0 END) AS Avg_Adj_205_Wt,
#SUM(CASE WHEN Birth_Weight>0 AND mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X')THEN Birth_Weight ELSE 0 END)/SUM(CASE WHEN Birth_Weight>0 AND mgt_code NOT IN('A','B','C','D','F','K','N','P','S','T','X')THEN 1 ELSE 0 END) AS Avg_Birth_Wt,
SUM(CASE WHEN ADG>0 THEN ADG ELSE 0 end)/SUM(CASE WHEN ADG>0 THEN 1 ELSE 0 end) as Avg_ADG,
SUM(CASE WHEN WDA>0 then WDA ELSE 0 end)/SUM(CASE WHEN WDA>0 then 1 ELSE 0 end) as Avg_WDA,
SUM(CASE WHEN Calving_Ease>=0 AND Calving_Ease<=4 THEN Calving_Ease ELSE 0 END)/COUNT(CASE WHEN Calving_Ease>=0 AND Calving_Ease<=4 THEN 1 ELSE 0 END) as Avg_Calving_Ease,
SUM(CASE WHEN Age_in_Days>0 AND Irr_Calf='F' THEN Age_in_Days ELSE 0 END )/SUM(CASE WHEN  Irr_Calf='F' Then 1 ELSE 0 END) As Avg_Age_In_Days,
SUM(CASE WHEN Frame_Score>0 THEN Frame_Score ELSE 0 END)/SUM(CASE WHEN Frame_Score>0 THEN 1 ELSE 0 END) AS Avg_Frame_Score
FROM (
 SELECT DISTINCTROW
		cattle_info_tbl.chaps_id as Calf_ID,
		weaning_tbl.manage_code as mgt_code,
		@sire_id:=cattle_info_tbl.sire_ID as Sire_ID,
		@sire_breed:=(select breed from cattle_info_tbl where chaps_id=@sire_id) as Sire_Breed,
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
		@adj205wt:=ROUND((((weight-@adj_birth_wt)/@age_in_days)*205)+@adj_birth_wt+@dam,1) as Adj_205_Wt,
		@rev_adj205wt:=
			CASE WHEN sex=0 THEN @adj205wt
				WHEN sex=1 THEN @adj205wt*0.95
				WHEN sex=2 THEN @adj205wt*1.05
				WHEN sex=3 THEN @adj205wt
		END as Rev_Adj205wt,
		@Avg_Age:=AverAge_Calf_Age(Birth_Start_Date,Birth_End_Date) as Avg_Age,
		@irr_calf:=
			CASE 
				WHEN @age_in_days>@Avg_Age+45 or @age_in_days<@Avg_Age-45 THEN 'T'
				ELSE 'F'
			END AS Irr_Calf,		
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
	where cattle_info_tbl.birth_date > Birth_Start_Date 
	AND cattle_info_tbl.birth_date < Birth_End_Date
	AND weaning_tbl.manage_code NOT IN('A','B','C','D','F','K','N','P','S','T','X')
	#AND measurement_tbl.entry_date<>'0000-00-00'
	#AND cattle_info_tbl.sex in (2,3)
)a
Group by sire_id;
END //
/*--------------Summary_Of_Cows_Present_Herd-----------*/
DROP PROCEDURE IF EXISTS Summary_Of_Cows_Present_In_Herd //
Create PROCEDURE Summary_Of_Cows_Present_In_Herd(Birth_Start_Date date,Birth_End_Date date,Bull_TurnOut_Date date)
BEGIN
DECLARE est_turn_date DATE;
set est_turn_date=Estimated_Bull_Turnout_Date(Birth_Start_Date, Birth_End_Date, Bull_Turnout_Date);
SELECT DISTINCT 
CASE b.Measures
	WHEN 'Total_Cows_Exposed' THEN 'Total_Cows_Exposed'
	WHEN 'Total_Cows_Kept_For_Calving' THEN 'Total_Cows_Kept_For_Calving'
	WHEN 'Number_Of_Cows_Aborted' THEN 'Number_Of_Cows_Aborted'
	WHEN 'Number_Of_Cows_Open' THEN 'Number_Of_Cows_Open'
	WHEN 'Number_Of_Cows_Calving' THEN 'Number_Of_Cows_Calving'
	WHEN 'Number_Of_Cows_Losing_Calf' THEN 'Number_Of_Cows_Losing_Calf'
	WHEN 'Number_Of_Cows_Weaning_Calf' THEN 'Number_Of_Cows_Weaning_Calf'
END as Measures,
CASE b. Measures
	WHEN 'Total_Cows_Exposed' THEN Total_Cows_Exposed
	WHEN 'Total_Cows_Kept_For_Calving' THEN Total_Cows_Kept_For_Calving
	WHEN 'Number_Of_Cows_Aborted' THEN Number_Of_Cows_Aborted
	WHEN 'Number_Of_Cows_Open' THEN Number_Of_Cows_Open
	WHEN 'Number_Of_Cows_Calving' THEN Number_Of_Cows_Calving
	WHEN 'Number_Of_Cows_Losing_Calf' THEN Number_Of_Cows_Losing_Calf
	WHEN 'Number_Of_Cows_Weaning_Calf' THEN Number_Of_Cows_Weaning_Calf
END as Count
FROM(
SELECT
(SELECT (select SUM(CASE WHEN weaning_tbl.manage_code='T' Then 0.5 ELSE 1 END)
from cattle_info_tbl INNER JOIN weaning_tbl ON cattle_info_tbl.chaps_id=weaning_tbl.chaps_id  where cattle_info_tbl.birth_date >= Birth_Start_Date AND  cattle_info_tbl.birth_date <=Birth_End_Date)
+
(select count(*) from owners_tbl where exit_date between est_turn_date AND DATE_ADD(est_turn_date, INTERVAL 365 DAY ))) as Total_Cows_Exposed,
@kept_for_calving:=(select SUM(CASE WHEN weaning_tbl.manage_code='T' Then 0.5 ELSE 1 END)
from cattle_info_tbl INNER JOIN weaning_tbl ON cattle_info_tbl.chaps_id=weaning_tbl.chaps_id  where cattle_info_tbl.birth_date >= Birth_Start_Date AND  cattle_info_tbl.birth_date <=Birth_End_Date) as Total_Cows_Kept_For_Calving,
@cows_aborted:=(select count(*) from weaning_tbl where entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date and manage_code ='B') as Number_Of_Cows_Aborted,
@cows_open:=(select count(*) from weaning_tbl where entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date and manage_code ='A') as Number_Of_Cows_Open,
@cows_calving:=@kept_for_calving-@cows_aborted-@cows_open as Number_Of_Cows_Calving,
@cows_losing_calves:=(select count(*) from weaning_tbl inner join cattle_info_tbl on weaning_tbl.chaps_id=cattle_info_tbl.chaps_id where cattle_info_tbl.birth_date >= Birth_Start_Date and  cattle_info_tbl.birth_date<=Birth_End_Date  and weaning_tbl.manage_code in ('C','D','F','K')) as Number_Of_Cows_Losing_Calf,
@cows_weaning_calves:=@cows_calving-@cows_losing_calves as Number_Of_Cows_Weaning_Calf
)a
cross join
(
	select 'Total_Cows_Exposed' as Measures
	union all select 'Total_Cows_Kept_For_Calving'
	union all select 'Number_Of_Cows_Aborted'
	union all select 'Number_Of_Cows_Open'
	union all select 'Number_Of_Cows_Calving'	
	union all select 'Number_Of_Cows_Losing_Calf'
	union all select 'Number_Of_Cows_Weaning_Calf'

)b;

END //

/*--------------Herd_Summary_Calving_Distribution_Table-----------*/
DROP PROCEDURE IF EXISTS Herd_Summary_Calving_Distribution_Table //
CREATE PROCEDURE Herd_Summary_Calving_Distribution_Table(Birth_Start_Date date,Birth_End_Date date,Bull_Turnout_Date date)
BEGIN
DECLARE est_turn_date DATE;
set est_turn_date=Estimated_Bull_Turnout_Date(Birth_Start_Date, Birth_End_Date, Bull_Turnout_Date);

SELECT a.Dam_Age, a.Calves_each_age,COALESCE(b.Early_Calves,0) as Early_Calves,COALESCE(c.First_21_Calves,0) as'1st 21 Calves' ,
COALESCE(d.Second_21_Calves,0) as '2nd 21 Calves',COALESCE(e.Third_21_Calves,0) as '3rd 21 Calves',
COALESCE(f.Fourth_21_Calves,0) as '4th 21 Calves',COALESCE(g.Late_Calves,0) as 'Late Calves',
COALESCE(h.open_Aborted,0) as 'Cows Open/aborted',
avg_date_each_age,avg_actual_wean_weight from

(SELECT cattle_info_tbl.cow_age as Dam_Age,COUNT(*) AS Calves_each_age,
CAST(from_unixtime(AVG(UNIX_TIMESTAMP(cattle_info_tbl.birth_date))) AS DATE) AS Avg_Date_Each_Age,
SUM(CASE WHEN measurement_tbl.weight>0 THEN  measurement_tbl.weight ELSE 0 END )/SUM(CASE WHEN measurement_tbl.weight>0 THEN  1 ELSE 0 END )as Avg_Actual_Wean_Weight
FROM cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
INNER JOIN measurement_tbl on measurement_tbl.chaps_id=cattle_info_tbl.chaps_id
WHERE
cattle_info_tbl.birth_date >= Birth_Start_Date
AND  cattle_info_tbl.birth_date <= Birth_End_Date
AND weaning_tbl.manage_code <>'A'
AND weaning_tbl.manage_code <>'B'
GROUP BY cattle_info_tbl.cow_age)a
Left JOIN

(SELECT cattle_info_tbl.cow_age as Dam_Age, COUNT(*) AS Early_Calves FROM cattle_info_tbl
 INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE cattle_info_tbl.sex in ('0','1','2','3')
AND cattle_info_tbl.birth_date >= Birth_Start_Date
AND  cattle_info_tbl.birth_date <= Birth_End_Date
AND weaning_tbl.manage_code <>'A'
AND weaning_tbl.manage_code <>'B'
AND cattle_info_tbl.birth_date < DATE_ADD(est_turn_date, INTERVAL 285 DAY )
GROUP BY cattle_info_tbl.cow_age
)b on a.Dam_Age=b.Dam_Age
LEFT JOIN

(SELECT cow_age as Dam_Age, COUNT(*) AS First_21_Calves FROM cattle_info_tbl 
INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE cattle_info_tbl.sex in ('0','1','2','3')
AND cattle_info_tbl.birth_date >= Birth_Start_Date
AND  cattle_info_tbl.birth_date <= Birth_End_Date
AND weaning_tbl.manage_code <>'A'
AND weaning_tbl.manage_code <>'B'
AND cattle_info_tbl.birth_date >=DATE_ADD(est_turn_date, INTERVAL 285 DAY )
AND cattle_info_tbl.birth_date <=DATE_ADD(est_turn_date, INTERVAL 285+20 DAY )
GROUP BY cow_age
)c on a.Dam_Age=c.Dam_Age
LEFT JOIN

(SELECT cow_age as Dam_Age, COUNT(*) AS Second_21_Calves  FROM cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE cattle_info_tbl.sex in ('0','1','2','3')
AND cattle_info_tbl.birth_date >= Birth_Start_Date
AND  cattle_info_tbl.birth_date <= Birth_End_Date
AND weaning_tbl.manage_code <>'A'
AND weaning_tbl.manage_code <>'B'
AND cattle_info_tbl.birth_date >=DATE_ADD(est_turn_date, INTERVAL 285+21 DAY )
AND cattle_info_tbl.birth_date <=DATE_ADD(est_turn_date, INTERVAL 285+41 DAY )
GROUP BY cow_age)d on a.Dam_Age=d.Dam_Age
LEFT JOIN
(SELECT cow_age Dam_Age, COUNT(*) AS Third_21_Calves FROM cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
AND cattle_info_tbl.birth_date >= Birth_Start_Date
AND  cattle_info_tbl.birth_date <= Birth_End_Date
AND weaning_tbl.manage_code <>'A'
AND weaning_tbl.manage_code <>'B'
AND cattle_info_tbl.birth_date >=DATE_ADD(est_turn_date, INTERVAL 285+42 DAY )
AND cattle_info_tbl.birth_date <=DATE_ADD(est_turn_date, INTERVAL 285+62 DAY )
GROUP BY cow_age)e on a.Dam_Age=e.Dam_Age
LEFT JOIN 

(SELECT cow_age Dam_Age, COUNT(*) AS  Fourth_21_Calves FROM cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
AND cattle_info_tbl.birth_date >= Birth_Start_Date
AND  cattle_info_tbl.birth_date <= Birth_End_Date
AND weaning_tbl.manage_code <>'A'
AND weaning_tbl.manage_code <>'B'
AND cattle_info_tbl.birth_date >=DATE_ADD(est_turn_date, INTERVAL 285+63 DAY )
AND cattle_info_tbl.birth_date <=DATE_ADD(est_turn_date, INTERVAL 285+83 DAY )
GROUP BY cow_age)f on a.Dam_Age=f.Dam_Age
LEFT JOIN

(SELECT cow_age Dam_Age, COUNT(*) AS Late_Calves  FROM cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
AND cattle_info_tbl.birth_date >= Birth_Start_Date
AND  cattle_info_tbl.birth_date <= Birth_End_Date
AND weaning_tbl.manage_code <>'A'
AND weaning_tbl.manage_code <>'B'
AND cattle_info_tbl.birth_date >DATE_ADD(est_turn_date, INTERVAL 285+84 DAY )
GROUP BY cow_age)g ON a.Dam_Age=g.Dam_Age
LEFT JOIN

(SELECT cow_age Dam_Age, COUNT(*) AS Open_Aborted  FROM cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
AND cattle_info_tbl.birth_date >= Birth_Start_Date
AND  cattle_info_tbl.birth_date <= Birth_End_Date
AND weaning_tbl.manage_code IN('A','B')
GROUP BY cow_age)h ON a.Dam_Age=h.Dam_Age;

END //

/*--------------Sex_Based_Calves_Born_Weighed_Unweighed_Count-----------*/
DROP PROCEDURE IF EXISTS Sex_Based_Calves_Born_Weighed_Unweighed_Count //
CREATE PROCEDURE Sex_Based_Calves_Born_Weighed_Unweighed_Count(Birth_Start_Date date,Birth_End_Date date)
BEGIN
/*Calves born Bulls*/
SELECT 
(SELECT COUNT(*)  FROM cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE weaning_tbl.manage_code NOT IN ('A','B') AND cattle_info_tbl.sex=1 AND
cattle_info_tbl.birth_date > Birth_Start_Date AND cattle_info_tbl.birth_date < Birth_End_Date)AS 'Calves born bulls',
/*Calves born Heifers*/
(SELECT COUNT(*)  FROM cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE weaning_tbl.manage_code NOT IN ('A','B') AND cattle_info_tbl.sex=2 AND
cattle_info_tbl.birth_date > Birth_Start_Date AND cattle_info_tbl.birth_date < Birth_End_Date) AS 'Calves born Heifers',

/*Calves born Steers*/
(SELECT COUNT(*)  FROM cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE weaning_tbl.manage_code NOT IN ('A','B') AND cattle_info_tbl.sex=3 AND
cattle_info_tbl.birth_date > Birth_Start_Date AND cattle_info_tbl.birth_date < Birth_End_Date) AS 'Calves born Steers',

/*Calves weighed bulls*/
(SELECT COUNT(*)  from cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
INNER JOIN measurement_tbl on weaning_tbl.chaps_id = measurement_tbl.chaps_id
WHERE weaning_tbl.manage_code NOT IN ('A','B') AND
cattle_info_tbl.sex=1 AND cattle_info_tbl.birth_weight > 0 AND 
cattle_info_tbl.birth_date > Birth_Start_Date AND cattle_info_tbl.birth_date < Birth_End_Date AND measurement_tbl.weight <> 0) AS 'Calves weighed bulls',

/*Calves weighed Heifers*/
(SELECT COUNT(*)  from cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
INNER JOIN measurement_tbl on weaning_tbl.chaps_id = measurement_tbl.chaps_id
WHERE weaning_tbl.manage_code NOT IN ('A','B') AND
cattle_info_tbl.sex=2 AND cattle_info_tbl.birth_weight > 0 AND 
cattle_info_tbl.birth_date > Birth_Start_Date AND cattle_info_tbl.birth_date < Birth_End_Date AND measurement_tbl.weight <> 0) AS 'Calves weighed Heifers',
/*Calves weighed Steers*/
(SELECT COUNT(*)  from cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
INNER JOIN measurement_tbl on weaning_tbl.chaps_id = measurement_tbl.chaps_id
WHERE weaning_tbl.manage_code NOT IN ('A','B') AND
cattle_info_tbl.sex=3 AND cattle_info_tbl.birth_weight > 0 AND 
cattle_info_tbl.birth_date > Birth_Start_Date AND cattle_info_tbl.birth_date < Birth_End_Date AND measurement_tbl.weight <> 0) AS 'Calves weighed Steers',

/*Calves Not weighed bulls*/
(SELECT COUNT(*)  from cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE weaning_tbl.manage_code ='X' AND
cattle_info_tbl.sex=1 AND cattle_info_tbl.birth_weight > 0 AND 
cattle_info_tbl.birth_date > Birth_Start_Date AND cattle_info_tbl.birth_date < Birth_End_Date)AS 'Calves not weighed bulls',

/*Calves Not weighed Heifers*/
(SELECT COUNT(*)  from cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE weaning_tbl.manage_code ='X' AND
cattle_info_tbl.sex=2 AND cattle_info_tbl.birth_weight > 0 AND 
cattle_info_tbl.birth_date > Birth_Start_Date AND cattle_info_tbl.birth_date < Birth_End_Date)AS 'Calves not weighed Heifers',
/*Calves Not weighed Steers*/
(SELECT COUNT(*)  from cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE weaning_tbl.manage_code ='X' AND
cattle_info_tbl.sex=3 AND cattle_info_tbl.birth_weight > 0 AND 
cattle_info_tbl.birth_date > Birth_Start_Date AND cattle_info_tbl.birth_date < Birth_End_Date ) AS 'Calves not weighed Steers';

END //
/*--------------Herd_Summary_Averages-----------*/
DROP PROCEDURE IF EXISTS Herd_Summary_Averages //
CREATE PROCEDURE Herd_Summary_Averages(Birth_Start_Date date,Birth_End_Date date)
BEGIN
	SELECT avg_age as Average_Age_At_Weaning,
			wt_2_day_gain as Weight_Per_Day_To_Weaning,
			sum(CASE WHEN birth_weight>0  THEN birth_weight ELSE 0 END)/SUM(CASE WHEN birth_weight>0 THEN 1 ELSE 0 END)as Average_Birth_Weight,
			#avg(birth_weight) as 'Average Birth Weight',
			avg_adj_205_day_wt as Average_Sex_Adjusted_Weight_205
	from (
		SELECT DISTINCTROW
		cattle_info_tbl.chaps_id,
		cattle_info_tbl.birth_weight,
		@avg_age:=Average_Calf_Age(Birth_Start_Date,Birth_End_Date ) as avg_age,
		Avg_Wt_Per_Day_Of_Age_To_Weaning(Birth_Start_Date,Birth_End_Date) as wt_2_day_gain,
		Adjusted_205_Day_Wt(Birth_Start_Date,Birth_End_Date ) as avg_adj_205_day_wt

	FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
	INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
	INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
	where cattle_info_tbl.birth_date >= Birth_Start_Date 
	AND cattle_info_tbl.birth_date <= Birth_End_Date 
	)a
;
END //
/*--------------Summary_Of_Cows_Leaving-----------*/
DROP PROCEDURE IF EXISTS Summary_Of_Cows_Leaving //
CREATE PROCEDURE Summary_Of_Cows_Leaving(Birth_Start_Date date,Birth_End_Date date, Bull_Turnout_Date date)
BEGIN
/*Cows Died*/
DECLARE est_turn_date DATE;
#DECLARE third_cow_date date;
set est_turn_date=Estimated_Bull_Turnout_Date(Birth_Start_Date, Birth_End_Date, Bull_Turnout_Date);

SELECT * FROM(
SELECT 
(SELECT COUNT(*) as Cows_Died from owners_tbl where exit_code='G' and exit_date between est_turn_date AND DATE_ADD(est_turn_date, INTERVAL 366 DAY )
 ) AS Cows_Died,

#/*Cows sold because of age*/
(SELECT COUNT(*) from owners_tbl where exit_code='H' and exit_date between est_turn_date AND DATE_ADD(est_turn_date, INTERVAL 365 DAY )  )as 'Cows Sold Because Of Age',

#/*Cows sold Because of physical defects*/
(SELECT COUNT(*) from owners_tbl where exit_code='J' and exit_date between est_turn_date AND DATE_ADD(est_turn_date, INTERVAL 365 DAY )) as 'Cows Sold Because of Physical Defects',

/*Cows sold because of poor fertility or open*/
(SELECT COUNT(*)from owners_tbl where exit_code='K' and exit_date between est_turn_date AND DATE_ADD(est_turn_date, INTERVAL 365 DAY )) as 'Cows sold because of poor fertility or open',

#/*Cows sold because of inferior calves*/
(SELECT COUNT(*) from owners_tbl where exit_code='L' and exit_date between est_turn_date AND DATE_ADD(est_turn_date, INTERVAL 365 DAY ) ) as 'Cows sold because of inferior calves',

/*Cows sold for replacement stock*/
(SELECT COUNT(*)  from owners_tbl where exit_code='R' and exit_date between est_turn_date AND DATE_ADD(est_turn_date, INTERVAL 365 DAY )  ) as 'Cows sold because of replacement stock',

/*Cows sold for unknown reason*/
(SELECT COUNT(*)  from owners_tbl where exit_code='Y'and exit_date between est_turn_date AND DATE_ADD(est_turn_date, INTERVAL 365 DAY ) ) as 'Cows sold for unknown reason'
)A;
END // 
/*--------------Herd_Comparison-----------*/
DROP PROCEDURE IF EXISTS Herd_Comparison //
Create Procedure Herd_Comparison(Birth_Start_Date date,Birth_End_Date date,Bull_Turnout_Date date)
BEGIN 
DECLARE est_turn_date DATE;
set est_turn_date=Estimated_Bull_Turnout_Date(Birth_Start_Date, Birth_End_Date, Bull_Turnout_Date);
SELECT 
CASE  b.Critical_Succes_Factors 
		WHEN "Calf_Production_Time" THEN 'Calf Production Time'
		WHEN "Frame_Score" THEN 'Frame Score'	
		WHEN "Birth_Weight" THEN 'Birth Weight'
		WHEN "Wt_Per_Day_Of_Age" THEN 'Weight Per Day Of Age'
		WHEN "Average_Daily_Gain" THEN 'Avergae Daily Gain In Weight'
		WHEN "Heifers_Calving_Early" THEN 'Percentage of Heifers Calving Early'
		WHEN  "Heifers_Calving_Within_21_Days" THEN 'Percentage of Heifers Calving Within 21 Days'
		WHEN  "Heifers_Calving_Within_42_Days" THEN 'Percentage of Heifers Calving Within 42 Days'
		WHEN  "Mature_Cows_Calving_Within_21_Days" THEN 'Percentage of Mature Cows Calving Within 21 Days'
		WHEN  "Mature_Cows_Calving_Within_42_Days" THEN 'Percentage of Mature Cows Calving Within 42 Days'
		WHEN "Replacements_Kept_To_Calve" THEN 'Replacements Kept To Calve'
		WHEN "Cow_Weight_At_Weaning" THEN 'Cow_Weight_At_Weaning'
		WHEN "Cow_Condition_Score_At_Weaning" THEN 'Cow_Condition_Score_At_Weaning'
		WHEN "Avg_Cow_Age" THEN 'Average Cow Age'
		WHEN "Adjusted_205_Day_Wt" THEN 'Adjusted 205 Day Weight'
		WHEN  "Actual_Weaning_Wt_Steers" THEN 'Actual Weaning Weight For Steer Calves'
		WHEN "Actual_Weaning_Wt_Heifers" THEN 'Actual Weaning Weight For Heifer Calves'
		WHEN "Actual_Weaning_Wt_Bulls" THEN 'Actual Weaning Weight For Bull Calves'
END as Critical_Success_Factors,
CASE  b.Critical_Succes_Factors 

		WHEN "Frame_Score" THEN Frame_Score
		WHEN "Calf_Production_Time" THEN Calf_Production_Time 
		WHEN "Birth_Weight" THEN Round(Birth_Weight,1)
		WHEN "Wt_Per_Day_Of_Age" THEN Round(Wt_Per_Day_Of_Age,2)
		WHEN "Average_Daily_Gain" THEN Round(Average_Daily_Gain,1)
		WHEN "Heifers_Calving_Early" THEN CONCAT(FORMAT(Heifers_Calving_Early*100,2)," %")
		WHEN  "Heifers_Calving_Within_21_Days" THEN CONCAT(FORMAT(Heifers_Calving_Within_21_Days*100,2)," %")
		WHEN  "Heifers_Calving_Within_42_Days" THEN CONCAT(FORMAT(Heifers_Calving_Within_42_Days*100,2)," %")
		WHEN  "Mature_Cows_Calving_Within_21_Days" THEN CONCAT(FORMAT(Mature_Cows_Calving_Within_21_Days*100,2)," %")
		WHEN  "Mature_Cows_Calving_Within_42_Days" THEN CONCAT(FORMAT(Mature_Cows_Calving_Within_42_Days*100,2)," %")
		WHEN "Replacements_Kept_To_Calve" THEN Replacements_Kept_To_Calve
		WHEN "Cow_Weight_At_Weaning" THEN Cow_Weight_At_Weaning
		WHEN "Cow_Condition_Score_At_Weaning" THEN Cow_Condition_Score_At_Weaning
		WHEN "Avg_Cow_Age" THEN Round(Avg_Cow_Age,1)
		WHEN "Adjusted_205_Day_Wt" THEN Round(Adjusted_205_Day_Wt,1)
		WHEN  "Actual_Weaning_Wt_Steers" THEN Round(Actual_Weaning_Wt_Steers,1)
		WHEN "Actual_Weaning_Wt_Heifers" THEN Round(Actual_Weaning_Wt_Heifers,1)
		WHEN "Actual_Weaning_Wt_Bulls" THEN Round(Actual_Weaning_Wt_Bulls,1)
END as Your_Herd_Performance
FROM(
	SELECT  Calf_Production_Time,Frame_Score,Birth_Weight,Wt_Per_Day_Of_Age,Average_Daily_Gain,
		    @early_heifer_calves/(@within_63days_heifer_calves+@after_63days_heifer_calves) as Heifers_Calving_Early,
			@within_21days_heifer_calves/(@within_63days_heifer_calves+@after_63days_heifer_calves) as Heifers_Calving_Within_21_Days,
			@within_42days_heifer_calves/(@within_63days_heifer_calves+@after_63days_heifer_calves) as Heifers_Calving_Within_42_Days,
			@within_21days_mature_calving/(@within_63days_mature_calving+@after_63days_mature_calving) as Mature_Cows_Calving_Within_21_Days,
			@within_42days_mature_calving/(@within_63days_mature_calving+@after_63days_mature_calving) as Mature_Cows_Calving_Within_42_Days,
			Replacements_Kept_To_Calve,		
			(select SUM(CASE WHEN wean_condition_score >0 THEN wean_condition_score ELSE 0 END)/SUM(CASE WHEN wean_condition_score >0 THEN 1 ELSE 0 END)
			from cow_breeding_tbl where cow_breeding_tbl.bull_turnout_date=Bull_Turnout_Date) as Cow_Condition_Score_At_Weaning,
			(select SUM(CASE WHEN wean_weight >0 THEN wean_weight ELSE 0 END)/SUM(CASE WHEN wean_weight >0 THEN 1 ELSE 0 END)
			from cow_breeding_tbl where cow_breeding_tbl.bull_turnout_date=Bull_Turnout_Date) Cow_Weight_At_Weaning,
			Avg_Cow_Age,Adjusted_205_Day_Wt,Actual_Weaning_Wt_Steers, Actual_Weaning_Wt_Heifers, Actual_Weaning_Wt_Bulls FROM(
				SELECT
					#ROUND(SUM(CASE WHEN DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date)>0 AND weaning_tbl.manage_code NOT IN  ('A','B','C','D') THEN DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date) ELSE 0 END )/SUM(CASE WHEN DATEDIFF(measurement_tbl.entry_date,cattle_info_tbl.birth_date)>0 AND weaning_tbl.manage_code NOT IN  ('A','B','C','D') THEN 1 ELSE 0 END ) ,1) AS Calf_Production_Time,
					@nursing_period:=Average_Calf_Age(Birth_Start_Date,Birth_End_Date ) AS Calf_Production_Time,
					SUM(CASE WHEN cattle_info_tbl.birth_weight>0 Then cattle_info_tbl.birth_weight ELSE 0 END)/SUM(CASE WHEN cattle_info_tbl.birth_weight>0 Then 1 ELSE 0 END) as Birth_Weight,
					Avg_Wt_Per_Day_Of_Age_To_Weaning(Birth_Start_Date,Birth_End_Date) as Wt_Per_Day_Of_Age,
					Average_Daily_Gain(Birth_Start_Date,Birth_End_Date) as Average_Daily_Gain,
					ROUND(SUM(CASE WHEN Frame_Score>0 THEN Frame_Score ELSE 0 END)/SUM(CASE WHEN Frame_Score>0 THEN 1 ELSE 0 END) ,1) AS Frame_Score,
					@early_heifer_calves:=Early_Calves(Birth_Start_Date,Birth_End_Date,est_turn_date,4,false),
					@within_21days_heifer_calves:=Calving_Within_21Days(Birth_Start_Date,Birth_End_Date,est_turn_date,4,false),
					@within_42days_heifer_calves:=Calving_Within_42Days(Birth_Start_Date,Birth_End_Date,est_turn_date,4,false),
					@within_63days_heifer_calves:=Calving_Within_63Days(Birth_Start_Date,Birth_End_Date,est_turn_date,4,false),
					@after_63days_heifer_calves:=Calving_After_63Days(Birth_Start_Date,Birth_End_Date,est_turn_date,4,false),
					@within_21days_mature_calving:=Calving_Within_21Days(Birth_Start_Date,Birth_End_Date,est_turn_date,4,true),
					@within_42days_mature_calving:=Calving_Within_42Days(Birth_Start_Date,Birth_End_Date,est_turn_date,4,true),
					@within_63days_mature_calving:=Calving_Within_63Days(Birth_Start_Date,Birth_End_Date,est_turn_date,4,true),
					@after_63days_mature_calving:=Calving_After_63Days(Birth_Start_Date,Birth_End_Date,est_turn_date,4,true),
					AVG(cattle_info_tbl.cow_age) as Avg_Cow_Age,
					SUM(CASE WHEN cattle_info_tbl.cow_age<3 THEN 1 ELSE 0 END ) as Replacements_Kept_To_Calve,
					Adjusted_205_Day_Wt(Birth_Start_Date, Birth_End_Date) as Adjusted_205_Day_Wt,
					Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,2) as Actual_Weaning_Wt_Heifers,
					Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,3) as Actual_Weaning_Wt_Steers,
					Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,1) as Actual_Weaning_Wt_Bulls
				FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
				INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
				INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
				WHERE cattle_info_tbl.birth_date >= Birth_Start_Date
				AND cattle_info_tbl.birth_date <= Birth_End_Date
		)x
)a
cross join
(
	select "Calf_Production_Time" as Critical_Succes_Factors	
	union all select "Wt_Per_Day_Of_Age"
	union all select "Birth_Weight"
	union all select "Average_Daily_Gain"	
	union all select "Heifers_Calving_Early"
	union all select  "Heifers_Calving_Within_21_Days"
	union all select  "Heifers_Calving_Within_42_Days"
	union all select  "Mature_Cows_Calving_Within_21_Days"
	union all select  "Mature_Cows_Calving_Within_42_Days"
	union all select "Avg_Cow_Age"
	union all select "Replacements_Kept_To_Calve"
	union all select "Cow_Weight_At_Weaning"

	union all select "Cow_Condition_Score_At_Weaning"
	union all select "Adjusted_205_Day_Wt"
	union all select "Actual_Weaning_Wt_Steers"
	union all select "Actual_Weaning_Wt_Heifers"
	union all select "Actual_Weaning_Wt_Bulls"
	union all select "Frame_Score"
	
)b;
END //
/*--------------SPA_Summary_Reprodcution_Performance-----------*/
DROP PROCEDURE IF EXISTS SPA_Summary_Reprodcution_Performance //
CREATE PROCEDURE SPA_Summary_Reprodcution_Performance(Birth_Start_Date date,Birth_End_Date date,Bull_Turn_Out_Date date)
BEGIN 
SELECT 
CASE  b.Critical_Succes_Factors 	
	WHEN "Pregnancy_Percentage" THEN 'Pregnancy Percentage'	
	WHEN "Pregnancy_Loss_Percentage" THEN 'Pregnancy Loss Percentage'
	WHEN "Calving_Percentage" THEN 'Calving Percentage'
	WHEN "Calf_Death_Loss" THEN 'Calf Death Loss'
	WHEN "Weaning_Percentage" THEN 'Calf Crop or Weaning Percentage'
	WHEN "Female_Replacement_Rate_Percentage" THEN 'Female Replacement Rate Percentage'
	WHEN "Calf_Death_Loss_Based_On_Calves_Born" THEN 'Calf Death Loss Based on # of Calves Born'
	WHEN "Pounds_Weaned_Per_Exposed_Female" THEN 'Pounds_Weaned_Per_Exposed_Female'
END as Critical_Success_Factors,
CASE  b.Critical_Succes_Factors 
	WHEN "Pregnancy_Percentage" THEN CONCAT(FORMAT(Pregnancy_Percentage* 100 ,2),' %')	
	WHEN "Pregnancy_Loss_Percentage" THEN CONCAT(FORMAT(Pregnancy_Loss_Percentage* 100 ,2),' %')
	WHEN "Calving_Percentage" THEN CONCAT(FORMAT(Calving_Percentage* 100 ,2),' %')
	WHEN "Calf_Death_Loss" THEN CONCAT(FORMAT(Calf_Death_Loss* 100 ,2),' %')
	WHEN "Weaning_Percentage" THEN CONCAT(FORMAT(Weaning_Percentage* 100 ,2),' %')
	WHEN "Female_Replacement_Rate_Percentage" THEN CONCAT(FORMAT(Female_Replacement_Rate_Percentage*100,2),'%')
	WHEN "Calf_Death_Loss_Based_On_Calves_Born" THEN CONCAT(FORMAT(Calf_Death_Loss_Based_On_Calves_Born*100,2),'%')
	WHEN "Pounds_Weaned_Per_Exposed_Female" THEN CONCAT(Pounds_Weaned_Per_Exposed_Female, ' lbs')
END as Your_Herd_Performance
FROM(
	SELECT 
	(@total_cows_kept_for_calving-@cows_open)/@denom as Pregnancy_Percentage,
	(@cows_aborted/(@total_cows_kept_for_calving+@cows_aborted)) as Pregnancy_Loss_Percentage,	
	(@total_cows_kept_for_calving-@cows_aborted-@cows_open)/@denom as Calving_Percentage,
	@cows_losing_calves/@denom as Calf_Death_Loss,
	(@cows_weaning_calves+@x-@f)/@denom as Weaning_Percentage,
	(@cows_losing_calves/@total_cows_calving) as Calf_Death_Loss_Based_On_Calves_Born,
	@total_wt/@denom as Pounds_Weaned_Per_Exposed_Female,
	@rep_calv as Female_Replacement_Rate_Percentage
	FROM(
		SELECT @total_cows_kept_for_calving:= (select SUM(CASE WHEN weaning_tbl.manage_code='T' Then 0.5 ELSE 1 END)
from cattle_info_tbl INNER JOIN weaning_tbl ON cattle_info_tbl.chaps_id=weaning_tbl.chaps_id  where cattle_info_tbl.birth_date >= Birth_Start_Date AND  cattle_info_tbl.birth_date <=Birth_End_Date),
		@total_cows_exposed:=@total_cows_kept_for_calving+(select count(*) from owners_tbl where exit_date between Bull_Turn_Out_Date AND DATE_ADD(Bull_Turn_Out_Date, INTERVAL 365 DAY ) ),
		@cows_aborted:=(select count(*) as 'Cows aborted' from weaning_tbl where entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date and manage_code ='B'),
		@cows_open:=(select count(*) as 'Cows open' from weaning_tbl where entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date and manage_code ='A'),
		@total_cows_calving:=@total_cows_kept_for_calving-@cows_aborted,
		@cows_losing_calves:=(select count(*) as 'Cows losing calves' from weaning_tbl inner join cattle_info_tbl on weaning_tbl.chaps_id=cattle_info_tbl.chaps_id where cattle_info_tbl.birth_date >= Birth_Start_Date and  cattle_info_tbl.birth_date<=Birth_End_Date  and weaning_tbl.manage_code in ('C','D','F','K')),
		@h:=(select count(*) as 'Cows_Sold_Because_Of_Age' from owners_tbl where exit_date between Bull_Turn_Out_Date AND DATE_ADD(Bull_Turn_Out_Date, INTERVAL 365 DAY ) and exit_code ='H'),
		@j:=(select count(*) as 'Cows_Sold_Because_Of_Physical_Defects' from owners_tbl where  exit_date between Bull_Turn_Out_Date AND DATE_ADD(Bull_Turn_Out_Date, INTERVAL 366 DAY ) and exit_code ='J'),
		@l:=(select count(*) as 'Cows_Sold_Because_Of_Inferior_Calves' from owners_tbl where exit_date between Bull_Turn_Out_Date AND DATE_ADD(Bull_Turn_Out_Date, INTERVAL 365 DAY ) and exit_code ='L'),
		@r:=(select count(*) as 'Cows_Sold_For_Replacement_Stock' from owners_tbl where exit_date between Bull_Turn_Out_Date AND DATE_ADD(Bull_Turn_Out_Date, INTERVAL 365 DAY ) and exit_code ='R'),
		@y:=(select count(*) as 'Cows_Sold_For_Unknown_Reason' from owners_tbl where exit_date between Bull_Turn_Out_Date AND DATE_ADD(Bull_Turn_Out_Date, INTERVAL 365 DAY ) and exit_code ='Y'),
		@cows_weaning_calves:=(select count(*) from measurement_tbl where weight>0 AND entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date),
		@f:=(select count(*) as 'foster or purchased calves'  from weaning_tbl where entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date and manage_code ='F'),
		@x:= (select count(*) as 'incomplete record in weaning tbl'  from weaning_tbl where entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date and manage_code ='X'),
		@denom:=@total_cows_exposed-@h-@j-@l-@r-@y,
		@total_wt:= (select sum(weight) from measurement_tbl where weight>0 AND entry_date >= Birth_Start_Date and entry_date <= Birth_End_Date),
		#@rep_calv:=(select count(*) from cattle_info_tbl where cow_age<3 and birth_date >= Birth_Start_Date AND birth_date <=Birth_End_Date)/(select count(*) from cattle_info_tbl WHERE birth_date >= Birth_Start_Date AND birth_date <=Birth_End_Date)
		@rep_calv:=(select count(*) from cattle_info_tbl where cow_age<3 and birth_date >= Birth_Start_Date AND birth_date <=Birth_End_Date)/@denom
		
	)x
)a
cross join
(
	select "Pregnancy_Percentage" as Critical_Succes_Factors
	union all select "Pregnancy_Loss_Percentage"
	union all select "Calving_Percentage"
	union all select "Calf_Death_Loss"
	union all select "Weaning_Percentage"	
	union all select "Female_Replacement_Rate_Percentage"
	union all select  "Calf_Death_Loss_Based_On_Calves_Born"
	union all select "Pounds_Weaned_Per_Exposed_Female"
)b;

END //

/*--------------SPA_Summary_Calving_Distribution-----------*/
DROP PROCEDURE IF EXISTS SPA_Summary_Calving_Distribution //
CREATE PROCEDURE SPA_Summary_Calving_Distribution(Birth_Start_Date date,Birth_End_Date date,Bull_Turnout_Date date) 
BEGIN
DECLARE est_turn_date DATE;
set est_turn_date=Estimated_Bull_Turnout_Date(Birth_Start_Date, Birth_End_Date, Bull_Turnout_Date);
SELECT DISTINCT
CASE  b.Critical_Succes_Factors 	
	WHEN 'Calves_Born_During_First_21_Days' THEN 'Calves_Born_During_First_21_Days'	
	WHEN 'Calves_Born_During_First_42_Days'	THEN 'Calves_Born_During_First_42_Days'
	WHEN 'Calves_Born_During_First_63_Days' THEN 'Calves_Born_During_First_63_Days'
	WHEN 'Calves_Born_After_First_63_Days'	THEN 'Calves_Born_After_First_63_Days'
	WHEN 'Avg_Age_at_Weaning'	THEN 'Avg_Age_at_Weaning'
	WHEN 'Actual_Weaning_Wts_Steers' THEN 'Actual_Weaning_Wts_Steers'
	WHEN 'Actual_Weaning_Wts_Heifers'  THEN 'Actual_Weaning_Wts_Heifers'
	WHEN 'Actual_Weaning_Wts_Bulls' THEN 'Actual_Weaning_Wts_Bulls'
	WHEN 'Avg_Weaning_Wts' THEN 'Avg_Weaning_Wts'
	WHEN 'Pounds_Weaned_Per_Exposed_Female' THEN 'Pounds_Weaned_Per_Exposed_Female'
END as Critical_Success_Factors,
CASE  b.Critical_Succes_Factors 
		WHEN 'Calves_Born_During_First_21_Days' THEN CONCAT(FORMAT(Calves_Born_During_First_21_Days * 100 ,2),' %')
		WHEN 'Calves_Born_During_First_42_Days' THEN CONCAT(FORMAT(Calves_Born_During_First_42_Days * 100 ,2),' %')
		WHEN 'Calves_Born_During_First_63_Days' THEN CONCAT(FORMAT(Calves_Born_During_First_63_Days * 100 ,2),' %')
		WHEN 'Calves_Born_After_First_63_Days' THEN CONCAT(FORMAT(Calves_Born_After_First_63_Days * 100 ,2),' %')
		WHEN 'Avg_Age_at_Weaning' THEN Round(Avg_Age_at_Weaning,1)
		WHEN 'Actual_Weaning_Wts_Heifers' THEN Round(Actual_Weaning_Wts_Heifers,1)
		WHEN 'Actual_Weaning_Wts_Steers' THEN Round(Actual_Weaning_Wts_Steers,1)
		WHEN 'Actual_Weaning_Wts_Bulls' THEN Round(Actual_Weaning_Wts_Bulls,1)
		WHEN 'Avg_Weaning_Wts' THEN Round(Avg_Weaning_Wts,1)
		WHEN 'Pounds_Weaned_Per_Exposed_Female' THEN 'Pounds_Weaned_Per_Exposed_Female'
END as Your_Herd_Performance
FROM(
	SELECT  @within_21days/(@within_63days+@after_63days) AS Calves_Born_During_First_21_Days,
			@within_42days/(@within_63days+@after_63days) AS Calves_Born_During_First_42_Days,
			@within_63days/(@within_63days+@after_63days) AS Calves_Born_During_First_63_Days,
			@after_63days/(@within_63days+@after_63days) AS Calves_Born_After_First_63_Days,
			Avg_Age_at_Weaning, Actual_Weaning_Wts_Steers, Actual_Weaning_Wts_Heifers, Actual_Weaning_Wts_Bulls, Avg_Weaning_Wts
	FROM(
		SELECT
		@early_calves:=Early_Calves(Birth_Start_Date, Birth_End_Date,est_turn_date,4,null)AS Early_Calves,
		@within_21days:=Calving_Within_21Days(Birth_Start_Date, Birth_End_Date,est_turn_date,4,null),
		@within_42days:=Calving_Within_42Days(Birth_Start_Date, Birth_End_Date,est_turn_date,4,null) ,
		@within_63days:=Calving_Within_63Days(Birth_Start_Date, Birth_End_Date,est_turn_date,4,null) ,
		@after_63days:=Calving_After_63Days(Birth_Start_Date, Birth_End_Date,est_turn_date,4,null) ,
		Average_Calf_Age(Birth_Start_Date, Birth_End_Date) as Avg_Age_at_Weaning,
		COALESCE(Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,3),0) as Actual_Weaning_Wts_Steers,
		COALESCE(Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,2),0) as Actual_Weaning_Wts_Heifers,
		COALESCE(Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,1),0) as Actual_Weaning_Wts_Bulls,
		COALESCE(Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,4),0) as Avg_Weaning_Wts
		FROM cattle_info_tbl INNER JOIN measurement_tbl ON (cattle_info_tbl.chaps_id = measurement_tbl.chaps_id)
			INNER JOIN weaning_tbl ON weaning_tbl.chaps_id=cattle_info_tbl.chaps_id
			INNER JOIN owners_tbl  ON owners_tbl.chaps_id=cattle_info_tbl.chaps_id
			where cattle_info_tbl.birth_date >=Birth_Start_Date
			AND cattle_info_tbl.birth_date <=     Birth_End_Date 
			AND measurement_tbl.entry_date<>'0000-00-00'
	)x
)a
cross join
(
	select 'Calves_Born_During_First_21_Days' as Critical_Succes_Factors
	union all select 'Calves_Born_During_First_42_Days'
	union all select 'Calves_Born_During_First_63_Days'
	union all select 'Calves_Born_After_First_63_Days'
	union all select 'Avg_Age_at_Weaning'	
	union all select 'Actual_Weaning_Wts_Steers'
	union all select 'Actual_Weaning_Wts_Heifers'
	union all select 'Actual_Weaning_Wts_Bulls'
	union all select  'Avg_Weaning_Wts'
	#union all select 'Pounds_Weaned_Per_Exposed_Female'
)b;
END //

/*--------------SPA_Summary_Production_Performance_Measures-----------*/
DROP PROCEDURE IF EXISTS SPA_Summary_Production_Performance_Measures //
CREATE PROCEDURE SPA_Summary_Production_Performance_Measures(Birth_Start_Date date,Birth_End_Date date,Bull_Turn_Out_Date date)
BEGIN 
SELECT 
CASE  b.Critical_Succes_Factors 	
	WHEN  "Avg Age at Weaning" THEN 'Avg Age at Weaning'
	WHEN "Actual Wean Weight- Steers" THEN 'Actual Wean Weight- Steers'
	WHEN "Actual Wean Weight- Heifers" THEN 'Actual Wean Weight- Heifers'
	WHEN "Actual Wean Weight- Bulls" THEN 'Actual Wean Weight- Bulls'
END as Critical_Success_Factors,
CASE  b.Critical_Succes_Factors 
	WHEN  "Avg Age at Weaning" THEN Average_Age_At_Weaing
	WHEN "Actual Wean Weight- Steers" THEN Actual_Weaning_Wt_Steers
	WHEN "Actual Wean Weight- Heifers" THEN Actual_Weaning_Wt_Heifers
	WHEN "Actual Wean Weight- Bulls" THEN Actual_Weaning_Wt_Bulls
END as Your_Herd_Performance
FROM(
	SELECT Average_Age_At_Weaing,Actual_Weaning_Wt_Heifers,Actual_Weaning_Wt_Steers,Actual_Weaning_Wt_Bulls

	FROM(
		SELECT @avg_age_at_weaning:= Average_Calf_Age(Birth_Start_Date,Birth_End_Date ) as Average_Age_At_Weaing,
		Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,2) as Actual_Weaning_Wt_Heifers,
		Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,3) as Actual_Weaning_Wt_Steers,
		Actual_Wean_Weight(Birth_Start_Date, Birth_End_Date,1) as Actual_Weaning_Wt_Bulls
		
	)x
)a
cross join
(
	select "Avg Age at Weaning" as Critical_Succes_Factors
	union all select "Actual Wean Weight- Steers"
	union all select "Actual Wean Weight- Heifers"
	union all select "Actual Wean Weight- Bulls"
	#union all select "Pounds Weaned Per Exposed Female"
)b;

END //
