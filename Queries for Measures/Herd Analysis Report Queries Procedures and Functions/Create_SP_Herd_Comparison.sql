DELIMITER \\
DROP PROCEDURE IF EXISTS Herd_Comparison \\
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
END \\

call Herd_Comparison('2013-1-1','2013-12-31','2012-08-01')\\

call Herd_Comparison('2012-1-1','2012-12-31','2011-08-08')
