DELIMITER //
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


CALL SPA_Summary_Production_Performance_Measures('2014-1-1','2014-12-31','2013-8-1')