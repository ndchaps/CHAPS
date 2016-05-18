DELIMITER \\
DROP FUNCTION Early_Calves \\
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

END \\
SELECT Early_Calves('2014-1-1','2014-12-31','2013-07-19',3,false)