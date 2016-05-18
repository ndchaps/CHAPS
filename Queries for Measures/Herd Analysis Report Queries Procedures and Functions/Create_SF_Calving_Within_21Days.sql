DELIMITER \\
DROP FUNCTION Calving_Within_21Days \\
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
END \\
SELECT Calving_Within_21Days('2014-1-1','2014-12-31','2013-07-19',4,null)