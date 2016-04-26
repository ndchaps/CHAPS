DELIMITER //
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
Call Sex_Based_Calves_Born_Weighed_Unweighed_Count('2012-1-1','2012-12-31');