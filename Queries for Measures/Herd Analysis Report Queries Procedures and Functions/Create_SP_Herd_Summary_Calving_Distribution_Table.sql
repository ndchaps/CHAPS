Delimiter //
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

CALL Herd_Summary_Calving_Distribution_Table('2014-1-1','2014-12-31','2013-08-01')//
#CALL Herd_Summary_Calving_Distribution_Table('2014-1-1','2014-12-31','2013-08-01')

