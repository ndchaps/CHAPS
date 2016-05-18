DELIMITER //
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

CALL Summary_Of_Cows_Leaving('2014-1-1','2014-12-31','2013-08-01');