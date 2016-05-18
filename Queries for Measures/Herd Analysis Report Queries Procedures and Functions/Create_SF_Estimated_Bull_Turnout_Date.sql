DELIMITER //
DROP FUNCTION Estimated_Bull_Turnout_Date //
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

select Estimated_Bull_Turnout_Date('2013-1-1','2013-12-31','2012-8-1')






