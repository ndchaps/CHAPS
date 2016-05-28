DELIMITER //
DROP PROCEDURE IF EXISTS Select_Calf_Birth_Info //
CREATE PROCEDURE Select_Calf_Birth_Info()
BEGIN 
	Select * from(
		select distinct cit.herd_id,cit.animal_id, cit.birth_date, cit.birth_weight, cit.dam_id , cit.cow_age, cit.breed, cit.reg_no, cit.reg_name,
			cit.elec_id, cit.sire_id, cit.sex, cit.calving_ease, cit.state, cit.sex_date, cit.lot_no,cit.picture,nt.notes
		from cattle_info_tbl cit inner join cattle_names_tbl cnt on cit.chaps_id=cnt.chaps_id 
		left join notes_tbl nt on cit.chaps_id=nt.chaps_id 
		where cnt.cattle_type='CA'
	)A;
END//



			