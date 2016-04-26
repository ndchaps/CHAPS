DELIMITER //
DROP PROCEDURE IF EXISTS Select_Calf_Birth_Info //
CREATE PROCEDURE Select_Calf_Birth_Info()
BEGIN 
	Select * from(
			select cit.herd_id, cit.chaps_id, cit.birth_date, cit.birth_weight, cit.dam_id , cit.cow_age, cit.breed, cit.reg_no, cit.reg_name,
					cit.elec_id, cit.sire_id, cit.sex, cit.calving_ease, cit.state, cit.sex_date, cit.lot_no,nt.note_type,nt.notes
			from cattle_info_tbl cit left join notes_tbl nt on cit.chaps_id=nt.chaps_id
	)A;
END//