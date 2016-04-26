DELIMITER //
DROP PROCEDURE IF EXISTS Select_Calf_Replacement_Info //
CREATE PROCEDURE Select_Calf_Replacement_Info()
BEGIN 
	Select * from(
			select  rt.chaps_id,rt.entry_date,rt.con,rt.back_fat,rt.rib_eye,rt.marbling,rt.`365_day_weight`,rt.scrotum_circm,scrotum_date,rt.pelvic_area,rt.pelvic_date,
					mt.lot_no,mt.weight,mt.hip_height,mt.frame_score,mt.`status`,mt.entry_type,nt.note_type,nt.notes
			from replacement_tbl rt	left join measurement_tbl mt on rt.chaps_id=mt.chaps_id and rt.entry_date=mt.entry_date
			left join notes_tbl nt on rt.chaps_id=nt.chaps_id
	)A;
END//