DELIMITER //
DROP PROCEDURE IF EXISTS Select_Calf_Replacement_Info //
CREATE PROCEDURE Select_Calf_Replacement_Info()
BEGIN 
	Select * from(
			select  rt.calf_id,rt.entry_date,mt.weight,rt.con as 'condition',mt.hip_height,mt.frame_score,rt.back_fat,rt.rib_eye,rt.marbling,rt.`365_day_weight`,
					rt.scrotum_circum,scrotum_date,rt.pelvic_area,rt.pelvic_date,mt.`status`,mt.entry_subtype,nt.note_type,nt.notes
			from replacement_tbl rt	left join measurement_tbl mt on rt.chaps_id=mt.chaps_id and rt.entry_date=mt.entry_date
			left join notes_tbl nt on rt.chaps_id=nt.chaps_id and mt.entry_type=nt.note_type
			where mt.entry_type='R'
	)A;
END//

