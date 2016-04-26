DELIMITER //
DROP PROCEDURE IF EXISTS Select_Calf_Feedlot_Info //
CREATE PROCEDURE Select_Calf_Feedlot_Info()
BEGIN 
	Select * from(
			select ft.chaps_id, ft.entry_date,ft.back_fat,ft.rib_eye,ft.marbling,
			mt.lot_no,mt.weight,mt.hip_height,mt.frame_score,mt.`status`,
			mt.entry_type,nt.note_type,nt.notes
			from feedlot_tbl ft	left join measurement_tbl mt on ft.chaps_id=mt.chaps_id and ft.entry_date=mt.entry_date
			left join notes_tbl nt on ft.chaps_id=nt.chaps_id 
	)A;
END//