DELIMITER //
DROP PROCEDURE IF EXISTS Select_Calf_Wean_Info //
CREATE PROCEDURE Select_Calf_Wean_Info()
BEGIN 
	Select * from(
			select  wt.calf_id, wt.entry_date,wt.manage_code,wt.contemp_grp,wt.muscle_grade,mt.lot_no,mt.weight,mt.hip_height,mt.frame_score,mt.`status`,
			mt.entry_type,nt.note_type,nt.notes
			from weaning_tbl wt inner join measurement_tbl mt on wt.chaps_id=mt.chaps_id and wt.entry_date=mt.entry_date
			left join notes_tbl nt on wt.chaps_id=nt.chaps_id
	)A;
END//