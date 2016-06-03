DELIMITER //
DROP PROCEDURE IF EXISTS Select_Calf_Background_Info //
CREATE PROCEDURE Select_Calf_Background_Info()
BEGIN 
	Select * from(
			select distinct bt.calf_id,mt.entry_date,mt.weight,mt.hip_height,mt.frame_score,mt.`status`,
		nt.note_type,nt.notes,mt.entry_subtype
			from background_tbl bt left join measurement_tbl mt on bt.chaps_id=mt.chaps_id and bt.entry_date=mt.entry_date
			left join notes_tbl nt on bt.chaps_id=nt.chaps_id and nt.note_type=mt.entry_type
			where mt.entry_type='BK'
	)A;
END//


