use chaps1_temporary;
DELIMITER //
DROP PROCEDURE IF EXISTS Select_Calf_Background_Info //
CREATE PROCEDURE Select_Calf_Background_Info()
BEGIN 
	Select * from(
			select bt.chaps_id,bt.entry_date,mt.lot_no,mt.weight,mt.hip_height,mt.frame_score,mt.`status`,mt.entry_type,nt.note_type,nt.notes
			from background_tbl bt left join measurement_tbl mt on bt.chaps_id=mt.chaps_id and bt.entry_date=mt.entry_date
			left join notes_tbl nt on bt.chaps_id=nt.chaps_id
	)A;
END//