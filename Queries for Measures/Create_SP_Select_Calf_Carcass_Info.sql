DELIMITER //
DROP PROCEDURE IF EXISTS Select_Calf_Carcass_Info //
CREATE PROCEDURE Select_Calf_Carcass_Info()
BEGIN 
	Select * from(
			select ct.chaps_id,ct.carcass_date,ct.fat_thickness,ct.kidney_kph,ct.rib_eye,ct.quality_grade,ct.marbling_score,
					ct.color,ct.texture_of_lean,ct.maturity,ct.conformance,ct.muscle_score,mt.lot_no,mt.weight,mt.hip_height,mt.frame_score,mt.`status`,
					mt.entry_type,nt.note_type,nt.notes					
			from carcass_tbl ct inner join	measurement_tbl mt on ct.chaps_id=mt.chaps_id and ct.carcass_date=mt.entry_date
			left join notes_tbl nt on ct.chaps_id=nt.chaps_id
	)A;
END//