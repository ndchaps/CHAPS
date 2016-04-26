use chaps1_temporary;
DELIMITER //
DROP PROCEDURE Select_Calf_Birth_Info //
CREATE PROCEDURE Select_Calf_Birth_Info()
BEGIN 
	Select * from(
			select cit.herd_id, cit.chaps_id, cit.birth_date, cit.birth_weight, cit.dam_id , cit.cow_age, cit.breed, cit.reg_no, cit.reg_name,
					cit.elec_id, cit.sire_id, cit.sex, cit.calving_ease, cit.state, cit.sex_date, cit.lot_no,
					wt.entry_date,wt.manage_code,wt.contemp_grp,wt.muscle_grade,mt.lot_no,mt.weight,mt.hip_height,mt.frame_score,mt.`status`,mt.entry_type,
					bt.entry_date,rt.entry_date,rt.con,rt.back_fat,rt.rib_eye,rt.marbling,rt.`365_day_weight`,rt.scrotum_circm,scrotum_date,rt.pelvic_area,rt.pelvic_date,
					ft.entry_date,ft.back_fat,ft.rib_eye,ft.marbling,ct.carcass_date,ct.hot_carcass_wt,ct.fat_thickness,ct.kidney_kph,ct.rib_eye,ct.quality_grade,ct.marbling_score,
					ct.color,ct.texture_of_lean,ct.maturity,ct.conformance,ct.muscle_score,nt.note_type,nt.notes
			from cattle_info_tbl cit
			left join weaning_tbl wt on cit.chaps_id=wt.chaps_id
			left join measurement_tbl mt on wt.chaps_id=mt.chaps_id
			left join background_tbl bt on cit.chaps_id=bt.chaps_id
			left join replacement_tbl rt on cit.chaps_id=rt.chaps_id
			left join feedlot_tbl ft on cit.chaps_id=ft.chaps_id
			left join carcass_tbl ct on cit.chaps_id=ct.chaps_id
			left join notes_tbl nt on cit.chaps_id=nt.chaps_id
	)A;
END//