ALTER TABLE cattle_info_tbl ADD COLUMN lot_no varchar(16);
ALTER TABLE measurement_tbl ADD COLUMN lot_no varchar(16);
ALTER TABLE weaning_tbl DROP COLUMN lot_no;
ALTER TABLE replacement_tbl ADD COLUMN scrotum_date DATE;
ALTER TABLE replacement_tbl ADD COLUMN pelvic_date DATE;