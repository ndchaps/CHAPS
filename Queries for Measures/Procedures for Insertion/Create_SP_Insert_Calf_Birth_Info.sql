use chaps1_unit_testing;
DELIMITER $$
DROP PROCEDURE IF EXISTS Insert_CalfBirthInfo_byPK $$
CREATE PROCEDURE Insert_CalfBirthInfo_byPK(Calf_ID VARCHAR(16),birth_date date,breed varchar(16),reg_no varchar(30),
					reg_name varchar(30),elec_id varchar(30),sire_id varchar(20),dam_id varchar(20),sex varchar(1),
					sex_date date,birth_weight float,
					calving_ease tinyint(4),herd_id varchar(30),birth_notes text)
BEGIN
#insert basic calf_information in cattle_info_table
INSERT into cattle_info_tbl(herd_id,animal_id,birth_date,breed,reg_no,reg_name,elec_id,sire_id,dam_id,sex,sex_date,birth_weight,calving_ease) 
            values         (herd_id,Calf_ID,birth_date,breed,reg_no,reg_name,elec_id,sire_id,dam_id,sex,sex_date,birth_weight,calving_ease);
SET @my_chaps_id= LAST_INSERT_ID();
#update this calf's mother's chaps id in the table based on the mother's calf id
Update cattle_info_tbl a join cattle_info_tbl b on a.dam_id=b.animal_id  set a.dam_chaps_id=b.chaps_id where a.chaps_id=LAST_INSERT_ID();
#update this calf's father's chaps id in the table based on the father's calf id
Update cattle_info_tbl a join cattle_info_tbl b on a.sire_id=b.animal_id  set a.sire_chaps_id=b.chaps_id where a.chaps_id=LAST_INSERT_ID();
#Update the cow_age for this calf entry
Update cattle_info_tbl a join cattle_info_tbl b on a.dam_id=b.animal_id  set a.cow_age=(Curdate()-b.birth_date) where a.chaps_id=LAST_INSERT_ID();
#insert the newly generated chaps id for this calf in cattle names table
insert into cattle_names_tbl(chaps_id,entry_date,cattle_name,cattle_type) values(LAST_INSERT_ID(),CURDATE(),calf_id,'CA');
#INSERT THE CHAPS_ID, BIRTH DATE as entry date and herd ID into ownser_table
insert into owners_tbl(chaps_id,start_date,herd_id) values(LAST_INSERT_ID(),birth_date,herd_id);
#insert calf's birth notes to notes table
INSERT into notes_tbl(chaps_id,animal_id,entry_date,note_type,notes) values(LAST_INSERT_ID(),calf_id,CURDATE(),'b',birth_notes);
END$$
DELIMITER ;