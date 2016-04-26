use chaps1_temporary;
DELIMITER //
DROP PROCEDURE Insert_CalfBirthInfo_byPK//
CREATE DEFINER=`root`@`localhost` PROCEDURE `Insert_CalfBirthInfo_byPK`(birth_date date,breed varchar(16),reg_no varchar(30),reg_name varchar(30),elec_id varchar(30),sire_id bigint(20),dam_id bigint(20),sex varchar(1),sex_date date,birth_weight float,calving_ease tinyint(4),herd_id varchar(30),cow_age tinyint(4),birth_notes text)
BEGIN
DECLARE my_chaps_id INT;
INSERT into cattle_info_tbl(birth_date,breed,reg_no,reg_name,elec_id,sire_id,dam_id,sex,sex_date,birth_weight,calving_ease,herd_id,cow_age) 
values(birth_date,breed,reg_no,reg_name,elec_id,sire_id,dam_id,sex,sex_date,birth_weight,calving_ease,herd_id,cow_age);
INSERT into notes_tbl(chaps_id,entry_date,note_type,notes) values(LAST_INSERT_ID(),CURDATE(),'b',birth_notes);
END//

#Call Insert_CalfBirthInfo_byPK('1990-03-06','AN','dummy',null,null,20177,0,2,'0000-01-01',53,1,'H38',2,'first birth notes');


