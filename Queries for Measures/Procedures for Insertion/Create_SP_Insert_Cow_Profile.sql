DELIMITER //
DROP PROCEDURE IF EXISTS Insert_NewCowProfileInfo //
CREATE PROCEDURE Insert_NewCowProfileInfo(input_herd_id varchar(16),input_cow_id varchar(20),input_calf_id_at_birth varchar(2),input_birth_date date,input_breed VARCHAR(16),
			input_reg_num VARCHAR(30),input_reg_name VARCHAR(30),input_elec_id VARCHAR(30),input_sire_id VARCHAR(20),input_dam_id VARCHAR(20),
			input_date_entered_herd date ,input_animal_source ENUM('S','P'),input_active_flag varchar(3),input_date_culled date,input_reason_culled varchar(2),input_cull_comments VARCHAR(100),input_profile_notes text)		
BEGIN
#------------IF THIS ANIMAL DOESN'T EXIST IN THE SYSTEM THEN ADD A NEW ENTRY IN CATTLE_INFO_TBL--------------#
	IF NOT EXISTS(select 1 from cattle_info_tbl where animal_id=input_calf_id_at_birth) THEN
		#insert basic cowinformation in cattle_info_table
		INSERT into cattle_info_tbl(herd_id,animal_id,birth_date,breed,reg_no,reg_name,elec_id,sire_id,dam_id,sex) 
		values(input_herd_id,input_cow_id,input_birth_date,input_breed,input_reg_num,input_reg_name,input_elec_id,input_sire_id,input_dam_id,'2');
		SET @my_chaps_id= LAST_INSERT_ID();
		#update this cow's mother's chaps id in the table based on the mother's calf id
		Update cattle_info_tbl a join cattle_info_tbl b on a.dam_id=b.animal_id  set a.dam_chaps_id=b.chaps_id where a.chaps_id= @my_chaps_id;
		#update this calf's father's chaps id in the table based on the father's calf id
		Update cattle_info_tbl a join cattle_info_tbl b on a.sire_id=b.animal_id  set a.sire_chaps_id=b.chaps_id where a.chaps_id= @my_chaps_id;
		#Update the cow_age for this calf entry
		Update cattle_info_tbl a join cattle_info_tbl b on a.dam_id=b.animal_id  set a.cow_age=(Curdate()-b.birth_date) where a.chaps_id= @my_chaps_id;
		#insert the newly generated chaps id and the calf name for this cow in cattle names table if the cow is new to the system
		IF NOT ISNULL(input_calf_id_at_birth) AND input_calf_id_at_birth!='' THEN 
			insert into cattle_names_tbl(chaps_id,entry_date,cattle_name,cattle_type) values( @my_chaps_id,CURDATE(),input_calf_id_at_birth,'CA');
		END IF;
		#insert the newly generated chaps id and cow name for this cow in cattle names table if the cow is new to the system
		insert into cattle_names_tbl(chaps_id,entry_date,cattle_name,cattle_type) values( @my_chaps_id,CURDATE(),input_cow_id,'CO');
		#INSERT THE CHAPS_ID, BIRTH DATE as entry date and herd ID into ownser_table
		insert into owners_tbl(chaps_id,start_date,herd_id,active_flag,animal_source) values(LAST_INSERT_ID(),input_date_entered_herd,herd_id,input_active_flag,input_animal_source);
		#insert this cow's profile notes to notes table
		INSERT into notes_tbl(chaps_id,animal_id,entry_date,note_type,notes) values(LAST_INSERT_ID(),input_cow_id,CURDATE(),'cp',input_profile_notes);
		#check if the user enters the cull information, if yes, enter it into the cull_tbl
		IF NOT ISNULL(input_date_culled) THEN
			INSERT into cull_tbl(chaps_id,animal_id,cull_date,cull_code,cull_comments) values(@my_chaps_id,input_cow_id,input_date_culled,input_reason_culled,input_cull_comments);
		END IF;
	ELSE 
		IF NOT ISNULL(input_calf_id_at_birth) THEN
			SET @my_chaps_id= (Select chaps_id from cattle_info_tbl where animal_id=input_calf_id_at_birth);
			UPDATE cattle_info_tbl SET animal_id=input_cow_id where chaps_id=@my_chaps_id;

			insert into cattle_names_tbl(chaps_id,entry_date,cattle_name,cattle_type) values(@my_chaps_id,CURDATE(),cow_id,'CO');

			insert into owners_tbl(chaps_id,start_date,herd_id,active_flag,animal_source) values(LAST_INSERT_ID(),input_date_entered_herd,input_herd_id,input_active_flag,input_animal_source);

			INSERT into notes_tbl(chaps_id,animal_id,entry_date,note_type,notes) values(LAST_INSERT_ID(),input_cow_id,CURDATE(),'cp',profile_notes);
			#check if the user enters the cull information, if yes, enter it into the cull_tbl
			IF NOT ISNULL(input_date_culled) THEN
				INSERT into cull_tbl(chaps_id,animal_id,cull_date,cull_code,cull_comments) values(@my_chaps_id,input_cow_id,input_date_culled,input_reason_culled,input_cull_comments);
			END IF;
		END IF;
	END IF;
END//
