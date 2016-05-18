SELECT (select count(*) from cattle_info_tbl where birth_date >= '2014-1-1' and birth_date <= '2014-12-31')
+
(select count(*) from owners_tbl where exit_date between '2013-8-1' AND DATE_ADD('2013-8-1', INTERVAL 365 DAY ) ),

/*Total Cows kept for Calving*/
(select count(*) as 'Cows kept for calving' from cattle_info_tbl where birth_date >= '2014-1-1' and birth_date <= '2014-12-31' ),

/*Number of cows aborted*/
(select count(*) as 'Cows aborted' from weaning_tbl 
where entry_date >= '2014-1-1' and entry_date <= '2014-12-31' and manage_code ='B'),

/*Number of cows open*/
(select count(*) as 'Cows open' from weaning_tbl 
where entry_date >= '2014-1-1' and entry_date <= '2014-12-31' and manage_code ='A'),


/*Number of cows losing calves*/
(select count(*) as 'Cows losing calves' from weaning_tbl 
where entry_date >= '2014-1-1' and entry_date <= '2014-12-31' and manage_code in ('C','D','F','K'));

(select count(*) as 'Cows losing calves' from weaning_tbl 
where entry_date >= '2014-1-1' and entry_date <= '2014-12-31' and manage_code not in ('A','B','C','D','F','K'));

# Calcing_Distribution_Table_Query

select a.Dam_Age, a.Calves_each_age,COALESCE(b.Early_Calves,0) as Early_Calves,avg_date_each_age,avg_actual_wean_weight from
(SELECT cattle_info_tbl.cow_age as Dam_Age,COUNT(*) AS Calves_each_age,
CAST(from_unixtime(AVG(UNIX_TIMESTAMP(cattle_info_tbl.birth_date))) AS DATE) AS avg_date_each_age,
AVG(measurement_tbl.weight )as avg_actual_wean_weight
FROM cattle_info_tbl INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
INNER JOIN measurement_tbl on measurement_tbl.chaps_id=cattle_info_tbl.chaps_id
WHERE cattle_info_tbl.sex in ('0','1','2','3')
AND cattle_info_tbl.birth_date >= '2014-1-1'
AND  cattle_info_tbl.birth_date <= '2014-12-31'
AND weaning_tbl.manage_code <>'A'
AND weaning_tbl.manage_code <>'B'
AND measurement_tbl.weight>0
GROUP BY cattle_info_tbl.cow_age)a
Left JOIN
(SELECT cattle_info_tbl.cow_age as Dam_Age, COUNT(*) AS Early_Calves FROM cattle_info_tbl
 INNER JOIN weaning_tbl on cattle_info_tbl.chaps_id = weaning_tbl.chaps_id
WHERE cattle_info_tbl.sex in ('0','1','2','3')
AND cattle_info_tbl.birth_date >= '2014-1-1'
AND  cattle_info_tbl.birth_date <= '2014-12-31'
AND weaning_tbl.manage_code <>'A'
AND weaning_tbl.manage_code <>'B'
AND cattle_info_tbl.birth_date < DATE_ADD('2013-07-19', INTERVAL 285 DAY )
GROUP BY cattle_info_tbl.cow_age
)b
on a.Dam_Age=b.Dam_Age
