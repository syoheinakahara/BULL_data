#! /bin/sh

rm floors.csv
echo "quest_id,stage_id,floor_id,total_size_min,total_size_max" > floors.csv
cat floors*.tmp >> floors.csv

rm enemy_habitats.csv
echo "quest_id,stage_id,floor_id,character_id,level,hp_coefficient,attack_coefficient,defense_correction,skill_set_id,fix_count,incidence,boss_flg,speed,drop_unit_id,drop_unit_level,drop_unit_rare,drop_unit_rate,hp_plus,attack_plus,defense_plus,heal_plus,bug,seq_no" > enemy_habitats.csv
cat enemy_habitats*.tmp >> enemy_habitats.csv
