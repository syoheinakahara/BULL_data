#! /bin/sh

FILE_NAME="enemy_habitat_states.csv"
rm ${FILE_NAME}
echo "quest_id,stage_id,floor_id,seq_no,state_no,character_id,level,speed,skill_set_id,hp_coefficient,attack_coefficient,defense_correction" > ${FILE_NAME}
cat enemy_habitat_S_*.tmp >> ${FILE_NAME}
