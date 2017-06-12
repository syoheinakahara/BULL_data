#! /bin/sh

FILE_NAME="enemy_skill_triggers.csv"
rm ${FILE_NAME}
echo "skill_trigger_id,first_time,invoke_rate,hp_threshold,hp_termination,invoke_times,invoke_weight,every_time_count,neighbor_dead_id,priority" > ${FILE_NAME}
cat enemy_skill_triggers_*.tmp >> ${FILE_NAME}
