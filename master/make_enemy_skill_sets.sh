#! /bin/sh

FILE_NAME="enemy_skill_sets.csv"
rm ${FILE_NAME}
echo "skill_set_id,skill_id,skill_trigger_id" > ${FILE_NAME}
cat enemy_skill_sets_*.tmp >> ${FILE_NAME}
