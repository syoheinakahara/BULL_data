#! /bin/sh

FILE_NAME="enemy_skills.csv"
rm ${FILE_NAME}
echo "enemy_skill_id,effect_param,valid_flg" > ${FILE_NAME}
cat enemy_skills_tmp_*.tmp >> ${FILE_NAME}
