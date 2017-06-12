#! /bin/sh

FILE_NAME="enemy_skills_i18n.csv"
rm ${FILE_NAME}
echo "enemy_skill_id,language,name,comment" > ${FILE_NAME}
cat enemy_skills_i18n_*.tmp >> ${FILE_NAME}
