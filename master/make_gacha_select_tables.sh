#! /bin/sh

FILE_NAME="gacha_select_tables.csv"
rm ${FILE_NAME}
echo "select_table_no,character_id,unit_level,weight,hp_plus,attack_plus,defense_plus,heal_plus,bug" > ${FILE_NAME}
cat gacha_select_tables_*.tmp >> ${FILE_NAME}
