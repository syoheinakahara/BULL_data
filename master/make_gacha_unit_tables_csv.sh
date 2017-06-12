#! /bin/sh

rm gacha_unit_tables.csv
echo "gacha_id,table_no,character_id,unit_level,weight,hp_plus,attack_plus,defense_plus,heal_plus,bug" > gacha_unit_tables.csv
cat gacha_unit_tables*.tmp >> gacha_unit_tables.csv
