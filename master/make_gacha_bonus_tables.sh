#! /bin/sh

FILE_NAME="gacha_bonus_tables.csv"
rm ${FILE_NAME}
echo "bonus_table_no,gift_id,weight" > ${FILE_NAME}
cat gacha_bonus_tables_*.tmp >> ${FILE_NAME}
