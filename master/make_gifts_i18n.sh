#! /bin/sh

FILE_NAME="gifts_i18n.csv"
rm ${FILE_NAME}
echo "gift_id,language,title,detail,message" > ${FILE_NAME}
cat gifts_i18n_*.tmp >> ${FILE_NAME}
