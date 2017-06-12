#! /bin/sh

FILE_NAME="gifts.csv"
rm ${FILE_NAME}
echo "gift_id,gift_type,gift_value_aapl,gift_value_goog,gift_value_amzn,gift_count_aapl,gift_count_goog,gift_count_amzn" > ${FILE_NAME}
cat gifts_tmp_*.tmp >> ${FILE_NAME}
