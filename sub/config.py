from openpyxl.utils import column_index_from_string

URL_COL      = column_index_from_string('AV')  # 48
DEST_COL     = URL_COL + 1                      # 49 = AW
GRANTED_COL  = column_index_from_string('K')   # 11 - Granted Number
PUB_COL      = column_index_from_string('M')   # 13 - Unexamined Pub. Number
JUSTIA_BASE  = "https://patents.justia.com/patent/"
MAX_CHARS    = 32767
ROW_HEIGHT   = 15
