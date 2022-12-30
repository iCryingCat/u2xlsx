local XLSX_GLOBAL_FOOD = require "global.food"

print(XLSX_GLOBAL_FOOD[1].city)
XLSX_GLOBAL_FOOD[1].city = '成都'
print(XLSX_GLOBAL_FOOD[1].city)
