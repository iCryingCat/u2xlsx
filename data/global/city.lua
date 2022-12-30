--[[ E:/code/c#/xlsx-exporter/xlsx/C1-测试.xlsx ]]--
-- 测试lua数据表
local ICITY = {
-- 索引
xid = 0,
-- 城市名
cityName = 1,
-- 省份
province = 2,
-- 特产
food = 3,
}
local XLSX_GLOBAL_CITY = {
[t1] = {
cityName = '广州',
province = '广东',
food = {
name = '煲仔饭',
type = '主食',
},
},
[t2] = {
cityName = '南昌',
province = '江西',
food = {
name = '脐橙',
type = '水果',
},
},
}

return XLSX_GLOBAL_CITY