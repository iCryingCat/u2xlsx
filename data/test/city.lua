--[[E:/code/c#/xlsx-exporter/assets/测试@test.xlsx]]--
-- 测试lua数据表
local ICITY = {
-- 索引
xid = 0,
-- 城市名
cityName = 1,
-- gdp
gdp = 2,
}
local XLSX_TEST_CITY = {
[1] = {
cityName = '广州',
gdp = 1000,
},
[2] = {
cityName = '上海',
gdp = 1200.3,
},
}

return XLSX_TEST_CITY