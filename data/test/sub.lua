--[[ E:/code/c#/xlsx-exporter/assets/测试@test.xlsx ]]--
-- 内嵌表
local ISUB = {
-- 索引
xid = 0,
-- 城市名
cityName = 1,
-- gdp
gdp = 2,
}
local XLSX_TEST_SUB = {
[1] = {
cityName = 'guangzhou',
gdp = 1000,
},
[2] = {
cityName = 'tianjing',
gdp = 1200.3,
},
}

return XLSX_TEST_SUB