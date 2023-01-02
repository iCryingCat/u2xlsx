-- 测试.xlsx
--[[
local ICITY = {
-- 索引
id = 0,
-- 城市名
cityName = 1,
-- 省份
province = 2,
-- 特产
food = 3,
-- 行政区
regions = 4,
}
--]]

-- 城市表
local XLSX_GLOBAL_CITY = {
    ['t1'] = {
        cityName = '广州',
        province = '广东',
        food = {
            [1] = {
                name = '煲仔饭',
                type = '主食',
            },
            [2] = {
                name = '脐橙',
                type = '水果',
            },
        },
        regions = {
            [0] = '天河区',
            [1] = '海珠区',
        },
    },
    ['t2'] = {
        cityName = '南昌',
        province = '江西',
        food = {
            [2] = {
                name = '脐橙',
                type = '水果',
            },
        },
        regions = {
            [0] = '西湖区',
            [1] = '新建区',
        },
    },
}

return XLSX_GLOBAL_CITY
