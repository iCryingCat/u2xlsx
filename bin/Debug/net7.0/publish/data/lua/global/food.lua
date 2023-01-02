-- 测试.xlsx
--[[
local IFOOD = {
-- 索引
id = 0,
-- 食物名
name = 1,
-- 类别
type = 2,
-- 关联城市
city = 3,
}
--]]

-- 特产表
local XLSX_GLOBAL_FOOD = {
    ['1'] = {
        name = '煲仔饭',
        type = '主食',
        city = {
            [t1] = {
                cityName = '广州',
                province = '广东',
                regions = {
                    [0] = '天河区',
                    [1] = '海珠区',
                },
            },
        },
    },
    ['2'] = {
        name = '脐橙',
        type = '水果',
        city = {
            [t2] = {
                cityName = '南昌',
                province = '江西',
                regions = {
                    [0] = '西湖区',
                    [1] = '新建区',
                },
            },
        },
    },
}

return XLSX_GLOBAL_FOOD
