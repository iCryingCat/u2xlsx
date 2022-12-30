--[[ E:/code/c#/xlsx-exporter/xlsx/测试.xlsx ]] --
-- 内嵌表
local IFOOD = {
    -- 索引
    xid = 0,
    -- 食物名
    name = 1,
    -- 类别
    type = 2,
    -- 关联城市
    city = 3,
}
local XLSX_GLOBAL_FOOD = {
    [1] = {
        name = '煲仔饭',
        type = '主食',
        city = {
            cityName = '广州',
            province = '广东',
        },
    },
    [2] = {
        name = '脐橙',
        type = '水果',
        city = {
            cityName = '南昌',
            province = '江西',
        },
    },
}

return XLSX_GLOBAL_FOOD
