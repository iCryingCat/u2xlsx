using XLSX.GLOBAL;
namespace XLSX.GLOBAL
{
    // 特产表
    public class FOOD
    {
        // 索引
        public string id { get; set; }
        // 食物名
        public string name { get; set; }
        // 类别
        public string type { get; set; }
        // 关联城市
        public XLSX.GLOBAL.CITY city { get; set; }

    }

}
