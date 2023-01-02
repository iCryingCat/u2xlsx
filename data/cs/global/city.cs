using XLSX.GLOBAL;
namespace XLSX.GLOBAL
{
    // 城市表
    public class CITY
    {
        // 索引
        public string id { get; set; }
        // 城市名
        public string cityName { get; set; }
        // 省份
        public string province { get; set; }
        // 特产
        public XLSX.GLOBAL.FOOD food { get; set; }
        // 行政区
        public List<string> regions { get; set; }

    }

}
