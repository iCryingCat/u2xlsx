using System.Text;

using Newtonsoft.Json;

namespace GFramework.Xlsx
{
    public class JsonStream
    {
        private FileStream stream;

        private string path;
        private string txt;

        public string Path { get => path; private set => path = value; }
        public string Txt { get => txt; private set => txt = value; }

        public JsonStream(string path)
        {
            if (!File.Exists(path))
                throw new Exception("file path {0} is not exist".Format(path));
            this.path = path;
            this.txt = File.ReadAllText(this.Path);
            this.stream = new FileStream(this.Path, FileMode.Open, FileAccess.ReadWrite);
        }

        //---- 读
        private string Read()
        {
            return this.txt;
        }

        public T Read<T>()
        {
            return JsonConvert.DeserializeObject<T>(this.txt);
        }

        public Dictionary<T1, T2> ToMap<T1, T2>()
        {
            return JsonConvert.DeserializeObject<Dictionary<T1, T2>>(this.txt);
        }

        //---- 写
        public void Write(string json)
        {
            this.Write(json, 0);
        }

        public void WriteToEnd(string json)
        {
            this.Write(json, this.stream.Length);
        }

        public void WriteToEnd<T>(T obj)
        {
            string json = JsonConvert.SerializeObject(obj);
            this.Write(json, this.stream.Length);
        }

        public void WriteToEnd<T1, T2>(Dictionary<T1, T2> keyMap)
        {
            string json = JsonConvert.SerializeObject(keyMap);
            this.Write(json, this.stream.Length);
        }

        private void Write(string json, long pos)
        {
            using (this.stream)
            {
                this.stream.Position = pos;
                byte[] bytes = Encoding.UTF8.GetBytes(json);
                this.stream.Write(bytes, 0, bytes.Length);
            }
        }
    }
}