using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public class ModifyCache
{
    // 缓存文件路径
    private const string CacheFileName = "modify_cache.txt";
    public readonly string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, CacheFileName);

    Dictionary<string, long> oldCache = new Dictionary<string, long>();
    Dictionary<string, long> newCache = new Dictionary<string, long>();

    public ModifyCache()
    {
        if (File.Exists(path))
        {
            using (StreamReader sr = new StreamReader(path, Encoding.UTF8))
            {
                string line;
                // 使用 ReadLine 循环读取，效率更高且更健壮
                while ((line = sr.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    string[] item2 = line.Split(new char[] { ':' }, 2); // 只分割一次，确保路径中包含冒号时不会出错

                    if (item2.Length == 2)
                    {
                        string cfgName = item2[0].Trim();
                        string timeStr = item2[1].Trim();

                        try
                        {
                            // 使用 long.Parse 兼容您当前的 long 类型时间戳
                            oldCache[cfgName] = long.Parse(timeStr);
                        }
                        catch (FormatException)
                        {
                            Console.WriteLine($"警告: 缓存文件 {CacheFileName} 中时间戳格式错误: {timeStr}");
                            // 忽略此条错误记录
                        }
                    }
                }
            }
        }
    }

    public bool IsDiff(string cfgName, long time)
    {
        newCache[cfgName] = time;
        if (oldCache.ContainsKey(cfgName))
        {
            return oldCache[cfgName] != time;
        }

        return true;
    }

    public void SaveCache()
    {
        using (StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8))
        {
            foreach (var item in newCache)
            {
                sw.WriteLine($"{item.Key}:{item.Value}");
            }
        }
    }
}