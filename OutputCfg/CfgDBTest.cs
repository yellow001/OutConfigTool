using System;
using System.Collections.Generic;
using System.IO;

public class CfgDBTest
{
    public static Dictionary<string, byte[]> cfgDataDic;

    public static void ReadAllCfg(string path)
    {
        if (File.Exists(path))
        {
            cfgDataDic = new Dictionary<string, byte[]>();
            using (FileStream sr = File.OpenRead(path))
            {
                byte[] allData = new byte[sr.Length];
                sr.Read(allData, 0, allData.Length);
                //Console.WriteLine("解前 \n" + allData);
                //Console.WriteLine("解后 \n" + CompressHelper.Decompress(allData));
                allData = CompressHelper.Decompress(allData);

                ByteArray ba = new ByteArray(allData);
                while (ba.CanRead())
                {
                    try
                    {
                        //表名
                        string tableName = ba.ReadString();

                        //长度
                        long length = ba.ReadInt64();

                        //数据
                        byte[] data = ba.ReadBytes((int)length);

                        if (cfgDataDic.ContainsKey(tableName))
                        {
                            Console.WriteLine("表名{0}有重复！！", tableName);
                        }

                        cfgDataDic[tableName] = data;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        break;
                    }
                }
                ba.Close();
            }
        }
    }

    public static void ReadCfg<T>(string tableName, Action<T> action) where T : BaseConfigModel, new()
    {
        if (cfgDataDic != null && cfgDataDic.ContainsKey(tableName))
        {
            ByteArray ba = new ByteArray(cfgDataDic[tableName]);

            while (ba.CanRead())
            {
                T obj = new T();
                obj.Read(ba);
                action(obj);
            }
            ba.Close();
        }
    }

    public static void ReadSingleCfg<T>(string path, string tableName, Action<T> action) where T : BaseConfigModel, new()
    {
        string realPath = Path.Combine(path, $"{tableName}.bytes");
        if (!File.Exists(realPath))
        {
            Console.WriteLine($"{tableName}.bytes 不存在！！！");
            return;
        }

        using (FileStream sr = File.OpenRead(realPath))
        {
            byte[] allData = new byte[sr.Length];
            sr.Read(allData, 0, allData.Length);
            //Console.WriteLine("解前 \n" + allData);
            //Console.WriteLine("解后 \n" + CompressHelper.Decompress(allData));
            allData = CompressHelper.Decompress(allData);

            using (ByteArray ba = new ByteArray(allData))
            {
                if (!ba.CanRead())
                {
                    return; // 空数据，直接返回
                }
                try
                {
                    //表名
                    string tbName = ba.ReadString();

                    if (tbName != tableName)
                    {
                        Console.WriteLine($"警告：表名'{tableName}'与文件数据中的表名'{tbName}'不一致！");
                        ba.Close();
                        return;
                    }

                    //长度
                    long length = ba.ReadInt64();

                    //数据
                    byte[] data = ba.ReadBytes((int)length);

                    // 3. 使用 using 封装内层 ByteArray (rowsBa)
                    using (ByteArray rowsBa = new ByteArray(data))
                    {
                        if (!rowsBa.CanRead()) return;

                        // 读取数据流开头的记录总数
                        int totalRecords = rowsBa.ReadInt32();

                        // 使用 for 循环精确读取 totalRecords 次
                        for (int i = 0; i < totalRecords; i++)
                        {
                            T obj = new T();
                            obj.Read(rowsBa);
                            action(obj);
                        }
                    } // rowsBa.Dispose() 自动调用
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"错误：读取配置表 {tableName} 失败。详情: {ex.ToString()}");
                }
            }// ba.Dispose() 自动调用
        }// sr.Dispose() 自动调用
    }
}
