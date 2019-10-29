using OutputCfg;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YUtility.Data {
    public class CfgDBTest {
        public static Dictionary<string, byte[]> cfgDataDic;

        public static void ReadAllCfg(string path) {
            if (File.Exists(path)) {
                cfgDataDic = new Dictionary<string, byte[]>();
                using (FileStream sr=File.OpenRead(path)) {
                    byte[] allData = new byte[sr.Length];
                    sr.Read(allData, 0, allData.Length);
                    //Console.WriteLine("解前 \n" + allData);
                    //Console.WriteLine("解后 \n" + CompressHelper.Decompress(allData));
                    allData = CompressHelper.Decompress(allData);

                    ByteArray ba = new ByteArray(allData);
                    while (ba.CanRead()) {
                        try {
                            //表名
                            string tableName = ba.ReadString();

                            //长度
                            long length = ba.ReadInt64();

                            //数据
                            byte[] data = ba.ReadBytes((int)length);

                            if (cfgDataDic.ContainsKey(tableName)) {
                                Console.WriteLine("表名{0}有重复！！",tableName);
                            }

                            cfgDataDic[tableName] = data;
                        }
                        catch (Exception ex) {
                            Console.WriteLine(ex.ToString());
                            break;
                        }
                    }
                }
            }
        }

        public static void ReadCfg<T>(string tableName,Action<T> action) where T:BaseConfigModel{
            if (cfgDataDic != null && cfgDataDic.ContainsKey(tableName)) {
                ByteArray ba = new ByteArray(cfgDataDic[tableName]);

                while (ba.CanRead()) {
                    T obj = Activator.CreateInstance<T>();
                    obj.Read(ba);
                    action(obj);
                }
                
            }
        }
    }
}
