using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YUtility.Data;
using YUtility.Data.Config;

namespace OutputCfg {
    class Program {
        static void Main(string[] args) {

            //string[] paths = Directory.GetFiles(AppSetting.Ins.GetValue("xlsDir"), "*.xl*");
            bool isExit = false;
            string cp = "";
            while (!isExit) {
                Console.Clear();
                Console.WriteLine("-------导表工具v1.0---------");
                Console.WriteLine("1 导出配置类以及配置数据");
                Console.WriteLine("2 仅导出配置数据");

                int num;

                string xp = AppSetting.Ins.GetValue("xlsCfgDir");

                cp = AppSetting.Ins.GetValue("cfgOutputDir");

                if (int.TryParse(Console.ReadLine(), out num)) {
                    switch (num) {
                        case 1:
                            XlsTool.AllXlsToCfgModel();
                            XlsTool.AllXlsToConfig(xp, cp);

                            isExit = true;
                            break;
                        case 2:
                            XlsTool.AllXlsToConfig(xp, cp);
                            isExit = true;
                            break;
                        default:
                            break;
                    }
                }

                Console.Read();
            }
            

            //CfgDBTest.ReadAllCfg(cp + Path.DirectorySeparatorChar + "cfg.bytes");

            //Dictionary<int, CResPath> dic = new Dictionary<int, CResPath>();

            //CfgDBTest.ReadCfg<CResPath>(CResPath.tableName, (obj) => {
            //    dic[obj.q_id] = obj;
            //});

            //Console.WriteLine(dic[1].q_path);
            //Console.ReadLine();
            //Console.ReadLine();
            //XlsToCsv.ConvertXlsToCSV(paths[0], "");
            //XlsToCsv.Test(paths[0]);
        }
    }
}
