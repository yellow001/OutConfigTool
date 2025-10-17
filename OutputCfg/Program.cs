using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutputCfg {
    class Program {
        static void Main(string[] args) {

            //string[] paths = Directory.GetFiles(AppSetting.Ins.GetValue("xlsDir"), "*.xl*");
            bool isExit = false;
            string cp = "";

            string csvp = "";

            while (!isExit) {
                Console.Clear();
                Console.WriteLine("-------导表工具v2.0---------");
                //Console.WriteLine("1 导出配置类以及配置数据(汇总一个配置文件)");
                Console.WriteLine("2 导出配置类以及配置数据(配置文件分开导出)");
                //Console.WriteLine("3 仅导出配置数据(汇总一个配置文件)");
                Console.WriteLine("4 仅导出配置数据(配置文件分开导出)");

                int num;

                string xp = AppSetting.Ins.GetValue("xlsCfgDir");

                cp = AppSetting.Ins.GetValue("cfgOutputDir");
                csvp = AppSetting.Ins.GetValue("xlsCfgCsvDir");

                if (int.TryParse(Console.ReadLine(), out num)) {
                    switch (num) {
                        case 1:
                            //XlsTool.AllXlsToCfgModel();
                            //XlsTool.XlsToOneConfig(xp, cp);

                            //isExit = true;
                            break;
                        case 2:
                            XlsTool.AllXlsToCfgModel();
                            XlsTool.XlsToSingleConfig(xp, cp);
                            XlsTool.XlsToCsv(xp, csvp);
                            isExit = true;
                            break;
                        case 3:
                            //XlsTool.XlsToOneConfig(xp, cp);
                            //isExit = true;
                            break;
                        case 4:
                            XlsTool.XlsToSingleConfig(xp, cp);
                            XlsTool.XlsToCsv(xp, csvp);
                            isExit = true;
                            break;
                        default:
                            break;
                    }
                }

                Console.Read();
            }


            //CfgDBTest.ReadAllCfg(cp + Path.DirectorySeparatorChar + "cfg.bytes");

            //Dictionary<int, CTestPath> dic = new Dictionary<int, CTestPath>();

            //CfgDBTest.ReadSingleCfg<CTestPath>(cp,CTestPath.tableName, (obj) =>
            //{
            //    dic[obj.id] = obj;
            //});

            //Console.WriteLine(dic[1].q_path);
            //Console.ReadLine();
            //Console.ReadLine();
            //XlsToCsv.ConvertXlsToCSV(paths[0], "");
            //XlsToCsv.Test(paths[0]);
        }
    }
}
