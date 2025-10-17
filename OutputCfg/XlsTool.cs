using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader;

public class XlsTool {

    static string tempPath = AppDomain.CurrentDomain.BaseDirectory + "/" + "ConfigModelTemple.cs";

    static string content = "";

    /// <summary>
    /// xls转成csv
    /// </summary>
    /// <param name="src"></param>
    /// <param name="dst"></param>
    public static void XlsToCSV(string src, string dst) {
        try {
            if ((src.EndsWith(".xls") || src.EndsWith(".xlsx"))&&!src.Contains("$")) {
                using (FileStream st = File.Open(src, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    IExcelDataReader reader = ExcelReaderFactory.CreateReader(st);
                    //Console.Write(reader.Name);
                    StringBuilder sb = new StringBuilder();
                    while (reader.Read()) {
                        for (int i = 0; i < reader.FieldCount; i++) {
                            sb.Append(reader.GetString(i));
                            if (i != reader.FieldCount - 1) {
                                sb.Append(",");

                            }
                        }
                        sb.Append("\r\n");
                    }

                    string p = src.Replace(".xlsx", ".csv");
                    p = p.Replace(".xls", "csv");

                    using (FileStream stream = File.Create(p)) {
                        byte[] content = Encoding.UTF8.GetBytes(sb.ToString());
                        stream.Write(content, 0, content.Length);
                    }
                    reader.Close();
                }
            }

            Console.WriteLine("xls转csv成功");
        }
        catch (Exception ex) {
            Console.WriteLine(ex.ToString());
        }


        Console.Read();
    }

    /// <summary>
    /// csv转成配置类
    /// </summary>
    public static void CvsToCfgModel() {


    }

    /// <summary>
    /// 直接把xls转成配置类
    /// </summary>
    public static void XlsToCfgModel(string src, string dst) {
        try {
            if (string.IsNullOrEmpty(content)) {

                if (!File.Exists(tempPath)) {
                    Console.WriteLine("模板文件不存在");
                    return;
                }

                StreamReader sr = new StreamReader(tempPath);

                content = sr.ReadToEnd();

                sr.Dispose();
            }

            if ((src.EndsWith(".xls") || src.EndsWith(".xlsx"))&&!src.Contains("$")) {
                using (FileStream st = File.Open(src, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    IExcelDataReader reader = ExcelReaderFactory.CreateReader(st);
                    //Console.Write(reader.Name);
                    StringBuilder sb = new StringBuilder();

                    List<string> attributeName = new List<string>();

                    List<string> atttibuteType = new List<string>();

                    List<string> attributeNote = new List<string>();

                    int attCount = 0;

                    //字段
                    if (reader.Read()) {
                        for (int i = 0; i < reader.FieldCount; i++) {
                            attributeName.Add(reader.GetString(i));
                        }
                    }

                    //类型
                    if (reader.Read()) {
                        for (int i = 0; i < reader.FieldCount; i++) {
                            string t = reader.GetString(i);
                            atttibuteType.Add(t);
                            if (!string.IsNullOrEmpty(t)) {
                                attCount++;
                            }
                        }
                    }

                    //注释
                    if (reader.Read()) {
                        for (int i = 0; i < reader.FieldCount; i++) {
                            attributeNote.Add(reader.GetString(i));
                        }
                    }

                    bool create = false;
                    if (File.Exists(dst)) {
                        //存在文件就要比较一下字段，看是否需要覆盖
                        using (StreamReader sr = new StreamReader(dst)) {
                            string allText = sr.ReadToEnd();

                            string tx = allText.Split(new string[] { "//member" }, StringSplitOptions.RemoveEmptyEntries)[1];
                            tx = tx.Split(new string[] { "//endMember" }, StringSplitOptions.RemoveEmptyEntries)[0];

                            //有多少个分号，就是有多少个字段 ( 减1是因为 tableName 不算 )
                            int oldAttCount = Regex.Matches(tx, ";").Count - 1;

                            //字段个数不相等，要重新生成
                            if (oldAttCount != attCount) {
                                create = true;
                            }
                            else {
                                //否则查看字段类型是否相同，不同的也要重新生成
                                int s = 0, e = 0;
                                for (int i = 0; i < attributeName.Count; i++) {
                                    if (!allText.Contains(attributeName[i]) && !string.IsNullOrEmpty(atttibuteType[i])) {
                                        create = true;
                                        break;
                                    }

                                    e = allText.IndexOf(attributeName[i]);

                                    string t = allText.Substring(s, e - s);
                                    if (!t.Contains(atttibuteType[i])) {
                                        create = true;
                                        break;
                                    }
                                    s = e;
                                }
                            }

                        }
                        if (create) { File.Delete(dst); }
                    }
                    else {
                        create = true;
                    }

                    if (create) {

                        Directory.CreateDirectory(Path.GetDirectoryName(dst));

                        using (StreamWriter sw = new StreamWriter(dst)) {

                            string[] contents = content.Split(new string[] { "//member" }, StringSplitOptions.RemoveEmptyEntries);

                            contents[0] = contents[0].Replace("ConfigModelTemple", Path.GetFileNameWithoutExtension(dst));

                            sb = new StringBuilder(contents[0]);

                            sb.Append("//member\n");

                            sb.Append("\t\tpublic static string tableName=\"");
                            sb.Append(Path.GetFileNameWithoutExtension(src).ToLower().Split('_')[0]);
                            sb.Append("\";\n");

                            for (int i = 0; i < attributeName.Count; i++) {

                                if (!string.IsNullOrEmpty(atttibuteType[i]) && AppSetting.Ins.GetValue("outputType").Contains(atttibuteType[i])) {

                                    sb.Append("\t\t/// <summary>\n");
                                    sb.Append("\t\t/// ");
                                    sb.Append(attributeNote[i]);
                                    sb.Append("\n\t\t/// </summary>\n");

                                    sb.Append("\t\tpublic ");

                                    switch (atttibuteType[i]) {
                                        case "int":
                                            sb.Append("int ");
                                            break;
                                        case "bool":
                                            sb.Append("bool ");
                                            break;
                                        case "long":
                                            sb.Append("long ");
                                            break;
                                        case "str":
                                            sb.Append("string ");
                                            break;
                                        case "list":
                                        case "list|int":
                                            sb.Append("List<int> ");
                                            break;
                                        case "list|long":
                                            sb.Append("List<long> ");
                                            break;
                                        case "list|str":
                                            sb.Append("List<string> ");
                                            break;
                                        case "dic":
                                        case "dic|int_int":
                                            sb.Append("Dictionary<int, int> ");
                                            break;
                                        case "dic|int_str":
                                            sb.Append("Dictionary<int, string> ");
                                            break;
                                        case "dic|str_int":
                                            sb.Append("Dictionary<string, int> ");
                                            break;
                                        case "dic|str_str":
                                            sb.Append("Dictionary<string, string> ");
                                            break;
                                        default:
                                            break;
                                    }

                                    sb.Append(attributeName[i]);
                                    sb.Append(";\n");
                                }
                            }

                            sb.Append("//endMember\n");

                            string[] readContents = contents[1].Split(new string[] { "//read" }, StringSplitOptions.RemoveEmptyEntries);
                            sb.Append(readContents[0]);

                            for (int i = 0; i < attributeName.Count; i++) {

                                if (!string.IsNullOrEmpty(atttibuteType[i])) {

                                    sb.Append("\t\t\t");
                                    sb.Append(attributeName[i]);
                                    sb.Append("=");

                                    switch (atttibuteType[i]) {
                                        case "int":
                                            sb.Append("ReadInt();\n");
                                            break;
                                        case "bool":
                                            sb.Append("ReadBool();\n");
                                            break;
                                        case "long":
                                            sb.Append("ReadLong();\n");
                                            break;
                                        case "str":
                                            sb.Append("ReadString();\n");
                                            break;
                                        case "list":
                                        case "list|int":
                                            sb.Append("ReadListInt();\n");
                                            break;
                                        case "list|long":
                                            sb.Append("ReadListLong();\n");
                                            break;
                                        case "list|str":
                                            sb.Append("ReadListStr();\n");
                                            break;
                                        case "dic":
                                        case "dic|int_int":
                                            sb.Append("ReadDicIntInt();\n");
                                            break;
                                        case "dic|int_str":
                                            sb.Append("ReadDicIntStr();\n");
                                            break;
                                        case "dic|str_int":
                                            sb.Append("ReadDicStrInt();\n");
                                            break;
                                        case "dic|str_str":
                                            sb.Append("ReadDicStrStr();\n");
                                            break;
                                        default:
                                            break;
                                    }
                                }
                            }

                            sb.Append(readContents[1]);

                            sw.Write(sb.ToString());

                            sb.Clear();
                        }
                    }
                }
            }

        }
        catch (Exception ex) {
            Console.WriteLine(ex.ToString());
        }
    }

    /// <summary>
    /// 所有的配置表汇总成一个数据文件
    /// </summary>
    /// <param name="path"></param>
    /// <param name="outPath"></param>
    public static void XlsToOneConfig(string path, string outPath) {
        //把文件夹下所有的xls转成一个总配置文件
        string[] paths = Directory.GetFiles(path, "*.xl*", SearchOption.AllDirectories);

        ByteArray ba = new ByteArray();

        for (int i = 0; i < paths.Length; i++) {
            byte[] data = XlsToConfig(paths[i]);
            if (data != null && data.LongLength > 0) {
                ba.Write(data);
            }
        }

        Directory.CreateDirectory(outPath);
        string cfgPath = outPath + Path.DirectorySeparatorChar + "cfg.bytes";
        if (File.Exists(cfgPath)) {
            File.Delete(cfgPath);
        }

        //Console.WriteLine("压前 \n"+ba.GetBuffer().ToString());
        //Console.WriteLine("压后 \n" + CompressHelper.Compress(ba.GetBuffer()));
        File.WriteAllBytes(cfgPath, CompressHelper.Compress(ba.GetBuffer()));
        
        ba.Close();

        Console.WriteLine("配置数据导出成功");
    }

    /// <summary>
    /// 每个配置文件导出自己的配置表
    /// </summary>
    /// <param name="path"></param>
    /// <param name="outPath"></param>
    public static void XlsToSingleConfig(string path, string outPath)
    {
        //把文件夹下所有的xls转成一个总配置文件
        string[] paths = Directory.GetFiles(path, "*.xl*", SearchOption.AllDirectories);

        ByteArray ba = new ByteArray();

        for (int i = 0; i < paths.Length; i++)
        {
            XlsToConfig(paths[i],true, outPath);
        }
    }

    public static byte[] XlsToConfig(string path,bool export = false,string outPath="") {
        //把xls转成字节流
        ByteArray ba = new ByteArray();

        try {
            if (File.Exists(path)&&!path.Contains("$") && (path.EndsWith(".xls") || path.EndsWith(".xlsx"))) {
                string fileName = Path.GetFileNameWithoutExtension(path);
                fileName = fileName.Split('_')[0];
                fileName = fileName.ToLower();

                ba.Write(fileName);

                //获取数据长度
                ByteArray baData = new ByteArray();

                // 所有数据行的数据
                List<byte[]> rowDatas = new List<byte[]>();

                using (FileStream st = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    IExcelDataReader reader = ExcelReaderFactory.CreateReader(st);
                    //Console.Write(reader.Name);
                    StringBuilder sb = new StringBuilder();

                    List<string> attributeName = new List<string>();

                    List<string> atttibuteType = new List<string>();

                    List<string> attributeNote = new List<string>();

                    int attCount = 0;

                    //字段
                    if (reader.Read()) {
                        for (int i = 0; i < reader.FieldCount; i++) {
                            attributeName.Add(reader.GetString(i));
                        }
                    }

                    //类型
                    if (reader.Read()) {
                        for (int i = 0; i < reader.FieldCount; i++) {
                            string t = reader.GetString(i);
                            atttibuteType.Add(t);
                            if (!string.IsNullOrEmpty(t)) {
                                attCount++;
                            }
                        }
                    }

                    //注释
                    if (reader.Read()) {
                        for (int i = 0; i < reader.FieldCount; i++) {
                            attributeNote.Add(reader.GetString(i));
                        }
                    }

                    //说明
                    if (reader.Read()) {

                    }

                    //数据
                    while (reader.Read()) {

                        ByteArray rowBa = new ByteArray(); // 临时存储单行数据

                        for (int i = 0; i < reader.FieldCount; i++) {
                            if (!string.IsNullOrEmpty(attributeName[i]) && !string.IsNullOrEmpty(atttibuteType[i])) {

                                string value = reader.GetValue(i).ToString();

                                switch (atttibuteType[i]) {
                                    case "int":
                                        rowBa.Write(int.Parse(value));
                                        break;
                                    case "bool":
                                        rowBa.Write(value.Equals("0") ? false : true);
                                        break;
                                    case "long":
                                        rowBa.Write(long.Parse(value));
                                        break;
                                    case "str":
                                    case "list":
                                    case "list|int":
                                    case "list|long":
                                    case "list|str":
                                    case "dic":
                                    case "dic|int_int":
                                    case "dic|int_str":
                                    case "dic|str_int":
                                    case "dic|str_str":
                                        // 【统一处理】对所有字符串进行换行符转义
                                        rowBa.Write(value.Replace("\r\n", "\\n").Replace("\n", "\\n"));
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }

                        // 缓存这行数据的字节流
                        rowDatas.Add(rowBa.GetBuffer());
                        rowBa.Close();
                    }
                }

                // 优化点 B：将记录总数写入 baData 的开头
                int recordCount = rowDatas.Count;
                baData.Write(recordCount); // 写入总记录数 (Int32)

                // 优化点 C：将所有缓存的行数据顺序写入 baData
                foreach (byte[] rowData in rowDatas)
                {
                    baData.Write(rowData); // 直接写入该行数据块
                }

                ba.Write(baData.GetLength());
                ba.Write(baData.GetBuffer());
                baData.Close();

                if (!export)
                {
                    return ba.GetBuffer();
                }
                else
                {
                    //导出单表配置数据文件
                    Directory.CreateDirectory(outPath);
                    string cfgPath = outPath + Path.DirectorySeparatorChar + $"{fileName}.bytes";
                    if (File.Exists(cfgPath))
                    {
                        File.Delete(cfgPath);
                    }
                    File.WriteAllBytes(cfgPath, CompressHelper.Compress(ba.GetBuffer()));

                    ba.Close();

                    Console.WriteLine($"配置数据 {fileName} 导出成功");
                }
            }
        }
        catch (Exception ex) {
            Console.WriteLine(ex.ToString());
        }
        return null;
    }

    /// <summary>
    /// 把文件夹下所有xls文件转成配置类
    /// </summary>
    /// <param name="path"></param>
    public static void AllXlsToCfgModel() {

        string outputPath = AppSetting.Ins.GetValue("cfgModelOutputDir");

        string xlsPath = AppSetting.Ins.GetValue("xlsCfgDir");

        string[] paths = Directory.GetFiles(xlsPath, "*.xl*", SearchOption.AllDirectories);

        for (int i = 0; i < paths.Length; i++) {
            string fileName = Path.GetFileNameWithoutExtension(paths[i]);
            fileName = fileName.Split('_')[0];
            string className = "C" + fileName.Substring(0, 1).ToUpper() + fileName.Substring(1);

            string relativePath = paths[i].Replace(xlsPath, "");

            relativePath = relativePath.Split(new string[] { fileName }, StringSplitOptions.RemoveEmptyEntries)[0];

            string truePath = outputPath + relativePath + className + ".cs";

            XlsToCfgModel(paths[i], truePath);
        }

        Console.WriteLine("配置类生成成功");
    }

    //public static void Test(string src) {
    //    Application application;
    //    Workbooks workbooks;
    //    Workbook workbook;
    //    application = new ApplicationClass();

    //    workbooks = application.Workbooks;

    //    workbook = workbooks.Open(src);

    //    string p = src.Replace(".xlsx", ".csv");
    //    //p = p.Replace(".xlsx", ".csv");

    //    application.Visible = false;
    //    application.DisplayAlerts = false;

    //    workbook.SaveAs(p,XlFileFormat.xlCSV);
    //}
}
