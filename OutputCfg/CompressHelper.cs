using ICSharpCode.SharpZipLib.BZip2;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public class CompressHelper {


    #region BZip
    /// <summary>
    /// 压缩
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public static string Compress(string input) {
        byte[] buffer = Encoding.UTF8.GetBytes(input);
        using (MemoryStream outputStream = new MemoryStream()) {
            using (BZip2OutputStream zipStream = new BZip2OutputStream(outputStream)) {
                zipStream.Write(buffer, 0, buffer.Length);
                zipStream.Close();
            }
            return Convert.ToBase64String(outputStream.ToArray());
        }

    }

    /// <summary>
    /// 压缩
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public static byte[] Compress(byte[] input) {
        using (MemoryStream outputStream = new MemoryStream()) {
            using (BZip2OutputStream zipStream = new BZip2OutputStream(outputStream)) {
                zipStream.Write(input, 0, input.Length);
                zipStream.Close();
            }
            return outputStream.ToArray();
        }
    }


    /// <summary>
    /// 解压缩
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public static string Decompress(string input) {
        string result = string.Empty;
        byte[] buffer = Convert.FromBase64String(input);
        using (Stream inputStream = new MemoryStream(buffer)) {
            BZip2InputStream zipStream = new BZip2InputStream(inputStream);

            using (StreamReader reader = new StreamReader(zipStream, Encoding.UTF8)) {
                //输出
                result = reader.ReadToEnd();
            }
        }

        return result;

    }

    /// <summary>
    /// 解压缩
    /// </summary>
    /// <param name="input"></param>
    /// <returns></returns>
    public static byte[] Decompress(byte[] input) {
        using (Stream inputStream = new MemoryStream(input)) {
            BZip2InputStream zipStream = new BZip2InputStream(inputStream);

            MemoryStream outputStream = new MemoryStream();
            int count = 0;

            byte[] data = new byte[4096];

            while ((count = zipStream.Read(data, 0, data.Length)) != 0) {
                outputStream.Write(data, 0, count);
            }

            byte[] result = new byte[outputStream.Length];
            result = outputStream.ToArray();

            outputStream.Dispose();

            return result;
        }
    }
    #endregion
}
