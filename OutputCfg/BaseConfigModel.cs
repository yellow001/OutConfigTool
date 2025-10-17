using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

public class BaseConfigModel
{

    private const char ListSeparator = ';';
    private const char DicKVSeparator = '_'; // Dictionary Key-Value 分隔符

    ByteArray ba;

    public virtual void Read(ByteArray b)
    {
        ba = b;
    }

    public int ReadInt()
    {
        return ba.ReadInt32();
    }

    public string ReadString()
    {
        return ba.ReadString();
    }

    public bool ReadBool()
    {
        return ba.ReadBool();
    }

    public long ReadLong()
    {
        return ba.ReadInt64();
    }

    public List<string> ReadListStr()
    {
        List<string> result = new List<string>();
        string data = ba.ReadString();
        if (string.IsNullOrEmpty(data))
        {
            return result;
        }
        string[] datas = data.Split(ListSeparator);
        result.AddRange(datas);

        return result;
    }

    public List<int> ReadListInt()
    {
        List<int> result = new List<int>();
        string data = ba.ReadString();
        if (string.IsNullOrEmpty(data))
        {
            return result;
        }
        string[] datas = data.Split(ListSeparator);
        foreach (var str in datas)
        {
            if (int.TryParse(str, out int value))
            {
                result.Add(value);
            }
        }
        return result;
    }

    public List<long> ReadListLong()
    {
        List<long> result = new List<long>();
        string data = ba.ReadString();
        if (string.IsNullOrEmpty(data))
        {
            return result;
        }
        string[] datas = data.Split(ListSeparator);
        foreach (var str in datas)
        {
            if (long.TryParse(str, out long value))
            {
                result.Add(value);
            }
        }
        return result;
    }

    public Dictionary<int, int> ReadDicIntInt()
    {
        Dictionary<int, int> result = new Dictionary<int, int>();
        string data = ba.ReadString();
        if (string.IsNullOrEmpty(data))
        {
            return result;
        }
        string[] datas = data.Split(ListSeparator);
        foreach (var str in datas)
        {
            string[] tmp = str.Split(DicKVSeparator);
            if (tmp.Length >= 2)
            {
                if (int.TryParse(tmp[0], out int key) && int.TryParse(tmp[1], out int value))
                {
                    result[key] = value;
                }
            }
        }
        return result;
    }

    public Dictionary<int, string> ReadDicIntStr()
    {
        Dictionary<int, string> result = new Dictionary<int, string>();
        string data = ba.ReadString();
        if (string.IsNullOrEmpty(data))
        {
            return result;
        }
        string[] datas = data.Split(ListSeparator);
        foreach (var str in datas)
        {
            string[] tmp = str.Split(DicKVSeparator);
            if (tmp.Length >= 2)
            {
                if (int.TryParse(tmp[0], out int key))
                {
                    result[key] = tmp[1];
                }
            }
        }
        return result;
    }

    public Dictionary<string, int> ReadDicStrInt()
    {
        Dictionary<string, int> result = new Dictionary<string, int>();
        string data = ba.ReadString();
        if (string.IsNullOrEmpty(data))
        {
            return result;
        }
        string[] datas = data.Split(ListSeparator);
        foreach (var str in datas)
        {
            string[] tmp = str.Split(DicKVSeparator);
            if (tmp.Length >= 2)
            {
                if (int.TryParse(tmp[1], out int value))
                {
                    result[tmp[0]] = value;
                }
            }
        }
        return result;
    }

    public Dictionary<string, string> ReadDicStrStr()
    {
        Dictionary<string, string> result = new Dictionary<string, string>();
        string data = ba.ReadString();
        if (string.IsNullOrEmpty(data))
        {
            return result;
        }
        string[] datas = data.Split(ListSeparator);
        foreach (var str in datas)
        {
            string[] tmp = str.Split(DicKVSeparator);
            if (tmp.Length >= 2)
            {
                result[tmp[0]] = tmp[1];
            }
        }
        return result;
    }
}


