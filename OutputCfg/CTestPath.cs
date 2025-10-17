using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

public class CTestPath : BaseConfigModel
{

    //member
	public static string tableName="q_test_path";
	/// <summary>
	/// 主键
	/// </summary>
	public int id;
	/// <summary>
	/// 名字
	/// </summary>
	public string name;
	/// <summary>
	/// 列表名
	/// </summary>
	public List<string> listName;
	/// <summary>
	/// 列表值
	/// </summary>
	public List<int> listValue;
	/// <summary>
	/// 字典值
	/// </summary>
	public Dictionary<int, int> dicX;
	//endMember


    public override void Read(ByteArray b)
    {
        base.Read(b);
        
		id=ReadInt();
		name=ReadString();
		listName=ReadListStr();
		listValue=ReadListInt();
		dicX=ReadDicIntInt();

    }

}

