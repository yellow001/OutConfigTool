using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YUtility.Data{
    public class CTest : BaseConfigModel {

//member
		public static string tableName="test";
		/// <summary>
		/// 主键id
		/// </summary>
		public int id;
		/// <summary>
		/// 名字
		/// </summary>
		public string name;
//endMember


        public override void Read(ByteArray b) {
            base.Read(b);

			id=ReadInt();
			name=ReadString();


            while (b.CanRead() && !b.ReadString().Equals("#")) { }
        }

    }
}

