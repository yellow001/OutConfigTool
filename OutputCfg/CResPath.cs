using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YUtility.Data.Config {
    public class CResPath : BaseConfigModel {

        //member
        public static string tableName = "respath";
        /// <summary>
        /// 主键
        /// </summary>
        public int q_id;
        /// <summary>
        /// 路径(Resource/AB  路径)
        /// </summary>
        public string q_path;
        //endMember


        public override void Read(ByteArray b) {
            base.Read(b);

            q_id = ReadInt();
            q_path = ReadString();


            while (b.CanRead() && !b.ReadString().Equals("#")) { }
        }

    }
}

