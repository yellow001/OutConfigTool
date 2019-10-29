using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YUtility.Data.Config {
    public class ConfigModelTemple : BaseConfigModel {

//member

        public override void Read(ByteArray b) {
            base.Read(b);

//read

            while (b.CanRead() && !b.ReadString().Equals("#")) { }
        }

    }
}

