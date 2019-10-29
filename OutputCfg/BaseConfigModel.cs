using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YUtility.Data{
    public class BaseConfigModel {
        ByteArray ba;

        public virtual void Read(ByteArray b) {
            ba = b;
        }

        public int ReadInt() {
            return ba.ReadInt32();
        }

        public string ReadString() {
            return ba.ReadString();
        }

        public bool ReadBool() {
            return ba.ReadBool();
        }

        public long ReadLong() {
            return ba.ReadInt64();
        }
    }
}


