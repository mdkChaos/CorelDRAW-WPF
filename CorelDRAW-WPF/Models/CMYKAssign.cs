using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorelDRAW_WPF.Models
{
    class CMYKAssign
    {
        public byte C { get;}
        public byte M { get; }
        public byte Y { get; }
        public byte K { get; }

        public CMYKAssign(byte c, byte m, byte y, byte k)
        {
            C = c;
            M = m;
            Y = y;
            K = k;
        }
    }
}
