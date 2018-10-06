using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorelDRAW_WPF.Models
{
    class RGBAssign
    {
        public byte R { get; }
        public byte G { get; }
        public byte B { get; }

        public RGBAssign(byte r, byte g, byte b)
        {
            R = r;
            G = g;
            B = b;
        }
    }
}
