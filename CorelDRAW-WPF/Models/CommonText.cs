using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorelDRAW_WPF.Models
{
    class CommonText
    {
        public double Left { get; }
        public double Bottom { get; }
        public string Text { get; }
        public string Font { get; }
        public float Size { get; }
        public string Name { get; }

        public CommonText(
            double left,
            double bottom,
            string text,
            string font,
            float size,
            string name
        )
        {
            Left = left;
            Bottom = bottom;
            Text = text;
            Font = font;
            Size = size;
            Name = name;
        }
    }
}
