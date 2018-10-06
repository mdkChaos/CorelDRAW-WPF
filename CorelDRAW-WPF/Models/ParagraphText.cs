using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorelDRAW_WPF.Models
{
    class ParagraphText : CommonText
    {
        public double Top { get; }
        public double Right { get; }

        public ParagraphText(
            double left,
            double top,
            double right,
            double bottom,
            string text,
            string font = "Arial",
            float size = 12,
            string name = "Name"
        ) : base(left, bottom, text, font, size, name)
        {
            Top = top;
            Right = right;
        }
    }
}
