using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorelDRAW_WPF.Models
{
    class ArtisticText : CommonText
    {
        public double SizeWidth { get; }

        public ArtisticText(
            double left,
            double bottom,
            string text,
            string font = "Arial",
            float size = 12,
            string name = "Name",
            double sizeWidth = 114.3,
            int alignment = 1
        ) : base(left, bottom, text, font, size, name, alignment)
        {
            SizeWidth = sizeWidth;
        }
    }
}
