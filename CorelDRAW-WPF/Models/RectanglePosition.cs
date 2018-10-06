using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorelDRAW_WPF.Models
{
    class RectanglePosition
    {
        public double PositionX { get; }
        public double PositionY { get; }

        public RectanglePosition(double positionX, double positionY)
        {
            PositionX = positionX;
            PositionY = positionY;
        }
    }
}
