﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Berry.Docx.Visual
{
    public class Margin
    {
        public Margin(double left, double top, double right, double bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

		public Margin() {
            Left = 0;
            Top = 0;
            Right = 0;
            Bottom = 0;
		}

        public double Left { get; set; } //= 0;
        public double Top { get; set; } //= 0;
        public double Right { get; set; } //= 0;
        public double Bottom { get; set; } //= 0;

        public override string ToString()
        {
            return "({Left}, {Top}, {Right}, {Bottom})";
        }
    }
}
