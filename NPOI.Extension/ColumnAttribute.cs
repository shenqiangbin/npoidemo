using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.Extension
{
    public class ColumnAttribute : Attribute
    {
        public int Index { get; internal set; }
        public double Title { get; internal set; }
    }
}
