using System;

namespace Excellency.Attributes
{
    public class CellSourceAttribute :Attribute
    {
        public string Source { get; set; }
        public CellSourceAttribute(string source)
        {
            Source = source;
        }
    }
}
