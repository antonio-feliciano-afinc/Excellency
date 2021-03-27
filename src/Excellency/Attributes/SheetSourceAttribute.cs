using System;

namespace Excellency.Attributes
{
    public class SheetSourceAttribute : Attribute
    {
        public string Source { get; set; }
        public SheetSourceAttribute(string source)
        {
            Source = source;
        }
    }
}
