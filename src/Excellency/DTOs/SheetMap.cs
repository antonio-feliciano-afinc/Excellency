using NPOI.SS.UserModel;
using System;

namespace Excellency.DTOs
{
    public class SheetMap
    {
        public SheetMap(string sheetName, Type type)
        {
            SheetName = sheetName;
            Type = type;
        }

        public string SheetName { get; private set; }
        public Type Type { get; private set; }
        internal ISheet Sheet {  get;  set; }
    }
}
