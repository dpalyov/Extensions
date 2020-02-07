using OfficeOpenXml.Table;
using System.Drawing;

namespace Extensions
{
    public class ExcelOptions
    {

        public string WorksheetName { get; set; }
        public Color HeaderBgColor { get; set; }

        public Color HeaderFontColor { get; set; }

        public bool HeaderBold { get; set; }

        public string ExcelTableName { get; set; }

        public TableStyles TableStyles { get; set; }
        public bool ShowFilter { get; set; }

        public bool ShowHeader { get; set; }

        public bool ShowStripes { get; set; }

        public int TableOffsetX { get; set; }

        public int TableOffsetY { get; set; }

    }
}