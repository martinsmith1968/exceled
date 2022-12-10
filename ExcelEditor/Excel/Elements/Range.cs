using ExcelEditor.Lib.Excel.Elements;

namespace ExcelEditor.Excel.Elements
{
    public class Range : IRange
    {
        public static string DefaultAddress = "A1";
        public string Reference { get; private set; }

        public int Row { get; private set; }

        public int Column { get; private set; }

        public static IRange Parse(string addressText)
        {
            if (string.IsNullOrEmpty(addressText))
                return null;

            var range = new Range()
            {
                Reference = addressText
            };

            return range;
        }
    }
}
