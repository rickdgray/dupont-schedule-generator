using OfficeOpenXml;

namespace DupontGenerator
{
    public static class Extensions
    {
        // https://stackoverflow.com/questions/9096176
        public static void SetTrueColumnWidth(this ExcelColumn column, double desiredWidth)
        {
            if (desiredWidth < 1)
            {
                column.Width = 12d * desiredWidth / 7;
                return;
            }

            column.Width = 5d / 7 + desiredWidth;
        }
    }
}
