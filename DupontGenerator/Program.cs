using DupontGenerator;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

ExcelPackage.License.SetNonCommercialPersonal("Rick Gray");

var dupontSchedule = new DupontRingList(10);

var nextYear = DateTime.Now.Year + 1;

using var package = new ExcelPackage();
var worksheet = package.Workbook.Worksheets.Add(nextYear.ToString());

var currentDate = new DateTime(nextYear, 1, 1);

while (currentDate.Year == nextYear)
{
    // day
    worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 1].Value = currentDate.Day;
    worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 1].Style.Font.Bold = true;
    worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

    // day of week
    worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 2].Value = currentDate.DayOfWeek.ToString().First();
    worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 2].Style.Font.Color.SetColor(Color.FromArgb(116, 116, 116));
    worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

    // weekend
    if (currentDate.DayOfWeek.ToString().First() == 'S')
    {
        worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(232, 232, 232));
        worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 2].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(232, 232, 232));
        worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 3].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(232, 232, 232));
    }

    // scheduled work
    worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 3].Value = dupontSchedule.GetNext();
    worksheet.Cells[currentDate.Day + 1, (currentDate.Month - 1) * 3 + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

    currentDate = currentDate.AddDays(1);
}

// months header
worksheet.Cells["A1:C1"].Merge = true;
worksheet.Cells["A1:C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["A1:C1"].Style.Font.Bold = true;
worksheet.Cells["A1:C1"].Value = "Jan";

worksheet.Cells["D1:F1"].Merge = true;
worksheet.Cells["D1:F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["D1:F1"].Style.Font.Bold = true;
worksheet.Cells["D1:F1"].Value = "Feb";

worksheet.Cells["G1:I1"].Merge = true;
worksheet.Cells["G1:I1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["G1:I1"].Style.Font.Bold = true;
worksheet.Cells["G1:I1"].Value = "Mar";

worksheet.Cells["J1:L1"].Merge = true;
worksheet.Cells["J1:L1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["J1:L1"].Style.Font.Bold = true;
worksheet.Cells["J1:L1"].Value = "Apr";

worksheet.Cells["M1:O1"].Merge = true;
worksheet.Cells["M1:O1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["M1:O1"].Style.Font.Bold = true;
worksheet.Cells["M1:O1"].Value = "May";

worksheet.Cells["P1:R1"].Merge = true;
worksheet.Cells["P1:R1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["P1:R1"].Style.Font.Bold = true;
worksheet.Cells["P1:R1"].Value = "Jun";

worksheet.Cells["S1:U1"].Merge = true;
worksheet.Cells["S1:U1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["S1:U1"].Style.Font.Bold = true;
worksheet.Cells["S1:U1"].Value = "Jul";

worksheet.Cells["V1:X1"].Merge = true;
worksheet.Cells["V1:X1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["V1:X1"].Style.Font.Bold = true;
worksheet.Cells["V1:X1"].Value = "Aug";

worksheet.Cells["Y1:AA1"].Merge = true;
worksheet.Cells["Y1:AA1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["Y1:AA1"].Style.Font.Bold = true;
worksheet.Cells["Y1:AA1"].Value = "Sep";

worksheet.Cells["AB1:AD1"].Merge = true;
worksheet.Cells["AB1:AD1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["AB1:AD1"].Style.Font.Bold = true;
worksheet.Cells["AB1:AD1"].Value = "Oct";

worksheet.Cells["AE1:AG1"].Merge = true;
worksheet.Cells["AE1:AG1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["AE1:AG1"].Style.Font.Bold = true;
worksheet.Cells["AE1:AG1"].Value = "Nov";

worksheet.Cells["AH1:AJ1"].Merge = true;
worksheet.Cells["AH1:AJ1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["AH1:AJ1"].Style.Font.Bold = true;
worksheet.Cells["AH1:AJ1"].Value = "Dec";

// vertical borders
for (int i = 0; i < 36; i += 3)
{
    worksheet.Cells[1, i + 1, 32, i + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
}

// horizontal borders
for (int i = 0; i < 32; i++)
{
    worksheet.Cells[i + 1, 1, i + 1, 36].Style.Border.BorderAround(ExcelBorderStyle.Thin);
}

// column widths to fit on one page
for (int col = 1; col <= 36; col++)
{
    if (col % 3 == 1)
    {
        worksheet.Column(col).SetTrueColumnWidth(2.43);
        continue;
    }

    if (col % 3 == 2)
    {
        worksheet.Column(col).SetTrueColumnWidth(1.86);
        continue;
    }

    worksheet.Column(col).SetTrueColumnWidth(3.71);
}

var ht = worksheet.HeaderFooter.OddHeader.Centered.AddText("Tanya's Schedule 2026");
ht.FontName = "Aptos Display";
ht.FontSize = 24;
ht.Bold = true;

// fake footer on row 33
worksheet.Cells["A33:C33"].Merge = true;
worksheet.Cells["A33:C33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["A33:C33"].Value = "D: Day";

worksheet.Cells["D33:F33"].Merge = true;
worksheet.Cells["D33:F33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["D33:F33"].Value = "N: Night";

worksheet.Cells["G33:I33"].Merge = true;
worksheet.Cells["G33:I33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
worksheet.Cells["G33:I33"].Value = "R: Relief";

worksheet.View.PageLayoutView = true;
worksheet.PrinterSettings.Orientation = eOrientation.Landscape;
worksheet.PrinterSettings.FitToPage = true;
worksheet.PrinterSettings.FitToWidth = 1;
worksheet.PrinterSettings.HorizontalCentered = true;

var excelFile = new FileInfo($"Dupont_{nextYear}.xlsx");
await package.SaveAsAsync(excelFile);
