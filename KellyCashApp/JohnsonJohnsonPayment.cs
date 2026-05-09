using ClosedXML.Excel;
using System.Globalization;

namespace KellyCashApp
{
    public class JohnsonJohnsonPayment
    {
        public static bool IsJohnsonJohnsonFormat(IXLWorksheet worksheet)
        {
            return worksheet.Cell(1, 1).GetString().Trim()
                .Equals("PeopleSoftInvoiceNumber", StringComparison.OrdinalIgnoreCase);
        }

        public static string Process(XLWorkbook workbook, IXLWorksheet worksheet, string inputPath)
        {
            int headerRow = 1;

            int weekEndingColumn = FindColumn(worksheet, headerRow, "WeekEndingDate");
            int lineTotalColumn = FindColumn(worksheet, headerRow, "LineTotal");

            if (weekEndingColumn == -1 || lineTotalColumn == -1)
                throw new Exception("Could not find WeekEndingDate or LineTotal for Johnson & Johnson format.");

            int nameColumn = FindColumn(worksheet, headerRow, "Name");

            NormalizeColumnsForJnj(worksheet, headerRow);

            int aggregateColumn = FindColumn(worksheet, headerRow, "Aggregate Line Total");
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            var totalsByWeekEnding = new Dictionary<string, decimal>();
            var lastRowByWeekEnding = new Dictionary<string, int>();

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string weekEnding = FormatWeekEndingDate(worksheet.Cell(row, weekEndingColumn));

                if (string.IsNullOrWhiteSpace(weekEnding))
                    continue;

                decimal lineTotal = GetDecimalValue(worksheet.Cell(row, lineTotalColumn));

                if (!totalsByWeekEnding.ContainsKey(weekEnding))
                    totalsByWeekEnding[weekEnding] = 0;

                totalsByWeekEnding[weekEnding] += lineTotal;
                lastRowByWeekEnding[weekEnding] = row;
            }

            foreach (var item in totalsByWeekEnding)
            {
                int targetRow = lastRowByWeekEnding[item.Key];
                var cell = worksheet.Cell(targetRow, aggregateColumn);

                cell.Value = item.Value;
                cell.Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

                if (item.Value < 0)
                    cell.Style.Font.FontColor = XLColor.Red;
            }

            worksheet.Columns().AdjustToContents();

            string downloadsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads"
            );

            string outputPath = GetUniqueOutputPath(downloadsPath, "Johnson & Johnson Fixed Fee Payment.xlsx");

            workbook.SaveAs(outputPath);

            return outputPath;
        }

        private static void NormalizeColumnsForJnj(IXLWorksheet worksheet, int headerRow)
        {
            int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

            string[] columnsToAdd =
            {
                "Invoice",
                "Amount",
                "Aggregate Line Total",
                "Notes"
            };

            int insertAt = lastColumn + 1;

            foreach (string columnName in columnsToAdd)
            {
                if (FindColumn(worksheet, headerRow, columnName) == -1)
                {
                    worksheet.Cell(headerRow, insertAt).Value = columnName;
                    insertAt++;
                }
            }
        }

        private static int FindColumn(IXLWorksheet worksheet, int headerRow, string headerName)
        {
            int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 100;

            for (int col = 1; col <= lastColumn; col++)
            {
                string headerText = worksheet.Cell(headerRow, col).GetString().Trim();

                if (headerText.Equals(headerName, StringComparison.OrdinalIgnoreCase))
                    return col;
            }

            return -1;
        }

        private static decimal GetDecimalValue(IXLCell cell)
        {
            string rawValue = cell.Value.ToString()
                .Replace("$", "")
                .Replace(",", "")
                .Trim();

            if (decimal.TryParse(rawValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return 0;
        }

        private static string FormatWeekEndingDate(IXLCell cell)
        {
            if (cell.Value.IsDateTime)
                return cell.GetDateTime().ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            string rawValue = cell.GetString().Trim();

            if (DateTime.TryParse(rawValue, out DateTime parsedDate))
                return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            return rawValue;
        }

        private static string GetUniqueOutputPath(string folderPath, string fileName)
        {
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);

            string outputPath = Path.Combine(folderPath, fileName);
            int counter = 1;

            while (File.Exists(outputPath))
            {
                outputPath = Path.Combine(folderPath, $"{fileNameWithoutExt} ({counter}){extension}");
                counter++;
            }

            return outputPath;
        }
    }
}