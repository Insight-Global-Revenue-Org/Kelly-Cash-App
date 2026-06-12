using ClosedXML.Excel;
using KellyCashApp.Models;
using System.Globalization;

namespace KellyCashApp.Processors.Allegis
{
    internal static class MicrosoftVms
    {
        public static Dictionary<string, List<MicrosoftVmsMatch>> Import(string filePath)
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var workbook = new XLWorkbook(stream);

            var worksheet = workbook.Worksheet(1);

            int headerRow = 1;

            int workerCol = FindColumn(worksheet, headerRow, "Worker Name");
            int invoiceCol = FindColumn(worksheet, headerRow, "Invoice #");
            int rtRateCol = FindColumn(worksheet, headerRow, "RT Rate");
            int otRateCol = FindColumn(worksheet, headerRow, "OT Rate");
            int dtRateCol = FindColumn(worksheet, headerRow, "DT Rate");
            int invoicedNetCol = FindColumn(worksheet, headerRow, "Invoiced Net");
            int invoicedUnitsCol = FindColumn(worksheet, headerRow, "Invoiced Units");

            var matches = new Dictionary<string, List<MicrosoftVmsMatch>>(StringComparer.OrdinalIgnoreCase);

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string rawWorkerName = worksheet.Cell(row, workerCol).GetString().Trim();
                string workerName = FormatName(rawWorkerName);

                string invoiceNumber = worksheet.Cell(row, invoiceCol).GetString().Trim();

                string vmsIdentifier = new string(invoiceNumber.Where(char.IsDigit).ToArray());

                if (vmsIdentifier.Length > 5)
                    vmsIdentifier = vmsIdentifier[^5..];

                if (string.IsNullOrWhiteSpace(workerName) || string.IsNullOrWhiteSpace(vmsIdentifier))
                    continue;

                string key = $"{workerName}|{vmsIdentifier}";

                decimal invoicedNet = GetDecimalValue(worksheet.Cell(row, invoicedNetCol));
                decimal invoicedUnits = GetDecimalValue(worksheet.Cell(row, invoicedUnitsCol));

                if (!matches.ContainsKey(key))
                    matches[key] = new List<MicrosoftVmsMatch>();

                if (!matches[key].Any())
                {
                    matches[key].Add(new MicrosoftVmsMatch(
                        VmsIdentifier: vmsIdentifier,
                        WorkerName: workerName,
                        AggregateInvoicedNet: invoicedNet,
                        Hours: invoicedUnits,
                        RtRate: GetDecimalValue(worksheet.Cell(row, rtRateCol)),
                        OtRate: GetDecimalValue(worksheet.Cell(row, otRateCol)),
                        DtRate: GetDecimalValue(worksheet.Cell(row, dtRateCol))
                    ));
                }
                else
                {
                    var existing = matches[key][0];

                    matches[key][0] = existing with
                    {
                        AggregateInvoicedNet = existing.AggregateInvoicedNet + invoicedNet,
                        Hours = existing.Hours + invoicedUnits
                    };
                }
            }

            return matches;
        }

        private static string FormatName(string input)
        {
            input = input.Trim();

            if (!input.Contains(","))
                return input;

            string[] parts = input.Split(',', 2);

            string last = ToTitle(parts[0].Trim());
            string first = ToTitle(parts[1].Trim());

            return $"{first} {last}".Trim();
        }

        private static string ToTitle(string value)
        {
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value.ToLower());
        }

        private static int FindHeaderRow(IXLWorksheet worksheet, string requiredHeader)
        {
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 100;

            for (int row = 1; row <= lastRow; row++)
            {
                for (int col = 1; col <= 100; col++)
                {
                    if (worksheet.Cell(row, col).GetString().Trim()
                        .Equals(requiredHeader, StringComparison.OrdinalIgnoreCase))
                    {
                        return row;
                    }
                }
            }

            return -1;
        }

        private static int FindColumn(IXLWorksheet worksheet, int headerRow, string headerName)
        {
            int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 100;

            for (int col = 1; col <= lastColumn; col++)
            {
                if (worksheet.Cell(headerRow, col).GetString().Trim()
                    .Equals(headerName, StringComparison.OrdinalIgnoreCase))
                {
                    return col;
                }
            }

            throw new Exception($"Could not find required VMS column: {headerName}");
        }

        private static DateTime GetDateValue(IXLCell cell)
        {
            if (cell.Value.IsDateTime)
                return cell.GetDateTime();

            return DateTime.TryParse(cell.GetString().Trim(), out DateTime parsed)
                ? parsed
                : DateTime.MinValue;
        }

        private static decimal GetDecimalValue(IXLCell cell)
        {
            string raw = cell.Value.ToString()
                .Replace("$", "")
                .Replace(",", "")
                .Replace("(", "-")
                .Replace(")", "")
                .Trim();

            return decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result)
                ? result
                : 0;
        }
    }
}