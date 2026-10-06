using ClosedXML.Excel;
using System.Globalization;

namespace KellyCashApp.Processors.Allegis
{
    internal static class AllegisTimesheet
    {
        public static List<AllegisTimesheetMatch> Import(string filePath)
        {
            var matches = new List<AllegisTimesheetMatch>();

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);

            int headerRow = FindHeaderRow(
                worksheet,
                "Last Name, First Name Middle Name");

            if (headerRow == -1)
                throw new Exception(
                    "Could not find timesheet header: " +
                    "Last Name, First Name Middle Name");

            int nameCol =
                FindColumn(
                    worksheet,
                    headerRow,
                    "Last Name, First Name Middle Name");

            int dateCol =
                FindColumn(
                    worksheet,
                    headerRow,
                    "Date");

            int hoursCol =
                FindColumn(
                    worksheet,
                    headerRow,
                    "Units");

            int rtRateCol =
                FindColumn(
                    worksheet,
                    headerRow,
                    "RT Rate");

            int otRateCol =
                FindColumn(
                    worksheet,
                    headerRow,
                    "OT Rate");

            int dtRateCol =
                FindColumn(
                    worksheet,
                    headerRow,
                    "DT Rate");

            int netCol =
                FindColumn(
                    worksheet,
                    headerRow,
                    "Estimated Net Supplier Payable");

            if (nameCol == -1 ||
                dateCol == -1 ||
                hoursCol == -1 ||
                rtRateCol == -1 ||
                otRateCol == -1 ||
                dtRateCol == -1 ||
                netCol == -1)
            {
                throw new Exception(
                    "Missing one or more required timesheet report columns.");
            }

            int lastRow =
                worksheet.LastRowUsed()?.RowNumber()
                ?? headerRow;

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string rawName =
                    worksheet.Cell(row, nameCol)
                        .GetString()
                        .Trim();

                if (string.IsNullOrWhiteSpace(rawName))
                    continue;

                DateTime date =
                    GetDateValue(
                        worksheet.Cell(row, dateCol));

                if (date == DateTime.MinValue)
                    continue;

                string name = FormatName(rawName);

                matches.Add(
                    new AllegisTimesheetMatch
                    {
                        ContractorName = name,
                        WeekEndingDate = date.Date,

                        Hours =
                            GetDecimalValue(
                                worksheet.Cell(row, hoursCol)),

                        RtRate =
                            GetDecimalValue(
                                worksheet.Cell(row, rtRateCol)),

                        OtRate =
                            GetDecimalValue(
                                worksheet.Cell(row, otRateCol)),

                        DtRate =
                            GetDecimalValue(
                                worksheet.Cell(row, dtRateCol)),

                        AggregateInvoicedNet =
                            GetDecimalValue(
                                worksheet.Cell(row, netCol))
                    });
            }

            return matches;
        }

        private static int FindHeaderRow(
            IXLWorksheet worksheet,
            string requiredHeader)
        {
            int lastRow =
                Math.Min(
                    worksheet.LastRowUsed()?.RowNumber() ?? 20,
                    20);

            for (int row = 1; row <= lastRow; row++)
            {
                if (FindColumn(
                    worksheet,
                    row,
                    requiredHeader) != -1)
                {
                    return row;
                }
            }

            return -1;
        }

        private static int FindColumn(
            IXLWorksheet worksheet,
            int headerRow,
            string headerName)
        {
            int lastColumn =
                worksheet.LastColumnUsed()?.ColumnNumber()
                ?? 100;

            for (int col = 1; col <= lastColumn; col++)
            {
                string header =
                    worksheet.Cell(headerRow, col)
                        .GetString()
                        .Trim();

                if (header.Equals(
                    headerName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return col;
                }
            }

            return -1;
        }

        private static string FormatName(string input)
        {
            input = input.Trim();

            if (!input.Contains(","))
                return input;

            string[] parts = input.Split(',', 2);

            string last =
                CultureInfo.CurrentCulture.TextInfo
                    .ToTitleCase(parts[0].Trim().ToLower());

            string first =
                CultureInfo.CurrentCulture.TextInfo
                    .ToTitleCase(parts[1].Trim().ToLower());

            return $"{first} {last}".Trim();
        }

        private static DateTime GetDateValue(IXLCell cell)
        {
            if (cell.Value.IsDateTime)
                return cell.GetDateTime();

            if (DateTime.TryParse(
                cell.GetString().Trim(),
                out DateTime parsed))
            {
                return parsed;
            }

            return DateTime.MinValue;
        }

        private static decimal GetDecimalValue(IXLCell cell)
        {
            string raw =
                cell.Value.ToString()
                    .Replace("$", "")
                    .Replace(",", "")
                    .Replace("(", "-")
                    .Replace(")", "")
                    .Trim();

            if (decimal.TryParse(
                raw,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out decimal result))
            {
                return result;
            }

            return 0;
        }
    }

    internal class AllegisTimesheetMatch
    {
        public string ContractorName { get; set; } = "";

        public DateTime WeekEndingDate { get; set; }

        public decimal Hours { get; set; }

        public decimal RtRate { get; set; }

        public decimal OtRate { get; set; }

        public decimal DtRate { get; set; }

        public decimal AggregateInvoicedNet { get; set; }
    }
}