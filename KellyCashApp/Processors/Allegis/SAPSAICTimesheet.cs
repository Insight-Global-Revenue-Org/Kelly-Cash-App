using ClosedXML.Excel;
using System.Globalization;

namespace KellyCashApp.Processors.Allegis
{
    internal static class SAPSAICTimesheet
    {
        public static List<SAPSAICTimesheetMatch> Import(string filePath)
        {
            var matches = new List<SAPSAICTimesheetMatch>();

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);

            // SAP/SAIC timesheet report headers are on row 2.
            const int headerRow = 2;

            int workerCol =
                FindColumn(
                    worksheet,
                    headerRow,
                    "Worker");

            int endDateCol =
                FindColumn(
                    worksheet,
                    headerRow,
                    "End Date");

            int rtRateCol =
                FindColumn(
                    worksheet,
                    headerRow,
                    "Bill Rate [ST/Hr]");

            int hoursCol =
                FindColumn(
                    worksheet,
                    headerRow,
                    "Time Sheet Billable Hours");

            int netAmountCol =
                FindColumn(
                    worksheet,
                    headerRow,
                    "Net Amount");

            if (workerCol == -1 ||
                endDateCol == -1 ||
                rtRateCol == -1 ||
                hoursCol == -1 ||
                netAmountCol == -1)
            {
                throw new Exception(
                    "Missing one or more required SAP/SAIC timesheet report columns.");
            }

            int lastRow =
                worksheet.LastRowUsed()?.RowNumber()
                ?? headerRow;

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string rawWorker =
                    worksheet.Cell(row, workerCol)
                        .GetString()
                        .Trim();

                if (string.IsNullOrWhiteSpace(rawWorker))
                    continue;

                DateTime endDate =
                    GetDateValue(
                        worksheet.Cell(row, endDateCol));

                if (endDate == DateTime.MinValue)
                    continue;

                string worker =
                    FormatName(rawWorker);

                matches.Add(
                    new SAPSAICTimesheetMatch
                    {
                        ContractorName = worker,
                        WeekEndingDate = endDate.Date,

                        Hours =
                            GetDecimalValue(
                                worksheet.Cell(row, hoursCol)),

                        RtRate =
                            GetDecimalValue(
                                worksheet.Cell(row, rtRateCol)),

                        AggregateInvoicedNet =
                            GetDecimalValue(
                                worksheet.Cell(row, netAmountCol))
                    });
            }

            return matches;
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

            // If report gives Last, First,
            // convert to First Last.
            if (!input.Contains(","))
                return input;

            string[] parts =
                input.Split(',', 2);

            string last =
                CultureInfo.CurrentCulture.TextInfo
                    .ToTitleCase(
                        parts[0].Trim().ToLower());

            string first =
                CultureInfo.CurrentCulture.TextInfo
                    .ToTitleCase(
                        parts[1].Trim().ToLower());

            return $"{first} {last}".Trim();
        }

        private static DateTime GetDateValue(
            IXLCell cell)
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

        private static decimal GetDecimalValue(
            IXLCell cell)
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

    internal class SAPSAICTimesheetMatch
    {
        public string ContractorName { get; set; } = "";

        public DateTime WeekEndingDate { get; set; }

        public decimal Hours { get; set; }

        public decimal RtRate { get; set; }

        public decimal AggregateInvoicedNet { get; set; }
    }
}