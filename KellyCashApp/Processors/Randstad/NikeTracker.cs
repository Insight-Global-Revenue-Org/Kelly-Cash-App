using ClosedXML.Excel;
using System.Globalization;

namespace KellyCashApp.Processors.Randstad
{
    internal class NikeTracker
    {
        public static NikeTrackerMatch? FindBestMatch(
            string trackerPath,
            string beelineId,
            string randstadWeekEnding,
            decimal aggregateAmountPaid)
        {
            if (string.IsNullOrWhiteSpace(trackerPath) || !File.Exists(trackerPath))
                return null;

            using var workbook = new XLWorkbook(trackerPath);

            if (!workbook.TryGetWorksheet("Beeline FF", out var worksheet))
                return null;

            int headerRow = FindHeaderRow(worksheet, "Beeline ID");

            if (headerRow == -1)
                return null;

            int beelineIdColumn = FindColumn(worksheet, headerRow, "Beeline ID");
            int submissionDateColumn = FindColumn(worksheet, headerRow, "Submission Date");
            int beelineAmountColumn = FindColumn(worksheet, headerRow, "Beeline Amt");
            int clientProjectNameColumn = FindColumn(worksheet, headerRow, "Client Project Name");

            if (beelineIdColumn == -1 ||
                submissionDateColumn == -1 ||
                beelineAmountColumn == -1 ||
                clientProjectNameColumn == -1)
            {
                return null;
            }

            if (!DateTime.TryParse(randstadWeekEnding, out DateTime targetDate))
                return null;

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            var candidates = new List<NikeTrackerMatch>();

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string currentBeelineId = worksheet.Cell(row, beelineIdColumn).GetString().Trim();

                if (!currentBeelineId.Equals(beelineId, StringComparison.OrdinalIgnoreCase))
                    continue;

                DateTime? submissionDate = GetDateValue(worksheet.Cell(row, submissionDateColumn));
                decimal beelineAmount = GetDecimalValue(worksheet.Cell(row, beelineAmountColumn));
                string clientProjectName = worksheet.Cell(row, clientProjectNameColumn).GetString().Trim();

                if (submissionDate == null || string.IsNullOrWhiteSpace(clientProjectName))
                    continue;

                candidates.Add(new NikeTrackerMatch
                {
                    BeelineId = currentBeelineId,
                    SubmissionDate = submissionDate.Value,
                    BeelineAmount = beelineAmount,
                    ClientProjectName = clientProjectName
                });
            }

            if (candidates.Count == 0)
                return null;

            var twoClosestDates = candidates
                .OrderBy(x => Math.Abs((x.SubmissionDate - targetDate).TotalDays))
                .Take(2)
                .ToList();

            return twoClosestDates
                .OrderBy(x => Math.Abs(x.BeelineAmount - aggregateAmountPaid))
                .FirstOrDefault();
        }

        private static int FindHeaderRow(IXLWorksheet worksheet, string requiredHeader)
        {
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 100;

            for (int row = 1; row <= lastRow; row++)
            {
                for (int col = 1; col <= 100; col++)
                {
                    string value = worksheet.Cell(row, col).GetString().Trim();

                    if (value.Equals(requiredHeader, StringComparison.OrdinalIgnoreCase))
                        return row;
                }
            }

            return -1;
        }

        private static int FindColumn(IXLWorksheet worksheet, int headerRow, string headerName)
        {
            for (int col = 1; col <= 100; col++)
            {
                string value = worksheet.Cell(headerRow, col).GetString().Trim();

                if (value.Equals(headerName, StringComparison.OrdinalIgnoreCase))
                    return col;
            }

            return -1;
        }

        private static DateTime? GetDateValue(IXLCell cell)
        {
            if (cell.Value.IsDateTime)
                return cell.GetDateTime();

            string value = cell.GetString().Trim();

            if (DateTime.TryParse(value, out DateTime parsedDate))
                return parsedDate;

            return null;
        }

        private static decimal GetDecimalValue(IXLCell cell)
        {
            string raw = cell.Value.ToString()
                .Replace("$", "")
                .Replace(",", "")
                .Trim();

            if (decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return 0;
        }
    }

    internal class NikeTrackerMatch
    {
        public string BeelineId { get; set; } = "";
        public DateTime SubmissionDate { get; set; }
        public decimal BeelineAmount { get; set; }
        public string ClientProjectName { get; set; } = "";
    }
}