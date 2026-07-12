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
            // Validate inputs
            if (string.IsNullOrWhiteSpace(trackerPath) || !File.Exists(trackerPath))
                return null;

            // Check if the file is locked
            using var workbook = new XLWorkbook(trackerPath);

            // Check if the "Beeline FF" worksheet exists
            if (!workbook.TryGetWorksheet("Beeline FF", out var worksheet))
                return null;

            // Find the header row containing "Beeline ID"
            int headerRow = FindHeaderRow(worksheet, "Beeline ID");

            // If the header row is not found, return null
            if (headerRow == -1)
                return null;

            // Find the column indices for the required headers
            int beelineIdColumn = FindColumn(worksheet, headerRow, "Beeline ID");
            int submissionDateColumn = FindColumn(worksheet, headerRow, "Submission Date");
            int beelineAmountColumn = FindColumn(worksheet, headerRow, "Beeline Amt");
            int clientProjectNameColumn = FindColumn(worksheet, headerRow, "Client Project Name");

            // If any of the required columns are not found, return null
            if (beelineIdColumn == -1 ||
                submissionDateColumn == -1 ||
                beelineAmountColumn == -1 ||
                clientProjectNameColumn == -1)
            {
                return null;
            }

            // Parse the randstadWeekEnding string into a DateTime object
            if (!DateTime.TryParse(randstadWeekEnding, out DateTime targetDate))
                return null;

            // Get the last used row in the worksheet
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            // Iterate through the rows and find candidates that match the Beeline ID
            var candidates = new List<NikeTrackerMatch>();

            // Iterate through the rows and find candidates that match the Beeline ID
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string currentBeelineId = worksheet.Cell(row, beelineIdColumn).GetString().Trim();

                if (!currentBeelineId.Equals(beelineId, StringComparison.OrdinalIgnoreCase))
                    continue;

                // Extract the relevant data from the row
                DateTime? submissionDate = GetDateValue(worksheet.Cell(row, submissionDateColumn));
                decimal beelineAmount = GetDecimalValue(worksheet.Cell(row, beelineAmountColumn));
                string clientProjectName = worksheet.Cell(row, clientProjectNameColumn).GetString().Trim();

                // If any of the required fields are missing, skip this row
                if (submissionDate == null || string.IsNullOrWhiteSpace(clientProjectName))
                    continue;

                // Add the candidate to the list
                candidates.Add(new NikeTrackerMatch
                {
                    BeelineId = currentBeelineId,
                    SubmissionDate = submissionDate.Value,
                    BeelineAmount = beelineAmount,
                    ClientProjectName = clientProjectName
                });
            }

            // If no candidates were found, return null
            if (candidates.Count == 0)
                return null;

            // Find the two closest dates to the target date
            var twoClosestDates = candidates
                .OrderBy(x => Math.Abs((x.SubmissionDate - targetDate).TotalDays))
                .Take(2)
                .ToList();

            // Return the candidate with the closest Beeline Amount to the aggregateAmountPaid
            return twoClosestDates
                .OrderBy(x => Math.Abs(x.BeelineAmount - aggregateAmountPaid))
                .FirstOrDefault();
        }

        // Helper function to find the row number of the header row that contains the specified requiredHeader string
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

        // Helper function to find the column index of the specified headerName in the given headerRow
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

        // Helper function to parse the cell value as a DateTime. If successful, it returns the DateTime; otherwise, it returns null.
        private static DateTime? GetDateValue(IXLCell cell)
        {
            if (cell.Value.IsDateTime)
                return cell.GetDateTime();

            string value = cell.GetString().Trim();

            if (DateTime.TryParse(value, out DateTime parsedDate))
                return parsedDate;

            return null;
        }

        // Helper function to parse the cell value as a decimal. If successful, it returns the decimal; otherwise, it returns 0.
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

    // This class represents a match found in the Nike Tracker data, containing the Beeline ID, submission date, Beeline amount, and client project name.
    internal class NikeTrackerMatch
    {
        public string BeelineId { get; set; } = "";
        public DateTime SubmissionDate { get; set; }
        public decimal BeelineAmount { get; set; }
        public string ClientProjectName { get; set; } = "";
    }
}