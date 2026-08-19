using ClosedXML.Excel;
using KellyCashApp.Models;
using System.Text.RegularExpressions;

namespace KellyCashApp.Processors.Kelly_Services
{
    internal static class JohnsonJohnsonVms
    {
        public static Dictionary<string, JohnsonJohnsonVmsMatch> Import(
            string filePath)
        {
            using var stream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite);

            using var workbook = new XLWorkbook(stream);

            var worksheet = workbook.Worksheet(1);

            int headerRow = FindHeaderRow(worksheet);

            int invoiceIdCol =
                FindColumn(worksheet, headerRow, "Invoice ID");

            int feeDescriptionCol =
                FindColumn(worksheet, headerRow, "Fee Description");

            var matches =
                new Dictionary<string, JohnsonJohnsonVmsMatch>(
                    StringComparer.OrdinalIgnoreCase);

            int lastRow =
                worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string invoiceId =
                    worksheet.Cell(row, invoiceIdCol)
                        .GetString()
                        .Trim();

                string feeDescription =
                    worksheet.Cell(row, feeDescriptionCol)
                        .GetString()
                        .Trim();

                if (string.IsNullOrWhiteSpace(invoiceId) ||
                    string.IsNullOrWhiteSpace(feeDescription))
                {
                    continue;
                }

                string workerName =
                ExtractWorkerName(feeDescription);

                // IMPORTANT:
                // Even if we could not parse a contractor name,
                // still keep the VMS record so Fee Description
                // can be written to Notes.
                matches[invoiceId] =
                    new JohnsonJohnsonVmsMatch(
                        InvoiceId: invoiceId,
                        WorkerName: workerName,
                        FeeDescription: feeDescription
                    );
            }

            return matches;
        }

        private static int FindHeaderRow(
            IXLWorksheet worksheet)
        {
            // First try row 1.
            if (FindColumn(
                worksheet,
                1,
                "Invoice ID",
                throwIfMissing: false) != -1)
            {
                return 1;
            }

            // If Invoice ID was not on row 1,
            // try row 2.
            if (FindColumn(
                worksheet,
                2,
                "Invoice ID",
                throwIfMissing: false) != -1)
            {
                return 2;
            }

            throw new Exception(
                "Could not find the 'Invoice ID' header " +
                "on row 1 or row 2 of the Johnson & Johnson VMS report.");
        }

        private static int FindColumn(
            IXLWorksheet worksheet,
            int headerRow,
            string headerName,
            bool throwIfMissing = true)
        {
            int lastColumn =
                worksheet.LastColumnUsed()?.ColumnNumber() ?? 100;

            for (int col = 1; col <= lastColumn; col++)
            {
                string value =
                    worksheet.Cell(headerRow, col)
                        .GetString()
                        .Trim();

                if (value.Equals(
                    headerName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return col;
                }
            }

            if (throwIfMissing)
            {
                throw new Exception(
                    $"Could not find required J&J VMS column: {headerName}");
            }

            return -1;
        }

        private static string ExtractWorkerName(
            string feeDescription)
        {
            if (string.IsNullOrWhiteSpace(feeDescription))
                return "";

            string value = Regex.Replace(
                feeDescription.Trim(),
                @"\s+",
                " ");

            /*
                Examples handled:

                Brett Norton December 2025
                Emma Williams Bonus Dec. 2025
                Hany Selim December 2025 180/184 Hours
                Zainab Khan January 2026 176 Hours

                Chloe Oxley 176 Hours December 2025
                Shannell Banks 340 Hours January 2026
                Brittany Winters 90 Hours OT January 2026
                Brittany Winters 35.33 Hours Ot October 2025
                Jeremy Verwey December 2025 167.49/184 ST Hours

                Stacey Davis WE 10/4/26
                Stacey Davis WE 12.13.25
                Stacey Davis 10.11.26

                Adam Vanderwalker worked 40 hours WE 3/23
                Isabel Henao Worked 40 Hours W.E. 2.28.26

                Sean Peterson Expense
                Kevin Tseng Jan Expenses 1/18/2026 - 1/24/2026
            */

            string months =
                @"January|February|March|April|May|June|" +
                @"July|August|September|October|November|December|" +
                @"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec";

            string pattern =
                $@"^(?<name>.+?)" +
                $@"(?=" +

                    // Emma Williams Bonus Dec. 2025
                    // Medhat Kamel Bonus
                    $@"\s+Bonus\b" +

                    // Adam Vanderwalker worked 40 hours
                    $@"|\s+worked\s+\d+(?:\.\d+)?\s+hours?\b" +

                    // John Wadkins 4 OT hours
                    // Jeremy Verwey 167.49/184 ST Hours
                    $@"|\s+\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?)?\s+(?:OT|ST|Straight\s+Time)\s+hours?\b" +

                    // Brittany Winters 35.33 Hours OT
                    $@"|\s+\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?)?\s+hours?\s+(?:OT|ST)\b" +

                    // Chloe Oxley 176 Hours
                    // Hany Selim 180/184 Hours
                    $@"|\s+\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?)?\s+hours?\b" +

                    // Sean Peterson Expense
                    $@"|\s+Expenses?\b" +

                    // Kevin Tseng Jan Expenses
                    $@"|\s+(?:{months})\.?\s+Expenses?\b" +

                    // Brett Norton December 2025
                    // Jeremy Verwey December 2025 ...
                    $@"|\s+(?:{months})\.?\s+\d{{4}}\b" +

                    // Stacey Davis WE 10/4/26
                    // Isabel Henao W.E. 2.28.26
                    $@"|\s+(?:WE|W\.E\.)\s+\d{{1,2}}[./-]\d{{1,2}}[./-]\d{{2,4}}\b" +

                    // Stacey Davis 10.11.26
                    $@"|\s+\d{{1,2}}[./-]\d{{1,2}}[./-]\d{{2,4}}\b" +

                $@")";

            Match match = Regex.Match(
                value,
                pattern,
                RegexOptions.IgnoreCase);

            if (!match.Success)
                return "";

            return ToTitleCase(
                match.Groups["name"]
                    .Value
                    .Trim());
        }

        private static string ToTitleCase(string value)
        {
            return System.Globalization.CultureInfo
                .CurrentCulture
                .TextInfo
                .ToTitleCase(value.ToLower());
        }
    }
}