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
                Supported examples:

                Brett Norton December 2025
                Emma Williams Bonus Dec. 2025
                Hany Selim December 2025 180/184 Hours
                Mital Patel January 2026 176 Hours

                Stacey Davis WE 10/4/26
                Stacey Davis WE 12.13.25
                Stacey Davis 10.11.26

                Ronald Eglentowicz 154.5 hours July 2023
                John Wadkins 4 OT hours July 2023
                Adam Vanderwalker worked 40 hours WE 3/23
                Sean Peterson Expense
            */

            string months =
                @"January|February|March|April|May|June|" +
                @"July|August|September|October|November|December|" +
                @"Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec";

            string[] patterns =
            {
        // Emma Williams Bonus Dec. 2025
        // Capture the name BEFORE the word "Bonus".
        $@"^(?<name>.+?)\s+Bonus\s+(?:{months})\.?\s+\d{{4}}\b",

        // Brett Norton December 2025
        // Hany Selim December 2025 180/184 Hours
        $@"^(?<name>.+?)\s+(?:{months})\.?\s+\d{{4}}\b",

        // Stacey Davis WE 10/4/26
        // Stacey Davis WE 12.13.25
        @"^(?<name>.+?)\s+WE\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b",

        // Stacey Davis 10.11.26
        // Stacey Davis 11.15.25
        @"^(?<name>.+?)\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b",

        // Adam Vanderwalker worked 40 hours WE 3/23
        @"^(?<name>.+?)\s+worked\s+\d+(?:\.\d+)?\s+hours?\b",

        // Ronald Eglentowicz 154.5 hours July 2023
        @"^(?<name>.+?)\s+\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?)?\s+hours?\b",

        // John Wadkins 4 OT hours July 2023
        // John Wadkins 155 ST hours July 2023
        @"^(?<name>.+?)\s+\d+(?:\.\d+)?\s+(?:OT|ST|Straight\s+Time)\s+hours?\b",

        // Sean Peterson Expense
        @"^(?<name>.+?)\s+Expense\b"
    };

            foreach (string pattern in patterns)
            {
                Match match = Regex.Match(
                    value,
                    pattern,
                    RegexOptions.IgnoreCase);

                if (match.Success)
                {
                    return ToTitleCase(
                        match.Groups["name"]
                            .Value
                            .Trim());
                }
            }

            // No recognized contractor-name pattern.
            // Return blank so the payment row receives
            // your light-red "needs review" highlighting.
            return "";
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