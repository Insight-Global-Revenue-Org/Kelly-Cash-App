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
    // ---------------------------------------------------------
    // BONUS
    //
    // Emma Williams Bonus Dec. 2025
    // Medhat Kamel Bonus
    // ---------------------------------------------------------

    $@"^(?<name>.+?)\s+Bonus(?:\s+(?:{months})\.?\s+\d{{4}})?\b",

    // ---------------------------------------------------------
    // WORKED HOURS
    //
    // Adam Vanderwalker worked 40 hours WE 3/23
    // Isabel Henao Worked 40 Hours W.E. 2.28.26
    // ---------------------------------------------------------

    @"^(?<name>.+?)\s+worked\s+\d+(?:\.\d+)?\s+hours?\b",

    // ---------------------------------------------------------
    // OT / ST / STRAIGHT TIME
    //
    // John Wadkins 4 OT hours July 2023
    // Brittany Winters 90 Hours OT January 2026
    // Jeremy Verwey 167.49/184 ST Hours
    // ---------------------------------------------------------

    @"^(?<name>.+?)\s+\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?)?\s+(?:OT|ST|Straight\s+Time)\s+hours?\b",

    // Some rows put OT/ST AFTER "Hours":
    // Brittany Winters 35.33 Hours Ot October 2025
    @"^(?<name>.+?)\s+\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?)?\s+hours?\s+(?:OT|ST)\b",

    // ---------------------------------------------------------
    // NORMAL HOURS
    //
    // Chloe Oxley 176 Hours December 2025
    // Shannell Banks 340 Hours January 2026
    // Laura Clagett 168 Hours December 2025
    // Hany Selim 180/184 Hours
    // ---------------------------------------------------------

    @"^(?<name>.+?)\s+\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?)?\s+hours?\b",

    // ---------------------------------------------------------
    // EXPENSE / EXPENSES
    //
    // Sean Peterson Expense
    // ---------------------------------------------------------

    @"^(?<name>.+?)\s+Expenses?\b",

    // Month before "Expense(s)"
    // Kevin Tseng Jan Expenses 1/18/2026 - 1/24/2026
    $@"^(?<name>.+?)\s+(?:{months})\.?\s+Expenses?\b",

    // ---------------------------------------------------------
    // MONTH + YEAR
    //
    // Brett Norton December 2025
    // Tejas Pawar November 2025
    // Zainab Khan January 2026 176 Hours
    //
    // IMPORTANT: This comes AFTER the hours patterns.
    // ---------------------------------------------------------

    $@"^(?<name>.+?)\s+(?:{months})\.?\s+\d{{4}}\b",

    // ---------------------------------------------------------
    // WE / W.E. + DATE
    //
    // Stacey Davis WE 10/4/26
    // Stacey Davis W.E. 12.13.25
    // ---------------------------------------------------------

    @"^(?<name>.+?)\s+(?:WE|W\.E\.)\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b",

    // ---------------------------------------------------------
    // PLAIN DATE
    //
    // Stacey Davis 10.11.26
    //
    // Keep this LAST because it's the broadest rule.
    // ---------------------------------------------------------

    @"^(?<name>.+?)\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b"
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