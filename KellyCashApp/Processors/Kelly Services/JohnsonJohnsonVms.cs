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
                Examples this is designed for:

                Ronald Eglentowicz 154.5 hours July 2023
                John Wadkins 4 OT hours July 2023
                Sean Peterson Expense
                Adam Vanderwalker worked 40 hours WE 3/23
                Jennifer Lithgow August 6th - Sept 5th 2023
            */

            string months =
                @"January|February|March|April|May|June|" +
                @"July|August|September|Sept|October|November|December|" +
                @"Jan|Feb|Mar|Apr|Jun|Jul|Aug|Oct|Nov|Dec";

            string pattern =
                $@"^(?<name>.+?)" +
                $@"(?=" +

                    // Example:
                    // Ronald Eglentowicz 154.5 hours July 2023
                    // John Wadkins 4 OT hours July 2023
                    $@"\s+(?:worked\s+)?\d+(?:\.\d+)?(?:\s*/\s*\d+(?:\.\d+)?)?\s+(?:OT\s+|ST\s+|Straight\s+Time\s+)?hours?\b" +

                    // Example:
                    // Sean Peterson Expense
                    $@"|\s+Expense\b" +

                    // Example:
                    // Hany Selim December 2025 180/184 Hours
                    // Mital Patel January 2026 176 Hours
                    $@"|\s+(?:{months})\b" +

                    // Example:
                    // Stacey Davis WE 10/4/26
                    // Stacey Davis WE 12.13.25
                    $@"|\s+WE\s+\d{{1,2}}[./-]\d{{1,2}}[./-]\d{{2,4}}\b" +

                    // Example:
                    // Stacey Davis 10.11.26
                    // Stacey Davis 11.15.25
                    $@"|\s+\d{{1,2}}[./-]\d{{1,2}}[./-]\d{{2,4}}\b" +

                $@")";

            Match match = Regex.Match(
                value,
                pattern,
                RegexOptions.IgnoreCase);

            if (!match.Success)
                return "";

            return ToTitleCase(
                match.Groups["name"].Value.Trim());
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