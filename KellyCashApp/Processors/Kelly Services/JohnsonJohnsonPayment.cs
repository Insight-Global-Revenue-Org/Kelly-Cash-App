using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Services;
using System.Globalization;
using System.Text.RegularExpressions;

namespace KellyCashApp.Processors.Kelly_Services
{
    internal static class JohnsonJohnsonPayment
    {
        private const int HeaderRow = 1;
        private const int FirstDataRow = 2;

        // Detects a Johnson & Johnson payment based on its fixed row-1 headers.
        public static bool IsJohnsonJohnsonFormat(IXLWorksheet worksheet)
        {
            int invoiceIdCol = FindColumn(
                worksheet,
                HeaderRow,
                "InvoiceID");

            int weekEndingCol = FindColumn(
                worksheet,
                HeaderRow,
                "WeekEndingDate");

            int itemGroupCol = FindColumn(
                worksheet,
                HeaderRow,
                "ItemGroup");

            int taxCol = FindColumn(
                worksheet,
                HeaderRow,
                "Tax");

            int lineTotalCol = FindColumn(
                worksheet,
                HeaderRow,
                "LineTotal");

            return invoiceIdCol != -1
                && weekEndingCol != -1
                && itemGroupCol != -1
                && taxCol != -1
                && lineTotalCol != -1;
        }

        public static string Process(
            XLWorkbook workbook,
            IXLWorksheet worksheet,
            string inputPath,
            Dictionary<string, List<OirMatch>> openInvoiceMatches,
            Dictionary<string, JohnsonJohnsonVmsMatch>? vmsMatches = null)
        {
            // ---------------------------------------------------------
            // Find required Johnson & Johnson payment columns.
            // Headers are always on row 1.
            // ---------------------------------------------------------

            int invoiceIdCol = FindColumn(
                worksheet,
                HeaderRow,
                "InvoiceID");

            int weekEndingCol = FindColumn(
                worksheet,
                HeaderRow,
                "WeekEndingDate");

            int itemGroupCol = FindColumn(
                worksheet,
                HeaderRow,
                "ItemGroup");

            int taxCol = FindColumn(
                worksheet,
                HeaderRow,
                "Tax");

            int lineTotalCol = FindColumn(
                worksheet,
                HeaderRow,
                "LineTotal");

            if (invoiceIdCol == -1 ||
                weekEndingCol == -1 ||
                itemGroupCol == -1 ||
                taxCol == -1 ||
                lineTotalCol == -1)
            {
                throw new Exception(
                    "Missing one or more required Johnson & Johnson payment columns.");
            }

            // Convert the OIR dictionary into individual rows
            // so we can search by name/date/amount.
            var oirRows = BuildOirRows(openInvoiceMatches);

            var outputRows = new List<JohnsonJohnsonOutputRow>();

            int lastRow =
                worksheet.LastRowUsed()?.RowNumber()
                ?? FirstDataRow;

            // Reuse your existing contractor name-change system.
            var nameChanges = Rename.LoadNameChanges();

            // ---------------------------------------------------------
            // Process each J&J payment row.
            // ---------------------------------------------------------

            for (int row = FirstDataRow;
                 row <= lastRow;
                 row++)
            {
                string invoiceId =
                    worksheet.Cell(row, invoiceIdCol)
                        .GetString()
                        .Trim();

                // Ignore completely blank rows.
                if (string.IsNullOrWhiteSpace(invoiceId))
                    continue;

                DateTime weekEndingDate =
                    GetDateValue(
                        worksheet.Cell(row, weekEndingCol));

                if (weekEndingDate == DateTime.MinValue)
                    continue;

                string formattedWeekEndingDate =
                    weekEndingDate.ToString(
                        "MM/dd/yyyy",
                        CultureInfo.InvariantCulture);

                string item =
                    worksheet.Cell(row, itemGroupCol)
                        .GetString()
                        .Trim();

                decimal tax =
                    GetDecimalValue(
                        worksheet.Cell(row, taxCol));

                decimal aggregateAmountPaid =
                    GetDecimalValue(
                        worksheet.Cell(row, lineTotalCol));

                // -----------------------------------------------------
                // Find contractor name through the J&J VMS report.
                //
                // Payment:
                // InvoiceID
                //
                //          ↓
                //
                // VMS:
                // Invoice ID -> Fee Description -> WorkerName
                // -----------------------------------------------------

                string name = "";
                string feeDescription = "";
                DateTime? feeMonth = null;

                if (vmsMatches != null &&
                    vmsMatches.TryGetValue(
                    invoiceId,
                        out JohnsonJohnsonVmsMatch? vmsMatch))
                {
                    name = vmsMatch.WorkerName;

                    feeDescription = vmsMatch.FeeDescription;

                    feeMonth = vmsMatch.FeeMonth;

                    name = Rename.ApplyNameChange(
                        name,
                        nameChanges);
                }

                string concat =
                    string.IsNullOrWhiteSpace(name)
                        ? ""
                        : $"{name} {formattedWeekEndingDate}";

                // -----------------------------------------------------
                // Build the possible OIR invoice pool.
                //
                // WeekEndingDate does NOT affect invoice matching.
                //
                // Once the contractor name is recovered from the
                // Johnson & Johnson VMS report, consider every OIR
                // invoice belonging to that contractor.
                // -----------------------------------------------------

                var possibleMatches =
                    string.IsNullOrWhiteSpace(name)
                        ? new List<OirLookupRow>()
                        : oirRows
                            .Where(x =>
                                x.Name.Equals(
                                    name,
                                    StringComparison.OrdinalIgnoreCase))
                            .ToList();

                // -----------------------------------------------------
                // If the VMS Fee Description gave us a month/year,
                // only consider OIR invoices from that same month/year.
                //
                // If no month/year was parsed, leave possibleMatches
                // unchanged so the old amount-only behavior is used.
                // -----------------------------------------------------

                if (feeMonth.HasValue)
                {
                    possibleMatches =
                        possibleMatches
                            .Where(x =>
                                x.WeekEndingDate.Month ==
                                    feeMonth.Value.Month
                                &&
                                x.WeekEndingDate.Year ==
                                    feeMonth.Value.Year)
                            .ToList();
                }

                // -----------------------------------------------------
                // Find ONE OIR invoice.
                //
                // Select whichever Amount Due is closest to LineTotal,
                // but only accept it if the difference is within 5%.
                //
                // WeekEndingDate does not affect matching.
                // -----------------------------------------------------

                OirLookupRow? bestMatch =
                    FindClosestInvoiceMatch(
                        possibleMatches,
                        aggregateAmountPaid);

                if (bestMatch != null)
                {
                    outputRows.Add(
                        new JohnsonJohnsonOutputRow
                        {
                            WeekEndingDate =
                                formattedWeekEndingDate,

                            Name =
                                name,

                            Invoice =
                                bestMatch.Invoice,

                            AmountDue =
                                bestMatch.AmountDue,

                            AggregateAmountPaid =
                                aggregateAmountPaid,

                            Tax =
                                tax,

                            // Always preserve the complete VMS Fee Description.
                            Notes =
                                feeDescription,

                            Item =
                                item,

                            Concat =
                                concat,

                            InvoiceNumber =
                                invoiceId
                        });
                }
                else
                {
                    outputRows.Add(
                        new JohnsonJohnsonOutputRow
                        {
                            WeekEndingDate =
                                formattedWeekEndingDate,

                            Name =
                                name,

                            Invoice =
                                "",

                            AmountDue =
                                0,

                            AggregateAmountPaid =
                                aggregateAmountPaid,

                            Tax =
                                tax,

                            // If the VMS lookup succeeded, always retain
                            // whatever Fee Description was present —
                            // even if no contractor name could be parsed.
                            Notes =
                                feeDescription,

                            Item =
                                item,

                            Concat =
                                concat,

                            InvoiceNumber =
                                invoiceId
                        });
                }
            }

            // ---------------------------------------------------------
            // Capture total before clearing the original worksheet.
            // ---------------------------------------------------------

            decimal total =
                outputRows.Sum(
                    x => x.AggregateAmountPaid);

            // Remove pictures/shapes before rebuilding the worksheet.
            foreach (var picture in worksheet.Pictures.ToList())
            {
                picture.Delete();
            }

            worksheet.Clear(XLClearOptions.All);

            worksheet.Style.Fill
                .SetBackgroundColor(
                    XLColor.NoColor);

            // ---------------------------------------------------------
            // Final Johnson & Johnson output columns.
            // ---------------------------------------------------------

            string[] headers =
            {
                "Week Ending Date",
                "Name",
                "Invoice",
                "Amount Due",
                "Aggregate Amount Paid",
                "Tax",
                "Notes",
                "Item",
                "Concat",
                "Invoice Number"
            };

            for (int col = 1;
                 col <= headers.Length;
                 col++)
            {
                worksheet.Cell(1, col).Value =
                    headers[col - 1];
            }

            // ---------------------------------------------------------
            // Write output rows.
            // ---------------------------------------------------------

            for (int i = 0;
                 i < outputRows.Count;
                 i++)
            {
                int outputRow = i + 2;

                var item = outputRows[i];

                worksheet.Cell(outputRow, 1).Value =
                    item.WeekEndingDate;

                worksheet.Cell(outputRow, 2).Value =
                    item.Name;

                worksheet.Cell(outputRow, 3).Value =
                    item.Invoice;

                worksheet.Cell(outputRow, 4).Value =
                    item.AmountDue;

                worksheet.Cell(outputRow, 5).Value =
                    item.AggregateAmountPaid;

                worksheet.Cell(outputRow, 6).Value =
                    item.Tax;

                worksheet.Cell(outputRow, 7).Value =
                    item.Notes;

                worksheet.Cell(outputRow, 8).Value =
                    item.Item;

                worksheet.Cell(outputRow, 9).Value =
                    item.Concat;

                worksheet.Cell(outputRow, 10).Value =
                    item.InvoiceNumber;
            }

            ApplyFormatting(
                worksheet,
                outputRows.Count + 1,
                headers.Length);

            // ---------------------------------------------------------
            // Save final file.
            // ---------------------------------------------------------

            string downloadsPath =
                Settings.GetRemittanceSavePath();

            string formattedTotal =
                total.ToString(
                    "$#,##0.00;($#,##0.00)",
                    CultureInfo.InvariantCulture);

            string processedDate =
                DateTime.Now.ToString(
                    "M.d.yyyy",
                    CultureInfo.InvariantCulture);

            string outputPath =
                GetUniqueOutputPath(
                    downloadsPath,
                    $"Johnson Johnson {processedDate} - {formattedTotal}.xlsx");

            workbook.SaveAs(outputPath);

            Analytics.LogRemittanceRun(
                $"Johnson Johnson - {formattedTotal}");

            return outputPath;
        }

        // =============================================================
        // OIR HELPERS
        // =============================================================

        private static List<OirLookupRow> BuildOirRows(
            Dictionary<string, List<OirMatch>> openInvoiceMatches)
        {
            var rows = new List<OirLookupRow>();

            foreach (var item in openInvoiceMatches)
            {
                // OIR dictionary keys look like:
                // John Smith 08/15/2026
                //
                // Separate the contractor name from the date.
                Match match = Regex.Match(
                    item.Key,
                    @"^(?<name>.+)\s(?<date>\d{1,2}/\d{1,2}/\d{2,4})$");

                if (!match.Success)
                    continue;

                string name =
                    match.Groups["name"]
                        .Value
                        .Trim();

                if (!DateTime.TryParse(
                    match.Groups["date"].Value,
                    out DateTime weekEnding))
                {
                    continue;
                }

                foreach (var oirMatch in item.Value)
                {
                    rows.Add(
                        new OirLookupRow
                        {
                            Name = name,
                            WeekEndingDate = weekEnding,
                            Invoice = oirMatch.DocumentNumber,
                            AmountDue = oirMatch.RemainingAmount,
                            Concat = item.Key
                        });
                }
            }

            return rows;
        }

        // =============================================================
        // SINGLE INVOICE AMOUNT MATCHING
        // =============================================================

        private static OirLookupRow? FindClosestInvoiceMatch(
            List<OirLookupRow> possibleMatches,
            decimal targetAmount)
        {
            const decimal allowedPercentDifference = 0.02m;

            if (possibleMatches.Count == 0)
                return null;

            if (targetAmount == 0)
                return null;

            // Find the ONE OIR invoice whose Amount Due
            // is closest to the J&J LineTotal.
            OirLookupRow? bestMatch =
                possibleMatches
                    .OrderBy(x =>
                        Math.Abs(
                            x.AmountDue - targetAmount))
                    .FirstOrDefault();

            if (bestMatch == null)
                return null;

            decimal difference =
                Math.Abs(
                    bestMatch.AmountDue -
                    targetAmount);

            decimal allowedDifference =
                Math.Abs(
                    targetAmount *
                    allowedPercentDifference);

            // Closest invoice still has to fall inside 5%.
            if (difference > allowedDifference)
                return null;

            return bestMatch;
        }

        // =============================================================
        // EXCEL HELPERS
        // =============================================================

        private static int FindColumn(
            IXLWorksheet worksheet,
            int headerRow,
            string headerName)
        {
            int lastColumn =
                worksheet.LastColumnUsed()
                    ?.ColumnNumber()
                ?? 100;

            for (int col = 1;
                 col <= lastColumn;
                 col++)
            {
                string headerText =
                    worksheet
                        .Cell(headerRow, col)
                        .GetString()
                        .Trim();

                if (headerText.Equals(
                    headerName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return col;
                }
            }

            return -1;
        }

        private static DateTime GetDateValue(
            IXLCell cell)
        {
            if (cell.Value.IsDateTime)
            {
                return cell.GetDateTime();
            }

            string raw =
                cell.GetString()
                    .Trim();

            if (DateTime.TryParse(
                raw,
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
                cell.Value
                    .ToString()
                    .Replace("$", "")
                    .Replace(",", "")
                    .Replace("(", "-")
                    .Replace(")", "")
                    .Trim();

            return decimal.TryParse(
                raw,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out decimal result)
                    ? result
                    : 0;
        }

        // =============================================================
        // FORMATTING
        // =============================================================

        private static void ApplyFormatting(
            IXLWorksheet worksheet,
            int lastRow,
            int lastColumn)
        {
            var range =
                worksheet.Range(
                    1,
                    1,
                    lastRow,
                    lastColumn);

            range.Style.Font.FontName =
                "Aptos Narrow";

            range.Style.Font.FontSize =
                9;

            range.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            range.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            range.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            worksheet.Row(1)
                .Style.Font.Bold = true;

            worksheet.Row(1)
                .Style.Fill.BackgroundColor =
                XLColor.FromHtml("#FCE4D6");

            worksheet.Row(1).Height = 15;

            // Currency formatting.
            worksheet.Column(4)
                .Style.NumberFormat.Format =
                "$#,##0.00;($#,##0.00)";

            worksheet.Column(5)
                .Style.NumberFormat.Format =
                "$#,##0.00;($#,##0.00)";

            worksheet.Column(6)
                .Style.NumberFormat.Format =
                "$#,##0.00;($#,##0.00)";

            for (int row = 2;
                 row <= lastRow;
                 row++)
            {
                worksheet.Row(row).Height = 13;
            }

            worksheet.Columns()
                .AdjustToContents();

            // Fixed widths similar to Microsoft.
            worksheet.Column(1).Width = 18;
            worksheet.Column(2).Width = 24;
            worksheet.Column(3).Width = 18;
            worksheet.Column(4).Width = 18;
            worksheet.Column(5).Width = 24;
            worksheet.Column(6).Width = 14;
            worksheet.Column(7).Width = 38;
            worksheet.Column(8).Width = 24;
            worksheet.Column(9).Width = 32;
            worksheet.Column(10).Width = 20;

            // Don't let Notes become ridiculously tall.
            worksheet.Column(7)
                .Style.Alignment.WrapText = false;

            // ---------------------------------------------------------
            // Highlight anything that did not successfully match.
            // ---------------------------------------------------------

            for (int row = 2;
      row <= lastRow;
      row++)
            {
                string name =
                    worksheet.Cell(row, 2)
                        .GetString()
                        .Trim();

                string invoice =
                    worksheet.Cell(row, 3)
                        .GetString()
                        .Trim();

                string item =
                    worksheet.Cell(row, 8)
                        .GetString()
                        .Trim();

                decimal amountDue =
                    GetDecimalValue(
                        worksheet.Cell(row, 4));

                // ---------------------------------------------------------
                // RULE 1:
                // Anything in Item other than exactly "Fees"
                // gets the very light blue fill.
                // ---------------------------------------------------------

                if (!item.Equals(
                    "Fees",
                    StringComparison.OrdinalIgnoreCase))
                {
                    worksheet.Cell(row, 8)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#DDEBF7");
                }

                // ---------------------------------------------------------
                // RULE 2:
                // If no contractor name could be parsed from the VMS
                // Fee Description, highlight the entire row very light red.
                //
                // Red takes priority over blue.
                // ---------------------------------------------------------

                if (string.IsNullOrWhiteSpace(name))
                {
                    worksheet.Range(
                        row,
                        1,
                        row,
                        lastColumn)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F4CCCC");
                }

                // ---------------------------------------------------------
                // RULE 3:
                // Contractor exists, but no OIR invoice was within 5%.
                // Keep your gray unmatched-row behavior.
                // ---------------------------------------------------------

                else if (string.IsNullOrWhiteSpace(invoice))
                {
                    worksheet.Range(
                        row,
                        1,
                        row,
                        lastColumn)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F2F2F2");

                    // Reapply the blue Item color because the gray row
                    // fill above overwrote it.
                    if (!item.Equals(
                        "Fees",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet.Cell(row, 8)
                            .Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#DDEBF7");
                    }
                }

                // Amount Due of zero stays red text.
                if (amountDue <= 0)
                {
                    worksheet.Cell(row, 4)
                        .Style.Font.FontColor =
                        XLColor.Red;
                }
            }

            worksheet.Range(
                1,
                1,
                lastRow,
                lastColumn)
                .SetAutoFilter();
        }

        // =============================================================
        // FILE SAVE HELPER
        // =============================================================

        private static string GetUniqueOutputPath(
            string folderPath,
            string fileName)
        {
            string name =
                Path.GetFileNameWithoutExtension(
                    fileName);

            string extension =
                Path.GetExtension(
                    fileName);

            string path =
                Path.Combine(
                    folderPath,
                    fileName);

            int counter = 1;

            while (File.Exists(path))
            {
                path =
                    Path.Combine(
                        folderPath,
                        $"{name} ({counter}){extension}");

                counter++;
            }

            return path;
        }

        // =============================================================
        // INTERNAL OUTPUT MODELS
        // =============================================================

        private class JohnsonJohnsonOutputRow
        {
            public string WeekEndingDate { get; set; } = "";

            public string Name { get; set; } = "";

            public string Invoice { get; set; } = "";

            public decimal AmountDue { get; set; }

            public decimal AggregateAmountPaid { get; set; }

            public decimal Tax { get; set; }

            public string Notes { get; set; } = "";

            public string Item { get; set; } = "";

            public string Concat { get; set; } = "";

            public string InvoiceNumber { get; set; } = "";
        }

        private class OirLookupRow
        {
            public string Name { get; set; } = "";

            public DateTime WeekEndingDate { get; set; }

            public string Invoice { get; set; } = "";

            public decimal AmountDue { get; set; }

            public string Concat { get; set; } = "";
        }
    }
}