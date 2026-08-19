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

            int groupId = 0;

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

if (vmsMatches != null &&
    vmsMatches.TryGetValue(
        invoiceId,
        out JohnsonJohnsonVmsMatch? vmsMatch))
{
    name = vmsMatch.WorkerName;

    feeDescription = vmsMatch.FeeDescription;

    name = Rename.ApplyNameChange(
        name,
        nameChanges);
}

                groupId++;

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
                // First try matching the full LineTotal.
                // -----------------------------------------------------

                var matches =
                    FindBestInvoiceCombination(
                        possibleMatches,
                        aggregateAmountPaid);

                // -----------------------------------------------------
                // If OIR invoice(s) were found, create one output row
                // per matched OIR invoice.
                //
                // This allows one J&J payment row to match multiple
                // OIR invoices, similar to Microsoft.
                // -----------------------------------------------------

                if (matches.Any())
                {
                    foreach (var match in matches)
                    {
                        outputRows.Add(
                            new JohnsonJohnsonOutputRow
                            {
                                WeekEndingDate =
                                    formattedWeekEndingDate,

                                Name = name,

                                Invoice =
                                    match.Invoice,

                                AmountDue =
                                    match.AmountDue,

                                AggregateAmountPaid =
                                    aggregateAmountPaid,

                                Tax = tax,

                                Notes = feeDescription,

                                Item = item,

                                Concat = concat,

                                InvoiceNumber =
                                    invoiceId,

                                GroupId = groupId
                            });
                    }
                }
                else
                {
                    // -------------------------------------------------
                    // Still output the payment row if no OIR match
                    // exists.
                    //
                    // This makes unmatched items visible rather than
                    // silently dropping them.
                    // -------------------------------------------------

                    outputRows.Add(
                        new JohnsonJohnsonOutputRow
                        {
                            WeekEndingDate =
                                formattedWeekEndingDate,

                            Name = name,

                            Invoice = "",

                            AmountDue = 0,

                            AggregateAmountPaid =
                                aggregateAmountPaid,

                            Tax = tax,

                            Notes =
                    !string.IsNullOrWhiteSpace(feeDescription)
                        ? feeDescription
                        : "No J&J VMS match found.",

                            Item = item,

                            Concat = concat,

                            InvoiceNumber =
                                invoiceId,

                            GroupId = groupId
                        });
                }
            }

            // ---------------------------------------------------------
            // Capture total before clearing the original worksheet.
            // ---------------------------------------------------------

            decimal total =
                outputRows.Sum(
                    x => x.AggregateAmountPaid);

            // If one payment matched multiple OIR invoices,
            // AggregateAmountPaid is repeated in outputRows.
            //
            // Therefore recalculate total by payment GroupId instead
            // of blindly summing every output row.
            total =
                outputRows
                    .GroupBy(x => x.GroupId)
                    .Sum(g =>
                        g.First().AggregateAmountPaid);

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

            // ---------------------------------------------------------
            // If one J&J payment row matched multiple OIR invoices,
            // merge the payment-level information vertically.
            //
            // Invoice + Amount Due stay separate because those belong
            // to individual OIR invoices.
            // ---------------------------------------------------------

            foreach (var group in
                     outputRows.GroupBy(x => x.GroupId))
            {
                int firstOutputRow =
                    outputRows.IndexOf(group.First()) + 2;

                int lastOutputRow =
                    outputRows.IndexOf(group.Last()) + 2;

                if (lastOutputRow > firstOutputRow)
                {
                    // Week Ending Date
                    worksheet.Range(
                        firstOutputRow,
                        1,
                        lastOutputRow,
                        1)
                        .Merge();

                    // Name
                    worksheet.Range(
                        firstOutputRow,
                        2,
                        lastOutputRow,
                        2)
                        .Merge();

                    // Aggregate Amount Paid
                    worksheet.Range(
                        firstOutputRow,
                        5,
                        lastOutputRow,
                        5)
                        .Merge();

                    // Tax
                    worksheet.Range(
                        firstOutputRow,
                        6,
                        lastOutputRow,
                        6)
                        .Merge();

                    // Notes
                    worksheet.Range(
                        firstOutputRow,
                        7,
                        lastOutputRow,
                        7)
                        .Merge();

                    // Item
                    worksheet.Range(
                        firstOutputRow,
                        8,
                        lastOutputRow,
                        8)
                        .Merge();

                    // Concat
                    worksheet.Range(
                        firstOutputRow,
                        9,
                        lastOutputRow,
                        9)
                        .Merge();

                    // Invoice Number / InvoiceID
                    worksheet.Range(
                        firstOutputRow,
                        10,
                        lastOutputRow,
                        10)
                        .Merge();
                }

                worksheet.Range(
                    firstOutputRow,
                    1,
                    lastOutputRow,
                    10)
                    .Style.Alignment.Vertical =
                    XLAlignmentVerticalValues.Center;
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
            Dictionary<string, List<OirMatch>>
                openInvoiceMatches)
        {
            var rows =
                new List<OirLookupRow>();

            foreach (var item in openInvoiceMatches)
            {
                // OIR dictionary keys look like:
                //
                // John Smith 08/15/2026
                //
                // Pull the contractor name and date back apart.
                Match match = Regex.Match(
                    item.Key,
                    @"^(?<name>.+)\s(?<date>\d{1,2}/\d{1,2}/\d{2,4})$");

                if (!match.Success)
                    continue;

                if (!DateTime.TryParse(
                    match.Groups["date"].Value,
                    out DateTime weekEnding))
                {
                    continue;
                }

                string name =
                    match.Groups["name"]
                        .Value
                        .Trim();

                foreach (var oirMatch in item.Value)
                {
                    rows.Add(
                        new OirLookupRow
                        {
                            Name = name,

                            WeekEndingDate =
                                weekEnding,

                            Invoice =
                                oirMatch.DocumentNumber,

                            AmountDue =
                                oirMatch.RemainingAmount,

                            Concat =
                                item.Key
                        });
                }
            }

            return rows;
        }

        // =============================================================
        // INVOICE COMBINATION MATCHING
        // =============================================================

        private static List<OirLookupRow>
            FindBestInvoiceCombination(
                List<OirLookupRow> possibleMatches,
                decimal targetAmount)
        {
            // Exact match tolerance = 10 cents.
            const decimal exactTolerance = 0.10m;

            // Similar to your Microsoft workflow.
            // Allows differences of up to 2% if no exact combination
            // exists.
            const decimal allowedPercentDifference = 0.02m;

            decimal amountTolerance =
                    Math.Abs(
                        targetAmount *
                        allowedPercentDifference);

            List<OirLookupRow> bestMatch =
                new();

            decimal bestDifference =
                decimal.MaxValue;

            int count =
                possibleMatches.Count;

            if (count == 0)
                return bestMatch;

            /*
                The bit-mask combination approach becomes extremely
                expensive with very large candidate sets.

                A contractor/week-ending combination should normally
                have only a small number of OIR invoices, but limiting
                to 20 also protects the application from accidentally
                attempting millions/billions of combinations.
            */
            if (count > 20)
            {
                possibleMatches =
                    possibleMatches
                        .OrderBy(x =>
                            Math.Abs(
                                x.AmountDue -
                                targetAmount))
                        .Take(20)
                        .ToList();

                count =
                    possibleMatches.Count;
            }

            int combinationCount =
                1 << count;

            for (int mask = 1;
                 mask < combinationCount;
                 mask++)
            {
                var currentGroup =
                    new List<OirLookupRow>();

                for (int i = 0;
                     i < count;
                     i++)
                {
                    if ((mask & (1 << i)) != 0)
                    {
                        currentGroup.Add(
                            possibleMatches[i]);
                    }
                }

                decimal groupTotal =
                    currentGroup.Sum(
                        x => x.AmountDue);

                decimal difference =
                    Math.Abs(
                        groupTotal -
                        targetAmount);

                // Exact match wins immediately.
                if (difference <= exactTolerance)
                {
                    return currentGroup;
                }

                // Otherwise remember the best match
                // within 2%.
                if (difference <= amountTolerance &&
                    difference < bestDifference)
                {
                    bestDifference =
                        difference;

                    bestMatch =
                        currentGroup;
                }
            }

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

                decimal amountDue =
                    GetDecimalValue(
                        worksheet.Cell(row, 4));

                // Gray unmatched rows.
                if (string.IsNullOrWhiteSpace(name) ||
                    string.IsNullOrWhiteSpace(invoice))
                {
                    worksheet.Range(
                        row,
                        1,
                        row,
                        lastColumn)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F2F2F2");
                }

                if (amountDue <= 0)
                {
                    worksheet.Cell(row, 4)
                        .Style.Font.FontColor =
                        XLColor.Red;
                }
            }

            // ---------------------------------------------------------
            // Total Aggregate Amount Paid.
            // ---------------------------------------------------------

            int totalRow =
                lastRow + 1;

            var totalCell =
                worksheet.Cell(
                    totalRow,
                    5);

            /*
                Some payment values may be merged because one payment
                matched multiple invoices.

                SUM still treats a merged region as the value in its
                first cell, so this remains safe.
            */
            totalCell.FormulaA1 =
                $"=SUM(E2:E{lastRow})";

            totalCell.Style.Font.FontName =
                "Aptos Narrow";

            totalCell.Style.Font.FontSize =
                9;

            totalCell.Style.Font.Bold =
                true;

            totalCell.Style.Fill.BackgroundColor =
                XLColor.Yellow;

            totalCell.Style.NumberFormat.Format =
                "$#,##0.00;($#,##0.00)";

            totalCell.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

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

            public int GroupId { get; set; }
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