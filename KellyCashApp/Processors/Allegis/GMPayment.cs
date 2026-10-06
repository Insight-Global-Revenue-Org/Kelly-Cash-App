using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Services;
using System.Globalization;
using System.Text.RegularExpressions;

namespace KellyCashApp.Processors.Allegis
{
    internal class GMPayment
    {
        private const int HeaderRow = 6;
        private const int FirstDataRow = 10;

        public static bool IsGMFormat(IXLWorksheet worksheet)
        {
            int customerCol =
                FindColumn(worksheet, HeaderRow, "Customer");

            if (customerCol == -1)
                return false;

            int lastRow =
                worksheet.LastRowUsed()?.RowNumber()
                ?? FirstDataRow;

            for (int row = FirstDataRow; row <= lastRow; row++)
            {
                string customer =
                    worksheet.Cell(row, customerCol)
                        .GetString()
                        .Trim();

                if (customer.Equals(
                    "GENERAL MOTORS (T&M) - USA",
                    StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static string Process(
            XLWorkbook workbook,
            IXLWorksheet worksheet,
            string inputPath,
            Dictionary<string, List<OirMatch>> openInvoiceMatches,
                List<GMExpenseMatch>? expenseMatches = null,
            List<AllegisTimesheetMatch>? timesheetMatches = null)
        {
            int gmInvoiceCol =
                FindColumn(
                    worksheet,
                    HeaderRow,
                    "Consolidated Invoice ID");

            int gmInvoiceIdCol =
                FindColumn(
                    worksheet,
                    HeaderRow,
                    "Invoice ID");

            int workerCol =
                FindColumn(
                    worksheet,
                    HeaderRow,
                    "Worker");
            int lineItemEndDateCol = FindColumn(worksheet, HeaderRow, "Invoice Line Item End Date");
            int aggregateAmountCol = FindColumn(worksheet, HeaderRow, "Total Invoice Line Item Amount (Supplier)");
            int taxCol = FindColumn(worksheet, HeaderRow, "Invoice Line Item Total Tax Amount (Supplier)");

            if (gmInvoiceCol == -1 ||
                gmInvoiceIdCol == -1 ||
                workerCol == -1 ||
                lineItemEndDateCol == -1 ||
                aggregateAmountCol == -1 ||
                taxCol == -1)
            {
                throw new Exception(
                    "Missing one or more required GM remittance columns.");
            }

            var oirRows = BuildOirRows(openInvoiceMatches);

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? FirstDataRow;

            var outputRows = new List<GMOutputRow>();

            var nameChanges = Rename.LoadNameChanges();

            for (int row = FirstDataRow; row <= lastRow; row++)
            {
               
                string rawWorker =
                    worksheet.Cell(row, workerCol).GetString().Trim();

                bool isExpense =
                rawWorker.Equals(
                    "NULL, NULL",
                StringComparison.OrdinalIgnoreCase);

                string type =
                    isExpense
                        ? "Expense"
                        : "Labor";

                if (string.IsNullOrWhiteSpace(rawWorker))
                    continue;

                DateTime lineItemEndDate =
                    GetDateValue(
                        worksheet.Cell(row, lineItemEndDateCol));

                if (lineItemEndDate == DateTime.MinValue)
                    continue;

                decimal aggregateAmount =
                    GetDecimalValue(
                        worksheet.Cell(row, aggregateAmountCol));

                decimal tax =
                    GetDecimalValue(
                        worksheet.Cell(row, taxCol));

                decimal preTaxAggregateAmount =
                    aggregateAmount - tax;

                string name;

                // GM expense rows do not contain the worker name.
                // Recover it from the configured GM Expense Report.
                if (isExpense)
                {
                    GMExpenseMatch? expenseMatch = expenseMatches?
                    .Where(x =>
                        Math.Abs(
                            (x.Date.Date - lineItemEndDate.Date).Days) <= 14
                        &&
                        x.Amount == aggregateAmount)
                    .OrderBy(x =>
                        Math.Abs(
                            (x.Date.Date - lineItemEndDate.Date).Days))
                    .FirstOrDefault();

                    if (expenseMatch != null)
                    {
                        name = FormatName(
                            expenseMatch.ContractorName);
                    }
                    else
                    {
                        // Leave unmatched expenses obvious in output.
                        name = "NULL, NULL";
                    }
                }
                else
                {
                    name = FormatName(rawWorker);
                }

                name = Rename.ApplyNameChange(
                    name,
                    nameChanges);

                string formattedLineItemEndDate =
                    lineItemEndDate.ToString("MM/dd/yyyy");

                string aggregationKey;

                if (isExpense &&
                    name.Equals(
                        "NULL, NULL",
                        StringComparison.OrdinalIgnoreCase))
                {
                    // Never combine unmatched expense rows.
                    aggregationKey = $"UNMATCHED_EXPENSE_{row}";
                }
                else
                {
                    aggregationKey =
                        $"{type}|{name.ToUpperInvariant()}|{formattedLineItemEndDate}";
                }

                if (lineItemEndDate == DateTime.MinValue)
                    continue;

                // Consolidated Invoice ID
                string gmInvoice =
                    worksheet.Cell(row, gmInvoiceCol)
                        .GetString()
                        .Trim();

                // Individual GM Invoice ID
                string gmInvoiceId =
                    worksheet.Cell(row, gmInvoiceIdCol)
                        .GetString()
                        .Trim();

                // VMS Identifier continues to come from Consolidated Invoice ID
                string vmsIdentifier =
                    new string(gmInvoice
                        .Where(char.IsDigit)
                        .ToArray());

                if (vmsIdentifier.Length > 5)
                    vmsIdentifier = vmsIdentifier[^5..];

                AllegisTimesheetMatch? timesheetMatch =
                    timesheetMatches?
                        .Where(x =>
                            x.ContractorName.Equals(
                 name,
                 StringComparison.OrdinalIgnoreCase)
             &&
             Math.Abs(
                 (x.WeekEndingDate.Date - lineItemEndDate.Date).Days) <= 1)
                .OrderBy(x =>
             Math.Abs(
                 (x.WeekEndingDate.Date - lineItemEndDate.Date).Days))
                    .FirstOrDefault();

                outputRows.Add(new GMOutputRow
                {
                    WeekEndingDate = formattedLineItemEndDate,
                    Name = name,

                    Invoice = "",
                    AmountDue = 0,

                    InvoiceLineItemEndDate = formattedLineItemEndDate,
                    AggregateInvoiceLineItemAmount = aggregateAmount,
                    Tax = tax,

                    Notes = "",

                    Type = type,

                    Concat = $"{name} {formattedLineItemEndDate}",

                    gmInvoice = gmInvoice,
                    VmsIdentifier = vmsIdentifier,
                    GMInvoiceId = gmInvoiceId,

                    AggregationKey = aggregationKey,

                    AggregateInvoicedNet =
                    timesheetMatch?.AggregateInvoicedNet ?? 0,

                    Hours =
                    timesheetMatch?.Hours ?? 0,

                    RtRate =
                    timesheetMatch?.RtRate ?? 0,

                    OtRate =
                    timesheetMatch?.OtRate ?? 0,

                    DtRate =
                    timesheetMatch?.DtRate ?? 0
                });

            }

            outputRows = outputRows
            .GroupBy(x => x.AggregationKey)
            .Select(group =>
            {
        var first = group.First();

        return new GMOutputRow
        {
            WeekEndingDate = first.WeekEndingDate,
            Name = first.Name,

            Invoice = "",
            AmountDue = 0,

            InvoiceLineItemEndDate = first.InvoiceLineItemEndDate,

            AggregateInvoiceLineItemAmount =
                group.Sum(x => x.AggregateInvoiceLineItemAmount),

            Tax = group.Sum(x => x.Tax),

            Notes = "",

            Type = first.Type,

            Concat =
                $"{first.Name} {first.WeekEndingDate}",

            gmInvoice = string.Join(
                ", ",
                group.Select(x => x.gmInvoice)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)),

                     VmsIdentifier = first.VmsIdentifier,

                        GMInvoiceId = string.Join(
                ", ",
                group.Select(x => x.GMInvoiceId)
        .Where(x => !string.IsNullOrWhiteSpace(x))
        .Distinct(StringComparer.OrdinalIgnoreCase)),

            AggregationKey = first.AggregationKey,

            AggregateInvoicedNet =
                group.Sum(x => x.AggregateInvoicedNet),

            Hours =
                group.Sum(x => x.Hours),

            RtRate = first.RtRate,
            OtRate = first.OtRate,
            DtRate = first.DtRate
        };
    })
    .ToList();

            var matchedInvoiceNumbers =
    new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var outputRow in outputRows)
            {
                if (!DateTime.TryParse(
                    outputRow.WeekEndingDate,
                    out DateTime centerpointDate))
                {
                    continue;
                }

                var match = oirRows
                    .Where(x =>
                        x.Name.Equals(
                            outputRow.Name,
                            StringComparison.OrdinalIgnoreCase)

                        && Math.Abs(
                            (x.WeekEndingDate.Date - centerpointDate.Date).Days) <= 1

                        && !string.IsNullOrWhiteSpace(x.Invoice)

                        && !matchedInvoiceNumbers.Contains(x.Invoice)

                        // Expense rows MUST match the OIR remaining amount exactly.
                        && (
                            !outputRow.Type.Equals(
                                "Expense",
                                StringComparison.OrdinalIgnoreCase)
                            ||
                            x.AmountDue ==
                                outputRow.AggregateInvoiceLineItemAmount
                        ))
                    .OrderBy(x =>
                        Math.Abs(
                             (x.WeekEndingDate.Date - centerpointDate.Date).Days))
                    .ThenBy(x =>
                        Math.Abs(
                            x.AmountDue -
                            outputRow.AggregateInvoiceLineItemAmount))
                    .FirstOrDefault();

                if (match == null)
                    continue;

                outputRow.Invoice = match.Invoice;
                outputRow.AmountDue = match.AmountDue;

                outputRow.Concat =
                    $"{outputRow.Name} {match.WeekEndingDate:MM/dd/yyyy}";

                matchedInvoiceNumbers.Add(match.Invoice);
            }

            decimal total = 0;

            for (int row = FirstDataRow; row <= lastRow; row++)
            {
                total += GetDecimalValue(worksheet.Cell(row, aggregateAmountCol));
            }

            foreach (var picture in worksheet.Pictures.ToList())
            {
                picture.Delete();
            }

            worksheet.Clear(XLClearOptions.All);
            worksheet.Style.Fill.SetBackgroundColor(XLColor.NoColor);

            // New name for output sheet :)
            worksheet.Name = "Reconciliation Notes";

            string[] headers =
            {
                "Week Ending Date",
                "Name",
                "Invoice",
                "Amount Due",
                "Aggregate Paid",
                "Tax",
                "Notes",
                "Type",
                "Concat",
                "Consolidated Invoice",
                "VMS Identifier",
                "GM Invoice ID",
                "Invoiced Net",
                "Hours",
                "RT Rate",
                "OT Rate",
                "DT Rate"
            };

            for (int col = 1; col <= headers.Length; col++)
                worksheet.Cell(1, col).Value = headers[col - 1];

            for (int i = 0; i < outputRows.Count; i++)
            {
                int row = i + 2;
                var item = outputRows[i];

                worksheet.Cell(row, 1).Value = item.InvoiceLineItemEndDate;
                worksheet.Cell(row, 2).Value = item.Name;
                worksheet.Cell(row, 3).Value = item.Invoice;
                worksheet.Cell(row, 4).Value = item.AmountDue;
                worksheet.Cell(row, 5).Value = item.AggregateInvoiceLineItemAmount;
                worksheet.Cell(row, 6).Value = item.Tax;
                worksheet.Cell(row, 7).Value = item.Notes;
                worksheet.Cell(row, 8).Value = item.Type;
                worksheet.Cell(row, 9).Value = item.Concat;
                worksheet.Cell(row, 10).Value = item.gmInvoice;
                worksheet.Cell(row, 11).Value = item.VmsIdentifier;
                worksheet.Cell(row, 12).Value = item.GMInvoiceId;
                worksheet.Cell(row, 13).Value = item.AggregateInvoicedNet;
                worksheet.Cell(row, 14).Value = item.Hours;
                worksheet.Cell(row, 15).Value = item.RtRate;
                worksheet.Cell(row, 16).Value = item.OtRate;
                worksheet.Cell(row, 17).Value = item.DtRate;

                worksheet.Row(row).AdjustToContents();
            }

            ApplyFormatting(worksheet, outputRows.Count + 1, headers.Length);

            string downloadsPath = Settings.GetRemittanceSavePath();


            string formattedTotal = total.ToString("$#,##0.00;($#,##0.00)", CultureInfo.InvariantCulture);
            string processedDate = DateTime.Now.ToString("M.d.yyyy", CultureInfo.InvariantCulture);

            string outputPath = GetUniqueOutputPath(
            downloadsPath,
            $"GM {processedDate} - {formattedTotal}.xlsx");

            workbook.SaveAs(outputPath);

            Analytics.LogRemittanceRun($"GM - {formattedTotal}");

            return outputPath;
        }

        private static List<OirLookupRow> BuildOirRows(Dictionary<string, List<OirMatch>> openInvoiceMatches)
        {
            var rows = new List<OirLookupRow>();

            foreach (var item in openInvoiceMatches)
            {
                Match match = Regex.Match(item.Key, @"^(?<name>.+)\s(?<date>\d{1,2}/\d{1,2}/\d{2,4})$");

                if (!match.Success)
                    continue;

                if (!DateTime.TryParse(match.Groups["date"].Value, out DateTime weekEnding))
                    continue;

                foreach (var oirMatch in item.Value)
                {
                    rows.Add(new OirLookupRow
                    {
                        Name = match.Groups["name"].Value.Trim(),
                        WeekEndingDate = weekEnding,
                        Invoice = oirMatch.DocumentNumber,
                        AmountDue = oirMatch.RemainingAmount,
                        Concat = item.Key
                    });
                }
            }

            return rows;
        }

        private static void ApplyFormatting(IXLWorksheet worksheet, int lastRow, int lastColumn)
        {
            var range = worksheet.Range(1, 1, lastRow, lastColumn);

            range.Style.Font.FontName = "Aptos Narrow";
            range.Style.Font.FontSize = 9;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            range.Style.Alignment.WrapText = true;
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            worksheet.Row(1).Style.Font.FontName = "Aptos Narrow";
            worksheet.Row(1).Style.Font.FontSize = 9;
            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FCE4D6");

            worksheet.Column(4).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            worksheet.Column(5).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            worksheet.Column(6).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

            for (int row = 2; row <= lastRow; row++)
            {
                worksheet.Row(row).Height = 13;
            }

            worksheet.Row(1).Height = 15;

            worksheet.Columns().AdjustToContents();

            worksheet.Column(1).Width = 22;  // Invoice Line Item End Date
            worksheet.Column(2).Width = 22;  // Name
            worksheet.Column(3).Width = 18;  // Invoice
            worksheet.Column(4).Width = 18;  // Amount Due
            worksheet.Column(5).Width = 28;  // Aggregate Paid
            worksheet.Column(6).Width = 14;  // Tax
            worksheet.Column(7).Width = 42;  // Notes
            worksheet.Column(7).Style.Alignment.WrapText = false;

            worksheet.Column(8).Width = 12;  // Type
            worksheet.Column(9).Width = 32;  // Concat
            worksheet.Column(10).Width = 20; // Consolidated Invoice
            worksheet.Column(11).Width = 12; // VMS Identifier
            worksheet.Column(12).Width = 20; // GM Invoice ID
            worksheet.Column(13).Width = 12; // Invoiced Net
            worksheet.Column(14).Width = 12; // Hours
            worksheet.Column(15).Width = 12; // RT Rate
            worksheet.Column(16).Width = 12; // OT Rate
            worksheet.Column(17).Width = 12; // DT Rate

            worksheet.Column(13).Style.NumberFormat.Format =
                "$#,##0.00;($#,##0.00)";

            worksheet.Column(14).Style.NumberFormat.Format =
                "0.00";

            worksheet.Column(15).Style.NumberFormat.Format =
                "$#,##0.00;($#,##0.00)";

            worksheet.Column(16).Style.NumberFormat.Format =
                "$#,##0.00;($#,##0.00)";

            worksheet.Column(17).Style.NumberFormat.Format =
                "$#,##0.00;($#,##0.00)";

            for (int row = 2; row <= lastRow; row++)
            {
                decimal amountDue = GetDecimalValue(worksheet.Cell(row, 4));

                if (amountDue <= 0)
                {
                    worksheet.Range(row, 1, row, 17)
                        .Style.Fill.BackgroundColor =
                        XLColor.FromHtml("#F2F2F2");
                }
            }

            worksheet.Range(1, 1, lastRow, lastColumn).SetAutoFilter();
        }

        private static List<OirLookupRow> FindBestInvoiceCombination(
    List<OirLookupRow> possibleMatches,
    decimal targetAmount)
        {
            const decimal vmsFeeTolerancePercent = 0.05m;

            decimal exactTolerance = 0.10m;
            decimal feeTolerance = Math.Abs(targetAmount * vmsFeeTolerancePercent);

            List<OirLookupRow> bestMatch = new();
            decimal bestDifference = decimal.MaxValue;

            int count = possibleMatches.Count;

            for (int mask = 1; mask < (1 << count); mask++)
            {
                var currentGroup = new List<OirLookupRow>();

                for (int i = 0; i < count; i++)
                {
                    if ((mask & (1 << i)) != 0)
                        currentGroup.Add(possibleMatches[i]);
                }

                decimal groupTotal = currentGroup.Sum(x => x.AmountDue);
                decimal difference = Math.Abs(groupTotal - targetAmount);

                if (difference <= exactTolerance)
                    return currentGroup;

                if (difference <= feeTolerance && difference < bestDifference)
                {
                    bestDifference = difference;
                    bestMatch = currentGroup;
                }
            }

            return bestMatch;
        }

        private static int FindColumn(IXLWorksheet worksheet, int headerRow, string headerName)
        {
            int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 50;

            for (int col = 1; col <= lastColumn; col++)
            {
                string headerText = worksheet.Cell(headerRow, col).GetString().Trim();

                if (headerText.Equals(headerName, StringComparison.OrdinalIgnoreCase))
                    return col;
            }

            return -1;
        }

        private static string FormatName(string input)
        {
            input = input.Trim();

            if (!input.Contains(","))
                return input;

            string[] parts = input.Split(',', 2);

            string last = ToTitle(parts[0].Trim());
            string first = ToTitle(parts[1].Trim());

            return $"{first} {last}".Trim();
        }

        private static string ToTitle(string value)
        {
            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value.ToLower());
        }

        private static DateTime GetDateValue(IXLCell cell)
        {
            if (cell.Value.IsDateTime)
                return cell.GetDateTime();

            if (DateTime.TryParse(cell.GetString().Trim(), out DateTime parsed))
                return parsed;

            return DateTime.MinValue;
        }

        private static decimal GetDecimalValue(IXLCell cell)
        {
            string raw = cell.Value.ToString()
                .Replace("$", "")
                .Replace(",", "")
                .Replace("(", "-")
                .Replace(")", "")
                .Trim();

            if (decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return 0;
        }

        private static string GetUniqueOutputPath(string folderPath, string fileName)
        {
            string name = Path.GetFileNameWithoutExtension(fileName);
            string ext = Path.GetExtension(fileName);

            string path = Path.Combine(folderPath, fileName);
            int counter = 1;

            while (File.Exists(path))
            {
                path = Path.Combine(folderPath, $"{name} ({counter}){ext}");
                counter++;
            }

            return path;
        }

        public class GMVmsMatch
        {
            public decimal AggregateInvoicedNet { get; set; }
            public decimal Hours { get; set; }
            public decimal RtRate { get; set; }
            public decimal OtRate { get; set; }
            public decimal DtRate { get; set; }
        }

        private class GMOutputRow
        {
            public string WeekEndingDate { get; set; } = "";
            public string Name { get; set; } = "";
            public string Invoice { get; set; } = "";
            public decimal AmountDue { get; set; }
            public decimal AggregateInvoiceLineItemAmount { get; set; }
            public decimal Tax { get; set; }
            public string Notes { get; set; } = "";
            public string Type { get; set; } = "";
            public string InvoiceLineItemEndDate { get; set; } = "";
            public string Concat { get; set; } = "";
            public string gmInvoice { get; set; } = "";
            public string VmsIdentifier { get; set; } = "";
            public string GMInvoiceId { get; set; } = "";

            // Internal only — used to control grouping.
            public string AggregationKey { get; set; } = "";

            public decimal AggregateInvoicedNet { get; set; }
            public decimal Hours { get; set; }
            public decimal RtRate { get; set; }
            public decimal OtRate { get; set; }
            public decimal DtRate { get; set; }
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