using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Services;
using System.Globalization;
using System.Text.RegularExpressions;

namespace KellyCashApp.Processors.Allegis
{
    internal class SAICPayment
    {
        private const int HeaderRow = 6;
        private const int FirstDataRow = 10;

        public static bool IsSAICFormat(IXLWorksheet worksheet)
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
                    "SAIC - (NSAI)",
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
            List<SAPSAICTimesheetMatch>? timesheetMatches = null)
        {
            int saicInvoiceCol = FindColumn(worksheet, HeaderRow, "Consolidated Invoice ID");
            int workerCol = FindColumn(worksheet, HeaderRow, "Worker");
            int lineItemEndDateCol = FindColumn(worksheet, HeaderRow, "Invoice Line Item End Date");
            int aggregateAmountCol = FindColumn(worksheet, HeaderRow, "Total Invoice Line Item Amount (Supplier)");
            int taxCol = FindColumn(worksheet, HeaderRow, "Invoice Line Item Total Tax Amount (Supplier)");

            if (saicInvoiceCol == -1 || workerCol == -1 || lineItemEndDateCol == -1 || aggregateAmountCol == -1 || taxCol == -1)
                throw new Exception("Missing one or more required SAIC remittance columns.");

            var oirRows = BuildOirRows(openInvoiceMatches);

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? FirstDataRow;

            var outputRows = new List<SAICOutputRow>();

            var nameChanges = Rename.LoadNameChanges();

            for (int row = FirstDataRow; row <= lastRow; row++)
            {
                string rawWorker = worksheet.Cell(row, workerCol).GetString().Trim();

                if (string.IsNullOrWhiteSpace(rawWorker))
                    continue;

                string name = FormatName(rawWorker);
                name = Rename.ApplyNameChange(name, nameChanges);
                DateTime lineItemEndDate = GetDateValue(worksheet.Cell(row, lineItemEndDateCol));

                if (lineItemEndDate == DateTime.MinValue)
                    continue;

                string formattedLineItemEndDate = lineItemEndDate.ToString("MM/dd/yyyy");


                decimal aggregateAmount = GetDecimalValue(worksheet.Cell(row, aggregateAmountCol));
                decimal tax = GetDecimalValue(worksheet.Cell(row, taxCol));
                decimal preTaxAggregateAmount = aggregateAmount - tax;
                string saicInvoice =
                worksheet.Cell(row, saicInvoiceCol).GetString().Trim();

                string vmsIdentifier =
                    new string(saicInvoice
                        .Where(char.IsDigit)
                        .ToArray());

                if (vmsIdentifier.Length > 5)
                    vmsIdentifier = vmsIdentifier[^5..];

                SAPSAICTimesheetMatch? timesheetMatch =
                    timesheetMatches?
                        .Where(x =>
            x.ContractorName.Equals(
                name,
                StringComparison.OrdinalIgnoreCase)

            && Math.Abs(
                (x.WeekEndingDate.Date -
                 lineItemEndDate.Date).Days) <= 1)
                        .OrderBy(x =>
            Math.Abs(
                (x.WeekEndingDate.Date -
                 lineItemEndDate.Date).Days))
                        .FirstOrDefault();

                outputRows.Add(new SAICOutputRow
                {
                    WeekEndingDate = formattedLineItemEndDate,
                    Name = name,

                    Invoice = "",
                    AmountDue = 0,

                    InvoiceLineItemEndDate = formattedLineItemEndDate,
                    AggregateInvoiceLineItemAmount = aggregateAmount,
                    Tax = tax,

                    Notes = "",

                    Concat = $"{name} {formattedLineItemEndDate}",

                    saicInvoice = saicInvoice,
                    VmsIdentifier = vmsIdentifier,

                    AggregateInvoicedNet =
                    timesheetMatch?.AggregateInvoicedNet ?? 0,

                    Hours =
                    timesheetMatch?.Hours ?? 0,

                    RtRate =
                    timesheetMatch?.RtRate ?? 0,

                    OtRate = 0,

                    DtRate = 0
                });

            }

            outputRows = outputRows
    .GroupBy(x => new
    {
        Name = x.Name.ToUpperInvariant(),
        x.WeekEndingDate
    })
    .Select(group =>
    {
        var first = group.First();

        return new SAICOutputRow
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

            Concat =
                $"{first.Name} {first.WeekEndingDate}",

            saicInvoice = string.Join(
                ", ",
                group.Select(x => x.saicInvoice)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)),

            VmsIdentifier = first.VmsIdentifier,

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

                        && !matchedInvoiceNumbers.Contains(x.Invoice))
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
                "Concat",
               "SAIC Invoice",
                "VMS Identifier",
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
                worksheet.Cell(row, 8).Value = item.Concat;
                worksheet.Cell(row, 9).Value = item.saicInvoice;
                worksheet.Cell(row, 10).Value = item.VmsIdentifier;
                worksheet.Cell(row, 11).Value = item.AggregateInvoicedNet;
                worksheet.Cell(row, 12).Value = item.Hours;
                worksheet.Cell(row, 13).Value = item.RtRate;
                worksheet.Cell(row, 14).Value = item.OtRate;
                worksheet.Cell(row, 15).Value = item.DtRate;

                worksheet.Row(row).AdjustToContents();
            }

            ApplyFormatting(worksheet, outputRows.Count + 1, headers.Length);

            string downloadsPath = Settings.GetRemittanceSavePath();


            string formattedTotal = total.ToString("$#,##0.00;($#,##0.00)", CultureInfo.InvariantCulture);
            string processedDate = DateTime.Now.ToString("M.d.yyyy", CultureInfo.InvariantCulture);

            string outputPath = GetUniqueOutputPath(
            downloadsPath,
            $"SAIC {processedDate} - {formattedTotal}.xlsx");

            workbook.SaveAs(outputPath);

            Analytics.LogRemittanceRun($"SAIC - {formattedTotal}");

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
            worksheet.Column(8).Width = 32;  // Concat
            worksheet.Column(9).Width = 20;  // SAIC Invoice
            worksheet.Column(10).Width = 12; // VMS Identifier
            worksheet.Column(11).Width = 12; // Invoiced Net
            worksheet.Column(12).Width = 12; // Hours
            worksheet.Column(13).Width = 12; // RT Rate
            worksheet.Column(14).Width = 12; // OT Rate
            worksheet.Column(15).Width = 12; // DT Rate;

            worksheet.Column(11).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            worksheet.Column(12).Style.NumberFormat.Format = "0.00";
            worksheet.Column(13).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            worksheet.Column(14).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            worksheet.Column(15).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

            for (int row = 2; row <= lastRow; row++)
            {
                decimal amountDue = GetDecimalValue(worksheet.Cell(row, 4));

                if (amountDue <= 0)
                {
                    worksheet.Range(row, 1, row, 15)
                        .Style.Fill.BackgroundColor = XLColor.FromHtml("#F2F2F2");
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

        public class SAICVmsMatch
        {
            public decimal AggregateInvoicedNet { get; set; }
            public decimal Hours { get; set; }
            public decimal RtRate { get; set; }
            public decimal OtRate { get; set; }
            public decimal DtRate { get; set; }
        }

        private class SAICOutputRow
        {
            public string WeekEndingDate { get; set; } = "";
            public string Name { get; set; } = "";
            public string Invoice { get; set; } = "";
            public decimal AmountDue { get; set; }
            public decimal AggregateInvoiceLineItemAmount { get; set; }
            public decimal Tax { get; set; }
            public string Notes { get; set; } = "";
            public string InvoiceLineItemEndDate { get; set; } = "";
            public string Concat { get; set; } = "";
            public string saicInvoice { get; set; } = "";
            public string VmsIdentifier { get; set; } = "";
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