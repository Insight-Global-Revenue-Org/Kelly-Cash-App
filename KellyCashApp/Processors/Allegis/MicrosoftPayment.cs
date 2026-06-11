using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Services;
using System.Globalization;
using System.Text.RegularExpressions;

namespace KellyCashApp.Processors.Allegis
{
    internal class MicrosoftPayment
    {
        private const int HeaderRow = 6;
        private const int FirstDataRow = 10;

        public static bool IsMicrosoftFormat(IXLWorksheet worksheet)
        {
            return worksheet.Cell(HeaderRow, 4).GetString().Trim()
                .Equals("Consolidated Invoice ID", StringComparison.OrdinalIgnoreCase)
                && worksheet.Cell(HeaderRow, 11).GetString().Trim()
                .Equals("Worker", StringComparison.OrdinalIgnoreCase)
                && worksheet.Cell(HeaderRow, 13).GetString().Trim()
                .Equals("Invoice Line Item End Date", StringComparison.OrdinalIgnoreCase);
        }

        public static string Process(
            XLWorkbook workbook,
            IXLWorksheet worksheet,
            string inputPath,
            Dictionary<string, List<OirMatch>> openInvoiceMatches,
            Dictionary<string, List<MicrosoftVmsMatch>>? vmsMatches = null)
        {
            int microsoftInvoiceCol = FindColumn(worksheet, HeaderRow, "Consolidated Invoice ID");
            int workerCol = FindColumn(worksheet, HeaderRow, "Worker");
            int lineItemEndDateCol = FindColumn(worksheet, HeaderRow, "Invoice Line Item End Date");
            int aggregateAmountCol = FindColumn(worksheet, HeaderRow, "Total Invoice Line Item Amount (Supplier)");
            int taxCol = FindColumn(worksheet, HeaderRow, "Invoice Line Item Total Tax Amount (Supplier)");

            if (microsoftInvoiceCol == -1 || workerCol == -1 || lineItemEndDateCol == -1 || aggregateAmountCol == -1 || taxCol == -1)
                throw new Exception("Missing one or more required Microsoft remittance columns.");

            var oirRows = BuildOirRows(openInvoiceMatches);

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? FirstDataRow;

            var outputRows = new List<MicrosoftOutputRow>();

            int groupId = 0;

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

                groupId++;
                string formattedLineItemEndDate = lineItemEndDate.ToString("MM/dd/yyyy");


                decimal aggregateAmount = GetDecimalValue(worksheet.Cell(row, aggregateAmountCol));
                decimal tax = GetDecimalValue(worksheet.Cell(row, taxCol));
                decimal preTaxAggregateAmount = aggregateAmount - tax;
                string microsoftInvoice = worksheet.Cell(row, microsoftInvoiceCol).GetString().Trim();

                string vmsIdentifier =
                    new string(microsoftInvoice
                        .Where(char.IsDigit)
                        .ToArray());

                if (vmsIdentifier.Length > 5)
                    vmsIdentifier = vmsIdentifier[^5..];

                string vmsLookupKey = $"{name}|{vmsIdentifier}";

                MicrosoftVmsMatch? vmsMatch = null;

                if (vmsMatches != null &&
                    vmsMatches.TryGetValue(vmsLookupKey, out var foundVmsRows))
                {
                    vmsMatch = foundVmsRows.FirstOrDefault();
                }

                DateTime monthStart = new DateTime(
                    lineItemEndDate.Year,
                    lineItemEndDate.Month,
                    1);

                DateTime monthEnd = lineItemEndDate.AddDays(14);

                var possibleMatches = oirRows
                    .Where(x =>
                        x.Name.Equals(name, StringComparison.OrdinalIgnoreCase)
                        && x.WeekEndingDate >= monthStart
                        && x.WeekEndingDate <= monthEnd)
                    .OrderBy(x => x.WeekEndingDate)
                    .ToList();

                bool matchedWithoutTax = false;

                var matches = FindBestInvoiceCombination(possibleMatches, aggregateAmount);

                if (!matches.Any() && tax != 0)
                {
                    matches = FindBestInvoiceCombination(possibleMatches, preTaxAggregateAmount);

                    if (matches.Any())
                        matchedWithoutTax = true;
                }

                if (matches.Any())
                {
                    foreach (var match in matches)
                    {
                        outputRows.Add(new MicrosoftOutputRow
                        {
                            WeekEndingDate = match.WeekEndingDate.ToString("MM/dd/yyyy"),
                            Name = name,
                            Invoice = match.Invoice,
                            AmountDue = match.AmountDue,
                            InvoiceLineItemEndDate = formattedLineItemEndDate,
                            AggregateInvoiceLineItemAmount = aggregateAmount,
                            Tax = tax,
                            Notes = matchedWithoutTax
                                ? "Sales Tax was not included in Amount Due!"
                                : "",
                            GroupId = groupId,
                            Concat = $"{name} {formattedLineItemEndDate}",
                            MicrosoftInvoice = microsoftInvoice,
                            VmsIdentifier = vmsIdentifier,
                            AggregateInvoicedNet = vmsMatch?.AggregateInvoicedNet ?? 0,
                            Hours = vmsMatch?.Hours ?? 0,
                            RtRate = vmsMatch?.RtRate ?? 0,
                            OtRate = vmsMatch?.OtRate ?? 0,
                            DtRate = vmsMatch?.DtRate ?? 0
                        });
                    }
                }
                else
                {
                    outputRows.Add(new MicrosoftOutputRow
                    {
                        WeekEndingDate = "",
                        Name = name,
                        Invoice = "",
                        AmountDue = 0,
                        InvoiceLineItemEndDate = formattedLineItemEndDate,
                        AggregateInvoiceLineItemAmount = aggregateAmount,
                        Tax = tax,
                        Notes = "",
                        GroupId = groupId,
                        Concat = $"{name} {formattedLineItemEndDate}",
                        MicrosoftInvoice = microsoftInvoice,
                        VmsIdentifier = vmsIdentifier,
                        AggregateInvoicedNet = vmsMatch?.AggregateInvoicedNet ?? 0,
                        Hours = vmsMatch?.Hours ?? 0,
                        RtRate = vmsMatch?.RtRate ?? 0,
                        OtRate = vmsMatch?.OtRate ?? 0,
                        DtRate = vmsMatch?.DtRate ?? 0
                    });
                }
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

            string[] headers =
            {
                "Invoice Line Item End Date",
                "Week Ending Date",
                "Name",
                "Invoice",
                "Amount Due",
                "Aggregate Paid",
                "Tax",
                "Notes",
                "Concat",
                "Microsoft Invoice",
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
                worksheet.Cell(row, 2).Value = item.WeekEndingDate;
                worksheet.Cell(row, 3).Value = item.Name;
                worksheet.Cell(row, 4).Value = item.Invoice;
                worksheet.Cell(row, 5).Value = item.AmountDue;
                worksheet.Cell(row, 6).Value = item.AggregateInvoiceLineItemAmount;
                worksheet.Cell(row, 7).Value = item.Tax;
                worksheet.Cell(row, 8).Value = item.Notes;
                worksheet.Cell(row, 9).Value = item.Concat;
                worksheet.Cell(row, 10).Value = item.MicrosoftInvoice;
                worksheet.Cell(row, 11).Value = item.VmsIdentifier;
                worksheet.Cell(row, 12).Value = item.AggregateInvoicedNet;
                worksheet.Cell(row, 13).Value = item.Hours;
                worksheet.Cell(row, 14).Value = item.RtRate;
                worksheet.Cell(row, 15).Value = item.OtRate;
                worksheet.Cell(row, 16).Value = item.DtRate;

                worksheet.Row(row).AdjustToContents();
            }

            foreach (var group in outputRows.GroupBy(x => x.GroupId))
            {
                int firstOutputRow = outputRows.IndexOf(group.First()) + 2;
                int lastOutputRow = outputRows.IndexOf(group.Last()) + 2;

                if (lastOutputRow > firstOutputRow)
                {
                    worksheet.Range(firstOutputRow, 1, lastOutputRow, 1).Merge();
                    worksheet.Range(firstOutputRow, 6, lastOutputRow, 6).Merge();
                    worksheet.Range(firstOutputRow, 7, lastOutputRow, 7).Merge();
                    worksheet.Range(firstOutputRow, 10, lastOutputRow, 10).Merge();
                    worksheet.Range(firstOutputRow, 11, lastOutputRow, 11).Merge();
                    worksheet.Range(firstOutputRow, 12, lastOutputRow, 12).Merge();
                    worksheet.Range(firstOutputRow, 13, lastOutputRow, 13).Merge();
                    worksheet.Range(firstOutputRow, 14, lastOutputRow, 14).Merge();
                    worksheet.Range(firstOutputRow, 15, lastOutputRow, 15).Merge();
                    worksheet.Range(firstOutputRow, 16, lastOutputRow, 16).Merge();
                }

                worksheet.Range(firstOutputRow, 1, lastOutputRow, 16)
                    .Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            ApplyFormatting(worksheet, outputRows.Count + 1, headers.Length);

            string downloadsPath = Settings.GetRemittanceSavePath();

          
            string formattedTotal = total.ToString("$#,##0.00;($#,##0.00)", CultureInfo.InvariantCulture);
            string processedDate = DateTime.Now.ToString("M.d.yyyy", CultureInfo.InvariantCulture);

            string outputPath = GetUniqueOutputPath(downloadsPath, $"Microsoft {processedDate} - {formattedTotal}.xlsx");

            workbook.SaveAs(outputPath);

            Analytics.LogRemittanceRun($"Microsoft - {formattedTotal}");

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

            worksheet.Column(5).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            worksheet.Column(6).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            worksheet.Column(7).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

            for (int row = 2; row <= lastRow; row++)
            {
                worksheet.Row(row).Height = 13;
            }

            worksheet.Row(1).Height = 15;

            worksheet.Columns().AdjustToContents();

            worksheet.Column(1).Width = 22;
            worksheet.Column(2).Width = 16;
            worksheet.Column(3).Width = 22;
            worksheet.Column(4).Width = 18;
            worksheet.Column(5).Width = 18;
            worksheet.Column(6).Width = 28;
            worksheet.Column(7).Width = 14;
            worksheet.Column(8).Width = 42;
            worksheet.Column(8).Style.Alignment.WrapText = false;
            worksheet.Column(9).Width = 32;
            worksheet.Column(10).Width = 20;
            worksheet.Column(11).Width = 12;
            worksheet.Column(12).Width = 12;
            worksheet.Column(13).Width = 12;
            worksheet.Column(14).Width = 12;
            worksheet.Column(15).Width = 12;
            worksheet.Column(16).Width = 12;

            worksheet.Column(12).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            worksheet.Column(13).Style.NumberFormat.Format = "0.00";
            worksheet.Column(14).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            worksheet.Column(15).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            worksheet.Column(16).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

            for (int row = 2; row <= lastRow; row++)
            {
                decimal amountDue = GetDecimalValue(worksheet.Cell(row, 5));

                if (amountDue <= 0)
                {
                    worksheet.Cell(row, 5).Style.Font.FontColor = XLColor.Red;

                    worksheet.Range(row, 1, row, 16)
                        .Style.Fill.BackgroundColor = XLColor.FromHtml("#F2F2F2");
                }
            }

            int totalRow = lastRow + 1;

            var totalCell = worksheet.Cell(totalRow, 6);

            totalCell.FormulaA1 = $"=SUM(F2:F{lastRow})";
            totalCell.Style.Font.FontName = "Aptos Narrow";
            totalCell.Style.Font.FontSize = 9;
            totalCell.Style.Font.Bold = true;
            totalCell.Style.Fill.BackgroundColor = XLColor.Yellow;
            totalCell.Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
            totalCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

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

        private class MicrosoftOutputRow
        {
            public string WeekEndingDate { get; set; } = "";
            public string Name { get; set; } = "";
            public string Invoice { get; set; } = "";
            public decimal AmountDue { get; set; }
            public decimal AggregateInvoiceLineItemAmount { get; set; }
            public decimal Tax { get; set; }
            public string Notes { get; set; } = "";
            public string InvoiceLineItemEndDate { get; set; } = "";
            public int GroupId { get; set; }
            public string Concat { get; set; } = "";
            public string MicrosoftInvoice { get; set; } = "";
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