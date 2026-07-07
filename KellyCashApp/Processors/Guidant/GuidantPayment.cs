using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Services;
using System.Globalization;

namespace KellyCashApp.Processors.Guidant
{
    internal class GuidantPayment
    {
        public static bool IsGuidantFormat(IXLWorksheet worksheet)
        {
            return FindGuidantHeaderRow(worksheet) != -1;
        }

        public static string Process(
            XLWorkbook workbook,
            IXLWorksheet worksheet,
            string inputPath,
            Dictionary<string, OirMatch> openInvoiceMatches)
        {
            decimal totalAcrossAllSheets = 0;

            var guidantSheets = workbook.Worksheets
                .Where(sheet => IsGuidantFormat(sheet))
                .ToList();

            foreach (var sheet in guidantSheets)
            {
                decimal sheetTotal = ProcessSingleSheet(sheet, openInvoiceMatches);
                totalAcrossAllSheets += sheetTotal;
            }

            string downloadsPath = Settings.GetRemittanceSavePath();

            string formattedTotal = totalAcrossAllSheets.ToString("N2", CultureInfo.InvariantCulture);

            string outputPath = GetUniqueOutputPath(
                downloadsPath,
                $"Guidant - {formattedTotal}.xlsx"
            );

            workbook.SaveAs(outputPath);

            Analytics.LogRemittanceRun($"Guidant - {formattedTotal}");

            return outputPath;
        }

        private static decimal ProcessSingleSheet(
            IXLWorksheet worksheet,
            Dictionary<string, OirMatch> openInvoiceMatches)
        {
            int headerRow = FindGuidantHeaderRow(worksheet);

            if (headerRow == -1)
            {
                throw new Exception($"Could not find Guidant header row on sheet: {worksheet.Name}");
            }

            int invoiceColumn = FindColumn(worksheet, headerRow, "Invoice");
            int invoiceIdColumn = FindColumn(worksheet, headerRow, "Invoice ID");
            int timesheetColumn = FindColumn(worksheet, headerRow, "Timesheet");
            int weekEndingColumn = FindColumn(worksheet, headerRow, "WeekEnding");
            int associateColumn = FindColumn(worksheet, headerRow, "Associate");
            int hoursTypeColumn = FindColumn(worksheet, headerRow, "Hours Type");
            int billHoursColumn = FindColumn(worksheet, headerRow, "Bill Hours");
            int billAmountColumn = FindColumn(worksheet, headerRow, "Bill Amount");

            if (invoiceColumn == -1 || invoiceIdColumn == -1 || timesheetColumn == -1 ||
                weekEndingColumn == -1 || associateColumn == -1 || hoursTypeColumn == -1 ||
                billHoursColumn == -1 || billAmountColumn == -1)
            {
                throw new Exception($"Missing one or more required Guidant columns on sheet: {worksheet.Name}");
            }

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            var nameChanges = Rename.LoadNameChanges();
            var outputRows = new List<GuidantOutputRow>();

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string rawName = worksheet.Cell(row, associateColumn).GetString().Trim();

                if (string.IsNullOrWhiteSpace(rawName))
                    continue;

                string formattedName = FormatName(rawName);
                formattedName = Rename.ApplyNameChange(formattedName, nameChanges);

                string formattedWeekEnding = FormatDate(worksheet.Cell(row, weekEndingColumn));
                string concat = $"{formattedName} {formattedWeekEnding}".Trim();

                string oirInvoice = "";
                decimal amountDue = 0;

                string hoursType = worksheet.Cell(row, hoursTypeColumn).GetString().Trim();
                decimal billAmount = GetDecimalValue(worksheet.Cell(row, billAmountColumn));

                bool matched = hoursType.Equals("EXP", StringComparison.OrdinalIgnoreCase)
                    ? TryMatchExpenseWithSevenDaySpread(formattedName, formattedWeekEnding, billAmount, openInvoiceMatches, out OirMatch match)
                    : TryMatchWithOneDaySpread(formattedName, formattedWeekEnding, openInvoiceMatches, out match);

                if (matched)
                {
                    oirInvoice = match.DocumentNumber;
                    amountDue = match.RemainingAmount;
                }

                outputRows.Add(new GuidantOutputRow
                {
                    WeekEndingDate = formattedWeekEnding,
                    Name = formattedName,
                    Invoice = oirInvoice,
                    AmountDue = amountDue,
                    AggregateAmountPaid = billAmount,
                    Notes = "",
                    Concat = concat,
                    GuidantInvoice = worksheet.Cell(row, invoiceColumn).GetString().Trim(),
                    InvoiceId = worksheet.Cell(row, invoiceIdColumn).GetString().Trim(),
                    TimesheetId = worksheet.Cell(row, timesheetColumn).GetString().Trim(),
                    HoursType = hoursType,
                    BillHours = GetDecimalValue(worksheet.Cell(row, billHoursColumn))
                });
            }

            var totalsByConcat = outputRows
                .GroupBy(r => r.Concat)
                .ToDictionary(g => g.Key, g => g.Sum(r => r.AggregateAmountPaid));

            foreach (var row in outputRows)
            {
                row.AggregateAmountPaid = totalsByConcat[row.Concat];
            }

            // sort alphabetically :)
            outputRows = outputRows
                .OrderBy(r => r.Name)
                .ThenBy(r => r.WeekEndingDate)
                .ToList();

            worksheet.Clear();

            string[] headers =
            {
                "Week Ending Date",
                "Name",
                "Invoice",
                "Amount Due",
                "Aggregate Amount Paid",
                "Notes",
                "Concat",
                "Hours Type",
                "Guidant Invoice",
                "Invoice ID",
                "Timesheet ID",
                "Bill Hours"
            };

            for (int col = 1; col <= headers.Length; col++)
                worksheet.Cell(1, col).Value = headers[col - 1];

            for (int i = 0; i < outputRows.Count; i++)
            {
                int row = i + 2;
                GuidantOutputRow item = outputRows[i];

                worksheet.Cell(row, 1).Value = item.WeekEndingDate;
                worksheet.Cell(row, 2).Value = item.Name;
                worksheet.Cell(row, 3).Value = item.Invoice;

                if (item.AmountDue != 0)
                    worksheet.Cell(row, 4).Value = item.AmountDue;

                worksheet.Cell(row, 5).Value = item.AggregateAmountPaid;
                worksheet.Cell(row, 6).Value = item.Notes;
                worksheet.Cell(row, 7).Value = item.Concat;
                worksheet.Cell(row, 8).Value = item.HoursType;
                worksheet.Cell(row, 9).Value = item.GuidantInvoice;
                worksheet.Cell(row, 10).Value = item.InvoiceId;
                worksheet.Cell(row, 11).Value = item.TimesheetId;
                worksheet.Cell(row, 12).Value = item.BillHours;
            }

            ApplyFormatting(worksheet, outputRows.Count + 1, headers.Length);

            return outputRows.Sum(r => r.AggregateAmountPaid);
        }

        private static bool TryMatchWithOneDaySpread(
            string formattedName,
            string formattedWeekEnding,
            Dictionary<string, OirMatch> openInvoiceMatches,
            out OirMatch match)
        {
            match = null!;

            if (!DateTime.TryParse(formattedWeekEnding, out DateTime baseDate))
                return false;

            for (int offset = -1; offset <= 1; offset++)
            {
                DateTime testDate = baseDate.AddDays(offset);
                string testDateString = testDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                string testKey = $"{formattedName} {testDateString}".Trim();

                if (openInvoiceMatches.TryGetValue(testKey, out match))
                    return true;
            }

            return false;
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

        private static string FormatDate(IXLCell cell)
        {
            if (cell.Value.IsDateTime)
                return cell.GetDateTime().ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            string rawValue = cell.GetString().Trim();

            if (DateTime.TryParse(rawValue, out DateTime parsedDate))
                return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            return rawValue;
        }

        // Expense matching allows for a ±7 day spread and a 5% margin of error in the amount due.
        private static bool TryMatchExpenseWithSevenDaySpread(
            string formattedName,
            string formattedWeekEnding,
            decimal billAmount,
            Dictionary<string, OirMatch> openInvoiceMatches,
            out OirMatch match)
        {
            match = null!;

            if (!DateTime.TryParse(formattedWeekEnding, out DateTime baseDate))
                return false;

            // First pass: exact amount match within ±7 days
            for (int offset = -7; offset <= 7; offset++)
            {
                DateTime testDate = baseDate.AddDays(offset);
                string testDateString = testDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                string testKey = $"{formattedName} {testDateString}".Trim();

                if (openInvoiceMatches.TryGetValue(testKey, out OirMatch exactMatch)
                    && Math.Abs(exactMatch.RemainingAmount - billAmount) == 0)
                {
                    match = exactMatch;
                    return true;
                }
            }

            // Second pass: within 5% margin within ±7 days
            for (int offset = -7; offset <= 7; offset++)
            {
                DateTime testDate = baseDate.AddDays(offset);
                string testDateString = testDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                string testKey = $"{formattedName} {testDateString}".Trim();

                if (openInvoiceMatches.TryGetValue(testKey, out OirMatch marginMatch)
                    && IsWithinFivePercent(marginMatch.RemainingAmount, billAmount))
                {
                    match = marginMatch;
                    return true;
                }
            }

            return false;
        }

        // Check if the oirAmount is within 5% of the paymentAmount
        private static bool IsWithinFivePercent(decimal oirAmount, decimal paymentAmount)
        {
            if (paymentAmount == 0)
                return oirAmount == 0;

            decimal difference = Math.Abs(oirAmount - paymentAmount);
            decimal allowedDifference = Math.Abs(paymentAmount) * 0.05m;

            return difference <= allowedDifference;
        }

        private static int FindGuidantHeaderRow(IXLWorksheet worksheet)
        {
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 10;

            for (int row = 1; row <= Math.Min(lastRow, 10); row++)
            {
                string firstHeader = worksheet.Cell(row, 1).GetString().Trim();
                string secondHeader = worksheet.Cell(row, 2).GetString().Trim();
                string thirdHeader = worksheet.Cell(row, 3).GetString().Trim();

                bool isGuidantHeader =
                    firstHeader.Equals("Customer", StringComparison.OrdinalIgnoreCase)
                    && secondHeader.Equals("Vendor", StringComparison.OrdinalIgnoreCase)
                    && thirdHeader.Equals("Con Invoice", StringComparison.OrdinalIgnoreCase);

                if (isGuidantHeader)
                    return row;
            }

            return -1;
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

        private static decimal GetDecimalValue(IXLCell cell)
        {
            string rawValue = cell.Value.ToString()
                .Replace("$", "")
                .Replace(",", "")
                .Replace("(", "-")
                .Replace(")", "")
                .Trim();

            if (decimal.TryParse(rawValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return 0;
        }

        private static void ApplyFormatting(IXLWorksheet worksheet, int lastRow, int lastColumn)
        {
            var range = worksheet.Range(1, 1, lastRow, lastColumn);

            // Main body formatting
            range.Style.Font.FontName = "Aptos Narrow";
            range.Style.Font.FontSize = 8;
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // Header row formatting
            worksheet.Row(1).Style.Font.FontName = "Aptos Narrow";
            worksheet.Row(1).Style.Font.FontSize = 9;
            worksheet.Row(1).Style.Font.Bold = true;
            worksheet.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FCE4D6");

            // Row heights
            for (int row = 2; row <= lastRow; row++)
            {
                worksheet.Row(row).Height = 13;
            }

            worksheet.Row(1).Height = 15;

            // Number formatting
            worksheet.Column(4).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)"; // Amount Due
            worksheet.Column(5).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)"; // Aggregate Amount Paid
            worksheet.Column(12).Style.NumberFormat.Format = "0.00";                 // Bill Hours

            // Make zero/negative aggregate payments red
            for (int row = 2; row <= lastRow; row++)
            {
                decimal aggregateAmountPaid = GetDecimalValue(worksheet.Cell(row, 5));

                if (aggregateAmountPaid <= 0)
                {
                    worksheet.Cell(row, 5).Style.Font.FontColor = XLColor.Red;
                }
            }

            // Center identifier/helper columns
            worksheet.Column(8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;  // Hours Type
            worksheet.Column(9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;  // Guidant Invoice
            worksheet.Column(10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Invoice ID
            worksheet.Column(11).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Timesheet ID
            worksheet.Column(12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Bill Hours

            // Freeze top row
            worksheet.SheetView.FreezeRows(1);

            // Auto-fit first, then manually widen important columns
            worksheet.Columns().AdjustToContents();

            worksheet.Column(3).Width = 16;   // Invoice
            worksheet.Column(4).Width = 14;   // Amount Due
            worksheet.Column(5).Width = 20;   // Aggregate Amount Paid
            worksheet.Column(6).Width = 22;   // Notes
            worksheet.Column(7).Width = 28;   // Concat
            worksheet.Column(8).Width = 14;   // Hours Type
            worksheet.Column(9).Width = 18;   // Guidant Invoice
            worksheet.Column(10).Width = 18;  // Invoice ID
            worksheet.Column(11).Width = 18;  // Timesheet ID
            worksheet.Column(12).Width = 12;  // Bill Hours

            // Filter
            worksheet.Range(1, 1, lastRow, lastColumn).SetAutoFilter();
        }

        private static string GetUniqueOutputPath(string folderPath, string fileName)
        {
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);

            string outputPath = Path.Combine(folderPath, fileName);
            int counter = 1;

            while (File.Exists(outputPath))
            {
                outputPath = Path.Combine(folderPath, $"{fileNameWithoutExt} ({counter}){extension}");
                counter++;
            }

            return outputPath;
        }

        private class GuidantOutputRow
        {
            public string WeekEndingDate { get; set; } = "";
            public string Name { get; set; } = "";
            public string Invoice { get; set; } = "";
            public decimal AmountDue { get; set; }
            public decimal AggregateAmountPaid { get; set; }
            public string Notes { get; set; } = "";
            public string Concat { get; set; } = "";
            public string GuidantInvoice { get; set; } = "";
            public string InvoiceId { get; set; } = "";
            public string TimesheetId { get; set; } = "";
            public string HoursType { get; set; } = "";
            public decimal BillHours { get; set; }
        }
    }
}