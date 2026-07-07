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
            return worksheet.Cell(1, 1).GetString().Trim().Equals("Customer", StringComparison.OrdinalIgnoreCase)
                && worksheet.Cell(1, 2).GetString().Trim().Equals("Vendor", StringComparison.OrdinalIgnoreCase)
                && worksheet.Cell(1, 3).GetString().Trim().Equals("Con Invoice", StringComparison.OrdinalIgnoreCase);
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
            int headerRow = 1;

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

                if (TryMatchWithOneDaySpread(formattedName, formattedWeekEnding, openInvoiceMatches, out OirMatch match))
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
                    AggregateAmountPaid = GetDecimalValue(worksheet.Cell(row, billAmountColumn)),
                    Notes = "",
                    Concat = concat,
                    GuidantInvoice = worksheet.Cell(row, invoiceColumn).GetString().Trim(),
                    InvoiceId = worksheet.Cell(row, invoiceIdColumn).GetString().Trim(),
                    TimesheetId = worksheet.Cell(row, timesheetColumn).GetString().Trim(),
                    HoursType = worksheet.Cell(row, hoursTypeColumn).GetString().Trim(),
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
                "Guidant Invoice",
                "Invoice ID",
                "Timesheet ID",
                "Hours Type",
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
                worksheet.Cell(row, 8).Value = item.GuidantInvoice;
                worksheet.Cell(row, 9).Value = item.InvoiceId;
                worksheet.Cell(row, 10).Value = item.TimesheetId;
                worksheet.Cell(row, 11).Value = item.HoursType;
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
            worksheet.Column(8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;  // Guidant Invoice
            worksheet.Column(9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;  // Invoice ID
            worksheet.Column(10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Timesheet ID
            worksheet.Column(11).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Hours Type
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
            worksheet.Column(8).Width = 18;   // Guidant Invoice
            worksheet.Column(9).Width = 18;   // Invoice ID
            worksheet.Column(10).Width = 18;  // Timesheet ID
            worksheet.Column(11).Width = 14;  // Hours Type
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