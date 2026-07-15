using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Services;
using System.Globalization;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;

namespace KellyCashApp.Processors.Leidos
{
    internal class LeidosPayment
    {
        public static bool IsLeidosFormat(string inputPath)
        {
            if (!Path.GetExtension(inputPath)
                .Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            try
            {
                using PdfDocument document = PdfDocument.Open(inputPath);

                if (document.NumberOfPages < 3)
                {
                    return false;
                }

                string thirdPageText = document
                    .GetPage(3)
                    .Text;

                return thirdPageText.Contains(
                           "leidos",
                           StringComparison.OrdinalIgnoreCase)
                       &&
                       thirdPageText.Contains(
                           "CONTROL NO",
                           StringComparison.OrdinalIgnoreCase)
                       &&
                       thirdPageText.Contains(
                           "AMOUNT PAID",
                           StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        public static string Process(
            string inputPath,
            Dictionary<string, OirMatch> openInvoiceMatches)
        {
            var outputRows = new List<LeidosOutputRow>();

            // The OIR dictionary is normally keyed by Name + Date.
            // Leidos is invoice-based, so create an index using DocumentNumber.
            Dictionary<string, OirMatch> matchesByInvoice =
                openInvoiceMatches.Values
                    .Where(x => !string.IsNullOrWhiteSpace(x.DocumentNumber))
                    .GroupBy(
                        x => NormalizeInvoiceNumber(x.DocumentNumber),
                        StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(
                        group => group.Key,
                        group => group.First(),
                        StringComparer.OrdinalIgnoreCase);

            using PdfDocument document = PdfDocument.Open(inputPath);

            foreach (var page in document.GetPages())
            {
                string text = NormalizePdfText(page.Text);

                /*
                 * Expected Leidos row:
                 *
                 * 05/04/2026 SLICS9000260578 6594246
                 * $6,279.20 $0.00 $6,279.20
                 */
                string amountPattern =
                    @"(?:-?\s*\$?[\d,]+\.\d{2}|\(\s*\$?[\d,]+\.\d{2}\s*\))";

                string rowPattern =
                    $@"(?<date>\d{{1,2}}/\d{{1,2}}/\d{{2,4}})\s+" +
                    @"(?<invoice>[A-Z0-9\-]+)\s+" +
                    @"(?<control>\d+)\s+" +
                    $@"(?<amount>{amountPattern})\s+" +
                    $@"(?<discount>{amountPattern})\s+" +
                    $@"(?<paid>{amountPattern})";

                MatchCollection matches = Regex.Matches(
                    text,
                    rowPattern,
                    RegexOptions.IgnoreCase);

                foreach (Match match in matches)
                {
                    string invoiceDate =
                        FormatDate(match.Groups["date"].Value);

                    string invoiceNumber =
                        match.Groups["invoice"].Value.Trim();

                    string controlNumber =
                        match.Groups["control"].Value.Trim();

                    decimal invoiceAmount =
                        GetDecimalValue(match.Groups["amount"].Value);

                    decimal discount =
                        GetDecimalValue(match.Groups["discount"].Value);

                    decimal amountPaid =
                        GetDecimalValue(match.Groups["paid"].Value);

                    string matchedInvoice = "";
                    decimal amountDue = 0;

                    string normalizedInvoice =
                        NormalizeInvoiceNumber(invoiceNumber);

                    if (matchesByInvoice.TryGetValue(
                        normalizedInvoice,
                        out OirMatch? oirMatch))
                    {
                        matchedInvoice = oirMatch.DocumentNumber;
                        amountDue = oirMatch.RemainingAmount;
                    }

                    outputRows.Add(new LeidosOutputRow
                    {
                        InvoiceDate = invoiceDate,
                        InvoiceNumber = invoiceNumber,
                        ControlNumber = controlNumber,
                        Invoice = matchedInvoice,
                        AmountDue = amountDue,
                        InvoiceAmount = invoiceAmount,
                        Discount = discount,
                        AmountPaid = amountPaid,
                        Notes = ""
                    });
                }
            }

            if (outputRows.Count == 0)
            {
                throw new InvalidOperationException(
                    "No Leidos payment rows were found in the PDF.");
            }

            outputRows = outputRows
                .OrderBy(x => x.InvoiceDate)
                .ThenBy(x => x.InvoiceNumber)
                .ToList();

            using var workbook = new XLWorkbook();

            IXLWorksheet worksheet =
                workbook.Worksheets.Add("Leidos Payment");

            string[] headers =
            {
                "Invoice Date",
                "Invoice Number",
                "Control Number",
                "Invoice",
                "Amount Due",
                "Invoice Amount",
                "Discount",
                "Amount Paid",
                "Notes"
            };

            for (int column = 1; column <= headers.Length; column++)
            {
                worksheet.Cell(1, column).Value = headers[column - 1];
            }

            for (int index = 0; index < outputRows.Count; index++)
            {
                int row = index + 2;
                LeidosOutputRow item = outputRows[index];

                worksheet.Cell(row, 1).Value = item.InvoiceDate;
                worksheet.Cell(row, 2).Value = item.InvoiceNumber;
                worksheet.Cell(row, 3).Value = item.ControlNumber;
                worksheet.Cell(row, 4).Value = item.Invoice;
                worksheet.Cell(row, 5).Value = item.AmountDue;
                worksheet.Cell(row, 6).Value = item.InvoiceAmount;
                worksheet.Cell(row, 7).Value = item.Discount;
                worksheet.Cell(row, 8).Value = item.AmountPaid;
                worksheet.Cell(row, 9).Value = item.Notes;
            }

            ApplyFormatting(
                worksheet,
                outputRows.Count + 1,
                headers.Length);

            string downloadsPath =
                Settings.GetRemittanceSavePath();

            decimal total =
                outputRows.Sum(x => x.AmountPaid);

            string formattedTotal = total.ToString(
                "$#,##0.00;($#,##0.00)",
                CultureInfo.InvariantCulture);

            string processedDate = DateTime.Now.ToString(
                "M.d.yyyy",
                CultureInfo.InvariantCulture);

            string outputPath = GetUniqueOutputPath(
                downloadsPath,
                $"Leidos {processedDate} - {formattedTotal}.xlsx");

            workbook.SaveAs(outputPath);

            Analytics.LogRemittanceRun(
                $"Leidos - {formattedTotal}");

            return outputPath;
        }

        private static void ApplyFormatting(
            IXLWorksheet worksheet,
            int lastRow,
            int lastColumn)
        {
            IXLRange range =
                worksheet.Range(1, 1, lastRow, lastColumn);

            range.Style.Font.FontName = "Aptos Narrow";
            range.Style.Font.FontSize = 9;

            range.Style.Border.OutsideBorder =
                XLBorderStyleValues.Thin;

            range.Style.Border.InsideBorder =
                XLBorderStyleValues.Thin;

            range.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;

            worksheet.Row(1).Style.Font.Bold = true;

            worksheet.Row(1)
                .Style
                .Fill
                .BackgroundColor = XLColor.FromHtml("#FCE4D6");

            worksheet.Column(5).Style.NumberFormat.Format =
                "$#,##0.00;($#,##0.00)";

            worksheet.Column(6).Style.NumberFormat.Format =
                "$#,##0.00;($#,##0.00)";

            worksheet.Column(7).Style.NumberFormat.Format =
                "$#,##0.00;($#,##0.00)";

            worksheet.Column(8).Style.NumberFormat.Format =
                "$#,##0.00;($#,##0.00)";

            for (int row = 2; row <= lastRow; row++)
            {
                string invoice =
                    worksheet.Cell(row, 4).GetString().Trim();

                decimal amountDue =
                    GetDecimalValue(
                        worksheet.Cell(row, 5).Value.ToString());

                decimal amountPaid =
                    GetDecimalValue(
                        worksheet.Cell(row, 8).Value.ToString());

                // Gray rows where no OIR invoice match was found.
                if (string.IsNullOrWhiteSpace(invoice))
                {
                    worksheet
                        .Range(row, 1, row, lastColumn)
                        .Style
                        .Fill
                        .BackgroundColor = XLColor.FromHtml("#F2F2F2");
                }

                if (amountDue <= 0)
                {
                    worksheet.Cell(row, 5)
                        .Style.Font.FontColor = XLColor.Red;
                }

                if (amountPaid <= 0)
                {
                    worksheet.Cell(row, 8)
                        .Style.Font.FontColor = XLColor.Red;
                }
            }

            worksheet.Columns().AdjustToContents();

            worksheet.Column(2).Width = 22;
            worksheet.Column(3).Width = 16;
            worksheet.Column(4).Width = 22;
            worksheet.Column(9).Width = 30;

            worksheet
                .Range(1, 1, lastRow, lastColumn)
                .SetAutoFilter();

            worksheet.SheetView.FreezeRows(1);
        }

        private static string NormalizePdfText(string text)
        {
            return Regex.Replace(
                    text,
                    @"\s+",
                    " ")
                .Trim();
        }

        private static string NormalizeInvoiceNumber(string value)
        {
            return Regex.Replace(
                    value ?? "",
                    @"[^A-Z0-9]",
                    "",
                    RegexOptions.IgnoreCase)
                .ToUpperInvariant();
        }

        private static string FormatDate(string value)
        {
            if (DateTime.TryParse(
                value,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out DateTime date))
            {
                return date.ToString(
                    "MM/dd/yyyy",
                    CultureInfo.InvariantCulture);
            }

            return value.Trim();
        }

        private static decimal GetDecimalValue(string value)
        {
            bool isNegative =
                value.Contains('(') && value.Contains(')');

            string raw = value
                .Replace("$", "")
                .Replace(",", "")
                .Replace("(", "")
                .Replace(")", "")
                .Trim();

            if (!decimal.TryParse(
                raw,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out decimal result))
            {
                return 0;
            }

            return isNegative ? -Math.Abs(result) : result;
        }

        private static string GetUniqueOutputPath(
            string folderPath,
            string fileName)
        {
            string name =
                Path.GetFileNameWithoutExtension(fileName);

            string extension =
                Path.GetExtension(fileName);

            string path =
                Path.Combine(folderPath, fileName);

            int counter = 1;

            while (File.Exists(path))
            {
                path = Path.Combine(
                    folderPath,
                    $"{name} ({counter}){extension}");

                counter++;
            }

            return path;
        }

        private sealed class LeidosOutputRow
        {
            public string InvoiceDate { get; set; } = "";

            public string InvoiceNumber { get; set; } = "";

            public string ControlNumber { get; set; } = "";

            public string Invoice { get; set; } = "";

            public decimal AmountDue { get; set; }

            public decimal InvoiceAmount { get; set; }

            public decimal Discount { get; set; }

            public decimal AmountPaid { get; set; }

            public string Notes { get; set; } = "";
        }
    }
}