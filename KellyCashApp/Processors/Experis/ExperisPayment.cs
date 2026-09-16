using ClosedXML.Excel;
using HtmlAgilityPack;
using KellyCashApp.Configuration;
using KellyCashApp.Services;
using MimeKit;
using System.Globalization;

namespace KellyCashApp.Processors.Experis
{
    internal static class ExperisPayment
    {
        // =============================================================
        // FORMAT DETECTION
        // =============================================================

        public static bool IsExperisFormat(string inputPath)
        {
            // Experis is currently the only payment processor
            // that accepts an .eml file.
            return Path.GetExtension(inputPath)
                .Equals(
                    ".eml",
                    StringComparison.OrdinalIgnoreCase);
        }


        // =============================================================
        // MAIN PROCESSOR
        // =============================================================

        public static string Process(string inputPath)
        {
            // ---------------------------------------------------------
            // 1. Load the EML file.
            // ---------------------------------------------------------

            MimeMessage message =
                MimeMessage.Load(inputPath);

            string htmlBody =
                message.HtmlBody ?? "";

            if (string.IsNullOrWhiteSpace(htmlBody))
            {
                throw new Exception(
                    "The Experis EML file did not contain an HTML email body.");
            }


            // ---------------------------------------------------------
            // 2. Extract all Experis invoice rows.
            // ---------------------------------------------------------

            List<ExperisPaymentRow> paymentRows =
                ExtractPaymentRows(htmlBody);

            if (paymentRows.Count == 0)
            {
                throw new Exception(
                    "No Experis invoice rows were found in the EML file.");
            }

            // ---------------------------------------------------------
            // Load optional Experis End Client mapping file.
            // ---------------------------------------------------------

            List<ExperisEndClientMapping> endClientMappings =
                LoadEndClientMappings();

            // ---------------------------------------------------------
            // 3. Create the output Excel workbook.
            // ---------------------------------------------------------

            using var workbook =
                new XLWorkbook();

            var worksheet =
                workbook.Worksheets.Add(
                    "Payment Details");


            // ---------------------------------------------------------
            // 4. Write standardized headers.
            // ---------------------------------------------------------

            string[] headers =
{
                    "Experis End Client",
                    "Experis Invoice Number",
                    "Payment Reference",
                    "Contractor Name",
                    "Week Ending Date",
                    "Invoice",
                    "Amount Due",
                    "Aggregate Amount Paid",
                    "Notes",
                    "Concat"
                };

            for (int col = 1;
                 col <= headers.Length;
                 col++)
            {
                worksheet.Cell(1, col).Value =
                    headers[col - 1];
            }


            // ---------------------------------------------------------
            // 5. Write the EML data into the output worksheet.
            // ---------------------------------------------------------

            for (int i = 0;
     i < paymentRows.Count;
     i++)
            {
                int outputRow = i + 2;

                ExperisPaymentRow paymentRow =
                    paymentRows[i];


                // -----------------------------------------------------
                // Determine Experis End Client from invoice identifier.
                // -----------------------------------------------------

                string endClient =
                    FindEndClient(
                        paymentRow.InvoiceNumber,
                        paymentRow.PaymentReference,
                        endClientMappings);


                // Experis End Client
                worksheet.Cell(outputRow, 1).Value =
                    endClient;

                // Experis Invoice Number
                worksheet.Cell(outputRow, 2).Value =
                    paymentRow.InvoiceNumber;

                // Payment Reference
                worksheet.Cell(outputRow, 3).Value =
                    paymentRow.PaymentReference;

                // Contractor Name
                worksheet.Cell(outputRow, 4).Value =
                    "";

                // Week Ending Date
                worksheet.Cell(outputRow, 5).Value =
                    "";

                // OIR Invoice
                worksheet.Cell(outputRow, 6).Value =
                    "";

                // Amount Due
                worksheet.Cell(outputRow, 7).Value =
                    "";

                // Aggregate Amount Paid
                worksheet.Cell(outputRow, 8).Value =
                    paymentRow.PaidAmount;

                // Notes
                worksheet.Cell(outputRow, 9).Value =
                    "";

                // Concat
                worksheet.Cell(outputRow, 10).Value =
                    "";
            }


            // ---------------------------------------------------------
            // 6. Apply formatting.
            // ---------------------------------------------------------

            ApplyFormatting(
                worksheet,
                paymentRows.Count + 1,
                headers.Length);


            // ---------------------------------------------------------
            // 7. Calculate total paid amount.
            // ---------------------------------------------------------

            decimal total =
                paymentRows.Sum(
                    x => x.PaidAmount);


            // ---------------------------------------------------------
            // 8. Create output filename.
            // ---------------------------------------------------------

            string savePath =
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
                    savePath,
                    $"Experis {processedDate} - {formattedTotal}.xlsx");


            // ---------------------------------------------------------
            // 9. Save and log.
            // ---------------------------------------------------------

            workbook.SaveAs(outputPath);

            Analytics.LogRemittanceRun(
                $"Experis - {formattedTotal}");

            return outputPath;
        }


        // =============================================================
        // EML / HTML PARSER
        // =============================================================

        private static List<ExperisPaymentRow> ExtractPaymentRows(
            string html)
        {
            var results =
                new List<ExperisPaymentRow>();

            var htmlDocument =
                new HtmlAgilityPack.HtmlDocument();

            htmlDocument.LoadHtml(html);


            // ---------------------------------------------------------
            // Every Experis remittance page contains a table whose
            // header includes both:
            //
            // Invoice Number
            // Paid Amount
            //
            // We locate all tables in the HTML and only process tables
            // containing those two headers.
            // ---------------------------------------------------------

            var tables =
                htmlDocument.DocumentNode
                    .SelectNodes("//table");

            if (tables == null)
                return results;


            foreach (HtmlNode table in tables)
            {
                string tableText =
                    HtmlEntity.DeEntitize(
                        table.InnerText);

                if (!tableText.Contains(
                        "Invoice Number",
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!tableText.Contains(
                        "Paid Amount",
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }


                // -----------------------------------------------------
                // The actual invoice records live inside a nested
                // five-column table:
                //
                // 1 Invoice Number
                // 2 Invoice Date
                // 3 Payment Reference
                // 4 Gross Amount
                // 5 Paid Amount
                // -----------------------------------------------------

                var rows =
                    table.SelectNodes(
                        ".//table//tr");

                if (rows == null)
                    continue;


                foreach (HtmlNode row in rows)
                {
                    var cells =
                        row.SelectNodes(
                            "./td");

                    if (cells == null ||
                        cells.Count != 5)
                    {
                        continue;
                    }


                    string invoiceNumber =
                        CleanCellText(
                            cells[0]);

                    string paymentReference =
                        CleanCellText(
                            cells[2]);

                    string paidAmountText =
                        CleanCellText(
                            cells[4]);


                    // Ignore headers or malformed rows.
                    if (string.IsNullOrWhiteSpace(
                        invoiceNumber))
                    {
                        continue;
                    }


                    if (!TryParseMoney(
                        paidAmountText,
                        out decimal paidAmount))
                    {
                        continue;
                    }


                    results.Add(
                    new ExperisPaymentRow
                    {
                        InvoiceNumber =
                            invoiceNumber,

                        PaymentReference =
                            paymentReference,

                        PaidAmount =
                            paidAmount
                    });
                }
            }


            return results;
        }


        // =============================================================
        // HTML HELPERS
        // =============================================================

        private static string CleanCellText(
            HtmlNode cell)
        {
            string value =
                HtmlEntity.DeEntitize(
                    cell.InnerText);

            return value
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();
        }


        private static bool TryParseMoney(
            string input,
            out decimal value)
        {
            string cleaned =
                input
                    .Replace("$", "")
                    .Replace(",", "")
                    .Replace("(", "-")
                    .Replace(")", "")
                    .Trim();

            return decimal.TryParse(
                cleaned,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out value);
        }


        // =============================================================
        // EXCEL FORMATTING
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


            // Header.
            worksheet.Row(1)
                .Style.Font.Bold = true;

            worksheet.Row(1)
                .Style.Fill.BackgroundColor =
                XLColor.FromHtml("#FCE4D6");

            worksheet.Row(1).Height =
                15;


            // Amount Due
            worksheet.Column(7)
                .Style.NumberFormat.Format =
                "$#,##0.00;[Red]($#,##0.00)";

            // Aggregate Amount Paid
            worksheet.Column(8)
                .Style.NumberFormat.Format =
                "$#,##0.00;[Red]($#,##0.00)";


            for (int row = 2;
                 row <= lastRow;
                 row++)
            {
                worksheet.Row(row).Height =
                    13;
            }


            worksheet.Columns()
                .AdjustToContents();


            worksheet.Column(1).Width = 30; // Experis End Client
            worksheet.Column(2).Width = 28; // Experis Invoice Number
            worksheet.Column(3).Width = 28; // Payment Reference
            worksheet.Column(4).Width = 24; // Contractor Name
            worksheet.Column(5).Width = 18; // Week Ending Date
            worksheet.Column(6).Width = 18; // Invoice
            worksheet.Column(7).Width = 18; // Amount Due
            worksheet.Column(8).Width = 24; // Aggregate Amount Paid
            worksheet.Column(9).Width = 38; // Notes
            worksheet.Column(10).Width = 32; // Concat


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

        private class ExperisEndClientMapping
        {
            public string Identifier
            {
                get;
                set;
            } = "";

            public string EndClient
            {
                get;
                set;
            } = "";
        }

        private static List<ExperisEndClientMapping>
    LoadEndClientMappings()
        {
            var mappings =
                new List<ExperisEndClientMapping>();


            // Get configured mapping file.
            string mappingPath =
                Settings.GetExperisEndClientMappingFilePath();


            // Mapping is optional.
            // If no valid file is configured, simply return
            // an empty list and Experis processing continues.
            if (string.IsNullOrWhiteSpace(mappingPath) ||
                !File.Exists(mappingPath))
            {
                return mappings;
            }


            using var workbook =
                new XLWorkbook(mappingPath);


            // Your mapping workbook uses this worksheet.
            IXLWorksheet worksheet;

            if (workbook.Worksheets.Any(
                x => x.Name.Equals(
                    "EXPERIS END CLIENTS",
                    StringComparison.OrdinalIgnoreCase)))
            {
                worksheet =
                    workbook.Worksheet(
                        "EXPERIS END CLIENTS");
            }
            else
            {
                // Fallback in case the sheet gets renamed.
                worksheet =
                    workbook.Worksheet(1);
            }


            // Based on the mapping file:
            //
            // A = INVOICE ID
            // B = IDENTIFIER
            // C = END CLIENT
            //
            // Row 1 = headers.

            int lastRow =
                worksheet.LastRowUsed()?.RowNumber()
                ?? 1;


            for (int row = 2;
                 row <= lastRow;
                 row++)
            {
                string identifier =
                    worksheet.Cell(row, 2)
                        .GetString()
                        .Trim();

                string endClient =
                    worksheet.Cell(row, 3)
                        .GetString()
                        .Trim();


                if (string.IsNullOrWhiteSpace(identifier) ||
                    string.IsNullOrWhiteSpace(endClient))
                {
                    continue;
                }


                string[] identifiers =
                    identifier.Split(
                        ',',
                        StringSplitOptions.RemoveEmptyEntries |
                        StringSplitOptions.TrimEntries);

                foreach (string individualIdentifier in identifiers)
                {
                    mappings.Add(
                        new ExperisEndClientMapping
                        {
                            Identifier = individualIdentifier,
                            EndClient = endClient
                        });
                }
            }


            return mappings;
        }

        private static string FindEndClient(
                    string invoiceNumber,
                    string paymentReference,
                    List<ExperisEndClientMapping> mappings)
        {
            // ---------------------------------------------------------
            // 1. Try matching against the Experis Invoice Number first.
            // ---------------------------------------------------------

            if (!string.IsNullOrWhiteSpace(invoiceNumber))
            {
                ExperisEndClientMapping? invoiceMatch =
                    mappings
                        .Where(x =>
                            !string.IsNullOrWhiteSpace(
                                x.Identifier))
                        .OrderByDescending(
                            x => x.Identifier.Length)
                        .FirstOrDefault(
                            x => invoiceNumber.Contains(
                                x.Identifier,
                                StringComparison.OrdinalIgnoreCase));

                if (invoiceMatch != null)
                {
                    return invoiceMatch.EndClient;
                }
            }


            // ---------------------------------------------------------
            // 2. If Invoice Number did not match, try Payment Reference.
            // ---------------------------------------------------------

            if (!string.IsNullOrWhiteSpace(paymentReference))
            {
                ExperisEndClientMapping? referenceMatch =
                    mappings
                        .Where(x =>
                            !string.IsNullOrWhiteSpace(
                                x.Identifier))
                        .OrderByDescending(
                            x => x.Identifier.Length)
                        .FirstOrDefault(
                            x => paymentReference.Contains(
                                x.Identifier,
                                StringComparison.OrdinalIgnoreCase));

                if (referenceMatch != null)
                {
                    return referenceMatch.EndClient;
                }
            }


            // No match in either field.
            return "";
        }

        // =============================================================
        // INTERNAL MODEL
        // =============================================================

        private class ExperisPaymentRow
        {
            public string InvoiceNumber
            {
                get;
                set;
            } = "";

            public string PaymentReference
            {
                get;
                set;
            } = "";

            public decimal PaidAmount
            {
                get;
                set;
            }
        }
    }
}