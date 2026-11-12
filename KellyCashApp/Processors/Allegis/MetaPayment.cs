using ClosedXML.Excel;
using KellyCashApp.Configuration;
using KellyCashApp.Models;
using KellyCashApp.Services;
using Microsoft.VisualBasic.FileIO;
using System.Globalization;
using System.Text.RegularExpressions;

namespace KellyCashApp.Processors.Allegis
{
    internal class MetaPayment
    {


        public static bool IsMetaFormat(string inputPath)
        {
            if (!Path.GetExtension(inputPath)
                .Equals(".csv", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            using var parser = new TextFieldParser(inputPath);
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            parser.HasFieldsEnclosedInQuotes = true;

            string[] headers = parser.ReadFields() ?? Array.Empty<string>();

            var headerNames = headers
                .Select(h => h.Trim().Trim('\uFEFF'))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            return headerNames.Contains("invoice_number")
                && headerNames.Contains("invoice_description")
                && headerNames.Contains("invoice_date")
                && headerNames.Contains("payment_amount")
                && headerNames.Contains("payment_reference_number")
                && headerNames.Contains("payment_type");
        }



        public static string Process(
            string inputPath,
            string registerPath,
            Dictionary<string, List<OirMatch>> openInvoiceMatches)

        {

            var registerMatches = LoadMetaRegister(registerPath);
            var oirRows = BuildOirRows(openInvoiceMatches);
            var nameChanges = Rename.LoadNameChanges();

            var outputRows = new List<MetaOutputRow>();


            using (var parser = new TextFieldParser(inputPath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = true;

                string[] csvHeaders = parser.ReadFields()
                    ?? throw new Exception("Meta CSV is empty.");

                var columns = csvHeaders
                    .Select((header, index) => new

                    {
                        Header = header.Trim().Trim('\uFEFF'),
                        Index = index
                    })
                    .ToDictionary(
                        x => x.Header,
                        x => x.Index,
                        StringComparer.OrdinalIgnoreCase);

                if (!columns.ContainsKey("invoice_number") ||
                    !columns.ContainsKey("payment_amount"))
                {
                    throw new Exception(
                        "Missing Meta invoice_number or payment_amount column.");
                }

                int invoiceCol = columns["invoice_number"];
                int amountCol = columns["payment_amount"];

                int csvRow = 1;

                while (!parser.EndOfData)
                {
                    csvRow++;

                    string[] fields = parser.ReadFields()
                        ?? Array.Empty<string>();

                    if (fields.Length <= Math.Max(invoiceCol, amountCol))
                        throw new Exception(
                            $"Meta CSV row {csvRow} has missing columns.");


                    string metaInvoice = NormalizeInvoiceNumber(
                        fields[invoiceCol]);


                    string itemType = metaInvoice.Length == 11
                        ? "Expense"
                        : "Labor";


                    string rawAmount = fields[amountCol].Trim();


                    if (string.IsNullOrWhiteSpace(metaInvoice))
                        throw new Exception(
                            $"Meta CSV row {csvRow} has no invoice number.");

                    if (!decimal.TryParse(
                        rawAmount,
                        NumberStyles.Currency,
                        CultureInfo.InvariantCulture,
                        out decimal paymentAmount))
                    {
                        throw new Exception(
                            $"Invalid payment amount on CSV row {csvRow}.");
                    }

                    string name = "";
                    string weekEnding = "";

                    if (registerMatches.TryGetValue(
                        metaInvoice, out var registerMatch))
                    {
                        name = registerMatch.Name;
                        name = Rename.ApplyNameChange(
                            name, nameChanges);

                        weekEnding = registerMatch.WeekEnding
                            .ToString("MM/dd/yyyy",
                                CultureInfo.InvariantCulture);
                    }

                    outputRows.Add(new MetaOutputRow
                    {
                        WeekEndingDate = weekEnding,
                        Name = name,
                        Invoice = "",
                        AmountDue = 0,
                        AggregateAmountPaid = paymentAmount,
                        ItemType = itemType,
                        Concat = string.IsNullOrWhiteSpace(name)
                            ? ""
                            : $"{name} {weekEnding}",
                        MetaInvoice = metaInvoice
                    });
                }
            }

            // ============================================================
            // AGGREGATE PAYMENTS BY CONTRACTOR AND WEEK ENDING
            // ============================================================


            outputRows = outputRows
                .GroupBy(x => new
                {
                    Key = string.IsNullOrWhiteSpace(x.Name)
                        ? $"UNMATCHED:{x.MetaInvoice.ToUpperInvariant()}|{x.ItemType}"
                        : $"{x.Name.ToUpperInvariant()}|{x.WeekEndingDate}|{x.ItemType}"
                })
                .Select(group =>
                {
                    var first = group.First();

                    decimal aggregatePaid =
                        group.Sum(x => x.AggregateAmountPaid);

                    // Look up each distinct Meta invoice only once,
                    // even if the CSV contains multiple payment lines.
                    var registerRows = group
                        .Select(x => x.MetaInvoice)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .Where(invoice => registerMatches.ContainsKey(invoice))
                        .Select(invoice => registerMatches[invoice])
                        .ToList();

                    decimal rtUnits = registerRows.Sum(x => x.RTUnits);
                    decimal otUnits = registerRows.Sum(x => x.OTUnits);
                    decimal dtUnits = registerRows.Sum(x => x.DTUnits);

                    decimal rtBillAmount =
                        registerRows.Sum(x => x.RTBillAmount);

                    decimal otBillAmount =
                        registerRows.Sum(x => x.OTBillAmount);

                    decimal dtBillAmount =
                        registerRows.Sum(x => x.DTBillAmount);

                    return new MetaOutputRow
                    {
                        WeekEndingDate = first.WeekEndingDate,
                        Name = first.Name,
                        Invoice = "",
                        AmountDue = 0,

                        AggregateAmountPaid = aggregatePaid,
                        ItemType = first.ItemType,

                        Notes = aggregatePaid < 0
                            ? "credit has been applied to payment"
                            : "",

                        Concat = first.Concat,

                        MetaInvoice = string.Join(", ",
                            group.Select(x => x.MetaInvoice)
                                .Distinct(StringComparer.OrdinalIgnoreCase)),

                        RTUnits = rtUnits,
                        RTBillRate = rtUnits != 0
                            ? rtBillAmount / rtUnits
                            : 0,

                        OTUnits = otUnits,
                        OTBillRate = otUnits != 0
                            ? otBillAmount / otUnits
                            : 0,

                        DTUnits = dtUnits,
                        DTBillRate = dtUnits != 0
                            ? dtBillAmount / dtUnits
                            : 0,

                        NetMSPFee =
                            registerRows.Sum(x => x.NetMSPFee),

                        NetSalesTax =
                            registerRows.Sum(x => x.NetSalesTax),

                        NetTotal =
                            registerRows.Sum(x => x.NetTotal)
                    };
                })
                .ToList();




            var matchedInvoiceNumbers =
    new HashSet<string>(StringComparer.OrdinalIgnoreCase);


            foreach (var outputRow in outputRows)
            {
                if (!DateTime.TryParseExact(
                    outputRow.WeekEndingDate,
                    "MM/dd/yyyy",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out DateTime metaDate))
                {
                    continue;
                }

                var match = oirRows
                    .Where(x =>
                        x.Name.Equals(
                            outputRow.Name,
                            StringComparison.OrdinalIgnoreCase)

                        && Math.Abs(
                            (x.WeekEndingDate.Date - metaDate.Date).Days) <= 1

                        && !string.IsNullOrWhiteSpace(x.Invoice)

                        && !matchedInvoiceNumbers.Contains(x.Invoice))

                    .OrderBy(x =>
                        Math.Abs(
                            (x.WeekEndingDate.Date - metaDate.Date).Days))

                    .ThenBy(x =>
                        Math.Abs(
                            x.AmountDue -
                            outputRow.AggregateAmountPaid))

                    .FirstOrDefault();

                if (match == null)
                    continue;

                outputRow.Invoice = match.Invoice;
                outputRow.AmountDue = match.AmountDue;

                outputRow.Concat =
                    $"{outputRow.Name} {match.WeekEndingDate:MM/dd/yyyy}";

                matchedInvoiceNumbers.Add(match.Invoice);
            }



            // ============================================================
            // GENERATE META RECONCILIATION WORKBOOK
            // ============================================================

            decimal total = outputRows.Sum(x => x.AggregateAmountPaid);

            using var workbook = new XLWorkbook();

            var worksheet = workbook.AddWorksheet("Reconciliation Notes");


            string[] headers =
            {
                "Week Ending Date",
                "Name",
                "Invoice",
                "Amount Due",
                "Aggregate Amount Paid",
                "Item Type",
                "Notes",
                "Concat",
                "Meta Invoice",
                "RT Units",
                "RT Bill Rate",
                "OT Units",
                "OT Bill Rate",
                "DT Units",
                "DT Bill Rate",
                "Net MSP Fee",
                "Net Sales Tax",
                "Net Total"
            };


            for (int col = 1; col <= headers.Length; col++)
            {
                worksheet.Cell(1, col).Value = headers[col - 1];
            }

            for (int i = 0; i < outputRows.Count; i++)
            {
                int row = i + 2;
                var item = outputRows[i];


                worksheet.Cell(row, 1).Value = item.WeekEndingDate;
                worksheet.Cell(row, 2).Value = item.Name;
                worksheet.Cell(row, 3).Value = item.Invoice;
                worksheet.Cell(row, 4).Value = item.AmountDue;
                worksheet.Cell(row, 5).Value = item.AggregateAmountPaid;
                worksheet.Cell(row, 6).Value = item.ItemType;
                worksheet.Cell(row, 7).Value = item.Notes;
                worksheet.Cell(row, 8).Value = item.Concat;
                worksheet.Cell(row, 9).Value = item.MetaInvoice;

                worksheet.Cell(row, 10).Value = item.RTUnits;
                worksheet.Cell(row, 11).Value = item.RTBillRate;

                worksheet.Cell(row, 12).Value = item.OTUnits;
                worksheet.Cell(row, 13).Value = item.OTBillRate;

                worksheet.Cell(row, 14).Value = item.DTUnits;
                worksheet.Cell(row, 15).Value = item.DTBillRate;

                worksheet.Cell(row, 16).Value = item.NetMSPFee;
                worksheet.Cell(row, 17).Value = item.NetSalesTax;
                worksheet.Cell(row, 18).Value = item.NetTotal;

            }

            ApplyFormatting(
                worksheet,
                outputRows.Count + 1,
                headers.Length);

            string downloadsPath = Settings.GetRemittanceSavePath();

            string formattedTotal = total.ToString(
                "$#,##0.00;($#,##0.00)",
                CultureInfo.InvariantCulture);

            string processedDate = DateTime.Now.ToString(
                "M.d.yyyy",
                CultureInfo.InvariantCulture);

            string outputPath = GetUniqueOutputPath(
                downloadsPath,
                $"Meta {processedDate} - {formattedTotal}.xlsx");

            workbook.SaveAs(outputPath);


            Analytics.LogRemittanceRun(
                $"Meta - {formattedTotal}");

            return outputPath;
        }

        // ============================================================
        // IMPORT META PAYMENT REGISTER
        // ============================================================


        private static Dictionary<string, MetaRegisterMatch>
            LoadMetaRegister(string registerPath)


        {
            var results = new Dictionary<string, MetaRegisterMatch>(
                StringComparer.OrdinalIgnoreCase);

            using var workbook = new XLWorkbook(registerPath);

            foreach (var sheet in workbook.Worksheets)
            {

                int headerRow = -1;
                int invoiceCol = -1;
                int contractorCol = -1;
                int weekEndingCol = -1;

                int payCodeCol = -1;
                int unitsCol = -1;
                int billRateCol = -1;
                int programFeeCol = -1;
                int taxAmountCol = -1;
                int netPayableCol = -1;


                for (int row = 1; row <= 25; row++)
                {

                    int inv = FindColumn(sheet, row, "Invoice Number");
                    int contractor = FindColumn(sheet, row, "Contractor");
                    int weekEnding = FindColumn(sheet, row, "Week Ending Date");

                    int payCode = FindColumn(sheet, row, "Pay Code");
                    int units = FindColumn(sheet, row, "Units");
                    int billRate = FindColumn(sheet, row, "Bill Rate");
                    int programFee = FindColumn(sheet, row, "Program Fee");
                    int taxAmount = FindColumn(sheet, row, "Tax Amount");
                    int netPayable = FindColumn(sheet, row, "Net Supplier Payable");

                    if (inv != -1 &&
                        contractor != -1 &&
                        weekEnding != -1 &&
                        payCode != -1 &&
                        units != -1 &&
                        billRate != -1 &&
                        programFee != -1 &&
                        taxAmount != -1 &&
                        netPayable != -1)
                    {
                        headerRow = row;

                        invoiceCol = inv;
                        contractorCol = contractor;
                        weekEndingCol = weekEnding;

                        payCodeCol = payCode;
                        unitsCol = units;
                        billRateCol = billRate;
                        programFeeCol = programFee;
                        taxAmountCol = taxAmount;
                        netPayableCol = netPayable;

                        break;
                    }


                }

                if (headerRow == -1)
                    continue;

                int lastRow = sheet.LastRowUsed()?.RowNumber() ?? headerRow;

                for (int row = headerRow + 1; row <= lastRow; row++)
                {

                    string invoiceNumber = NormalizeInvoiceNumber(
                        sheet.Cell(row, invoiceCol).GetString());


                    string contractor =
                        sheet.Cell(row, contractorCol).GetString().Trim();

                    if (string.IsNullOrWhiteSpace(invoiceNumber))
                        continue;

                    if (string.IsNullOrWhiteSpace(contractor))
                        throw new Exception(
                            $"Meta register row {row}: Contractor is blank.");

                    var weekEndingCell = sheet.Cell(row, weekEndingCol);

                    DateTime weekEnding;

                    if (weekEndingCell.Value.IsDateTime)
                    {
                        // Excel stores this as an actual date.
                        weekEnding = weekEndingCell.GetDateTime();
                    }
                    else if (!DateTime.TryParseExact(
                        weekEndingCell.GetString().Trim(),
                        new[] { "M/d/yyyy", "MM/dd/yyyy", "M/d/yy" },
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out weekEnding))
                    {
                        throw new Exception(
                            $"Meta register row {row}: Invalid Week Ending Date.");
                    }


                    string name = FormatName(contractor);


                    string payCode = sheet.Cell(row, payCodeCol)
                        .GetString().Trim().ToUpperInvariant();

                    decimal units = GetDecimalValue(
                        sheet.Cell(row, unitsCol));

                    decimal billRate = GetDecimalValue(
                        sheet.Cell(row, billRateCol));

                    decimal programFee = GetDecimalValue(
                        sheet.Cell(row, programFeeCol));

                    decimal taxAmount = GetDecimalValue(
                        sheet.Cell(row, taxAmountCol));

                    decimal netPayable = GetDecimalValue(
                        sheet.Cell(row, netPayableCol));

                    // Create a register record the first time
                    // we encounter this Meta invoice number.
                    if (!results.TryGetValue(invoiceNumber, out var entry))
                    {
                        entry = new MetaRegisterMatch
                        {
                            Name = name,
                            WeekEnding = weekEnding.Date
                        };

                        results.Add(invoiceNumber, entry);
                    }
                    else
                    {
                        // All rows for the same invoice should
                        // belong to the same contractor and week.
                        if (!entry.Name.Equals(
                                name, StringComparison.OrdinalIgnoreCase) ||
                            entry.WeekEnding != weekEnding.Date)
                        {
                            throw new Exception(
                                $"Conflicting Meta register records for invoice {invoiceNumber}.");
                        }
                    }

                    // Accumulate units and weighted bill-rate amounts
                    // only for the pay codes we care about.
                    switch (payCode)
                    {
                        case "RT":
                            entry.RTUnits += units;
                            entry.RTBillAmount += units * billRate;
                            break;

                        case "OT":
                            entry.OTUnits += units;
                            entry.OTBillAmount += units * billRate;
                            break;

                        case "DT":
                            entry.DTUnits += units;
                            entry.DTBillAmount += units * billRate;
                            break;
                    }

                    // Always sum these amounts across ALL register
                    // rows, regardless of their Pay Code.
                    entry.NetMSPFee += programFee;
                    entry.NetSalesTax += taxAmount;
                    entry.NetTotal += netPayable;

                }
            }

            if (results.Count == 0)
                throw new Exception(
                    "No Meta Payment Register records were found.");

            return results;
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


            private static void ApplyFormatting(
                IXLWorksheet worksheet, int lastRow, int lastColumn)
            {
                var range = worksheet.Range(1, 1, lastRow, lastColumn);

                range.Style.Font.FontName = "Aptos Narrow";
                range.Style.Font.FontSize = 9;

                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                range.Style.Alignment.Vertical =
                    XLAlignmentVerticalValues.Center;

                // Header styling - same peach color as SAIC.
                worksheet.Row(1).Style.Font.Bold = true;
                worksheet.Row(1).Style.Fill.BackgroundColor =
                    XLColor.FromHtml("#FCE4D6");

                worksheet.Row(1).Height = 15;


            // Currency formatting.
            int[] currencyColumns = { 4, 5, 11, 13, 15, 16, 17, 18 };

            foreach (int col in currencyColumns)
            {
                worksheet.Column(col).Style.NumberFormat.Format =
                    "$#,##0.00;($#,##0.00)";
            }

            // Units formatting.
            int[] unitColumns = { 10, 12, 14 };

            foreach (int col in unitColumns)
            {
                worksheet.Column(col).Style.NumberFormat.Format = "0.00";
            }



            worksheet.Column(1).Width = 22;  // Week Ending Date
            worksheet.Column(2).Width = 24;  // Name
            worksheet.Column(3).Width = 20;  // Invoice
            worksheet.Column(4).Width = 18;  // Amount Due
            worksheet.Column(5).Width = 26;  // Aggregate Amount Paid
            worksheet.Column(6).Width = 16;  // Item Type
            worksheet.Column(7).Width = 42;  // Notes
            worksheet.Column(8).Width = 34;  // Concat
            worksheet.Column(9).Width = 35;  // Meta Invoice

            worksheet.Column(10).Width = 14; // RT Units
            worksheet.Column(11).Width = 18; // RT Bill Rate
            worksheet.Column(12).Width = 14; // OT Units
            worksheet.Column(13).Width = 18; // OT Bill Rate
            worksheet.Column(14).Width = 14; // DT Units
            worksheet.Column(15).Width = 18; // DT Bill Rate
            worksheet.Column(16).Width = 18; // Net MSP Fee
            worksheet.Column(17).Width = 18; // Net Sales Tax
            worksheet.Column(18).Width = 20; // Net Total





            for (int row = 2; row <= lastRow; row++)
            {
                worksheet.Row(row).Height = 15;

                // Read Item Type from column F.
                string itemType = worksheet.Cell(row, 6)
                    .GetString()
                    .Trim();

                if (itemType.Equals(
                    "Expense",
                    StringComparison.OrdinalIgnoreCase))
                {
                    // Light grey for Expense rows.
                    worksheet.Range(row, 1, row, lastColumn)
                        .Style.Fill.BackgroundColor =
                            XLColor.FromHtml("#F2F2F2");
                }
                else
                {
                    // No shading for Labor rows.
                    worksheet.Range(row, 1, row, lastColumn)
                        .Style.Fill.BackgroundColor = XLColor.NoColor;
                }

                // Aggregate Amount Paid is column E.
                decimal aggregatePaid =
                    worksheet.Cell(row, 5).GetValue<decimal>();

                if (aggregatePaid < 0)
                {
                    worksheet.Cell(row, 5)
                        .Style.Font.FontColor = XLColor.Red;
                }
            }



            range.SetAutoFilter();
                worksheet.SheetView.FreezeRows(1);
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


        private static string NormalizeInvoiceNumber(string raw)
        {
            raw = raw.Trim();

            if (string.IsNullOrWhiteSpace(raw))
                return "";

            if (Regex.IsMatch(
                raw,
                @"^[+-]?\d+(?:\.\d+)?[Ee][+-]?\d+$"))
            {
                if (decimal.TryParse(
                    raw,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out decimal scientificNumber))
                {
                    if (scientificNumber != decimal.Truncate(scientificNumber))
                        throw new Exception(
                            $"Invoice number is not a whole number: {raw}");

                    return scientificNumber.ToString(
                        "0",
                        CultureInfo.InvariantCulture);
                }

                throw new Exception(
                    $"Unable to parse scientific invoice number: {raw}");
            }

            // Remove a decimal suffix only if it is entirely zero.
            if (decimal.TryParse(
                raw,
                NumberStyles.Number,
                CultureInfo.InvariantCulture,
                out decimal number))
            {
                if (number == decimal.Truncate(number))
                {
                    return number.ToString(
                        "0",
                        CultureInfo.InvariantCulture);
                }
            }

            // Preserve nonnumeric invoice identifiers.
            return raw;
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


        private class MetaOutputRow
        {
            public string WeekEndingDate { get; set; } = "";
            public string Name { get; set; } = "";
            public string Invoice { get; set; } = "";
            public decimal AmountDue { get; set; }
            public decimal AggregateAmountPaid { get; set; }
            public string ItemType { get; set; } = "Labor";
            public string Notes { get; set; } = "";
            public string Concat { get; set; } = "";
            public string MetaInvoice { get; set; } = "";

            public decimal RTUnits { get; set; }
            public decimal RTBillRate { get; set; }

            public decimal OTUnits { get; set; }
            public decimal OTBillRate { get; set; }

            public decimal DTUnits { get; set; }
            public decimal DTBillRate { get; set; }

            public decimal NetMSPFee { get; set; }
            public decimal NetSalesTax { get; set; }
            public decimal NetTotal { get; set; }
        }


        private class MetaRegisterMatch
        {
            public string Name { get; set; } = "";
            public DateTime WeekEnding { get; set; }

            public decimal RTUnits { get; set; }
            public decimal RTBillAmount { get; set; }

            public decimal OTUnits { get; set; }
            public decimal OTBillAmount { get; set; }

            public decimal DTUnits { get; set; }
            public decimal DTBillAmount { get; set; }

            public decimal NetMSPFee { get; set; }
            public decimal NetSalesTax { get; set; }
            public decimal NetTotal { get; set; }
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