using ClosedXML.Excel;

namespace KellyCashApp
{
    internal class MonumentPayment
    {
        public static bool IsMonumentFormat(IXLWorksheet worksheet)
        {
            return worksheet.Cell(1, 1).GetString().Trim()
                .Equals("Invoice Number", StringComparison.OrdinalIgnoreCase)
                && worksheet.Cell(1, 4).GetString().Trim()
                .Equals("Amount Paid", StringComparison.OrdinalIgnoreCase);
        }
        public static string Process(XLWorkbook workbook, IXLWorksheet worksheet, string inputPath)
        {
            int headerRow = 1;

            int invoiceColumn = FindColumn(worksheet, headerRow, "Invoice Number");
            int amountPaidColumn = FindColumn(worksheet, headerRow, "Amount Paid");
            int detailAmountColumn = FindColumn(worksheet, headerRow, "Detail Amount");

            if (invoiceColumn == -1 || amountPaidColumn == -1)
                throw new Exception("Missing required Monument columns.");

            // Monument Column Normalization
            NormalizeColumns(worksheet, headerRow);

            int aggregateColumn = FindColumn(worksheet, headerRow, "Aggregate Amount");

            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

            var totalsByInvoice = new Dictionary<string, decimal>();
            var lastRowByInvoice = new Dictionary<string, int>();

            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string invoice = worksheet.Cell(row, invoiceColumn).GetString().Trim();

                if (string.IsNullOrWhiteSpace(invoice))
                    continue;

                decimal amount = GetDecimalValue(worksheet.Cell(row, amountPaidColumn));

                if (!totalsByInvoice.ContainsKey(invoice))
                    totalsByInvoice[invoice] = 0;

                totalsByInvoice[invoice] += amount;
                lastRowByInvoice[invoice] = row;
            }

            foreach (var item in totalsByInvoice)
            {
                int targetRow = lastRowByInvoice[item.Key];
                var cell = worksheet.Cell(targetRow, aggregateColumn);

                cell.Value = item.Value;
                cell.Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

                if (item.Value < 0)
                    cell.Style.Font.FontColor = XLColor.Red;
            }

            worksheet.Columns().AdjustToContents();

            string downloadsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads"
            );

            string outputPath = GetUniqueOutputPath(downloadsPath, "Monument Payment Processed.xlsx");

            workbook.SaveAs(outputPath);

            return outputPath;
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

        private static void NormalizeColumns(IXLWorksheet worksheet, int headerRow)
        {
            string[] columnsToAdd =
            {
        "Aggregate Amount",
        "Notes"
    };

            int insertAt = (worksheet.LastColumnUsed()?.ColumnNumber() ?? 1) + 1;

            foreach (string column in columnsToAdd)
            {
                if (FindColumn(worksheet, headerRow, column) == -1)
                {
                    worksheet.Cell(headerRow, insertAt).Value = column;
                    insertAt++;
                }
            }
        }

        private static decimal GetDecimalValue(IXLCell cell)
        {
            string raw = cell.Value.ToString()
                .Replace("$", "")
                .Replace(",", "")
                .Trim();

            if (decimal.TryParse(raw, out decimal result))
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
    }
}