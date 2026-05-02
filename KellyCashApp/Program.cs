using ClosedXML.Excel;
using System.Globalization;

Console.ForegroundColor = ConsoleColor.White;



Console.ForegroundColor = ConsoleColor.White;

Console.WriteLine(@"
  _  __    _ _         ____                  _               
 | |/ /___| | |_   _  / ___|  ___ _ ____   _(_) ___ ___  ___ 
 | ' // _ \ | | | | | \___ \ / _ \ '__\ \ / / |/ __/ _ \/ __|
 | . \  __/ | | |_| |  ___) |  __/ |   \ V /| | (_|  __/\__ \
 |_|\_\___|_|_|\__, | |____/ \___|_|    \_/ |_|\___\___||___/
               |___/                                         
");

Console.ResetColor();

Thread.Sleep(100);

Console.WriteLine("A Week-Ending Line Total Aggregate Script\n");
Console.WriteLine("--------------------------------------------------\n");

Console.WriteLine("Paste the full file path of the Remittance Payment:");
string? inputPath = Console.ReadLine()?.Trim().Trim('"');

if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
{
    Console.WriteLine("File not found. Please check the path and try again.");
    return;
}

try
{
    using var workbook = new XLWorkbook(inputPath);
    var worksheet = workbook.Worksheet("Payment Details");

    int headerRow = 13;

    int amountColumn = FindColumn(worksheet, headerRow, "Amount");
    int weekEndingColumn = FindColumn(worksheet, headerRow, "Week Ending Date");
    int contractorColumn = FindColumn(worksheet, headerRow, "Name");
    int lineTotalColumn = FindColumn(worksheet, headerRow, "Line Total");

    if (amountColumn == -1 || weekEndingColumn == -1 || contractorColumn == -1 || lineTotalColumn == -1)
    {
        Console.WriteLine("Could not find one or more required columns.");
        Console.WriteLine($"Amount: {amountColumn}");
        Console.WriteLine($"Week Ending: {weekEndingColumn}");
        Console.WriteLine($"Contractor Name: {contractorColumn}");
        Console.WriteLine($"Line Total: {lineTotalColumn}");
        return;
    }

    Console.WriteLine($"Found Amount column: {amountColumn}");
    Console.WriteLine($"Found Week Ending column: {weekEndingColumn}");
    Console.WriteLine($"Found Contractor Name column: {contractorColumn}");
    Console.WriteLine($"Found Line Total column: {lineTotalColumn}");

    int aggregateColumn = amountColumn + 1;
    int lastRow = 23;
    int lastColumn = 21;

    // Shift columns after Amount one column to the right
    for (int row = 1; row <= lastRow; row++)
    {
        for (int col = lastColumn; col >= aggregateColumn; col--)
        {
            worksheet.Cell(row, col + 1).CopyFrom(worksheet.Cell(row, col));
            worksheet.Cell(row, col).Clear();
        }
    }

    // After shifting, Line Total moved one column to the right
    if (lineTotalColumn >= aggregateColumn)
        lineTotalColumn++;

    worksheet.Cell(headerRow, aggregateColumn).Value = "Aggregate Line Total";
    worksheet.Cell(headerRow, aggregateColumn).Style =
        worksheet.Cell(headerRow, amountColumn).Style;

    // Extend row 12 design
    worksheet.Cell(headerRow - 1, aggregateColumn).Style =
        worksheet.Cell(headerRow - 1, amountColumn).Style;

    // First pass: calculate totals and remember the LAST row for each Contractor + Week Ending
    var totalsByContractorAndWeek = new Dictionary<string, decimal>();
    var lastRowByContractorAndWeek = new Dictionary<string, int>();

    for (int row = headerRow + 1; row < lastRow; row++)
    {
        string contractor = worksheet.Cell(row, contractorColumn).GetString().Trim();
        string weekEnding = worksheet.Cell(row, weekEndingColumn).GetString().Trim();

        if (string.IsNullOrWhiteSpace(contractor) || string.IsNullOrWhiteSpace(weekEnding))
            continue;

        decimal lineTotal = GetDecimalValue(worksheet.Cell(row, lineTotalColumn));
        string key = $"{contractor}|{weekEnding}";

        if (!totalsByContractorAndWeek.ContainsKey(key))
            totalsByContractorAndWeek[key] = 0;

        totalsByContractorAndWeek[key] += lineTotal;
        lastRowByContractorAndWeek[key] = row;
    }

    // Second pass: write aggregate total ONLY to the last row for each group
    foreach (var item in totalsByContractorAndWeek)
    {
        string key = item.Key;
        decimal aggregateTotal = item.Value;
        int targetRow = lastRowByContractorAndWeek[key];

        var cell = worksheet.Cell(targetRow, aggregateColumn);

        cell.Value = aggregateTotal;

        // Copy base style first
        cell.Style = worksheet.Cell(targetRow, lineTotalColumn).Style;

        // Apply accounting format to ALL values
        cell.Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

        // If negative → make red
        if (aggregateTotal < 0)
        {
            cell.Style.Font.FontColor = XLColor.Red;
        }
    }

    string downloadsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        "Downloads"
    );

    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
    string outputPath = Path.Combine(downloadsPath, $"{fileNameWithoutExt}_Updated.xlsx");

    workbook.SaveAs(outputPath);

    Console.WriteLine();
    Console.WriteLine("Week-Ending Line Totals Calculated Per Invoice Line Item");
    Console.WriteLine("Done!");
    Console.WriteLine($"Updated file saved to: {outputPath}");
}
catch (Exception ex)
{
    Console.WriteLine("Something went wrong:");
    Console.WriteLine(ex.Message);
}

static int FindColumn(IXLWorksheet worksheet, int headerRow, string headerName)
{
    for (int col = 1; col <= 100; col++)
    {
        string headerText = worksheet.Cell(headerRow, col).Value.ToString().Trim();

        if (headerText.Equals(headerName, StringComparison.OrdinalIgnoreCase))
            return col;
    }

    return -1;
}

static decimal GetDecimalValue(IXLCell cell)
{
    string rawValue = cell.Value.ToString()
        .Replace("$", "")
        .Replace(",", "")
        .Trim();

    if (decimal.TryParse(rawValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
        return result;

    return 0;
}