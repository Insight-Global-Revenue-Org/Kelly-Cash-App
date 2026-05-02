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

bool loading = true;

Task spinner = Task.Run(() =>
{
    char[] frames = { '/', '-', '\\', '|' };
    int i = 0;

    while (loading)
    {
        Console.Write($"\rProcessing payment file... {frames[i++ % frames.Length]}");
        Thread.Sleep(120);
    }
});

try
{
    using var workbook = new XLWorkbook(inputPath);
    var worksheet = workbook.Worksheet("Payment Details");

    int headerRow = 13;

    int weekEndingColumn = FindColumn(worksheet, headerRow, "Week Ending Date");
    int contractorColumn = FindColumn(worksheet, headerRow, "Name");
    int lineTotalColumn = FindColumn(worksheet, headerRow, "Line Total");
    int locationDescriptionColumn = FindColumn(worksheet, headerRow, "Location Description");

    if (weekEndingColumn == -1 || contractorColumn == -1 || lineTotalColumn == -1 || locationDescriptionColumn == -1)
    {
        Console.WriteLine("Could not find Week Ending Date, Name, or Line Total.");
        return;
    }

    NormalizeColumns(worksheet, headerRow, contractorColumn);

    // Re-find columns after normalization
    int invoiceColumn = FindColumnAfter(worksheet, headerRow, "Invoice", contractorColumn);
    int amountColumn = FindColumnAfter(worksheet, headerRow, "Amount", contractorColumn);
    int aggregateColumn = FindColumn(worksheet, headerRow, "Aggregate Line Total");
    int notesColumn = FindColumn(worksheet, headerRow, "Notes");

    weekEndingColumn = FindColumn(worksheet, headerRow, "Week Ending Date");
    contractorColumn = FindColumn(worksheet, headerRow, "Name");
    lineTotalColumn = FindColumn(worksheet, headerRow, "Line Total");

    if (amountColumn == -1 || aggregateColumn == -1 || weekEndingColumn == -1 || contractorColumn == -1 || lineTotalColumn == -1)
    {
        Console.WriteLine("Could not find one or more required columns.");
        Console.WriteLine($"Amount: {amountColumn}");
        Console.WriteLine($"Week Ending: {weekEndingColumn}");
        Console.WriteLine($"Contractor Name: {contractorColumn}");
        Console.WriteLine($"Line Total: {lineTotalColumn}");
        return;
    }

    int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

    loading = false;
    spinner.Wait();

    Console.Write("\r" + new string(' ', 60)); 
    Console.WriteLine(); 

    Console.WriteLine($"Found Amount column: {amountColumn}");
    Console.WriteLine($"Found Week Ending column: {weekEndingColumn}");
    Console.WriteLine($"Found Contractor Name column: {contractorColumn}");
    Console.WriteLine($"Found Line Total column: {lineTotalColumn}");
    Console.WriteLine($"Found Aggregate column: {aggregateColumn}");

    // First pass: calculate totals up to last row for each Contractor + Week Ending
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

    // Second pass: write aggregate total to the last row for each contractor
    foreach (var item in totalsByContractorAndWeek)
    {
        string key = item.Key;
        decimal aggregateTotal = item.Value;
        int targetRow = lastRowByContractorAndWeek[key];

        var cell = worksheet.Cell(targetRow, aggregateColumn);

        cell.Value = aggregateTotal;

        cell.Style = worksheet.Cell(targetRow, lineTotalColumn).Style;

        cell.Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";

        // Are credits applied? If so, make negative totals red so MSP can identify
        if (aggregateTotal < 0)
        {
            cell.Style.Font.FontColor = XLColor.Red;
        }
    }

    string downloadsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        "Downloads"
    );

    string companyName = worksheet.Cell(headerRow + 1, locationDescriptionColumn).GetString().Trim();

    if (string.IsNullOrWhiteSpace(companyName))
    {
        companyName = "Remittance Payment";
    }

    companyName = MakeSafeFileName(companyName);

    decimal totalAmount = GetDecimalValue(worksheet.Cell(lastRow, lineTotalColumn));
    string formattedTotal = totalAmount.ToString("N2", CultureInfo.InvariantCulture);

    string outputPath = GetUniqueOutputPath(downloadsPath, $"{companyName} - {formattedTotal}.xlsx");

    workbook.SaveAs(outputPath);



    Console.WriteLine();
    Console.WriteLine("Week-Ending Line Totals Calculated Per Invoice Line Item");
    Console.WriteLine("Done!");
    Console.WriteLine($"Updated file saved to: {outputPath}");
}
catch (Exception ex)
{
    loading = false;
    spinner.Wait();
    Console.WriteLine();

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

static void NormalizeColumns(IXLWorksheet worksheet, int headerRow, int contractorColumn)
{
    string[] desiredColumns =
    {
        "Invoice",
        "Amount",
        "Aggregate Line Total",
        "Notes"
    };

    int insertAt = contractorColumn + 1;

    foreach (string columnName in desiredColumns)
    {
        string currentHeader = worksheet.Cell(headerRow, insertAt).Value.ToString().Trim();

        if (!currentHeader.Equals(columnName, StringComparison.OrdinalIgnoreCase))
        {
            InsertBlankColumn(worksheet, insertAt, columnName, headerRow);
        }

        insertAt++;
    }
}

static void InsertBlankColumn(IXLWorksheet worksheet, int targetColumn, string headerName, int headerRow)
{
    int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;
    int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? targetColumn;

    for (int row = 1; row <= lastRow; row++)
    {
        for (int col = lastColumn; col >= targetColumn; col--)
        {
            worksheet.Cell(row, col + 1).CopyFrom(worksheet.Cell(row, col));
            worksheet.Cell(row, col).Clear();
        }
    }

    worksheet.Cell(headerRow, targetColumn).Value = headerName;
    worksheet.Cell(headerRow, targetColumn).Style =
        worksheet.Cell(headerRow, targetColumn - 1).Style;

    worksheet.Cell(headerRow - 1, targetColumn).Style =
        worksheet.Cell(headerRow - 1, targetColumn - 1).Style;
}

static int FindColumnAfter(IXLWorksheet worksheet, int headerRow, string headerName, int afterColumn)
{
    for (int col = afterColumn + 1; col <= 100; col++)
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

static string MakeSafeFileName(string fileName)
{
    foreach (char invalidChar in Path.GetInvalidFileNameChars())
    {
        fileName = fileName.Replace(invalidChar, '_');
    }

    return fileName.Trim();
}

static string GetUniqueOutputPath(string folderPath, string fileName)
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