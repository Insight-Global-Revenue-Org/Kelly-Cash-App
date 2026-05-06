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

Console.WriteLine("A Week-Ending Line Total Aggregate Script");
Console.WriteLine("──────────────────────────────────────────────────");
var openInvoiceMatches = new Dictionary<string, OirMatch>();

string? inputPath = null;
int defaultMenuOption = 0;
int fixedMenuTop = Console.CursorTop;

// Main application workflow loop.
// Menu navigation, OIR imports, remittance processing (further workflows may be added)
while (true)
{
    int selected = ShowMenu(new[]
    {
        "Import Open Invoice Report",
        "Process Remittance Payment",
        "Exit"
    }, defaultMenuOption, fixedMenuTop);

    if (selected == 2)
        break;

    int promptTop = fixedMenuTop + 6;

    // Open Invoice Report (OIR) import workflow.
    // After importing, saves invoice/document amount mappings into system memory
    if (selected == 0)
    {
        string? oirPath = PromptForFilePath(
            "Paste the full file path of the Open Invoice Report:",
            promptTop
        );

        if (string.IsNullOrWhiteSpace(oirPath) || !File.Exists(oirPath))
        {
            ClearArea(promptTop, 4);
            Console.SetCursorPosition(0, promptTop);
            Console.WriteLine("File not found. Please check the path and try again.\n");
            continue;
        }

        bool oirLoading = true;

        Task oirSpinner = Task.Run(() =>
        {
            char[] frames = { '/', '-', '\\', '|' };
            int i = 0;

            while (oirLoading)
            {
                Console.SetCursorPosition(0, promptTop);
                Console.Write($"Importing Open Invoice Report... {frames[i++ % frames.Length]}   ");
                Thread.Sleep(120);
            }
        });

        try
        {
            openInvoiceMatches = LoadOpenInvoiceReport(oirPath);
        }
        catch (Exception ex)
        {
            oirLoading = false;
            oirSpinner.Wait();

            ClearArea(promptTop, 8);
            Console.SetCursorPosition(0, promptTop);

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Something went wrong while importing the OIR:");
            Console.ResetColor();

            Console.WriteLine(ex.Message);
            Console.WriteLine();
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 0;
            continue;
        }
        finally
        {
            oirLoading = false;
            oirSpinner.Wait();

            ClearArea(promptTop, 8);
        }

        ClearArea(promptTop, 4);
        Console.SetCursorPosition(0, promptTop);
        Console.WriteLine($"Imported {openInvoiceMatches.Count} OIR matches into memory.\n");
        Console.WriteLine("Press 'Enter' to process your payment");

        Thread.Sleep(1000); 
        defaultMenuOption = 1;
        continue;
    }

    // Remittance processing workflow.
    // Normalize columns, calculate aggregates,
    // auto-match invoices with documented amount, and generates the formatted output file.
    inputPath = PromptForFilePath(
        "Paste the full file path of the Remittance Payment:",
        promptTop
    );

    if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
    {
        ClearArea(promptTop, 4);
        Console.SetCursorPosition(0, promptTop);
        Console.WriteLine("File not found. Please check the path and try again.\n");
        continue;
    }

    bool loading = true;

    Task spinner = Task.Run(() =>
    {
        char[] frames = { '/', '-', '\\', '|' };
        int i = 0;

        while (loading)
        {
            Console.SetCursorPosition(0, promptTop);
            Console.Write($"Processing payment file... {frames[i++ % frames.Length]}   ");
            Thread.Sleep(120);
        }
    });

    try
{
    using var stream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        // Open remittance workbook using the installed ClosedXML dependancy
        // FileShare.ReadWrite allows the system to process while the file is open elsewhere.
        using var workbook = new XLWorkbook(stream);
    var worksheet = workbook.Worksheet("Payment Details");

    int headerRow = 13;
     
    // This is where we dynamically locate the header columns
    int weekEndingColumn = FindColumn(worksheet, headerRow, "Week Ending Date");
    int contractorColumn = FindColumn(worksheet, headerRow, "Name");
    int lineTotalColumn = FindColumn(worksheet, headerRow, "Line Total");
    int locationDescriptionColumn = FindColumn(worksheet, headerRow, "Location Description");

    if (weekEndingColumn == -1 || contractorColumn == -1 || lineTotalColumn == -1 || locationDescriptionColumn == -1)
    {
            loading = false;
            spinner.Wait();

            Console.WriteLine("Could not find Week Ending Date, Name, Line Total, or Location Description.");
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 1;
            continue;
        }
    // Normalize worksheet layout by ensuring required automation columns exist
    // will remain in a consistent position for the downstream processing.
    NormalizeColumns(worksheet, headerRow, contractorColumn);
    InsertBlankColumn(worksheet, contractorColumn + 1, "Name_", headerRow);

    // Re-find columns after normalization
    int invoiceColumn = FindColumnAfter(worksheet, headerRow, "Invoice", contractorColumn);
    int amountColumn = FindColumnAfter(worksheet, headerRow, "Amount", contractorColumn);
    int aggregateColumn = FindColumn(worksheet, headerRow, "Aggregate Line Total");
    int notesColumn = FindColumn(worksheet, headerRow, "Notes");

    weekEndingColumn = FindColumn(worksheet, headerRow, "Week Ending Date");
    contractorColumn = FindColumn(worksheet, headerRow, "Name");
    int nameFormattedColumn = FindColumnAfter(worksheet, headerRow, "Name_", contractorColumn);
    lineTotalColumn = FindColumn(worksheet, headerRow, "Line Total");

    // Add Concat column after Line Total
    InsertBlankColumn(worksheet, lineTotalColumn + 1, "Concat", headerRow);

    // Re-finding Line Total after inserting Concat
    lineTotalColumn = FindColumn(worksheet, headerRow, "Line Total");

    // And finding the new Concat column again for the downstream matching
    int concatColumn = FindColumnAfter(worksheet, headerRow, "Concat", lineTotalColumn);

    if (amountColumn == -1 || aggregateColumn == -1 || weekEndingColumn == -1 || contractorColumn == -1 || lineTotalColumn == -1)
    {
        Console.WriteLine("Could not find one or more required columns.");
        Console.WriteLine($"Amount: {amountColumn}");
        Console.WriteLine($"Week Ending: {weekEndingColumn}");
        Console.WriteLine($"Contractor Name: {contractorColumn}");
            loading = false;
            spinner.Wait();

            Console.WriteLine($"Line Total: {lineTotalColumn}");
            Console.WriteLine("Press any key to return to the menu...");
            Console.ReadKey(true);

            defaultMenuOption = 1;
            continue;
        }

    int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

    for (int row = headerRow + 1; row <= lastRow; row++)
    {
        string rawName = worksheet.Cell(row, contractorColumn).GetString().Trim();

        if (string.IsNullOrWhiteSpace(rawName))
            continue;

        string formattedName = FormatName(rawName);
        string formattedWeekEnding = FormatWeekEndingDate(worksheet.Cell(row, weekEndingColumn));

        // Generates composite lookup key used for our automatic OIR invoice matching.
        string concatValue = $"{formattedName} {formattedWeekEnding}".Trim();

        worksheet.Cell(row, nameFormattedColumn).Value = formattedName;
        worksheet.Cell(row, concatColumn).Value = concatValue;

     
    }

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

        string concatValue = worksheet.Cell(targetRow, concatColumn).GetString().Trim();

        // Auto-match our remittance entries against the users imported OIR invoice data.
        if (openInvoiceMatches.TryGetValue(concatValue, out OirMatch match))
        {
            worksheet.Cell(targetRow, invoiceColumn).Value = match.DocumentNumber;
            worksheet.Cell(targetRow, amountColumn).Value = match.RemainingAmount;

            worksheet.Cell(targetRow, amountColumn).Style =
                worksheet.Cell(targetRow, lineTotalColumn).Style;

            worksheet.Cell(targetRow, amountColumn).Style.NumberFormat.Format = "$#,##0.00;($#,##0.00)";
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

    // Generate output filename using company name and payment total
    string outputPath = GetUniqueOutputPath(downloadsPath, $"{companyName} - {formattedTotal}.xlsx");

    int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? lineTotalColumn;

    if (worksheet.AutoFilter.IsEnabled)
    {
        worksheet.AutoFilter.Clear();
    }

    worksheet.Range(headerRow, 1, lastRow, lastColumn).SetAutoFilter();

    worksheet.Column(invoiceColumn).Width = 16;
    worksheet.Column(amountColumn).Width = 14;
    worksheet.Column(aggregateColumn).Width = 20;
    worksheet.Column(notesColumn).Width = 24;
    worksheet.Column(nameFormattedColumn).Width = 18;
    worksheet.Column(concatColumn).Width = 28;

    workbook.SaveAs(outputPath);



    Console.WriteLine();
    Console.WriteLine("Week-Ending Line Totals Calculated Per Invoice Line Item");
    Console.WriteLine("All available open invoices have been automatically matched!");
    Console.WriteLine("\nPress 'Enter' to process another payment.");
    Console.WriteLine($"Updated file saved to: {outputPath}");

        Thread.Sleep(1000); // wait 1 second
        defaultMenuOption = 1;
        continue;
    }
    catch (Exception ex)
    {
        loading = false;
        spinner.Wait();
        Console.WriteLine();

        Console.WriteLine("Something went wrong:");
        Console.WriteLine(ex.Message);
    }

}

Console.WriteLine();
Console.WriteLine("Press any key to exit...");
Console.ReadKey(true);

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

    var designCell = worksheet.Cell(headerRow - 1, targetColumn);
    var leftCell = worksheet.Cell(headerRow - 1, targetColumn - 1);

    // Copy style first
    designCell.Style = leftCell.Style;

    // Remove the shared border from the LEFT column
    leftCell.Style.Border.RightBorder = XLBorderStyleValues.None;
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

static string FormatName(string input)
{
    // Expected: "DOE, JOHN"
    if (!input.Contains(","))
        return input;

    var parts = input.Split(',');

    if (parts.Length < 2)
        return input;

    string last = parts[0].Trim().ToLower();
    string first = parts[1].Trim().ToLower();

    // Capitalize first letters
    last = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(last);
    first = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(first);

    return $"{first} {last}";
}

static string FormatWeekEndingDate(IXLCell cell)
{
    if (cell.Value.IsDateTime)
    {
        return cell.GetDateTime().ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
    }

    string rawValue = cell.GetString().Trim();

    if (DateTime.TryParse(rawValue, out DateTime parsedDate))
    {
        return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
    }

    return rawValue;
}

// The main console gui menu rendering function. Displays a list of options, highlights the currently selected option,
// and allows the user to navigate with the up/down arrow keys and select with Enter.
static int ShowMenu(string[] options, int defaultSelected, int menuTop)
{
    int selected = defaultSelected;
    int menuWidth = options.Max(o => o.Length) + 4;

    int neededRows = options.Length + 2;

    if (menuTop + neededRows >= Console.BufferHeight)
    {
        menuTop = Math.Max(0, Console.BufferHeight - neededRows);
    }

    Console.CursorVisible = false;

    DrawFullMenu(options, selected, menuTop, menuWidth);

    while (true)
    {
        ConsoleKey key = Console.ReadKey(true).Key;

        if (key == ConsoleKey.Enter)
            break;

        int oldSelected = selected;

        if (key == ConsoleKey.UpArrow)
            selected = selected == 0 ? options.Length - 1 : selected - 1;

        if (key == ConsoleKey.DownArrow)
            selected = selected == options.Length - 1 ? 0 : selected + 1;

        if (selected != oldSelected)
        {
            DrawFullMenu(options, selected, menuTop, menuWidth);
        }
    }

    Console.ResetColor();
    Console.CursorVisible = true;

    int afterMenu = Math.Min(menuTop + neededRows, Console.BufferHeight - 1);
    Console.SetCursorPosition(0, afterMenu);
    Console.WriteLine();

    return selected;
}

static void DrawFullMenu(string[] options, int selected, int menuTop, int menuWidth)
{
    Console.ResetColor();
    Console.SetCursorPosition(0, menuTop);
    Console.WriteLine("Select".PadRight(menuWidth));

    for (int i = 0; i < options.Length; i++)
    {
        Console.SetCursorPosition(0, menuTop + i + 1);

        string line = i == selected
            ? $"> {options[i]}"
            : $"  {options[i]}";

        if (i == selected)
        {
            Console.ForegroundColor = ConsoleColor.Black;
            Console.BackgroundColor = ConsoleColor.Gray;
            Console.Write(line.PadRight(menuWidth));
            Console.ResetColor();
        }
        else
        {
            Console.Write(line.PadRight(menuWidth));
        }
    }
}

// Loads and parses the Open Invoice Report (OIR)
// into an in-memory lookup dictionary.
static Dictionary<string, OirMatch> LoadOpenInvoiceReport(string filePath)
{
    var matches = new Dictionary<string, OirMatch>();

    using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    using var workbook = new XLWorkbook(stream);
    var worksheet = workbook.Worksheets.First();

    int headerRow = FindHeaderRow(worksheet, "Consultant");

    if (headerRow == -1)
        throw new Exception("Could not find Consultant header in Open Invoice Report.");

    int consultantColumn = FindColumn(worksheet, headerRow, "Consultant");
    int serviceEndColumn = FindColumn(worksheet, headerRow, "Service End");
    int documentNumberColumn = FindColumn(worksheet, headerRow, "Document Number");
    int remainingAmountColumn = FindColumn(worksheet, headerRow, "Remaining Amount");

    if (consultantColumn == -1 || serviceEndColumn == -1 || documentNumberColumn == -1 || remainingAmountColumn == -1)
        throw new Exception("Could not find Consultant, Service End, Document Number, or Remaining Amount in OIR.");

    int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

    for (int row = headerRow + 1; row <= lastRow; row++)
    {
        string consultant = worksheet.Cell(row, consultantColumn).GetString().Trim();
        string serviceEnd = FormatWeekEndingDate(worksheet.Cell(row, serviceEndColumn));

        if (string.IsNullOrWhiteSpace(consultant) || string.IsNullOrWhiteSpace(serviceEnd))
            continue;

        string concatKey = $"{consultant} {serviceEnd}".Trim();
        string documentNumber = worksheet.Cell(row, documentNumberColumn).GetString().Trim();
        decimal remainingAmount = GetDecimalValue(worksheet.Cell(row, remainingAmountColumn));

        if (!matches.ContainsKey(concatKey))
        {
            matches.Add(concatKey, new OirMatch(documentNumber, remainingAmount));
        }
    }

    return matches;
}
static string? PromptForFilePath(string prompt, int startLine)
{
    Console.CursorVisible = true;

    ClearArea(startLine, 6);

    Console.SetCursorPosition(0, startLine);
    Console.WriteLine(prompt);
    Console.Write("> ");

    int inputLine = Console.CursorTop;

    string? path = Console.ReadLine()?.Trim().Trim('"');

    ClearArea(startLine, 6);

    return path;
}

// HELPER FUNCTION: Console region cleanup, prevent some UI duplication artifacts I encountered
static void ClearArea(int startLine, int numberOfLines)
{
    for (int i = 0; i < numberOfLines; i++)
    {
        int line = startLine + i;
        if (line < 0 || line >= Console.BufferHeight) continue;

        Console.SetCursorPosition(0, line);
        Console.Write(new string(' ', Console.WindowWidth - 1));
    }

    Console.SetCursorPosition(0, startLine);
}
// May remove later, not currently being used
static void ClearLine(int line)
{
    if (line < 0 || line >= Console.BufferHeight) return;

    Console.SetCursorPosition(0, line);
    Console.Write(new string(' ', Console.WindowWidth - 1));
    Console.SetCursorPosition(0, line);
}

//HELPER FUNCTION: Dynamically locate spreadsheet header rows
static int FindHeaderRow(IXLWorksheet worksheet, string requiredHeader)
{
    int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 100;

    for (int row = 1; row <= lastRow; row++)
    {
        for (int col = 1; col <= 100; col++)
        {
            string cellText = worksheet.Cell(row, col).GetString().Trim();

            if (cellText.Equals(requiredHeader, StringComparison.OrdinalIgnoreCase))
                return row;
        }
    }

    return -1;
}

// Represents a single OIR invoice mapping entry used during auto-matching (may expand upon if more columns are needed)
record OirMatch(string DocumentNumber, decimal RemainingAmount);