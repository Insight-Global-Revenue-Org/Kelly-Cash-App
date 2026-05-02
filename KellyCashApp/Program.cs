using ClosedXML.Excel;

Console.WriteLine("Paste the full file path of the Payment Details Excel file:");
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
    int amountColumn = -1;

    for (int col = 1; col <= 100; col++)
    {
        string headerText = worksheet.Cell(headerRow, col).Value.ToString().Trim();

        if (headerText.Equals("Amount", StringComparison.OrdinalIgnoreCase))
        {
            amountColumn = col;
            break;
        }
    }

    if (amountColumn == -1)
    {
        Console.WriteLine("Could not find the 'Amount' column on row 13.");
        return;
    }

    Console.WriteLine($"Found Amount column at column number: {amountColumn}");

    int aggregateColumn = amountColumn + 1;
    int lastRow = 23; // Your sample has totals on row 23
    int lastColumn = 21; // Column U in your file

    // Manually shift columns K through U one column to the right
    for (int row = 1; row <= lastRow; row++)
    {
        for (int col = lastColumn; col >= aggregateColumn; col--)
        {
            worksheet.Cell(row, col + 1).CopyFrom(worksheet.Cell(row, col));
            worksheet.Cell(row, col).Clear();
        }
    }

    worksheet.Cell(headerRow, aggregateColumn).Value = "Aggregate Line Total";
    worksheet.Cell(headerRow, aggregateColumn).Style =
        worksheet.Cell(headerRow, amountColumn).Style;

    string downloadsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        "Downloads"
    );

    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
    string outputPath = Path.Combine(downloadsPath, $"{fileNameWithoutExt}_Updated.xlsx");

    workbook.SaveAs(outputPath);

    Console.WriteLine();
    Console.WriteLine("Done!");
    Console.WriteLine($"Updated file saved to: {outputPath}");
}
catch (Exception ex)
{
    Console.WriteLine("Something went wrong:");
    Console.WriteLine(ex.Message);
}