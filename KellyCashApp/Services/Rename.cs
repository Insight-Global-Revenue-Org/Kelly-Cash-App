using ClosedXML.Excel;
using KellyCashApp.Configuration;

namespace KellyCashApp.Services
{
    internal class Rename
    {
        public static Dictionary<string, string> LoadNameChanges()
        {
            string nameChangeFilePath = Settings.GetNameChangeFilePath();

            var nameChanges = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(nameChangeFilePath))
                return nameChanges;

            if (!File.Exists(nameChangeFilePath))
                return nameChanges;

            string extension = Path.GetExtension(nameChangeFilePath).ToLower();

            if (extension == ".xlsx" || extension == ".xls")
            {
                LoadNameChangesFromExcel(nameChangeFilePath, nameChanges);
            }
            else
            {
                LoadNameChangesFromText(nameChangeFilePath, nameChanges);
            }

            return nameChanges;
        }

        private static void LoadNameChangesFromText(string filePath, Dictionary<string, string> nameChanges)
        {
            foreach (string line in File.ReadAllLines(filePath))
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                string[] parts = line.Split(',', 2);

                if (parts.Length < 2)
                    continue;

                AddNameChange(parts[0], parts[1], nameChanges);
            }
        }

        private static void LoadNameChangesFromExcel(string filePath, Dictionary<string, string> nameChanges)
        {
            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.First();

            foreach (var row in worksheet.RowsUsed())
            {
                string oldName = row.Cell(1).GetString().Trim();
                string newName = row.Cell(2).GetString().Trim();

                AddNameChange(oldName, newName, nameChanges);
            }
        }

        private static void AddNameChange(string oldName, string newName, Dictionary<string, string> nameChanges)
        {
            oldName = oldName.Trim();
            newName = newName.Trim();

            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName))
                return;

            nameChanges[oldName] = newName;
        }

        public static string ApplyNameChange(string formattedName, Dictionary<string, string> nameChanges)
        {
            if (nameChanges.TryGetValue(formattedName, out string? newName))
                return newName;

            return formattedName;
        }
    }
}