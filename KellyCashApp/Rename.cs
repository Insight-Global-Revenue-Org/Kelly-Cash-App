namespace KellyCashApp
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

            foreach (string line in File.ReadAllLines(nameChangeFilePath))
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                string[] parts = line.Split(',', 2);

                if (parts.Length < 2)
                    continue;

                string oldName = parts[0].Trim();
                string newName = parts[1].Trim();

                if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName))
                    continue;

                nameChanges[oldName] = newName;
            }

            return nameChanges;
        }

        public static string ApplyNameChange(string formattedName, Dictionary<string, string> nameChanges)
        {
            if (nameChanges.TryGetValue(formattedName, out string? newName))
            {
                return newName;
            }

            return formattedName;
        }
    }
}