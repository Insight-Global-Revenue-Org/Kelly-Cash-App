public static class ConsoleUi
{
    public static void ResetPage(int startLine = 0)
    {
        // Clear the Windows Terminal
        Console.Write("\x1b[3J");

        Console.ResetColor();
        Console.CursorVisible = false;

        int linesToClear = Console.WindowHeight - startLine;

        for (int i = 0; i < linesToClear; i++)
        {
            Console.SetCursorPosition(0, startLine + i);
            Console.Write(new string(' ', Console.WindowWidth - 1));
        }

        Console.SetCursorPosition(0, startLine);
        Console.CursorVisible = true;
    }

    // <summary> test
    public static int DrawHeader()
    {
        Console.ForegroundColor = ConsoleColor.White;

        // Moved this UI (and its helper function) to it's own public class :) If you are reading this, Hello!
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

        return Console.CursorTop;
    }
}