using SearchText.Services;

namespace SearchText;
public class Program
{
    static void Main(string[] args)
    {
        ConsoleKeyInfo cki = default;
        string Result = string.Empty;
        string? rootPath = string.Empty;
        string? searchTerm = string.Empty;
        int searchedFilesCount = 0;

        do
        {
            Print("Enter the path where you want to search for the texts in the files inside it");
            rootPath = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(rootPath))
            {
                Print("please enter a valid string", ConsoleColor.Red);
                continue;
            }

            Print("Please enter the term you want to search for");
            searchTerm = Console.ReadLine();
            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                Print("please enter a valid string", ConsoleColor.Red);
                continue;
            }

            try
            {
                Print("");
                Print("Please wait until the results are announced and do not close the program");
                Print("");

                using Search searchService = new Search();
                Result = searchService.SearchIntoFiles(searchTerm, rootPath, out searchedFilesCount);

                if (string.IsNullOrEmpty(Result))
                    Print("No results found", ConsoleColor.Blue);
                else
                    Print($"{searchedFilesCount} files were searched = {Result}" );
            }
            catch (Exception exp)
            {
                Print($"an error ocurred on searching the searchTerm : message : {exp.Message} - stackTrace : {exp.StackTrace}", ConsoleColor.Red);
            }

            Print("Press the Escape (Esc) key to quit: \n", ConsoleColor.Yellow);
            cki = Console.ReadKey();

        } while (cki.Key != ConsoleKey.Escape);
    }
    internal static void Print(string text) => Console.WriteLine(text);
    internal static void Print(string text, ConsoleColor? color = null)
    {
        if (color is object)
            Console.ForegroundColor = (ConsoleColor)color;

        Console.WriteLine(text);
        Console.ResetColor();
    }
}