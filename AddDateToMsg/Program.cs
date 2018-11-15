using System;

namespace AddDateToMsg
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0 && MessageHandler.IsMsgFile(args[0]))
            {
                MessageHandler.ProcessMsgList(args);
            }
            else if (args.Length > 0 && System.IO.Directory.Exists(args[0]))
            {
                MessageHandler.ProcessCompleteFolder(args[0]);
            }
            else if (args.Length == 0)
            {
                string path = System.IO.Directory.GetCurrentDirectory();
                MessageHandler.ProcessCompleteFolder(path);
            }
            else
            {
                HandleConsoleWindow.ShowConsoleWindow();
                Console.WriteLine("The given Argument is not valid. Please use a *.msg file or a folder path!");
                Console.WriteLine($"Argument: {args[0]}");
                Console.WriteLine();
                Console.WriteLine("Please press any key to close this application...");
                Console.ReadKey();
            }
        }
    }
}