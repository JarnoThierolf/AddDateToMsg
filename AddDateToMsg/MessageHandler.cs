using iwantedue;
using System;

namespace AddDateToMsg
{
    class MessageHandler
    {
        public static void ProcessMsgList(string[] files)
        {
            foreach (string filename in files)
            {
                if (IsMsgFile(filename))
                {
                    OutlookStorage.Message message = new OutlookStorage.Message(filename);
                    DateTime date = message.ReceivedOrSentTime;
                    message.Dispose();
                    AddDateToMsgName(filename, date);
                }
            }
        }

        public static bool IsMsgFile(string filename)
        {
            return filename.Substring(filename.Length - 4) == ".msg";
        }

        public static void ProcessCompleteFolder(string path)
        {
            string[] files = System.IO.Directory.GetFiles(path);
            ProcessMsgList(files);
        }

        public static void AddDateToMsgName(string filename, DateTime date)
        {
            // round minutes
            if (date.Second >= 30)
                date = date.AddSeconds(30);

            string dateText = date.ToString("yyyyMMdd-HHmm_");

            // stop if dateText is allready in filename
            if (filename.Contains(dateText))
                return;

            int index = filename.LastIndexOf('\\') + 1;
            string newFilename = filename.Insert(index, dateText);

            if (System.IO.File.Exists(newFilename))
            {
                HandleConsoleWindow.ShowConsoleWindow();
                Console.WriteLine($"Die Datei: \"{newFilename}\" ist bereits vorhanden!");
                Console.WriteLine($"Die Datei: \"{filename}\" ist demnach doppelt und sollte gelöscht werden.");
                Console.WriteLine();
                Console.WriteLine("Please press any key to continue...");
                Console.ReadKey();
            }
            else
            {
                System.IO.File.Move(filename, newFilename);
            }
        }
    }
}
