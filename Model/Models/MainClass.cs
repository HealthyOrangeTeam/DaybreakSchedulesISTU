using System.Globalization;
using Document = Spire.Doc.Document;
using Section = Spire.Doc.Section;
using Spire.Doc;
using Spire.Doc.Documents;
using Model.Models.Entities;
using Model.Interface;

namespace Model.Models
{
    public class MainClass
    {
        public IDocumentRead Read { get; set; }
        public ISiteParse SiteParse { get; set; }

        #region тоже надо!!
        public void Process()
        {
            using var watcher = new FileSystemWatcher("temp\\");

            watcher.NotifyFilter = NotifyFilters.Attributes
                                 | NotifyFilters.CreationTime
                                 | NotifyFilters.DirectoryName
                                 | NotifyFilters.FileName
                                 | NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.Security
                                 | NotifyFilters.Size;
            watcher.Changed += OnChanged;
            watcher.Filter = "*.docx";
            watcher.IncludeSubdirectories = false;
            watcher.EnableRaisingEvents = true;



            while (true)
            {
                if(Console.ReadKey().Key == ConsoleKey.Spacebar)
                {
                    SiteParse.Parse();
                }
            }
        
        }

        private void OnChanged(object sender, FileSystemEventArgs e)
        {
            if (e.ChangeType != WatcherChangeTypes.Changed)
            {
                return;
            }
            Console.WriteLine($"Changed: {e.FullPath}");

            Read.Parse(e.FullPath);
                    Console.WriteLine("END PARSE");
        }
        #endregion
    }
}
