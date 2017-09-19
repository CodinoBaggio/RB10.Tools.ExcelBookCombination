using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RB10.Tools.ExcelBookCombination
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if(0 < args.Length && args[0].ToLower() == "/s")
            {
                Execute();
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new FolderSelectForm());
            }
        }

        static void Execute()
        {
            foreach (var folder in System.IO.Directory.GetDirectories(Properties.Settings.Default.TargetFolder))
            {
                var bc = new BookCombination(folder, Properties.Settings.Default.OutputParentFolder);
                bc.Run();
            }
        }
    }
}
