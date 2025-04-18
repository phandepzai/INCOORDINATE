using System;
using System.Windows.Forms;

namespace ENTER_COORDINATE
{
    static class Program
    {
        /// <summary>
        /// Entry point chính của ứng dụng.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
