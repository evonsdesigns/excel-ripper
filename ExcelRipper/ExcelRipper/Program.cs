//
// EvonsDesigns Excel-Ripper Application for Windows
// The following is copyright 2012 EvonsDesigns
// Author: Joe Evans (evonsdesigns@gmail.com)
//
using System;
using System.Windows.Forms;

namespace ExcelRipper
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
