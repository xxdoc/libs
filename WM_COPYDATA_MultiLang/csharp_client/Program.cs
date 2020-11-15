using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace wm_copydata
{
    static class Program
    {
        public static string[] cmdline_args;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            cmdline_args = args;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
