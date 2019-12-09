using System;
using System.Collections.Generic;
 
using System.Windows.Forms;

namespace ad_optimization
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {

           // Form1 _form = new Form1();

            //if (args.Length > 0) {

            //    if (args[0].ToString() == "debug") {

            //        _form.isDEBUG = true;

            //    }
            
            //}

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
