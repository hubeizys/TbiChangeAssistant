using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DevExpress.UserSkins;
using DevExpress.Skins;

namespace TbiCA
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {

            DateTime dt = Convert.ToDateTime("2017-11-5 12:22:05");
            DateTime nowt = DateTime.Now;//现在的时间
            if (dt <= nowt)
            {
                MessageBox.Show("程序已经过期了");
                System.Environment.Exit(0);
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            BonusSkins.Register();
            SkinManager.EnableFormSkins();
            Application.Run(new Form1());
        }
    }
}
