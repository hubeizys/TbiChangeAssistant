using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TbiCA
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            //MessageBox.Show(config.AppSettings.Settings["dir"].Value);
            config.AppSettings.Settings["dir"].Value = "zhu ge  haoshuai";
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }

        private void barButtonItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }

        private void ribbonControl1_SelectedPageChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(ribbonControl1.SelectedPage.Name);    

            if(ribbonControl1.SelectedPage.PageIndex == 0)
            {
                this.xtraTabControl1.SelectedTabPageIndex = 0;
            }

            if (ribbonControl1.SelectedPage.PageIndex == 1)
            {
                this.xtraTabControl1.SelectedTabPageIndex = 1;
            }
        }
    }
}
