using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TestGetFiles;

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


            try
            {
                using (NetUtilityLib.ExcelHelper exc = new NetUtilityLib.ExcelHelper("dst.xlsx"))
                {
                    DataTable dt = exc.ExcelToDataTable("sheet1", true);
                    /*
                    DataRow dr = dt.NewRow();
                    dr["dst"] = add_dst_dir.SelectedPath;
                    dt.Rows.Add(dr);
                    exc.DataTableToExcel(dt, "sheet1", true);
                    */
                    searchLookUpEdit1.Properties.DataSource = dt;
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
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

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if(config_dlg.ShowDialog() == DialogResult.OK)
            {
                // MessageBox.Show(config_dlg.FileName);
                this.default_label.Text = config_dlg.FileName;
            }
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            if (dst_dlg.ShowDialog() == DialogResult.OK)
            {
                searchLookUpEdit1.Properties.NullValuePrompt = dst_dlg.SelectedPath;
            }
       }

        private void button2_Click(object sender, EventArgs e)
        {
            if(add_dst_dir.ShowDialog() == DialogResult.OK)
            {
                searchLookUpEdit3.Properties.NullText = add_dst_dir.SelectedPath;
            }
            try
            {
                using (NetUtilityLib.ExcelHelper exc = new NetUtilityLib.ExcelHelper("dst.xlsx"))
                {
                    DataTable dt = exc.ExcelToDataTable("sheet1", true);
                    DataRow dr = dt.NewRow();
                    dr["dst"] = add_dst_dir.SelectedPath;
                    dt.Rows.Add(dr);
                    exc.DataTableToExcel(dt, "sheet1", true);
                    searchLookUpEdit3.Properties.DataSource = dt;
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (add_src_dir.ShowDialog() == DialogResult.OK)
            {
                searchLookUpEdit2.Properties.NullText = add_src_dir.SelectedPath;
            }

            try
            {
                /*
                using (NetUtilityLib.ExcelHelper exc = new NetUtilityLib.ExcelHelper("工作簿1.xlsx"))
                {
                    DataTable dt = exc.ExcelToDataTable("dst", true);

                    DataRow dr = dt.NewRow();
                    dr["dst"] = add_src_dir.SelectedPath;
                    dt.Rows.Add(dr);
                    searchLookUpEdit2.Properties.DataSource = dt;
                }*/
                using (NetUtilityLib.ExcelHelper exc = new NetUtilityLib.ExcelHelper("src.xlsx")) 
                {
                    DataTable dt = exc.ExcelToDataTable("sheet1", true);
                    DataRow dr = dt.NewRow();
                    dr["src"] = add_src_dir.SelectedPath;
                    dt.Rows.Add(dr);

                    exc.DataTableToExcel(dt, "sheet1", true);
                    searchLookUpEdit2.Properties.DataSource = dt;
                }
                
            }
            catch(Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        private void searchLookUpEdit2_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void searchLookUpEdit1_Click(object sender, EventArgs e)
        {


        }

        private void searchLookUpEdit1_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {

        }

        private void searchLookUpEdit1_Popup(object sender, EventArgs e)
        {

        }

        private void xtraTabControl1_Selected(object sender, DevExpress.XtraTab.TabPageEventArgs e)
        {
            try
            {
                using (NetUtilityLib.ExcelHelper exc = new NetUtilityLib.ExcelHelper("dst.xlsx"))
                {
                    DataTable dt = exc.ExcelToDataTable("sheet1", true);
                    /*
                    DataRow dr = dt.NewRow();
                    dr["dst"] = add_dst_dir.SelectedPath;
                    dt.Rows.Add(dr);
                    exc.DataTableToExcel(dt, "sheet1", true);
                    */
                    searchLookUpEdit1.Properties.DataSource = dt;
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(textEdit1.Text))
            {
                Directory.CreateDirectory(textEdit1.Text);
            }
            string filename = searchLookUpEdit1.Properties.NullValuePrompt;
            string fils = bsGetFiles.GetFiles(new DirectoryInfo(filename), "*.TBI");
            // MessageBox.Show(fils);

            string[] sArray = fils.Split(';');

            foreach (string i in sArray)
            {
                string[] text_arr = i.Split('.');
                if (text_arr.Length != 2)
                { continue; }
                string file_path = Path.GetDirectoryName(i);

                string real_file_name = file_path + "\\" + textEdit1.Text + "\\" + Path.GetFileNameWithoutExtension(i) + ".png";


                string real_dir = Path.GetDirectoryName(i) + "\\" + textEdit1.Text;
                if (!Directory.Exists(real_dir))
                {
                    Directory.CreateDirectory(real_dir);
                }


                System.IO.File.Copy(i, real_file_name);


                // System.IO.File.Move(i, text_arr[0] + ".TBI");
            }
        }
    }
}
