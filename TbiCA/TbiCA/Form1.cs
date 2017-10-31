using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
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

            if (ribbonControl1.SelectedPage.PageIndex == 0)
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
            if (config_dlg.ShowDialog() == DialogResult.OK)
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
            if (add_dst_dir.ShowDialog() == DialogResult.OK)
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
            catch (Exception err)
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



        /// <summary>
        /// 将DataTable中数据写入到CSV文件中
        /// </summary>
        /// <param name="dt">提供保存数据的DataTable</param>
        /// <param name="fileName">CSV的文件路径</param>
        public void SaveCSV(DataTable dt, string fileName)
        {
            FileStream fs = new FileStream(fileName, System.IO.FileMode.Append, System.IO.FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Unicode);
            string data = "";
            //写出列名称
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                data += dt.Columns[i].ColumnName.ToString();
                if (i < dt.Columns.Count - 1)
                {
                    data += "\t";
                }
            }
            // sw.WriteLine(data);
            sw.Write(data);
            sw.Write("\n");
            //写出各行数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data = "";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    data += dt.Rows[i][j].ToString();
                    if (j < dt.Columns.Count - 1)
                    {
                        data += "\t";
                    }
                }
                sw.Write(data);
                sw.Write("\n");
            }
            sw.Close();
            fs.Close();
            MessageBox.Show("CSV文件保存成功！");
        }


        /// <summary>
        /// 将CSV文件的数据读取到DataTable中
        /// </summary>
        /// <param name="fileName">CSV文件路径</param>
        /// <returns>返回读取了CSV数据的DataTable</returns>
        public DataTable OpenCSV(string fileName)
        {
            DataTable dt = new DataTable();
            FileStream fs = new FileStream(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            StreamReader sr = new StreamReader(fs, System.Text.Encoding.Default);
            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] aryLine;
            //标示列数
            int columnCount = 0;
            //标示是否是读取的第一行
            bool IsFirst = true;
            //逐行读取CSV中的数据
            int line = 0;
            while ((strLine = sr.ReadLine()) != null)
            {

                if (line == 2)
                {
                    //aryLine = strLine.Split('\\');
                    line++;
                }
                else if (line < 3)
                {
                    line++;
                    continue;
                }
                aryLine = strLine.Split('\t');

                columnCount = aryLine.Length;
                if (IsFirst == true)
                {
                    IsFirst = false;

                    //创建列
                    for (int i = 0; i < columnCount; i++)
                    {
                        DataColumn dc = new DataColumn(aryLine[i]);
                        dt.Columns.Add(dc);
                    }
                }
                else
                {
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < columnCount; j++)
                    {
                        dr[j] = aryLine[j];
                    }
                    dt.Rows.Add(dr);
                }
            }
            sr.Close();
            fs.Close();
            return dt;
        }


        private void simpleButton3_Click(object sender, EventArgs e)
        {
            string file_path = searchLookUpEdit1.Properties.NullValuePrompt;
            /*
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
            */
            try
            {
                //string[] will_create_arr_pngs = { };


                // 1.0  读入 配置文件,  根据CVS 中的 picture 找到 需要处理的图片
                string ext_name = Path.GetExtension(default_label.Text);
                if (ext_name != ".csv")
                {
                    MessageBox.Show("文件格式不正确");
                    return;
                }
                DataTable dt = this.OpenCSV(default_label.Text);

                foreach (DataRow dr in dt.Rows)
                {

                    // 把 在配置文件中 老图片 和将要修改的图片 都放置到队列中去
                    #region 处理 新图片
                    string pic_value = dr["新图片"].ToString();
                    // 分割字符串
                    string[] arr_pic = pic_value.Split('|');
                    List<string> will_create_arr_pngs = new List<string>();
                    List<string> old_tbi_list = new List<string>();
                    foreach (string src_pic in arr_pic)
                    {
                        // 原始pic   分割 然后依次处理
                        // 分割原始pic
                        string[] src_2_pic = src_pic.Split(':');
                        if (src_2_pic.Length != 4)
                        {
                            continue;
                        }
                        string guolv = "\"!;";
                        src_2_pic[0] = src_2_pic[0].Trim(guolv.ToCharArray());
                        Console.WriteLine("src_2_pic== !{0}!    == ", src_2_pic[0]);

                        // 根据 用户填的名字  生产一个 文件名
                        string real_file_name = file_path + "\\" + textEdit1.Text + "\\" + "web" + "\\" + src_2_pic[0] + ".png";
                        will_create_arr_pngs.Add(real_file_name);

                        string old_file_name = file_path + "\\" + src_2_pic[0] + ".tbi";
                        old_tbi_list.Add(old_file_name);
                    }
                    #endregion


                    #region 创建目标文件夹
                    if (!Directory.Exists(textEdit1.Text))
                    {
                        Directory.CreateDirectory(textEdit1.Text);
                    }
                    #endregion

                    // 1.1   tbi 转为 png
                    // 1.2  并把文件放入指定的文件夹中
                    #region 放入指定的文件夹中

                    string real_dir = file_path + "\\" + textEdit1.Text + "\\" + "web" + "\\";
                    if (!Directory.Exists(real_dir))
                    {
                        Directory.CreateDirectory(real_dir);
                    }
                    for (int nn = 0; nn < will_create_arr_pngs.Count; nn++)
                    {
                        System.IO.File.Copy(old_tbi_list[nn], will_create_arr_pngs[nn]);
                    }

                    #endregion


                    // 先在datatable 中修改好 需要修改的内容
                    #region 修改配置文件中的内容

                    string old_dscription = dr["宝贝描述"].ToString();
                    old_dscription = old_dscription.Replace(" (", "_");
                    old_dscription = old_dscription.Replace(")", "");
                    old_dscription = old_dscription.Replace("\"", "");
                    //old_dscription = old_dscription.Replace();
                    old_dscription = old_dscription.Trim();
                    MessageBox.Show(old_dscription);

                    WebBrowser wb2 = new WebBrowser();
                    wb2.Navigate("about:blank");
                    wb2.Document.Write(old_dscription);
                    wb2.DocumentText = old_dscription;
                    HtmlDocument doc2 = wb2.Document;

                    for (int nn = 0; nn < will_create_arr_pngs.Count; nn++)
                    {
                        HtmlElement img_ele = doc2.CreateElement("img");
                        img_ele.SetAttribute("width", "1");
                        img_ele.SetAttribute("height", "1");
                        img_ele.SetAttribute("src", will_create_arr_pngs[nn]);
                        doc2.Body.AppendChild(img_ele);
                    }
                    dr["宝贝描述"] = "";
                    foreach (HtmlElement et in doc2.GetElementsByTagName("body"))
                    {
                        //Console.WriteLine(et.InnerHtml);

                        dr["宝贝描述"] += et.InnerHtml;
                    }
                    MessageBox.Show(dr["宝贝描述"].ToString());
                    #endregion

                }

                // 等到文件都处理结束了在去写  csv 文件
                // 2.1 在文件夹中创建一个 复制 一个【处理前的文件】 作为处理后的文件的 文本
                #region 创建一个空的模板文件
                string dst_temp_dir = file_path + "\\" + textEdit1.Text;
                string dst_temp_file = dst_temp_dir + "\\" + "template.csv";
                if (!Directory.Exists(dst_temp_dir))
                {
                    Directory.CreateDirectory(dst_temp_dir);
                }
                System.IO.File.Copy("template.csv", dst_temp_file);

                // 2.2 转好的datatable 放进去
                this.SaveCSV(dt, dst_temp_file);
                #endregion
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }

        }

        private void barButtonItem4_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

            // 1 转换每个图片 
            // 1.0  读入 配置文件,  根据CVS 中的 picture 找到 需要处理的图片
            // 给每个文件 基础文件变换名字  例如[c28267bcb20d431532f48d945614623e:1:0: ---- > c28267bcb20d431532f48d945614623e.tbi]
            // 1.1   tbi 转为 png
            // 1.2  放入指定的文件夹中
            // 创建指定文件夹
            // 1.3  把路径记录下来  [路径A-----]
            // 1.3.1 把以前的路径 名字转换一下  放入数组 Arr_DIR 中
            // 1.3.2 把新建的文件名放进数组 ARR_dIR 中

            // 2 图标信息记录在目标描述里面
            // 2.1 在文件夹中创建一个 复制 一个【处理前的文件】 作为处理后的文件的 文本
            // 2.2 转好的文件放进去


            webBrowser1.Navigate("about:blank");
            webBrowser1.Document.Write("<IMG align=middle src=\"FILE:///D:\\6月份完整数据包\\淘宝数据库\\情趣挑逗\\contentPic\\601049-601051 水晶套聚品\\images\\601049 (1).jpg\"><IMG align=middle src=\"FILE:///D:\\6月份完整数据包\\淘宝数据库\\情趣挑逗\\contentPic\\601049-601051 水晶套聚品\\images\\601049 (2).jpg\"><IMG align=middle src=\"FILE:///D:\\6月份完整数据包\\淘宝数据库\\情趣挑逗\\contentPic\\601049-601051 水晶套聚品\\images\\601049 (3).jpg\"><IMG align=\"middle\" src=\"FILE:///D:\\6月份完整数据包\\淘宝数据库\\情趣挑逗\\contentPic\\601049-601051 水晶套聚品\\images\\601049 (4).jpg\"><IMG align=middle src=\"FILE:///D:\\6月份完整数据包\\淘宝数据库\\情趣挑逗\\contentPic\\601049-601051 水晶套聚品\\images\\601049 (5).jpg\">");
            webBrowser1.DocumentText = "<IMG align=middle src=\"FILE:///D:\\6月份完整数据包\\淘宝数据库\\情趣挑逗\\contentPic\\601049-601051 水晶套聚品\\images\\601049 (1).jpg\"><IMG align=middle src=\"FILE:///D:\\6月份完整数据包\\淘宝数据库\\情趣挑逗\\contentPic\\601049-601051 水晶套聚品\\images\\601049 (2).jpg\"><IMG align=middle src=\"FILE:///D:\\6月份完整数据包\\淘宝数据库\\情趣挑逗\\contentPic\\601049-601051 水晶套聚品\\images\\601049 (3).jpg\"><IMG align=\"middle\" src=\"FILE:///D:\\6月份完整数据包\\淘宝数据库\\情趣挑逗\\contentPic\\601049-601051 水晶套聚品\\images\\601049 (4).jpg\"><IMG align=middle src=\"FILE:///D:\\6月份完整数据包\\淘宝数据库\\情趣挑逗\\contentPic\\601049-601051 水晶套聚品\\images\\601049 (5).jpg\">";
            HtmlDocument doc = webBrowser1.Document;
            HtmlElementCollection elemColl = doc.GetElementsByTagName("IMG");
            foreach (HtmlElement elem in elemColl)
            {
                Console.WriteLine(elem.GetAttribute("src"));
            }

            WebBrowser wb2 = new WebBrowser();
            wb2.Navigate("about:blank");
            wb2.Document.Write("<html><head></head><Body></Body></html>");
            wb2.DocumentText = "<html><head></head><Body></Body></html>";
            HtmlDocument doc2 = wb2.Document;

            HtmlElement img_ele = doc2.CreateElement("img");
            //img_ele.InnerText = "sdasdsd";
            doc2.Body.AppendChild(img_ele);

            foreach (HtmlElement et in doc2.GetElementsByTagName("body"))
            {
                Console.WriteLine(et.InnerHtml);
            }
        }

        private void te_willdonefile_Click(object sender, EventArgs e)
        {
            if (ofd_2willdone.ShowDialog() == DialogResult.OK)
            {
                //Console.WriteLine();
                MessageBox.Show(ofd_2willdone.FileName);
                te_willdonefile.Text = ofd_2willdone.FileName;
            }
        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {

            Thread t = new Thread(()=> {
                //try
                //{
                    // 1.0  读入 配置文件,  根据CVS 中的 picture 找到 需要处理的图片
                    string ext_name = Path.GetExtension(te_willdonefile.Text);
                    if (ext_name != ".csv")
                    {
                        MessageBox.Show("文件格式不正确");
                        return;
                    }
                    DataTable dt = this.OpenCSV(te_willdonefile.Text);


                    //Console.WriteLine( dt.Rows[0]["宝贝描述"].ToString());



                    foreach (DataRow dr in dt.Rows)
                    {
                        string _l_baobeimiaoshu = "";

                        _l_baobeimiaoshu = dr["宝贝描述"].ToString();
                        //Console.WriteLine(dr["宝贝描述"].ToString());

                        // _l_baobeimiaoshu.Split("</table>");
                        string[] sarr = Regex.Split(_l_baobeimiaoshu, "</table>");

                        /*
                        foreach (string i in sarr)
                        {
                            Console.WriteLine("\n aa= " + i);
                        }*/



                        #region 截取dom树 并进行md5 加密
                        wb_willdone.Invoke(new Action(
                            () =>
                            {
                                WebBrowser wb_willdone = new WebBrowser();
                                wb_willdone.Navigate("about:blank");
                                try { string _l_willdonstr = sarr[1];
                                    _l_willdonstr = _l_willdonstr.Replace("\"\"", "\"");
                                    //Console.WriteLine("_l_willdonstr =========  " + _l_willdonstr);
                                    textBox_result.BeginInvoke(new Action(() => {
                                        textBox_result.ForeColor = Color.Black;
                                        textBox_result.AppendText(DateTime.Now.ToString("HH:mm:ss  "));
                                        textBox_result.AppendText("正在处理dom =========  " + _l_willdonstr);
                                        textBox_result.AppendText(Environment.NewLine);
                                        textBox_result.ScrollToCaret();
                                    }));
                                    wb_willdone.Document.Write(_l_willdonstr);
                                    wb_willdone.DocumentText = _l_willdonstr;
                                }
                                catch (Exception err)
                                {
                                    int cur_line = (dt.Rows.IndexOf(dr) + 4);
                                    MessageBox.Show("图片描述的格式似乎不正确" +err.Message + "第： " + cur_line + "行");
                                }
                                

                                HtmlDocument doc_willdone = wb_willdone.Document;
                                string _temp_one_line = "";
                                int temp_num = 0;
                                foreach (HtmlElement et in doc_willdone.GetElementsByTagName("img"))
                                {
                                    string _tmp_src = et.GetAttribute("src");
                                    //Console.WriteLine(_tmp_src + "md5: =" + System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(_tmp_src, "MD5"));

                                    textBox_result.BeginInvoke(new Action(() => {
                                        textBox_result.ForeColor = Color.Red;
                                        textBox_result.AppendText(DateTime.Now.ToString("HH:mm:ss  "));
                                        textBox_result.AppendText("正在处理的图片是 =========  " + _tmp_src);
                                        textBox_result.AppendText(Environment.NewLine);
                                        textBox_result.ScrollToCaret();
                                    }));
                                    string _temp_md5 = System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(_tmp_src, "MD5");
                                    _temp_one_line += string.Format("{0}:1:{1}:|{2};", _temp_md5, temp_num++, _tmp_src);
                                }

                                dr["新图片"] = _temp_one_line;
                                // Console.WriteLine("_temp_one_line ===== " + _temp_one_line);
                                textBox_result.BeginInvoke(new Action(() => {
                                    textBox_result.ForeColor = Color.Green;
                                    textBox_result.AppendText(DateTime.Now.ToString("HH:mm:ss  "));
                                    textBox_result.AppendText("合成的新图片为 =========  " + _temp_one_line);
                                    textBox_result.AppendText(Environment.NewLine);
                                    textBox_result.ScrollToCaret();
                                }));
                                try
                                {
                                    #endregion
                                    lc_result.Invoke(new Action(() =>
                                    {
                                        lc_result.Text = string.Format("已经完成 {0}/{1}", dt.Rows.IndexOf(dr), dt.Rows.Count);
                                        lc_result.Refresh();
                                    }));
                                }
                                catch (Exception err)
                                {
                                    Console.WriteLine("aaaa ", err.Message);
                                }

                            }
                            ));

                    }

                    string dst_temp_dir = Path.GetDirectoryName(te_willdonefile.Text);
                    string goal_filename = Path.GetRandomFileName();
                    string dst_temp_file = dst_temp_dir + "\\" + "新转化" + goal_filename + ".csv";

                    Console.WriteLine("dst_temp_file === " + dst_temp_file);

                    if (!Directory.Exists(dst_temp_dir))
                    {
                        Directory.CreateDirectory(dst_temp_dir);
                    }
                    System.IO.File.Copy("template.csv", dst_temp_file);

                    // 2.2 转好的datatable 放进去
                    this.SaveCSV(dt, dst_temp_file);
                    /*
                }

                catch (Exception err)
                {
                    MessageBox.Show("err : " + err.Message);
                }*/
            });
            t.Start();
            
        }
    }
}
