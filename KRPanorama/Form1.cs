using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using HtmlAgilityPack;
using Ookii.Dialogs.WinForms;
using System.Management;
using Sunny.UI;
using ExcelDataReader;


namespace KRPanorama
{
    public partial class Form1 : UIForm
    {
        public Form1()
        {
            InitializeComponent();
            frm1 = this;
        }
        public static Form1 frm1;

        public string path;//单文件路径
        public string ofolder;//打开路径
        public string sfolder;//保存路径
        public DirectoryInfo[] fnames;//打开的名称
        public string oPath;//打开时使用的路径
        public Thread th;//进程
        List<FileSystemInfo> files = new List<FileSystemInfo>();
        DataTableCollection tableCollection;
        Tools tool = new Tools();



        delegate void SafeSetText(string strMsg);
        public void SetText(string strMsg)//代理文本更新
        {
            SafeSetText objSet = delegate (string str)
            {
                uiRichTextBox1.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")+" - " + str);
            };
            this.Invoke(objSet, new object[] { strMsg });
        }
        /// <summary>
        /// 切片
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton1_Click(object sender, EventArgs e)
        {
            string pro = Thread.GetDomain().BaseDirectory + "krpanotools64.exe";
            if (ofolder.IsNullOrWhiteSpace())
            {
                UIMessageBox.Show("请先选择打开目录","提示",UIStyle.Blue, UIMessageBoxButtons.OK,false);
            }
            else if(File.Exists(pro))
            {
                string[] folders = Directory.GetDirectories(ofolder, "*", SearchOption.TopDirectoryOnly);
                foreach (var fold in folders)
                {
                    th = new Thread(delegate () {
                        SetText("start:" + fold + "\n");
                        string path = "(\"" + pro + "\" makepano -config=templates/vtour-multires.config askforxmloverwrite=false \"" + fold + "\\*\")";//bat路径
                        tool.bat(path);
                        SetText(fold + " processing done.\n");
                    });
                    th.IsBackground = true;
                    th.Start();
                }
            }
            else {
                UIMessageBox.Show("未找到krpanotools64.exe,请将krpano复制到程序目录", "提示", UIStyle.Blue, UIMessageBoxButtons.OK, false);
            }
        }
        /// <summary>
        /// 修改文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton2_Click(object sender, EventArgs e)
        {
            if (ofolder.IsNullOrWhiteSpace())
            {
                
                UIMessageBox.Show("请先选择打开目录", "提示",UIStyle.Blue, UIMessageBoxButtons.OK,false);
            }
            else
            {
                files.Clear();
                files = tool.fileDirector(ofolder);
                foreach (FileSystemInfo file in files)
                {
                    FileInfo fil = file as FileInfo;
                    tool.reple(fil);
                    SetText(fil.FullName + " edit done\n");
                }
            }
        }
        /// <summary>
        /// 打开匹配文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton3_Click(object sender, EventArgs e)
        {
            var of = new VistaOpenFileDialog();
            of.Filter = "Excel(*.xlsx,*.xls)|*.xlsx;*.xls|all|*.*";
            DialogResult result = of.ShowDialog();
            if (result == DialogResult.OK)
            {
                using (var stream = System.IO.File.Open(of.FileName, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        DataSet results = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        tableCollection = results.Tables;
                        
                        uiDataGridView1.DataSource = tableCollection[0];
                        uiComboBox1.Items.Clear();
                        uiComboBox2.Items.Clear();
                        foreach (DataTable table in tableCollection)
                        {
                            
                            foreach (DataColumn item in table.Columns)
                            {

                                uiComboBox1.Items.Add(item.ColumnName);
                                uiComboBox2.Items.Add(item.ColumnName);
                            }                            
                            //uiComboBox1.Items.Add(table.TableName);
                        }
                        uiComboBox1.Text = uiComboBox1.Items[0].ToString();
                        uiComboBox2.Text = uiComboBox1.Items[1].ToString();
                    }
                }
            }

        }
        /// <summary>
        /// 标准化文件名
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton4_Click(object sender, EventArgs e)
        {
            if (sfolder.IsNullOrWhiteSpace()||ofolder.IsNullOrWhiteSpace())
            {
                UIMessageBox.Show("请先设置保存目录和打开目录", "提示", UIStyle.Blue, UIMessageBoxButtons.OK,false);
            }
            uiButton8_Click(sender, e);
            if (uiTreeView1.Nodes.Count ==0)
            {
                UIMessageBox.Show("未找到全景，请先切片或重新选择目录", "提示", UIStyle.Blue, UIMessageBoxButtons.OK, false);
            }
            else if(tableCollection is null)
            {
                UIMessageBox.Show("请先打开文件", "提示", UIStyle.Blue, UIMessageBoxButtons.OK,false);
            }
            else
            {
                string name = uiComboBox1.Text;
                string code = uiComboBox2.Text;
                Dictionary<string, string> data = new Dictionary<string, string>();
                List<string> names = tableCollection[0].AsEnumerable().Select(it => it.Field<string>(name)).ToList();
                List<string> codes = tableCollection[0].AsEnumerable().Select(it => it.Field<string>(code)).ToList();
                for (var i =0;i<names.Count;i++)
                {
                    string file = ofolder + "\\" + names[i] + "\\vtour\\tour.xml";
                    if (System.IO.File.Exists(file))
                    {
                        string folder = sfolder + "\\" + codes[i];
                        Directory.CreateDirectory(folder);
                        Directory.Move(ofolder + "\\" + names[i] + "\\vtour", folder + "\\vtour");
                    }
                    else
                    {
                        SetText(names[i]+"not exsit\n");
                    }
                }
                UIMessageBox.Show("完成", "提示", UIStyle.Blue, UIMessageBoxButtons.OK, false);
                //var d=uiDataGridView1.Columns[name].Index;
                //var s = uiDataGridView1.Columns[code].Index;

                //List<string> names = dt.AsEnumerable().Select(it => it.Field<int>(0)).ToList();
            }
        }
        /// <summary>
        /// 打开程序时打开上一次的路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            string InputText = Properties.Settings.Default.InputText;
            string OutputText = Properties.Settings.Default.OutputText;
            
            if (!InputText.IsNullOrWhiteSpace())
            {
                uiLabel3.Text = Properties.Settings.Default.InputText;
                ofolder = uiLabel3.Text;
            }
            if (!OutputText.IsNullOrWhiteSpace())
            {
                uiLabel4.Text = Properties.Settings.Default.OutputText;
                sfolder = uiLabel4.Text;
            }

            //uiDataGridView1.DataSource = Properties.Settings.Default.ExcelFile;
        }
       /// <summary>
       /// 设置打开目录
       /// </summary>
       /// <param name="sender"></param>
       /// <param name="e"></param>
        private void uiButton5_Click(object sender, EventArgs e)
        {
            var ofd = new VistaFolderBrowserDialog();
            DialogResult result = ofd.ShowDialog();
            if (result == DialogResult.OK)
            {
                ofolder = ofd.SelectedPath;
                oPath = ofd.SelectedPath;
                uiLabel3.Text = ofolder;
                Properties.Settings.Default.InputText = ofolder;
                Exist(ofolder);
            }
        }
        /// <summary>
        /// 设置保存目录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton6_Click(object sender, EventArgs e)
        {
            var ofd = new VistaFolderBrowserDialog();
            DialogResult result = ofd.ShowDialog();
            if (result == DialogResult.OK)
            {
                sfolder = ofd.SelectedPath;
                uiLabel4.Text = sfolder;
                Properties.Settings.Default.OutputText = sfolder;
            }
        }
        /// <summary>
        /// 清理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton7_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(ofolder)) { return; }
            string[] dirs = Directory.GetDirectories(ofolder, "vtour", SearchOption.AllDirectories);
            foreach (string dir in dirs)
            {
                th = new Thread(delegate () {
                    Directory.Delete(dir, true);
                    SetText("delete:" + dir + "\n");
                });
                th.IsBackground = true;
                th.Start();
            }
        }
        /// <summary>
        /// 判断全景切片是否存在，并将存在的添加进treeview
        /// </summary>
        /// <param name="fnames"></param>
        public void Exist(string ofolder)
        {
            if (string.IsNullOrEmpty(ofolder)) { return; }
            DirectoryInfo directoryInfo = new DirectoryInfo(ofolder);
            fnames = directoryInfo.GetDirectories("*", SearchOption.TopDirectoryOnly);
            string file="";
            uiTreeView1.Nodes.Clear();
            foreach (DirectoryInfo f in fnames) {
                file = f.FullName + "\\vtour\\tour.xml";

                if (System.IO.File.Exists(file))
                {
                    uiTreeView1.Nodes.Add(f.Name);
                }
            }
                
        }
        /// <summary>
        /// 打开全景查看
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton9_Click(object sender, EventArgs e)
        {
            string pro = Thread.GetDomain().BaseDirectory + "krpano Tools.exe";
            TreeNode sNode = uiTreeView1.SelectedNode;
            if (sNode is null) { return; }
            string file = oPath + "\\" + sNode.Text + "\\vtour\\tour.xml";
            if (!File.Exists(pro))
            {
                UIMessageBox.Show("未找到krpano Tools.exe,请将krpano复制到程序目录", "提示", UIStyle.Blue, UIMessageBoxButtons.OK, false);
                return;
            }
            else if (System.IO.File.Exists(file))
            {
                string path = "(\"" + Thread.GetDomain().BaseDirectory + "krpano Tools.exe\" \"" +file+ "\")";//命令
                th = new Thread(delegate () {
                    tool.bat(path);
                });
                th.IsBackground = true;
                th.Start();
                //System.Drawing.Icon.ExtractAssociatedIcon("\"" + Thread.GetDomain().BaseDirectory + "krpano Tools.exe\"").ToBitmap().Save(Thread.GetDomain().BaseDirectory+"\\krpano.ico");
            }
            else
            {
                UIMessageBox.Show("文件不存在，请刷新", "提示", UIStyle.Blue, UIMessageBoxButtons.OK,false);
            }
        }

        private void uiTreeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }
        /// <summary>
        /// 双击打开全景
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiTreeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            uiButton9_Click(sender, e);
        }
        /// <summary>
        /// 刷新全景打开目录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton8_Click(object sender, EventArgs e)
        {
            Exist(ofolder);
            oPath = ofolder;
        }
        /// <summary>
        /// 刷新全景保存目录的
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton10_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(sfolder)) { return; }
            Exist(sfolder);
            oPath = sfolder;
        }

        private void uiRichTextBox1_TextChanged(object sender, EventArgs e)
        {
            uiRichTextBox1.SelectionStart = uiRichTextBox1.Text.Length;
            uiRichTextBox1.ScrollToCaret();
        }

        /// <summary>
        /// 自动编号
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            e.Row.HeaderCell.Value = string.Format("{0}", e.Row.Index + 1);           
        }
        /// <summary>
        /// 保存记录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.Save();
            //保存内容到txt
            if (!uiRichTextBox1.Text.IsNullOrWhiteSpace())
            {
                Directory.CreateDirectory(Thread.GetDomain().BaseDirectory + "LOG");
                string path = Thread.GetDomain().BaseDirectory + "\\LOG\\" + DateTime.Now.ToString("yyyy-MM-dd") + ".log";
                StreamWriter sw = File.AppendText(path);
                sw.Write(uiRichTextBox1.Text);
                //释放资源
                sw.Close();
            }
        }
    }
}
