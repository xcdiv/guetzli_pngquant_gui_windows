using IWshRuntimeLibrary;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using System.Xml.Serialization;
using DataMining.Common;
using System.Threading;
using RoleDomain.Common;

namespace ad_optimization
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataTable dt = new DataTable();
        List<PicList> list = new List<PicList>();
        public bool isDEBUG = false;
        public bool isUpdate = false;
        public int ProcessorCount = 1;

        private void Form1_Load(object sender, EventArgs e)
        {
            
            
            if (System.Configuration.ConfigurationSettings.AppSettings["debug"] != null && System.Configuration.ConfigurationSettings.AppSettings["debug"] == "true")
            {
                isDEBUG = true;

            }
            
            
            
            CreateShortcutOnDesktop();

            ProcessorCount = Environment.ProcessorCount;
            if (ProcessorCount>1) {

                ProcessorCount = ProcessorCount / 2;

            }

            this.ProcessorCount_val.Text = ProcessorCount.ToString();


            this.pictureBox1.SendToBack();//将背景图片放到最下面
            this.panel1.BackColor = Color.Transparent;//将Panel设为透明
            //this.panel1.Parent = this.pictureBox1;//将panel父控件设为背景图片控件
            this.panel1.BringToFront();//将panel放在前面
            this.panel1.Visible = false;

            this.Text = "河南有线广告素材优化工具" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

            using (StreamWriter sw = new StreamWriter(Application.StartupPath + @"\ad_optimization.txt"))
            {

                sw.Write(
                    System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
                    );
                sw.Close();
            }

            this.saveExcel.Filter = "Office Excel2007|*.xls";
        }

        private void openFile_Click(object sender, EventArgs e)
        {

            //判断保存的路径是否存在
            //if (!Directory.Exists(Application.StartupPath + @"\Test\Debug1"))
            //{
            //    //创建路径
            //    Directory.CreateDirectory(Application.StartupPath + @"\Test\Debug1");
            //}
            //设置默认打开路径(项目安装路径+Test\Debug1\)
            //openFileDialog1.InitialDirectory = Application.StartupPath + @"\Test\Debug1";
            //设置打开标题、后缀
            openFileDialog1.Title = "请选择导入jpg文件;png文件";
            openFileDialog1.Filter = "广告素材(*.png;*.jpg)|*.png;*.jpg|jpg文件(*.jpg)|*.jpg|png文件(*.png)|*.png|所有文件(*.*)|*.*";
            openFileDialog1.Multiselect = true;

            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            //if (openFileDialog1.ShowDialog() == DialogResult.OK)
            //{


            // }

            string[] picArr = openFileDialog1.FileNames;

            load_file(picArr, true);

            pictureBox1.BackColor = Color.Transparent;


        }

        private void load_file(string[] picArr, bool isadd)
        {


            //pictureBox1.Image=Image.FromFile(picArr[0]);

            if (picArr.Length > 0)
            {
                if (!isadd)
                {
                    list = new List<PicList>();
                }

                foreach (string pic in picArr)
                {
                    FileInfo fi = new FileInfo(pic);
                    int size = 0;
                    using (StreamReader sr = new StreamReader(fi.FullName))
                    {

                        size = sr.ReadToEnd().Length;
                        sr.Close();
                    }


                    if (!list.Exists(x => x.Filename == fi.FullName))
                    {
                        list.Add(new PicList()
                        {
                            Filename = fi.FullName
                            ,
                            Size = size
                            ,
                            Ext = fi.Extension

                        });
                    }
                }


            }
            this.dataGridView1.ResetBindings();
            this.dataGridView1.DataSource = list.ToArray();
            this.dataGridView1.Invalidate();
        }

        List<string> cmd_list = new List<string>();



        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

          

            for (int i = 0; i < list.Count; i++)
            {
                // _ManualEvents[i] = new ManualResetEvent(false);
                ThreadPool.QueueUserWorkItem(new WaitCallback(TCallback), i);
                //刷新UI
                //  SyncdataGridView();
            }

            // WaitHandle.WaitAll(_ManualEvents);
            //for (int i = 1; i <= cycleNum; i++)
            //{
            //    ThreadPool.QueueUserWorkItem(new WaitCallback(TCallback), i.ToString());
            //}
            //Console.WriteLine("主线程执行！");
            //Console.WriteLine("主线程结束！");
            //ThreadEvent.WaitOne();

            lock (locker)
            {
                while (runningThreads > 0)
                {
                    Monitor.Wait(locker);
                }
            }

         
        }


        private void backgroundWorker1_DoWorkCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //如果用户取消了当前操作就关闭窗口。
            if (e.Cancelled)
            {
                this.Close();
            }

            this.button2.Visible = true;
            this.button2.Text = "优化";

            //计算已经结束，需要禁用取消按钮。
            //this.btnCancel.Enabled = false;

            //计算过程中的异常会被抓住，在这里可以进行处理。
            if (e.Error != null)
            {
                Type errorType = e.Error.GetType();
                switch (errorType.Name)
                {
                    case "ArgumentNullException":
                    case "MyException":
                        //do something.
                        break;
                    default:
                        //do something.
                        break;
                }
            }

            //计算结果信息：e.Result
            //use it do something.

            timer1.Stop();

            this.label1.Text = cmd_list.Count.ToString() + "/" + list.Count.ToString();

            this.dataGridView1.ResetBindings();
            this.dataGridView1.DataSource = list.ToArray();

            progressBar1.Maximum = list.Count;
            progressBar1.Value = cmd_list.Count;

            int s_size = 0;
            int o_size = 0;
            foreach (PicList pic in list)
            {
                s_size += pic.Size;
                o_size += pic.Size_Out;
            }

            MessageBox.Show("执行完成！减少" + ((s_size - o_size) / 1024) + "KB");
        }

        const int cycleNum = 10;
        static int cnt = 10;
        static AutoResetEvent ThreadEvent = new AutoResetEvent(false);


        public void TCallback(object obj)
        {
            int i = int.Parse(obj.ToString());
            cnt -= 1;
            PicList pic = list[i];
            FileInfo fi = new FileInfo(pic.Filename);
            string aFirstName = pic.Filename.Substring(pic.Filename.LastIndexOf("\\") + 1, (pic.Filename.LastIndexOf(".") - pic.Filename.LastIndexOf("\\") - 1)); //文件名
            FileInfo fiOut = new FileInfo(Path.Combine(this.outDir.Text, aFirstName + fi.Extension));
            string _status = "";

            list[i].FilenameOut = fiOut.FullName;

            string _cmd = "";
            if (outBMP.Checked)
            {
                if (IntPtr.Size == 8)
                {
                    _cmd = "\"" + Application.StartupPath + "\\cmd\\guetzli_windows_x86-64.exe\" --quality 84 \"" + fi.FullName + " " + fiOut.FullName + "\"";

                    _status = cmd.cmd_exec(_cmd);
                }
                else
                {
                    _cmd = "\"" + Application.StartupPath + "\\cmd\\guetzli_windows_x86.exe\" --quality 84 \"" + fi.FullName + " " + fiOut.FullName + "\"";

                    _status = cmd.cmd_exec(_cmd);

                }

                try
                {
                    using (Bitmap bmp = new Bitmap(fiOut.FullName))
                    {
                        FileInfo fiOut_BMP = new FileInfo(Path.Combine(this.outDir.Text, aFirstName + ".bmp"));
                        if (fiOut_BMP.Exists)
                        {
                            fiOut_BMP.Delete();
                        }
                        bmp.Save(fiOut_BMP.FullName, ImageFormat.Bmp);
                        list[i].FilenameOut = fiOut_BMP.FullName;
                        list[i].Status = "生成成功";
                    }
                }
                catch (Exception err)
                {
                    list[i].Status = "生成失败";

                }


            }
            else
            {
                if (fi.Extension.ToLower() == ".jpg")
                {

                    if (IntPtr.Size == 8)
                    {
                        _cmd = "\"" + Application.StartupPath + "\\cmd\\guetzli_windows_x86-64.exe\" --quality 84 \"" + fi.FullName + "\" \"" + fiOut.FullName + "\"";
                        _status = cmd.cmd_exec(_cmd);
                    }
                    else
                    {
                        _cmd = "\"" + Application.StartupPath + "\\cmd\\guetzli_windows_x86.exe\" --quality 84 \"" + fi.FullName + "\" \"" + fiOut.FullName + "\"";
                        _status = cmd.cmd_exec(_cmd);

                    }
                }
                else if (fi.Extension.ToLower() == ".png")
                {

                    //fi.CopyTo(fiOut.FullName, true);
                    //                       _cmd = "\"" + Application.StartupPath + "\\cmd\\optipng.exe\" -o7 \"" + fiOut.FullName + "\"";
                    _cmd = "\"" + Application.StartupPath + "\\cmd\\pngquant.exe\" --force --verbose --quality=45-85 \"" + fi.FullName + "\" -o \"" + fiOut.FullName + "\"";

                    _status = cmd.cmd_exec(_cmd);
                }


                if (fiOut.Exists)
                {

                    int size = 0;
                    using (StreamReader sr = new StreamReader(fiOut.FullName))
                    {

                        size = sr.ReadToEnd().Length;
                        sr.Close();
                    }
                    pic.Size_Out = size;

                    if (pic.Size_Out == 0)
                    {
                        list[i].Status = "生成失败";
                    }
                    else if (list[i].Size_Out >= pic.Size)
                    {
                        fi.CopyTo(fiOut.FullName, true);
                        list[i].Status = "未优化";
                    }
                    else
                    {
                        list[i].Status = "生成成功";
                    }


                }
                else
                {

                    list[i].Status = "生成失败";

                }
            }



            cmd_list.Add(_cmd);


            //if (cnt == 0)
            //{
            //    ThreadEvent.Set();
            //}
        //    _ManualEvents[i].Set();

            lock (locker)
            {
                runningThreads--;
                Monitor.Pulse(locker);
            }
        }

        ManualResetEvent[] _ManualEvents;

        int runningThreads = 0;
        object locker = new object();
        
        
        private void button2_Click(object sender, EventArgs e)
        {

            if (!Directory.Exists(this.outDir.Text))
            {

                Directory.CreateDirectory(this.outDir.Text);
            }

            this.button2.Text = "开始优化";
            this.button2.Visible = false;

            BackgroundWorker backgroundWorker1 = new BackgroundWorker();
            backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            backgroundWorker1.RunWorkerCompleted += backgroundWorker1_DoWorkCompleted;
            backgroundWorker1.RunWorkerAsync();

            ThreadPool.SetMinThreads(1, 1);
            ThreadPool.SetMaxThreads(int.Parse(ProcessorCount_val.Text), int.Parse(ProcessorCount_val.Text));



            cmd_list = new List<string>();
            timer1.Enabled = true;
            timer1.Start();
            //_ManualEvents = new ManualResetEvent[list.Count];


            runningThreads = list.Count;
           

            //Console.WriteLine("线程池终止！");
            //Console.ReadKey();



            if (isDEBUG)
            {

                XmlSerializer xs = new XmlSerializer(typeof(List<PicList>));

                using (Stream stream = new FileStream(Environment.CurrentDirectory + "\\data.XML", FileMode.Create,

                  FileAccess.Write, FileShare.Read))
                {

                    xs.Serialize(stream, list);

                    stream.Close();
                }


                xs = new XmlSerializer(typeof(List<string>));

                using (Stream stream = new FileStream(Environment.CurrentDirectory + "\\cmd_log.XML", FileMode.Create,

                   FileAccess.Write, FileShare.Read))
                {

                    xs.Serialize(stream, cmd_list);

                    stream.Close();
                }
            }
            //this.dataGridView1.Rows[0].Selected = true;
            //this.dataGridView1.Invalidate();
        }

 

        private void button1_Click(object sender, EventArgs e)
        {
            //folderBrowserDialog1.ShowDialog();

            string defaultPath = "";

            //打开的文件夹浏览对话框上的描述  
            folderBrowserDialog1.Description = "请选择一个文件夹";
            //是否显示对话框左下角 新建文件夹 按钮，默认为 true  
            folderBrowserDialog1.ShowNewFolderButton = false;
            //首次defaultPath为空，按FolderBrowserDialog默认设置（即桌面）选择  
            if (defaultPath != "")
            {
                //设置此次默认目录为上一次选中目录  
                folderBrowserDialog1.SelectedPath = defaultPath;
            }
            //按下确定选择的按钮  
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                //记录选中的目录  
                defaultPath = folderBrowserDialog1.SelectedPath;
            }


            this.outDir.Text = defaultPath;
        }

        private void About_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("使用有问题，请联系QQ:32191100 \n email:cdiv@qq.com");
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            list = new List<PicList>();
            this.dataGridView1.ResetBindings();
            //this.dataGridView1.Invalidate();
        }



        private void CreateShortcutOnDesktop()
        {
            //添加引用 (com->Windows Script Host Object Model)，using IWshRuntimeLibrary;
            String shortcutPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "河南有线广告素材优化工具.lnk");
            if (!System.IO.File.Exists(shortcutPath))
            {
                // 获取当前应用程序目录地址
                String exePath = Process.GetCurrentProcess().MainModule.FileName;
                IWshShell shell = new WshShell();
                // 确定是否已经创建的快捷键被改名了
                foreach (var item in Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "*.lnk"))
                {
                    WshShortcut tempShortcut = (WshShortcut)shell.CreateShortcut(item);
                    if (tempShortcut.TargetPath == exePath)
                    {
                        return;
                    }
                }
                WshShortcut shortcut = (WshShortcut)shell.CreateShortcut(shortcutPath);
                shortcut.TargetPath = exePath;
                shortcut.Arguments = "";// 参数  
                shortcut.Description = "河南有线广告素材优化工具";
                shortcut.WorkingDirectory = Environment.CurrentDirectory;//程序所在文件夹，在快捷方式图标点击右键可以看到此属性  
                shortcut.IconLocation = exePath;//图标，该图标是应用程序的资源文件  
                //shortcut.Hotkey = "CTRL+SHIFT+W";//热键，发现没作用，大概需要注册一下  
                shortcut.WindowStyle = 1;
                shortcut.Save();
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {


        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {


            //if (e.Data.GetDataPresent(DataFormats.FileDrop))
            //{


            //    string[] picArr = (string[])e.Data.GetData(DataFormats.FileDrop);

            //    load_file(picArr);

            //    //string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            //    ////得到拖进来的路径,取第一个文件
            //    //string path = files[0];
            //    //try
            //    //{
            //    //    textBox3.Text = path;
            //    //}
            //    //catch (Exception ex)
            //    //{
            //    //    MessageBox.Show(ex.Message);
            //    //    return;
            //    //}
            //}
        }

        private void dataGridView1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = e.AllowedEffect;

                string[] picArr = (string[])e.Data.GetData(DataFormats.FileDrop);

                load_file(picArr, true);

            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            this.dataGridView1.ResetBindings();
            this.dataGridView1.DataSource = list.ToArray();

            progressBar1.Maximum = list.Count;
            progressBar1.Value = cmd_list.Count;

            this.label1.Text = cmd_list.Count.ToString() + "/" + list.Count.ToString();

        }

        private void saveDBFileFpExcel_Click(object sender, EventArgs e)
        {
            saveExcel.ShowDialog();
        }
        //文件导出
        ExportFiles exportfile = new ExportFiles();
        private void saveExcel_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                if (saveExcel.FileName.ToString().Length > 1)
                {
                    DataSet _excel = new DataSet();
                    _excel.Tables.Add(DataView2DataTable.DataGridView2DataTable(dataGridView1));
                    exportfile.DataSetToExcel(_excel, saveExcel.FileName.ToString(), "[数据源]"); //MissionName.Text +
                    MessageBox.Show("保存成功！");
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("对不起！数据尚未生成或数据结果有误！\n错误消息：" + error.Message.ToString());
            }
        }

        private void 检查更新ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            HTTPGET_UTF8.GetWebContent("");
        }



    }
}
