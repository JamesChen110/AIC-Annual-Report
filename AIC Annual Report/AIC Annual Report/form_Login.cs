using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AIC_Annual_Report
{
    public partial class form_Login : Form
    {

        private Thread thdMain;
        //public string strCurLog;
        public UploadDataMain datamain = null;

        public BuildDataMain builddatamain = null;

        public string strTelCode;

        public Boolean blRunBuiildData=false;
        //public ExcelHelper dfds = null;

        public string strCurLog;
        public string strRunType;

        public form_Login()
        {
            InitializeComponent();
            this.FormClosing += new FormClosingEventHandler(Form_Closing);
           
            datamain = new UploadDataMain();
            
        }
        private void Form_Closing(object sender, CancelEventArgs e)
        {
            WriteLineToTXT(Path.Combine(Directory.GetCurrentDirectory(), "path.txt"), textBox_FilePath.Text);
            System.Environment.Exit(0);
            //MessageBox.Show("This is the first thing I want know!");
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            //this.Visible = false;


            //datamain.Start_testing();
            //this.Close();

            this.Activate();
            textBox_FilePath.Text = strGetConfigFilePath(Path.Combine(Directory.GetCurrentDirectory(), "path.txt"));
            
            //comboBox_BuildYear.DataSource = new List<string> { DateTime.Now.AddYears(-2).ToString("yyyy"), DateTime.Now.AddYears(-1).ToString("yyyy"), DateTime.Now.AddYears(0).ToString("yyyy") };
            comboBox_BuildYear.Text = DateTime.Now.AddYears(-1).ToString("yyyy");

            //WriteLineToTXT(Path.Combine(Directory.GetCurrentDirectory(), "path.txt"),"123.txt");

        }


        //private void InitializeComponent()
        //{
        //    this.SuspendLayout();
        //    // 
        //    // Form1
        //    // 
        //    this.ClientSize = new System.Dradwing.Size(284, 261);
        //    this.Name = "Form1";
        //    this.Load += new System.EventHandler(this.Form1_Load_1);
        //    this.ResumeLayout(false);

        //}

        //多线程
        private void ThreadMainStreamingServiceOpenWeb()
        {


            datamain.OpenWeb();
        }
        private void ThreadMainStreamingServiceBuildData()
        {
            blRunBuiildData = true;
            
            builddatamain.Main_BuildData();

        }


        private void ThreadMainStreamingServiceLogin()
        {
            datamain.Start_testing(strTelCode);

        }
        //多线程


        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            button_OpenWebAndGetTelode.Enabled = true;
            btn_Login.Enabled = true;

        }

        private void text_telephoneCode_TextChanged(object sender, EventArgs e)
        {
        
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void label1_Click_2(object sender, EventArgs e)
        {

        }

        private void button_OpenWebAndGetTelCode(object sender, EventArgs e)
        {


            thdMain = new Thread(new ThreadStart(ThreadMainStreamingServiceOpenWeb));
            thdMain.IsBackground = false;
            thdMain.SetApartmentState(ApartmentState.STA);

            thdMain.Start();
           
            //this.Activate();
            button_OpenWebAndGetTelode.Enabled = false;


        }

        private void btn_Login_Click(object sender, EventArgs e)
        {


            //strTelCode = this.text_telephoneCode.Text;
            thdMain = new Thread(new ThreadStart(ThreadMainStreamingServiceLogin));
            thdMain.IsBackground = false;
            thdMain.Start();

            btn_Login.Enabled = false;


            //当前的store 已经运行，需要从待上传的store 清单去除
            datamain.lstStores.Remove(this.comboBox_SelectStore.Text);
            //datamain.lstStores = new List<string>();
            comboBox_SelectStore.DataSource = null;
            comboBox_SelectStore.DataSource = datamain.lstStores;
            comboBox_SelectStore.Text = datamain.strCurStore;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }

        private void timer_Log_Tick(object sender, EventArgs e)
        {
            if (blRunBuiildData)
            {
                this.textBox_Log.Text = builddatamain.strCurLog;
                this.label_BuildResult.Text = builddatamain.strBuildResult;
                if (this.label_BuildResult.Text == "Fail")
                {
                    this.label_BuildResult.ForeColor = System.Drawing.Color.Red;

                }
                else
                {

                    this.label_BuildResult.ForeColor = System.Drawing.Color.Green;
                }
                    
            }
            else
            {

                this.textBox_Log.Text = datamain.strCurLog;
            }
            
          
        }

        private void textBox_Log_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnLoadData_Click(object sender, EventArgs e)
        {
            btnLoadData.Enabled = false;
            datamain.strUploadDataFilePath = textBox_FilePath.Text;
            datamain.strDataType = comboBox_DataType.Text;
            datamain.LoadingData();
            comboBox_SelectStore.DataSource = datamain.lstStores;
            btnLoadData.Enabled = true;

        }

        private void btnSelectStore_Click(object sender, EventArgs e)
        {

            try
            {
                if (!datamain.lstStores.Contains(comboBox_SelectStore.Text.ToString()))
                {
                    throw new Exception("或者这个门店刚刚已运行");
                }
                datamain.strCurStore = comboBox_SelectStore.Text.ToString();
                datamain.strDataType = comboBox_DataType.Text.ToString();
                text_City.Text = datamain.dicStore[comboBox_SelectStore.Text.ToString()].strCity;
                text_Province.Text = datamain.dicStore[comboBox_SelectStore.Text.ToString()].strProvince;
                text_UserName.Text = datamain.dicStore[comboBox_SelectStore.Text.ToString()].strUserName;
                text_Password.Text = datamain.dicStore[comboBox_SelectStore.Text.ToString()].strPassword;
                textBox_id.Text = datamain.dicStore[comboBox_SelectStore.Text.ToString()].strIDNo;


                datamain.strCurCity = text_City.Text;
                datamain.strCurProvince = text_Province.Text;
                btn_Login.Enabled = true;
                button_OpenWebAndGetTelode.Enabled = true;
                //datamain.Config();
            }
            catch (Exception ex)
            {
                datamain.strCurLog = "这个门店不存在：" + ex.Message;

                text_City.Text = "";
                text_Province.Text = "";
                text_UserName.Text = "";
                text_Password.Text = "";
                textBox_id.Text ="";


            }


        }

        private void comboBox_SelectStore_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox_DataType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void textBox_FilePath_TextChanged(object sender, EventArgs e)
        {

        }
        private string strGetConfigFilePath(string in_strSettingsFilepath)
        {
            //Look for the config file
            if (System.IO.File.Exists(in_strSettingsFilepath))
            {
                String strText = System.IO.File.ReadAllText(in_strSettingsFilepath);
                if (strText == null)
                    return "";
                else
                    return strText.Trim();
            }
            else
                return "";
        }
        public void WriteLineToTXT(String txtPath,string strTxt)
        {
            System.IO.File.Delete(txtPath);
            System.IO.File.WriteAllText(txtPath, strTxt);
        }

        private void btnBuildData_Click(object sender, EventArgs e)
        {
            builddatamain = new BuildDataMain();
            builddatamain.strDataType = comboBox_builddatatype.Text;
            thdMain = new Thread(new ThreadStart(ThreadMainStreamingServiceBuildData));
            thdMain.IsBackground = false;
            thdMain.SetApartmentState(ApartmentState.STA);

            thdMain.Start();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
