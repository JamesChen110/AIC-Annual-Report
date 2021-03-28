namespace AIC_Annual_Report
{
    partial class form_Login
{
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
            this.components = new System.ComponentModel.Container();
            this.btn_Login = new System.Windows.Forms.Button();
            this.label_Store = new System.Windows.Forms.Label();
            this.comboBox_SelectStore = new System.Windows.Forms.ComboBox();
            this.button_OpenWebAndGetTelode = new System.Windows.Forms.Button();
            this.textBox_Log = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.timer_Log = new System.Windows.Forms.Timer(this.components);
            this.btnLoadData = new System.Windows.Forms.Button();
            this.btnSelectStore = new System.Windows.Forms.Button();
            this.label_UserName = new System.Windows.Forms.Label();
            this.text_UserName = new System.Windows.Forms.TextBox();
            this.label_Province = new System.Windows.Forms.Label();
            this.text_Province = new System.Windows.Forms.TextBox();
            this.text_City = new System.Windows.Forms.TextBox();
            this.label_City = new System.Windows.Forms.Label();
            this.label_password = new System.Windows.Forms.Label();
            this.text_Password = new System.Windows.Forms.TextBox();
            this.comboBox_DataType = new System.Windows.Forms.ComboBox();
            this.label_datatype = new System.Windows.Forms.Label();
            this.textBox_FilePath = new System.Windows.Forms.TextBox();
            this.label_FilePath = new System.Windows.Forms.Label();
            this.label_id = new System.Windows.Forms.Label();
            this.textBox_id = new System.Windows.Forms.TextBox();
            this.btnBuildData = new System.Windows.Forms.Button();
            this.comboBox_builddatatype = new System.Windows.Forms.ComboBox();
            this.tabControl_BuildData = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label_BuildDataType = new System.Windows.Forms.Label();
            this.label_AICYear = new System.Windows.Forms.Label();
            this.comboBox_BuildYear = new System.Windows.Forms.ComboBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label_BuildResult = new System.Windows.Forms.Label();
            this.tabControl_BuildData.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_Login
            // 
            this.btn_Login.Location = new System.Drawing.Point(235, 402);
            this.btn_Login.Name = "btn_Login";
            this.btn_Login.Size = new System.Drawing.Size(100, 29);
            this.btn_Login.TabIndex = 3;
            this.btn_Login.Text = "主页面数据填报";
            this.btn_Login.UseVisualStyleBackColor = false;
            this.btn_Login.UseWaitCursor = true;
            this.btn_Login.Click += new System.EventHandler(this.btn_Login_Click);
            // 
            // label_Store
            // 
            this.label_Store.AutoSize = true;
            this.label_Store.Location = new System.Drawing.Point(130, 108);
            this.label_Store.Name = "label_Store";
            this.label_Store.Size = new System.Drawing.Size(43, 13);
            this.label_Store.TabIndex = 8;
            this.label_Store.Text = "门店：";
            // 
            // comboBox_SelectStore
            // 
            this.comboBox_SelectStore.FormattingEnabled = true;
            this.comboBox_SelectStore.Location = new System.Drawing.Point(235, 105);
            this.comboBox_SelectStore.Name = "comboBox_SelectStore";
            this.comboBox_SelectStore.Size = new System.Drawing.Size(100, 21);
            this.comboBox_SelectStore.TabIndex = 9;
            this.comboBox_SelectStore.SelectedIndexChanged += new System.EventHandler(this.comboBox_SelectStore_SelectedIndexChanged);
            // 
            // button_OpenWebAndGetTelode
            // 
            this.button_OpenWebAndGetTelode.Location = new System.Drawing.Point(235, 349);
            this.button_OpenWebAndGetTelode.Name = "button_OpenWebAndGetTelode";
            this.button_OpenWebAndGetTelode.Size = new System.Drawing.Size(100, 28);
            this.button_OpenWebAndGetTelode.TabIndex = 11;
            this.button_OpenWebAndGetTelode.Text = "打开网站";
            this.button_OpenWebAndGetTelode.UseVisualStyleBackColor = true;
            this.button_OpenWebAndGetTelode.Click += new System.EventHandler(this.button_OpenWebAndGetTelCode);
            // 
            // textBox_Log
            // 
            this.textBox_Log.BackColor = System.Drawing.SystemColors.Window;
            this.textBox_Log.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox_Log.Location = new System.Drawing.Point(2, 506);
            this.textBox_Log.Multiline = true;
            this.textBox_Log.Name = "textBox_Log";
            this.textBox_Log.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox_Log.Size = new System.Drawing.Size(613, 175);
            this.textBox_Log.TabIndex = 12;
            this.textBox_Log.TextChanged += new System.EventHandler(this.textBox_Log_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(-1, 490);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 13);
            this.label1.TabIndex = 13;
            this.label1.Text = "当前状态：";
            this.label1.Click += new System.EventHandler(this.label1_Click_2);
            // 
            // timer_Log
            // 
            this.timer_Log.Enabled = true;
            this.timer_Log.Interval = 1000;
            this.timer_Log.Tick += new System.EventHandler(this.timer_Log_Tick);
            // 
            // btnLoadData
            // 
            this.btnLoadData.Location = new System.Drawing.Point(379, 30);
            this.btnLoadData.Name = "btnLoadData";
            this.btnLoadData.Size = new System.Drawing.Size(94, 22);
            this.btnLoadData.TabIndex = 14;
            this.btnLoadData.Text = "加载数据源...";
            this.btnLoadData.UseVisualStyleBackColor = true;
            this.btnLoadData.Click += new System.EventHandler(this.btnLoadData_Click);
            // 
            // btnSelectStore
            // 
            this.btnSelectStore.Location = new System.Drawing.Point(379, 105);
            this.btnSelectStore.Name = "btnSelectStore";
            this.btnSelectStore.Size = new System.Drawing.Size(55, 23);
            this.btnSelectStore.TabIndex = 15;
            this.btnSelectStore.Text = "确定";
            this.btnSelectStore.UseVisualStyleBackColor = true;
            this.btnSelectStore.Click += new System.EventHandler(this.btnSelectStore_Click);
            // 
            // label_UserName
            // 
            this.label_UserName.AutoSize = true;
            this.label_UserName.Location = new System.Drawing.Point(130, 150);
            this.label_UserName.Name = "label_UserName";
            this.label_UserName.Size = new System.Drawing.Size(55, 13);
            this.label_UserName.TabIndex = 17;
            this.label_UserName.Text = "用户名：";
            // 
            // text_UserName
            // 
            this.text_UserName.Location = new System.Drawing.Point(235, 150);
            this.text_UserName.Name = "text_UserName";
            this.text_UserName.Size = new System.Drawing.Size(238, 20);
            this.text_UserName.TabIndex = 18;
            // 
            // label_Province
            // 
            this.label_Province.AutoSize = true;
            this.label_Province.Location = new System.Drawing.Point(130, 271);
            this.label_Province.Name = "label_Province";
            this.label_Province.Size = new System.Drawing.Size(43, 13);
            this.label_Province.TabIndex = 19;
            this.label_Province.Text = "省份：";
            // 
            // text_Province
            // 
            this.text_Province.Location = new System.Drawing.Point(235, 269);
            this.text_Province.Name = "text_Province";
            this.text_Province.Size = new System.Drawing.Size(114, 20);
            this.text_Province.TabIndex = 20;
            // 
            // text_City
            // 
            this.text_City.Location = new System.Drawing.Point(235, 309);
            this.text_City.Name = "text_City";
            this.text_City.Size = new System.Drawing.Size(114, 20);
            this.text_City.TabIndex = 21;
            // 
            // label_City
            // 
            this.label_City.AutoSize = true;
            this.label_City.Location = new System.Drawing.Point(130, 309);
            this.label_City.Name = "label_City";
            this.label_City.Size = new System.Drawing.Size(43, 13);
            this.label_City.TabIndex = 22;
            this.label_City.Text = "城市：";
            // 
            // label_password
            // 
            this.label_password.AutoSize = true;
            this.label_password.Location = new System.Drawing.Point(130, 190);
            this.label_password.Name = "label_password";
            this.label_password.Size = new System.Drawing.Size(43, 13);
            this.label_password.TabIndex = 23;
            this.label_password.Text = "密码：";
            // 
            // text_Password
            // 
            this.text_Password.Location = new System.Drawing.Point(235, 190);
            this.text_Password.Name = "text_Password";
            this.text_Password.Size = new System.Drawing.Size(238, 20);
            this.text_Password.TabIndex = 24;
            // 
            // comboBox_DataType
            // 
            this.comboBox_DataType.FormattingEnabled = true;
            this.comboBox_DataType.Items.AddRange(new object[] {
            "法人",
            "分支机构"});
            this.comboBox_DataType.Location = new System.Drawing.Point(235, 30);
            this.comboBox_DataType.Name = "comboBox_DataType";
            this.comboBox_DataType.Size = new System.Drawing.Size(100, 21);
            this.comboBox_DataType.TabIndex = 25;
            this.comboBox_DataType.Text = "法人";
            this.comboBox_DataType.SelectedIndexChanged += new System.EventHandler(this.comboBox_DataType_SelectedIndexChanged);
            // 
            // label_datatype
            // 
            this.label_datatype.AutoSize = true;
            this.label_datatype.Location = new System.Drawing.Point(130, 33);
            this.label_datatype.Name = "label_datatype";
            this.label_datatype.Size = new System.Drawing.Size(96, 13);
            this.label_datatype.TabIndex = 26;
            this.label_datatype.Text = "法人/分支机构：";
            // 
            // textBox_FilePath
            // 
            this.textBox_FilePath.Location = new System.Drawing.Point(235, 70);
            this.textBox_FilePath.Name = "textBox_FilePath";
            this.textBox_FilePath.Size = new System.Drawing.Size(284, 20);
            this.textBox_FilePath.TabIndex = 27;
            this.textBox_FilePath.Text = "C:\\U1sers\\c0c04nc\\Documents\\Project\\合规年报\\AIC Annual Report Build Data\\result\\分支机构" +
    " UploadCorporateData.xlsx";
            this.textBox_FilePath.TextChanged += new System.EventHandler(this.textBox_FilePath_TextChanged);
            // 
            // label_FilePath
            // 
            this.label_FilePath.AutoSize = true;
            this.label_FilePath.Location = new System.Drawing.Point(133, 70);
            this.label_FilePath.Name = "label_FilePath";
            this.label_FilePath.Size = new System.Drawing.Size(79, 13);
            this.label_FilePath.TabIndex = 28;
            this.label_FilePath.Text = "数据源路径：";
            // 
            // label_id
            // 
            this.label_id.AutoSize = true;
            this.label_id.Location = new System.Drawing.Point(130, 229);
            this.label_id.Name = "label_id";
            this.label_id.Size = new System.Drawing.Size(79, 13);
            this.label_id.TabIndex = 29;
            this.label_id.Text = "身份证号码：";
            // 
            // textBox_id
            // 
            this.textBox_id.Location = new System.Drawing.Point(235, 229);
            this.textBox_id.Name = "textBox_id";
            this.textBox_id.Size = new System.Drawing.Size(238, 20);
            this.textBox_id.TabIndex = 30;
            // 
            // btnBuildData
            // 
            this.btnBuildData.Location = new System.Drawing.Point(228, 119);
            this.btnBuildData.Name = "btnBuildData";
            this.btnBuildData.Size = new System.Drawing.Size(75, 23);
            this.btnBuildData.TabIndex = 31;
            this.btnBuildData.Text = "生成数据";
            this.btnBuildData.UseVisualStyleBackColor = true;
            this.btnBuildData.Click += new System.EventHandler(this.btnBuildData_Click);
            // 
            // comboBox_builddatatype
            // 
            this.comboBox_builddatatype.FormattingEnabled = true;
            this.comboBox_builddatatype.Items.AddRange(new object[] {
            "法人",
            "分支机构",
            "全部"});
            this.comboBox_builddatatype.Location = new System.Drawing.Point(228, 64);
            this.comboBox_builddatatype.Name = "comboBox_builddatatype";
            this.comboBox_builddatatype.Size = new System.Drawing.Size(85, 21);
            this.comboBox_builddatatype.TabIndex = 32;
            this.comboBox_builddatatype.Text = "法人";
            // 
            // tabControl_BuildData
            // 
            this.tabControl_BuildData.Controls.Add(this.tabPage1);
            this.tabControl_BuildData.Controls.Add(this.tabPage2);
            this.tabControl_BuildData.Location = new System.Drawing.Point(-2, 1);
            this.tabControl_BuildData.Name = "tabControl_BuildData";
            this.tabControl_BuildData.SelectedIndex = 0;
            this.tabControl_BuildData.Size = new System.Drawing.Size(617, 486);
            this.tabControl_BuildData.TabIndex = 33;
            this.tabControl_BuildData.Tag = "生成数据";
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.Transparent;
            this.tabPage1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tabPage1.Controls.Add(this.label_BuildResult);
            this.tabPage1.Controls.Add(this.label_BuildDataType);
            this.tabPage1.Controls.Add(this.label_AICYear);
            this.tabPage1.Controls.Add(this.comboBox_BuildYear);
            this.tabPage1.Controls.Add(this.comboBox_builddatatype);
            this.tabPage1.Controls.Add(this.btnBuildData);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(609, 460);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "生成数据";
            this.tabPage1.Click += new System.EventHandler(this.tabPage1_Click);
            // 
            // label_BuildDataType
            // 
            this.label_BuildDataType.AutoSize = true;
            this.label_BuildDataType.Location = new System.Drawing.Point(145, 67);
            this.label_BuildDataType.Name = "label_BuildDataType";
            this.label_BuildDataType.Size = new System.Drawing.Size(60, 13);
            this.label_BuildDataType.TabIndex = 35;
            this.label_BuildDataType.Text = "分支/法人";
            // 
            // label_AICYear
            // 
            this.label_AICYear.AutoSize = true;
            this.label_AICYear.Location = new System.Drawing.Point(145, 23);
            this.label_AICYear.Name = "label_AICYear";
            this.label_AICYear.Size = new System.Drawing.Size(31, 13);
            this.label_AICYear.TabIndex = 34;
            this.label_AICYear.Text = "年份";
            // 
            // comboBox_BuildYear
            // 
            this.comboBox_BuildYear.FormattingEnabled = true;
            this.comboBox_BuildYear.Location = new System.Drawing.Point(228, 20);
            this.comboBox_BuildYear.Name = "comboBox_BuildYear";
            this.comboBox_BuildYear.Size = new System.Drawing.Size(85, 21);
            this.comboBox_BuildYear.TabIndex = 33;
            this.comboBox_BuildYear.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged_1);
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.Transparent;
            this.tabPage2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tabPage2.Controls.Add(this.comboBox_DataType);
            this.tabPage2.Controls.Add(this.textBox_id);
            this.tabPage2.Controls.Add(this.btn_Login);
            this.tabPage2.Controls.Add(this.label_id);
            this.tabPage2.Controls.Add(this.label_Store);
            this.tabPage2.Controls.Add(this.label_FilePath);
            this.tabPage2.Controls.Add(this.comboBox_SelectStore);
            this.tabPage2.Controls.Add(this.textBox_FilePath);
            this.tabPage2.Controls.Add(this.button_OpenWebAndGetTelode);
            this.tabPage2.Controls.Add(this.label_datatype);
            this.tabPage2.Controls.Add(this.btnLoadData);
            this.tabPage2.Controls.Add(this.text_Password);
            this.tabPage2.Controls.Add(this.btnSelectStore);
            this.tabPage2.Controls.Add(this.label_password);
            this.tabPage2.Controls.Add(this.label_UserName);
            this.tabPage2.Controls.Add(this.label_City);
            this.tabPage2.Controls.Add(this.text_UserName);
            this.tabPage2.Controls.Add(this.text_City);
            this.tabPage2.Controls.Add(this.label_Province);
            this.tabPage2.Controls.Add(this.text_Province);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(609, 460);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "上报数据";
            // 
            // label_BuildResult
            // 
            this.label_BuildResult.AutoSize = true;
            this.label_BuildResult.Font = new System.Drawing.Font("Microsoft Sans Serif", 40F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label_BuildResult.Location = new System.Drawing.Point(205, 288);
            this.label_BuildResult.Name = "label_BuildResult";
            this.label_BuildResult.Size = new System.Drawing.Size(0, 63);
            this.label_BuildResult.TabIndex = 37;
            this.label_BuildResult.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.label_BuildResult.Click += new System.EventHandler(this.label3_Click);
            // 
            // form_Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(617, 701);
            this.Controls.Add(this.tabControl_BuildData);
            this.Controls.Add(this.textBox_Log);
            this.Controls.Add(this.label1);
            this.ForeColor = System.Drawing.Color.Black;
            this.Name = "form_Login";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Tag = "登录页面";
            this.Text = "AIC年报工具";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl_BuildData.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

    }

        #endregion
        private System.Windows.Forms.Button btn_Login;
        private System.Windows.Forms.Label label_Store;
        private System.Windows.Forms.ComboBox comboBox_SelectStore;
        private System.Windows.Forms.Button button_OpenWebAndGetTelode;
        private System.Windows.Forms.TextBox textBox_Log;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Timer timer_Log;
        private System.Windows.Forms.Button btnLoadData;
        private System.Windows.Forms.Button btnSelectStore;
        private System.Windows.Forms.Label label_UserName;
        private System.Windows.Forms.TextBox text_UserName;
        private System.Windows.Forms.Label label_Province;
        private System.Windows.Forms.TextBox text_Province;
        private System.Windows.Forms.TextBox text_City;
        private System.Windows.Forms.Label label_City;
        private System.Windows.Forms.Label label_password;
        private System.Windows.Forms.TextBox text_Password;
        private System.Windows.Forms.ComboBox comboBox_DataType;
        private System.Windows.Forms.Label label_datatype;
        private System.Windows.Forms.TextBox textBox_FilePath;
        private System.Windows.Forms.Label label_FilePath;
        private System.Windows.Forms.Label label_id;
        private System.Windows.Forms.TextBox textBox_id;
        private System.Windows.Forms.Button btnBuildData;
        private System.Windows.Forms.ComboBox comboBox_builddatatype;
        private System.Windows.Forms.TabControl tabControl_BuildData;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label_BuildDataType;
        private System.Windows.Forms.Label label_AICYear;
        private System.Windows.Forms.ComboBox comboBox_BuildYear;
        private System.Windows.Forms.Label label_BuildResult;
    }
}

