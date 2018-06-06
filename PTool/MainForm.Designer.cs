namespace PTool
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.lbPumpPort = new CCWin.SkinControl.SkinLabel();
            this.cbToolingPort = new CCWin.SkinControl.SkinComboBox();
            this.cbPumpPort = new CCWin.SkinControl.SkinComboBox();
            this.lbToolingPort = new CCWin.SkinControl.SkinLabel();
            this.btnStopPumpChannel1 = new CCWin.SkinControl.SkinButton();
            this.btnStartPumpChannel1 = new CCWin.SkinControl.SkinButton();
            this.skinLabel3 = new CCWin.SkinControl.SkinLabel();
            this.lbRate = new CCWin.SkinControl.SkinLabel();
            this.cbPumpType = new CCWin.SkinControl.SkinComboBox();
            this.skinLabel2 = new CCWin.SkinControl.SkinLabel();
            this.WavelinePanel = new System.Windows.Forms.Panel();
            this.skinGroupBox1 = new CCWin.SkinControl.SkinGroupBox();
            this.tbRateChannel1 = new System.Windows.Forms.TextBox();
            this.lbPChannel1 = new CCWin.SkinControl.SkinLabel();
            this.skinLabel13 = new CCWin.SkinControl.SkinLabel();
            this.lbWeightChannel1 = new CCWin.SkinControl.SkinLabel();
            this.skinLabel10 = new CCWin.SkinControl.SkinLabel();
            this.skinLabel9 = new CCWin.SkinControl.SkinLabel();
            this.gbParameter = new CCWin.SkinControl.SkinGroupBox();
            this.skinLabel7 = new CCWin.SkinControl.SkinLabel();
            this.skinLabel8 = new CCWin.SkinControl.SkinLabel();
            this.tbToolingNo = new CCWin.SkinControl.SkinTextBox();
            this.tbPumpNo = new CCWin.SkinControl.SkinTextBox();
            this.skinMenuStrip1 = new CCWin.SkinControl.SkinMenuStrip();
            this.文件ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.操作ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuInitTooling = new System.Windows.Forms.ToolStripMenuItem();
            this.menuTareTooling = new System.Windows.Forms.ToolStripMenuItem();
            this.gbChannel2 = new CCWin.SkinControl.SkinGroupBox();
            this.cbPumpPort2 = new CCWin.SkinControl.SkinComboBox();
            this.tbRateChannel2 = new System.Windows.Forms.TextBox();
            this.btnStopPumpChannel2 = new CCWin.SkinControl.SkinButton();
            this.WavelinePanel2 = new System.Windows.Forms.Panel();
            this.btnStartPumpChannel2 = new CCWin.SkinControl.SkinButton();
            this.cbToolingPort2 = new CCWin.SkinControl.SkinComboBox();
            this.lbPChannel2 = new CCWin.SkinControl.SkinLabel();
            this.skinLabel4 = new CCWin.SkinControl.SkinLabel();
            this.lbWeightChannel2 = new CCWin.SkinControl.SkinLabel();
            this.skinLabel6 = new CCWin.SkinControl.SkinLabel();
            this.skinLabel11 = new CCWin.SkinControl.SkinLabel();
            this.skinLabel12 = new CCWin.SkinControl.SkinLabel();
            this.skinLabel15 = new CCWin.SkinControl.SkinLabel();
            this.skinLabel16 = new CCWin.SkinControl.SkinLabel();
            this.skinLabel17 = new CCWin.SkinControl.SkinLabel();
            this.skinGroupBox1.SuspendLayout();
            this.gbParameter.SuspendLayout();
            this.skinMenuStrip1.SuspendLayout();
            this.gbChannel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbPumpPort
            // 
            this.lbPumpPort.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.lbPumpPort.AutoSize = true;
            this.lbPumpPort.BackColor = System.Drawing.Color.Transparent;
            this.lbPumpPort.BorderColor = System.Drawing.Color.White;
            this.lbPumpPort.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.lbPumpPort.Location = new System.Drawing.Point(210, 59);
            this.lbPumpPort.Name = "lbPumpPort";
            this.lbPumpPort.Size = new System.Drawing.Size(51, 20);
            this.lbPumpPort.TabIndex = 0;
            this.lbPumpPort.Text = "泵串口";
            // 
            // cbToolingPort
            // 
            this.cbToolingPort.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbToolingPort.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.cbToolingPort.FormattingEnabled = true;
            this.cbToolingPort.Location = new System.Drawing.Point(98, 54);
            this.cbToolingPort.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.cbToolingPort.Name = "cbToolingPort";
            this.cbToolingPort.Size = new System.Drawing.Size(106, 26);
            this.cbToolingPort.TabIndex = 1;
            this.cbToolingPort.WaterText = "";
            this.cbToolingPort.SelectedIndexChanged += new System.EventHandler(this.cbToolingPort_SelectedIndexChanged);
            // 
            // cbPumpPort
            // 
            this.cbPumpPort.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbPumpPort.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.cbPumpPort.FormattingEnabled = true;
            this.cbPumpPort.Location = new System.Drawing.Point(280, 54);
            this.cbPumpPort.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.cbPumpPort.Name = "cbPumpPort";
            this.cbPumpPort.Size = new System.Drawing.Size(106, 26);
            this.cbPumpPort.TabIndex = 1;
            this.cbPumpPort.WaterText = "";
            this.cbPumpPort.SelectedIndexChanged += new System.EventHandler(this.cbPumpPort_SelectedIndexChanged);
            // 
            // lbToolingPort
            // 
            this.lbToolingPort.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.lbToolingPort.AutoSize = true;
            this.lbToolingPort.BackColor = System.Drawing.Color.Transparent;
            this.lbToolingPort.BorderColor = System.Drawing.Color.White;
            this.lbToolingPort.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.lbToolingPort.Location = new System.Drawing.Point(11, 59);
            this.lbToolingPort.Name = "lbToolingPort";
            this.lbToolingPort.Size = new System.Drawing.Size(65, 20);
            this.lbToolingPort.TabIndex = 0;
            this.lbToolingPort.Text = "工装串口";
            // 
            // btnStopPumpChannel1
            // 
            this.btnStopPumpChannel1.BackColor = System.Drawing.Color.Transparent;
            this.btnStopPumpChannel1.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.btnStopPumpChannel1.DownBack = null;
            this.btnStopPumpChannel1.Enabled = false;
            this.btnStopPumpChannel1.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.btnStopPumpChannel1.Location = new System.Drawing.Point(566, 119);
            this.btnStopPumpChannel1.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.btnStopPumpChannel1.MouseBack = null;
            this.btnStopPumpChannel1.Name = "btnStopPumpChannel1";
            this.btnStopPumpChannel1.NormlBack = null;
            this.btnStopPumpChannel1.Size = new System.Drawing.Size(66, 45);
            this.btnStopPumpChannel1.TabIndex = 5;
            this.btnStopPumpChannel1.Text = "停止";
            this.btnStopPumpChannel1.UseVisualStyleBackColor = false;
            this.btnStopPumpChannel1.Click += new System.EventHandler(this.btnStopPumpChannel1_Click);
            // 
            // btnStartPumpChannel1
            // 
            this.btnStartPumpChannel1.BackColor = System.Drawing.Color.Transparent;
            this.btnStartPumpChannel1.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.btnStartPumpChannel1.DownBack = null;
            this.btnStartPumpChannel1.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.btnStartPumpChannel1.Location = new System.Drawing.Point(490, 119);
            this.btnStartPumpChannel1.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.btnStartPumpChannel1.MouseBack = null;
            this.btnStartPumpChannel1.Name = "btnStartPumpChannel1";
            this.btnStartPumpChannel1.NormlBack = null;
            this.btnStartPumpChannel1.Size = new System.Drawing.Size(66, 45);
            this.btnStartPumpChannel1.TabIndex = 4;
            this.btnStartPumpChannel1.Text = "开始";
            this.btnStartPumpChannel1.UseVisualStyleBackColor = false;
            this.btnStartPumpChannel1.Click += new System.EventHandler(this.btnStartPumpChannel1_Click);
            // 
            // skinLabel3
            // 
            this.skinLabel3.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel3.AutoSize = true;
            this.skinLabel3.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel3.BorderColor = System.Drawing.Color.White;
            this.skinLabel3.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel3.Location = new System.Drawing.Point(574, 56);
            this.skinLabel3.Name = "skinLabel3";
            this.skinLabel3.Size = new System.Drawing.Size(44, 20);
            this.skinLabel3.TabIndex = 0;
            this.skinLabel3.Text = "mL\\h";
            // 
            // lbRate
            // 
            this.lbRate.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.lbRate.AutoSize = true;
            this.lbRate.BackColor = System.Drawing.Color.Transparent;
            this.lbRate.BorderColor = System.Drawing.Color.White;
            this.lbRate.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.lbRate.Location = new System.Drawing.Point(392, 58);
            this.lbRate.Name = "lbRate";
            this.lbRate.Size = new System.Drawing.Size(45, 20);
            this.lbRate.TabIndex = 0;
            this.lbRate.Text = "速  率";
            // 
            // cbPumpType
            // 
            this.cbPumpType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbPumpType.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.cbPumpType.FormattingEnabled = true;
            this.cbPumpType.Location = new System.Drawing.Point(215, 40);
            this.cbPumpType.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.cbPumpType.Name = "cbPumpType";
            this.cbPumpType.Size = new System.Drawing.Size(131, 26);
            this.cbPumpType.TabIndex = 1;
            this.cbPumpType.WaterText = "";
            this.cbPumpType.SelectedIndexChanged += new System.EventHandler(this.cbPumpType_SelectedIndexChanged);
            // 
            // skinLabel2
            // 
            this.skinLabel2.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel2.AutoSize = true;
            this.skinLabel2.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel2.BorderColor = System.Drawing.Color.White;
            this.skinLabel2.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel2.Location = new System.Drawing.Point(392, 44);
            this.skinLabel2.Name = "skinLabel2";
            this.skinLabel2.Size = new System.Drawing.Size(65, 20);
            this.skinLabel2.TabIndex = 0;
            this.skinLabel2.Text = "产品序号";
            // 
            // WavelinePanel
            // 
            this.WavelinePanel.BackColor = System.Drawing.Color.White;
            this.WavelinePanel.Font = new System.Drawing.Font("Microsoft YaHei", 9F);
            this.WavelinePanel.ForeColor = System.Drawing.Color.Black;
            this.WavelinePanel.Location = new System.Drawing.Point(3, 173);
            this.WavelinePanel.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.WavelinePanel.Name = "WavelinePanel";
            this.WavelinePanel.Size = new System.Drawing.Size(633, 710);
            this.WavelinePanel.TabIndex = 4;
            this.WavelinePanel.Paint += new System.Windows.Forms.PaintEventHandler(this.WavelinePanel_Paint);
            // 
            // skinGroupBox1
            // 
            this.skinGroupBox1.BackColor = System.Drawing.Color.Transparent;
            this.skinGroupBox1.BorderColor = System.Drawing.Color.White;
            this.skinGroupBox1.Controls.Add(this.WavelinePanel);
            this.skinGroupBox1.Controls.Add(this.tbRateChannel1);
            this.skinGroupBox1.Controls.Add(this.cbPumpPort);
            this.skinGroupBox1.Controls.Add(this.btnStopPumpChannel1);
            this.skinGroupBox1.Controls.Add(this.btnStartPumpChannel1);
            this.skinGroupBox1.Controls.Add(this.cbToolingPort);
            this.skinGroupBox1.Controls.Add(this.lbPChannel1);
            this.skinGroupBox1.Controls.Add(this.skinLabel13);
            this.skinGroupBox1.Controls.Add(this.lbWeightChannel1);
            this.skinGroupBox1.Controls.Add(this.skinLabel10);
            this.skinGroupBox1.Controls.Add(this.lbRate);
            this.skinGroupBox1.Controls.Add(this.skinLabel9);
            this.skinGroupBox1.Controls.Add(this.lbPumpPort);
            this.skinGroupBox1.Controls.Add(this.skinLabel3);
            this.skinGroupBox1.Controls.Add(this.lbToolingPort);
            this.skinGroupBox1.Font = new System.Drawing.Font("Microsoft YaHei", 14F);
            this.skinGroupBox1.ForeColor = System.Drawing.Color.Black;
            this.skinGroupBox1.Location = new System.Drawing.Point(9, 212);
            this.skinGroupBox1.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.skinGroupBox1.Name = "skinGroupBox1";
            this.skinGroupBox1.Padding = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.skinGroupBox1.RectBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.skinGroupBox1.RoundStyle = CCWin.SkinClass.RoundStyle.All;
            this.skinGroupBox1.Size = new System.Drawing.Size(640, 885);
            this.skinGroupBox1.TabIndex = 2;
            this.skinGroupBox1.TabStop = false;
            this.skinGroupBox1.Text = "1道泵";
            this.skinGroupBox1.TitleBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.skinGroupBox1.TitleRectBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.skinGroupBox1.TitleRoundStyle = CCWin.SkinClass.RoundStyle.All;
            // 
            // tbRateChannel1
            // 
            this.tbRateChannel1.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.tbRateChannel1.Location = new System.Drawing.Point(458, 52);
            this.tbRateChannel1.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.tbRateChannel1.Name = "tbRateChannel1";
            this.tbRateChannel1.Size = new System.Drawing.Size(106, 25);
            this.tbRateChannel1.TabIndex = 4;
            this.tbRateChannel1.Text = "50";
            this.tbRateChannel1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.OnRateKeyPress);
            // 
            // lbPChannel1
            // 
            this.lbPChannel1.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.lbPChannel1.AutoSize = true;
            this.lbPChannel1.BackColor = System.Drawing.Color.Transparent;
            this.lbPChannel1.BorderColor = System.Drawing.Color.White;
            this.lbPChannel1.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.lbPChannel1.Location = new System.Drawing.Point(303, 122);
            this.lbPChannel1.Name = "lbPChannel1";
            this.lbPChannel1.Size = new System.Drawing.Size(39, 20);
            this.lbPChannel1.TabIndex = 0;
            this.lbPChannel1.Text = "-----";
            // 
            // skinLabel13
            // 
            this.skinLabel13.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel13.AutoSize = true;
            this.skinLabel13.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel13.BorderColor = System.Drawing.Color.White;
            this.skinLabel13.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel13.Location = new System.Drawing.Point(223, 122);
            this.skinLabel13.Name = "skinLabel13";
            this.skinLabel13.Size = new System.Drawing.Size(60, 20);
            this.skinLabel13.TabIndex = 0;
            this.skinLabel13.Text = "机器P值";
            // 
            // lbWeightChannel1
            // 
            this.lbWeightChannel1.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.lbWeightChannel1.AutoSize = true;
            this.lbWeightChannel1.BackColor = System.Drawing.Color.Transparent;
            this.lbWeightChannel1.BorderColor = System.Drawing.Color.White;
            this.lbWeightChannel1.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.lbWeightChannel1.Location = new System.Drawing.Point(96, 121);
            this.lbWeightChannel1.Name = "lbWeightChannel1";
            this.lbWeightChannel1.Size = new System.Drawing.Size(39, 20);
            this.lbWeightChannel1.TabIndex = 0;
            this.lbWeightChannel1.Text = "-----";
            // 
            // skinLabel10
            // 
            this.skinLabel10.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel10.AutoSize = true;
            this.skinLabel10.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel10.BorderColor = System.Drawing.Color.White;
            this.skinLabel10.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel10.Location = new System.Drawing.Point(14, 121);
            this.skinLabel10.Name = "skinLabel10";
            this.skinLabel10.Size = new System.Drawing.Size(65, 20);
            this.skinLabel10.TabIndex = 0;
            this.skinLabel10.Text = "工装读数";
            // 
            // skinLabel9
            // 
            this.skinLabel9.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel9.AutoSize = true;
            this.skinLabel9.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel9.BorderColor = System.Drawing.Color.White;
            this.skinLabel9.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel9.Location = new System.Drawing.Point(176, 121);
            this.skinLabel9.Name = "skinLabel9";
            this.skinLabel9.Size = new System.Drawing.Size(26, 20);
            this.skinLabel9.TabIndex = 0;
            this.skinLabel9.Text = "kg";
            // 
            // gbParameter
            // 
            this.gbParameter.BackColor = System.Drawing.Color.Transparent;
            this.gbParameter.BorderColor = System.Drawing.Color.White;
            this.gbParameter.Controls.Add(this.skinLabel7);
            this.gbParameter.Controls.Add(this.skinLabel8);
            this.gbParameter.Controls.Add(this.skinLabel2);
            this.gbParameter.Controls.Add(this.cbPumpType);
            this.gbParameter.Controls.Add(this.tbToolingNo);
            this.gbParameter.Controls.Add(this.tbPumpNo);
            this.gbParameter.Font = new System.Drawing.Font("Microsoft YaHei", 12F);
            this.gbParameter.ForeColor = System.Drawing.Color.Black;
            this.gbParameter.Location = new System.Drawing.Point(9, 110);
            this.gbParameter.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.gbParameter.Name = "gbParameter";
            this.gbParameter.Padding = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.gbParameter.RectBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.gbParameter.RoundStyle = CCWin.SkinClass.RoundStyle.All;
            this.gbParameter.Size = new System.Drawing.Size(1288, 98);
            this.gbParameter.TabIndex = 2;
            this.gbParameter.TabStop = false;
            this.gbParameter.Text = "泵参数设置";
            this.gbParameter.TitleBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.gbParameter.TitleRectBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.gbParameter.TitleRoundStyle = CCWin.SkinClass.RoundStyle.All;
            // 
            // skinLabel7
            // 
            this.skinLabel7.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel7.AutoSize = true;
            this.skinLabel7.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel7.BorderColor = System.Drawing.Color.White;
            this.skinLabel7.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel7.Location = new System.Drawing.Point(782, 44);
            this.skinLabel7.Name = "skinLabel7";
            this.skinLabel7.Size = new System.Drawing.Size(65, 20);
            this.skinLabel7.TabIndex = 0;
            this.skinLabel7.Text = "工装编号";
            // 
            // skinLabel8
            // 
            this.skinLabel8.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel8.AutoSize = true;
            this.skinLabel8.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel8.BorderColor = System.Drawing.Color.White;
            this.skinLabel8.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel8.Location = new System.Drawing.Point(120, 45);
            this.skinLabel8.Name = "skinLabel8";
            this.skinLabel8.Size = new System.Drawing.Size(65, 20);
            this.skinLabel8.TabIndex = 0;
            this.skinLabel8.Text = "机器型号";
            // 
            // tbToolingNo
            // 
            this.tbToolingNo.BackColor = System.Drawing.Color.Transparent;
            this.tbToolingNo.DownBack = null;
            this.tbToolingNo.Icon = null;
            this.tbToolingNo.IconIsButton = false;
            this.tbToolingNo.IconMouseState = CCWin.SkinClass.ControlState.Normal;
            this.tbToolingNo.IsPasswordChat = '\0';
            this.tbToolingNo.IsSystemPasswordChar = false;
            this.tbToolingNo.Lines = new string[0];
            this.tbToolingNo.Location = new System.Drawing.Point(872, 38);
            this.tbToolingNo.Margin = new System.Windows.Forms.Padding(0);
            this.tbToolingNo.MaxLength = 32767;
            this.tbToolingNo.MinimumSize = new System.Drawing.Size(38, 49);
            this.tbToolingNo.MouseBack = null;
            this.tbToolingNo.MouseState = CCWin.SkinClass.ControlState.Normal;
            this.tbToolingNo.Multiline = true;
            this.tbToolingNo.Name = "tbToolingNo";
            this.tbToolingNo.NormlBack = null;
            this.tbToolingNo.Padding = new System.Windows.Forms.Padding(7, 9, 7, 9);
            this.tbToolingNo.ReadOnly = false;
            this.tbToolingNo.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.tbToolingNo.Size = new System.Drawing.Size(274, 49);
            // 
            // 
            // 
            this.tbToolingNo.SkinTxt.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbToolingNo.SkinTxt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbToolingNo.SkinTxt.Font = new System.Drawing.Font("Microsoft YaHei", 9F);
            this.tbToolingNo.SkinTxt.Location = new System.Drawing.Point(7, 9);
            this.tbToolingNo.SkinTxt.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.tbToolingNo.SkinTxt.Multiline = true;
            this.tbToolingNo.SkinTxt.Name = "BaseText";
            this.tbToolingNo.SkinTxt.Size = new System.Drawing.Size(262, 31);
            this.tbToolingNo.SkinTxt.TabIndex = 0;
            this.tbToolingNo.SkinTxt.WaterColor = System.Drawing.Color.FromArgb(((int)(((byte)(127)))), ((int)(((byte)(127)))), ((int)(((byte)(127)))));
            this.tbToolingNo.SkinTxt.WaterText = "";
            this.tbToolingNo.TabIndex = 2;
            this.tbToolingNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.tbToolingNo.WaterColor = System.Drawing.Color.FromArgb(((int)(((byte)(127)))), ((int)(((byte)(127)))), ((int)(((byte)(127)))));
            this.tbToolingNo.WaterText = "";
            this.tbToolingNo.WordWrap = true;
            // 
            // tbPumpNo
            // 
            this.tbPumpNo.BackColor = System.Drawing.Color.Transparent;
            this.tbPumpNo.DownBack = null;
            this.tbPumpNo.Icon = null;
            this.tbPumpNo.IconIsButton = false;
            this.tbPumpNo.IconMouseState = CCWin.SkinClass.ControlState.Normal;
            this.tbPumpNo.IsPasswordChat = '\0';
            this.tbPumpNo.IsSystemPasswordChar = false;
            this.tbPumpNo.Lines = new string[0];
            this.tbPumpNo.Location = new System.Drawing.Point(482, 37);
            this.tbPumpNo.Margin = new System.Windows.Forms.Padding(0);
            this.tbPumpNo.MaxLength = 32767;
            this.tbPumpNo.MinimumSize = new System.Drawing.Size(38, 49);
            this.tbPumpNo.MouseBack = null;
            this.tbPumpNo.MouseState = CCWin.SkinClass.ControlState.Normal;
            this.tbPumpNo.Multiline = true;
            this.tbPumpNo.Name = "tbPumpNo";
            this.tbPumpNo.NormlBack = null;
            this.tbPumpNo.Padding = new System.Windows.Forms.Padding(7, 9, 7, 9);
            this.tbPumpNo.ReadOnly = false;
            this.tbPumpNo.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.tbPumpNo.Size = new System.Drawing.Size(275, 49);
            // 
            // 
            // 
            this.tbPumpNo.SkinTxt.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbPumpNo.SkinTxt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbPumpNo.SkinTxt.Font = new System.Drawing.Font("Noto Sans CJK SC Regular", 8.999999F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbPumpNo.SkinTxt.Location = new System.Drawing.Point(7, 9);
            this.tbPumpNo.SkinTxt.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.tbPumpNo.SkinTxt.Multiline = true;
            this.tbPumpNo.SkinTxt.Name = "BaseText";
            this.tbPumpNo.SkinTxt.Size = new System.Drawing.Size(262, 30);
            this.tbPumpNo.SkinTxt.TabIndex = 0;
            this.tbPumpNo.SkinTxt.WaterColor = System.Drawing.Color.FromArgb(((int)(((byte)(127)))), ((int)(((byte)(127)))), ((int)(((byte)(127)))));
            this.tbPumpNo.SkinTxt.WaterText = "";
            this.tbPumpNo.TabIndex = 2;
            this.tbPumpNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.tbPumpNo.WaterColor = System.Drawing.Color.FromArgb(((int)(((byte)(127)))), ((int)(((byte)(127)))), ((int)(((byte)(127)))));
            this.tbPumpNo.WaterText = "";
            this.tbPumpNo.WordWrap = true;
            // 
            // skinMenuStrip1
            // 
            this.skinMenuStrip1.Arrow = System.Drawing.Color.Black;
            this.skinMenuStrip1.Back = System.Drawing.Color.White;
            this.skinMenuStrip1.BackRadius = 4;
            this.skinMenuStrip1.BackRectangle = new System.Drawing.Rectangle(10, 10, 10, 10);
            this.skinMenuStrip1.Base = System.Drawing.Color.FromArgb(((int)(((byte)(105)))), ((int)(((byte)(200)))), ((int)(((byte)(254)))));
            this.skinMenuStrip1.BaseFore = System.Drawing.Color.Black;
            this.skinMenuStrip1.BaseForeAnamorphosis = false;
            this.skinMenuStrip1.BaseForeAnamorphosisBorder = 4;
            this.skinMenuStrip1.BaseForeAnamorphosisColor = System.Drawing.Color.White;
            this.skinMenuStrip1.BaseHoverFore = System.Drawing.Color.White;
            this.skinMenuStrip1.BaseItemAnamorphosis = true;
            this.skinMenuStrip1.BaseItemBorder = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinMenuStrip1.BaseItemBorderShow = true;
            this.skinMenuStrip1.BaseItemDown = ((System.Drawing.Image)(resources.GetObject("skinMenuStrip1.BaseItemDown")));
            this.skinMenuStrip1.BaseItemHover = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinMenuStrip1.BaseItemMouse = ((System.Drawing.Image)(resources.GetObject("skinMenuStrip1.BaseItemMouse")));
            this.skinMenuStrip1.BaseItemPressed = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinMenuStrip1.BaseItemRadius = 4;
            this.skinMenuStrip1.BaseItemRadiusStyle = CCWin.SkinClass.RoundStyle.All;
            this.skinMenuStrip1.BaseItemSplitter = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinMenuStrip1.DropDownImageSeparator = System.Drawing.Color.FromArgb(((int)(((byte)(197)))), ((int)(((byte)(197)))), ((int)(((byte)(197)))));
            this.skinMenuStrip1.Font = new System.Drawing.Font("Microsoft YaHei", 11F);
            this.skinMenuStrip1.Fore = System.Drawing.Color.Black;
            this.skinMenuStrip1.HoverFore = System.Drawing.Color.White;
            this.skinMenuStrip1.ItemAnamorphosis = true;
            this.skinMenuStrip1.ItemBorder = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinMenuStrip1.ItemBorderShow = true;
            this.skinMenuStrip1.ItemHover = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinMenuStrip1.ItemPressed = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(148)))), ((int)(((byte)(212)))));
            this.skinMenuStrip1.ItemRadius = 4;
            this.skinMenuStrip1.ItemRadiusStyle = CCWin.SkinClass.RoundStyle.All;
            this.skinMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.文件ToolStripMenuItem,
            this.操作ToolStripMenuItem});
            this.skinMenuStrip1.Location = new System.Drawing.Point(4, 28);
            this.skinMenuStrip1.Name = "skinMenuStrip1";
            this.skinMenuStrip1.Padding = new System.Windows.Forms.Padding(8, 3, 0, 3);
            this.skinMenuStrip1.RadiusStyle = CCWin.SkinClass.RoundStyle.All;
            this.skinMenuStrip1.Size = new System.Drawing.Size(1295, 30);
            this.skinMenuStrip1.SkinAllColor = true;
            this.skinMenuStrip1.TabIndex = 3;
            this.skinMenuStrip1.Text = "skinMenuStrip1";
            this.skinMenuStrip1.TitleAnamorphosis = true;
            this.skinMenuStrip1.TitleColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(228)))), ((int)(((byte)(236)))));
            this.skinMenuStrip1.TitleRadius = 4;
            this.skinMenuStrip1.TitleRadiusStyle = CCWin.SkinClass.RoundStyle.All;
            // 
            // 文件ToolStripMenuItem
            // 
            this.文件ToolStripMenuItem.Name = "文件ToolStripMenuItem";
            this.文件ToolStripMenuItem.Size = new System.Drawing.Size(51, 24);
            this.文件ToolStripMenuItem.Text = "文件";
            // 
            // 操作ToolStripMenuItem
            // 
            this.操作ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuInitTooling,
            this.menuTareTooling});
            this.操作ToolStripMenuItem.Name = "操作ToolStripMenuItem";
            this.操作ToolStripMenuItem.Size = new System.Drawing.Size(51, 24);
            this.操作ToolStripMenuItem.Text = "设置";
            // 
            // menuInitTooling
            // 
            this.menuInitTooling.Name = "menuInitTooling";
            this.menuInitTooling.Size = new System.Drawing.Size(168, 24);
            this.menuInitTooling.Text = "工装初始化";
            this.menuInitTooling.Click += new System.EventHandler(this.menuInitTooling_Click);
            // 
            // menuTareTooling
            // 
            this.menuTareTooling.Name = "menuTareTooling";
            this.menuTareTooling.Size = new System.Drawing.Size(168, 24);
            this.menuTareTooling.Text = "工装读数清零";
            this.menuTareTooling.Click += new System.EventHandler(this.menuTareTooling_Click);
            // 
            // gbChannel2
            // 
            this.gbChannel2.BackColor = System.Drawing.Color.Transparent;
            this.gbChannel2.BorderColor = System.Drawing.Color.White;
            this.gbChannel2.Controls.Add(this.cbPumpPort2);
            this.gbChannel2.Controls.Add(this.tbRateChannel2);
            this.gbChannel2.Controls.Add(this.btnStopPumpChannel2);
            this.gbChannel2.Controls.Add(this.WavelinePanel2);
            this.gbChannel2.Controls.Add(this.btnStartPumpChannel2);
            this.gbChannel2.Controls.Add(this.cbToolingPort2);
            this.gbChannel2.Controls.Add(this.lbPChannel2);
            this.gbChannel2.Controls.Add(this.skinLabel4);
            this.gbChannel2.Controls.Add(this.lbWeightChannel2);
            this.gbChannel2.Controls.Add(this.skinLabel6);
            this.gbChannel2.Controls.Add(this.skinLabel11);
            this.gbChannel2.Controls.Add(this.skinLabel12);
            this.gbChannel2.Controls.Add(this.skinLabel15);
            this.gbChannel2.Controls.Add(this.skinLabel16);
            this.gbChannel2.Controls.Add(this.skinLabel17);
            this.gbChannel2.Enabled = false;
            this.gbChannel2.Font = new System.Drawing.Font("Microsoft YaHei", 14F);
            this.gbChannel2.ForeColor = System.Drawing.Color.Black;
            this.gbChannel2.Location = new System.Drawing.Point(657, 210);
            this.gbChannel2.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.gbChannel2.Name = "gbChannel2";
            this.gbChannel2.Padding = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.gbChannel2.RectBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.gbChannel2.RoundStyle = CCWin.SkinClass.RoundStyle.All;
            this.gbChannel2.Size = new System.Drawing.Size(640, 885);
            this.gbChannel2.TabIndex = 2;
            this.gbChannel2.TabStop = false;
            this.gbChannel2.Text = "2道泵";
            this.gbChannel2.TitleBorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.gbChannel2.TitleRectBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(208)))), ((int)(((byte)(255)))));
            this.gbChannel2.TitleRoundStyle = CCWin.SkinClass.RoundStyle.All;
            // 
            // cbPumpPort2
            // 
            this.cbPumpPort2.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbPumpPort2.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.cbPumpPort2.FormattingEnabled = true;
            this.cbPumpPort2.Location = new System.Drawing.Point(272, 54);
            this.cbPumpPort2.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.cbPumpPort2.Name = "cbPumpPort2";
            this.cbPumpPort2.Size = new System.Drawing.Size(106, 26);
            this.cbPumpPort2.TabIndex = 1;
            this.cbPumpPort2.WaterText = "";
            this.cbPumpPort2.SelectedIndexChanged += new System.EventHandler(this.cbPumpPort2_SelectedIndexChanged);
            // 
            // tbRateChannel2
            // 
            this.tbRateChannel2.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.tbRateChannel2.Location = new System.Drawing.Point(450, 52);
            this.tbRateChannel2.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.tbRateChannel2.Name = "tbRateChannel2";
            this.tbRateChannel2.Size = new System.Drawing.Size(106, 25);
            this.tbRateChannel2.TabIndex = 4;
            this.tbRateChannel2.Text = "50";
            this.tbRateChannel2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.OnRateKeyPress);
            // 
            // btnStopPumpChannel2
            // 
            this.btnStopPumpChannel2.BackColor = System.Drawing.Color.Transparent;
            this.btnStopPumpChannel2.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.btnStopPumpChannel2.DownBack = null;
            this.btnStopPumpChannel2.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.btnStopPumpChannel2.Location = new System.Drawing.Point(560, 119);
            this.btnStopPumpChannel2.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.btnStopPumpChannel2.MouseBack = null;
            this.btnStopPumpChannel2.Name = "btnStopPumpChannel2";
            this.btnStopPumpChannel2.NormlBack = null;
            this.btnStopPumpChannel2.Size = new System.Drawing.Size(66, 45);
            this.btnStopPumpChannel2.TabIndex = 5;
            this.btnStopPumpChannel2.Text = "停止";
            this.btnStopPumpChannel2.UseVisualStyleBackColor = false;
            this.btnStopPumpChannel2.Click += new System.EventHandler(this.btnStopPumpChannel2_Click);
            // 
            // WavelinePanel2
            // 
            this.WavelinePanel2.BackColor = System.Drawing.Color.White;
            this.WavelinePanel2.Font = new System.Drawing.Font("Microsoft YaHei", 9F);
            this.WavelinePanel2.ForeColor = System.Drawing.Color.Black;
            this.WavelinePanel2.Location = new System.Drawing.Point(3, 173);
            this.WavelinePanel2.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.WavelinePanel2.Name = "WavelinePanel2";
            this.WavelinePanel2.Size = new System.Drawing.Size(633, 710);
            this.WavelinePanel2.TabIndex = 4;
            this.WavelinePanel2.Paint += new System.Windows.Forms.PaintEventHandler(this.WavelinePanel2_Paint);
            // 
            // btnStartPumpChannel2
            // 
            this.btnStartPumpChannel2.BackColor = System.Drawing.Color.Transparent;
            this.btnStartPumpChannel2.ControlState = CCWin.SkinClass.ControlState.Normal;
            this.btnStartPumpChannel2.DownBack = null;
            this.btnStartPumpChannel2.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.btnStartPumpChannel2.Location = new System.Drawing.Point(482, 119);
            this.btnStartPumpChannel2.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.btnStartPumpChannel2.MouseBack = null;
            this.btnStartPumpChannel2.Name = "btnStartPumpChannel2";
            this.btnStartPumpChannel2.NormlBack = null;
            this.btnStartPumpChannel2.Size = new System.Drawing.Size(66, 45);
            this.btnStartPumpChannel2.TabIndex = 4;
            this.btnStartPumpChannel2.Text = "开始";
            this.btnStartPumpChannel2.UseVisualStyleBackColor = false;
            this.btnStartPumpChannel2.Click += new System.EventHandler(this.btnStartPumpChannel2_Click);
            // 
            // cbToolingPort2
            // 
            this.cbToolingPort2.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbToolingPort2.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.cbToolingPort2.FormattingEnabled = true;
            this.cbToolingPort2.Location = new System.Drawing.Point(98, 54);
            this.cbToolingPort2.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.cbToolingPort2.Name = "cbToolingPort2";
            this.cbToolingPort2.Size = new System.Drawing.Size(106, 26);
            this.cbToolingPort2.TabIndex = 1;
            this.cbToolingPort2.WaterText = "";
            this.cbToolingPort2.SelectedIndexChanged += new System.EventHandler(this.cbToolingPort2_SelectedIndexChanged);
            // 
            // lbPChannel2
            // 
            this.lbPChannel2.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.lbPChannel2.AutoSize = true;
            this.lbPChannel2.BackColor = System.Drawing.Color.Transparent;
            this.lbPChannel2.BorderColor = System.Drawing.Color.White;
            this.lbPChannel2.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.lbPChannel2.Location = new System.Drawing.Point(303, 122);
            this.lbPChannel2.Name = "lbPChannel2";
            this.lbPChannel2.Size = new System.Drawing.Size(39, 20);
            this.lbPChannel2.TabIndex = 0;
            this.lbPChannel2.Text = "-----";
            // 
            // skinLabel4
            // 
            this.skinLabel4.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel4.AutoSize = true;
            this.skinLabel4.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel4.BorderColor = System.Drawing.Color.White;
            this.skinLabel4.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel4.Location = new System.Drawing.Point(223, 122);
            this.skinLabel4.Name = "skinLabel4";
            this.skinLabel4.Size = new System.Drawing.Size(60, 20);
            this.skinLabel4.TabIndex = 0;
            this.skinLabel4.Text = "机器P值";
            // 
            // lbWeightChannel2
            // 
            this.lbWeightChannel2.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.lbWeightChannel2.AutoSize = true;
            this.lbWeightChannel2.BackColor = System.Drawing.Color.Transparent;
            this.lbWeightChannel2.BorderColor = System.Drawing.Color.White;
            this.lbWeightChannel2.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.lbWeightChannel2.Location = new System.Drawing.Point(96, 121);
            this.lbWeightChannel2.Name = "lbWeightChannel2";
            this.lbWeightChannel2.Size = new System.Drawing.Size(39, 20);
            this.lbWeightChannel2.TabIndex = 0;
            this.lbWeightChannel2.Text = "-----";
            // 
            // skinLabel6
            // 
            this.skinLabel6.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel6.AutoSize = true;
            this.skinLabel6.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel6.BorderColor = System.Drawing.Color.White;
            this.skinLabel6.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel6.Location = new System.Drawing.Point(14, 121);
            this.skinLabel6.Name = "skinLabel6";
            this.skinLabel6.Size = new System.Drawing.Size(65, 20);
            this.skinLabel6.TabIndex = 0;
            this.skinLabel6.Text = "工装读数";
            // 
            // skinLabel11
            // 
            this.skinLabel11.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel11.AutoSize = true;
            this.skinLabel11.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel11.BorderColor = System.Drawing.Color.White;
            this.skinLabel11.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel11.Location = new System.Drawing.Point(384, 56);
            this.skinLabel11.Name = "skinLabel11";
            this.skinLabel11.Size = new System.Drawing.Size(45, 20);
            this.skinLabel11.TabIndex = 0;
            this.skinLabel11.Text = "速  率";
            // 
            // skinLabel12
            // 
            this.skinLabel12.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel12.AutoSize = true;
            this.skinLabel12.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel12.BorderColor = System.Drawing.Color.White;
            this.skinLabel12.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel12.Location = new System.Drawing.Point(176, 121);
            this.skinLabel12.Name = "skinLabel12";
            this.skinLabel12.Size = new System.Drawing.Size(26, 20);
            this.skinLabel12.TabIndex = 0;
            this.skinLabel12.Text = "kg";
            // 
            // skinLabel15
            // 
            this.skinLabel15.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel15.AutoSize = true;
            this.skinLabel15.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel15.BorderColor = System.Drawing.Color.White;
            this.skinLabel15.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel15.Location = new System.Drawing.Point(208, 59);
            this.skinLabel15.Name = "skinLabel15";
            this.skinLabel15.Size = new System.Drawing.Size(51, 20);
            this.skinLabel15.TabIndex = 0;
            this.skinLabel15.Text = "泵串口";
            // 
            // skinLabel16
            // 
            this.skinLabel16.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel16.AutoSize = true;
            this.skinLabel16.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel16.BorderColor = System.Drawing.Color.White;
            this.skinLabel16.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel16.Location = new System.Drawing.Point(554, 56);
            this.skinLabel16.Name = "skinLabel16";
            this.skinLabel16.Size = new System.Drawing.Size(44, 20);
            this.skinLabel16.TabIndex = 0;
            this.skinLabel16.Text = "mL\\h";
            // 
            // skinLabel17
            // 
            this.skinLabel17.ArtTextStyle = CCWin.SkinControl.ArtTextStyle.None;
            this.skinLabel17.AutoSize = true;
            this.skinLabel17.BackColor = System.Drawing.Color.Transparent;
            this.skinLabel17.BorderColor = System.Drawing.Color.White;
            this.skinLabel17.Font = new System.Drawing.Font("Microsoft YaHei", 10F);
            this.skinLabel17.Location = new System.Drawing.Point(11, 59);
            this.skinLabel17.Name = "skinLabel17";
            this.skinLabel17.Size = new System.Drawing.Size(65, 20);
            this.skinLabel17.TabIndex = 0;
            this.skinLabel17.Text = "工装串口";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1303, 1039);
            this.Controls.Add(this.gbChannel2);
            this.Controls.Add(this.skinGroupBox1);
            this.Controls.Add(this.gbParameter);
            this.Controls.Add(this.skinMenuStrip1);
            this.Font = new System.Drawing.Font("Noto Sans CJK SC Regular", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.skinMenuStrip1;
            this.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.Text = "压力测试1.0";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.skinGroupBox1.ResumeLayout(false);
            this.skinGroupBox1.PerformLayout();
            this.gbParameter.ResumeLayout(false);
            this.gbParameter.PerformLayout();
            this.skinMenuStrip1.ResumeLayout(false);
            this.skinMenuStrip1.PerformLayout();
            this.gbChannel2.ResumeLayout(false);
            this.gbChannel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private CCWin.SkinControl.SkinLabel lbPumpPort;
        private CCWin.SkinControl.SkinComboBox cbPumpPort;
        private CCWin.SkinControl.SkinLabel lbToolingPort;
        private CCWin.SkinControl.SkinComboBox cbToolingPort;
        private CCWin.SkinControl.SkinLabel lbRate;
        private CCWin.SkinControl.SkinComboBox cbPumpType;
        private CCWin.SkinControl.SkinLabel skinLabel2;
        private CCWin.SkinControl.SkinLabel skinLabel3;
        private System.Windows.Forms.Panel WavelinePanel;
        private CCWin.SkinControl.SkinButton btnStopPumpChannel1;
        private CCWin.SkinControl.SkinButton btnStartPumpChannel1;
        private CCWin.SkinControl.SkinGroupBox skinGroupBox1;
        private CCWin.SkinControl.SkinGroupBox gbParameter;
        private CCWin.SkinControl.SkinLabel skinLabel7;
        private CCWin.SkinControl.SkinLabel skinLabel8;
        private CCWin.SkinControl.SkinTextBox tbToolingNo;
        private CCWin.SkinControl.SkinTextBox tbPumpNo;
        private CCWin.SkinControl.SkinLabel lbPChannel1;
        private CCWin.SkinControl.SkinLabel skinLabel13;
        private CCWin.SkinControl.SkinLabel lbWeightChannel1;
        private CCWin.SkinControl.SkinLabel skinLabel10;
        private CCWin.SkinControl.SkinLabel skinLabel9;
        private CCWin.SkinControl.SkinMenuStrip skinMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 文件ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 操作ToolStripMenuItem;
        private System.Windows.Forms.TextBox tbRateChannel1;
        private CCWin.SkinControl.SkinGroupBox gbChannel2;
        private CCWin.SkinControl.SkinComboBox cbPumpPort2;
        private CCWin.SkinControl.SkinButton btnStopPumpChannel2;
        private System.Windows.Forms.Panel WavelinePanel2;
        private CCWin.SkinControl.SkinButton btnStartPumpChannel2;
        private CCWin.SkinControl.SkinComboBox cbToolingPort2;
        private CCWin.SkinControl.SkinLabel lbPChannel2;
        private CCWin.SkinControl.SkinLabel skinLabel4;
        private CCWin.SkinControl.SkinLabel lbWeightChannel2;
        private CCWin.SkinControl.SkinLabel skinLabel6;
        private CCWin.SkinControl.SkinLabel skinLabel11;
        private CCWin.SkinControl.SkinLabel skinLabel12;
        private CCWin.SkinControl.SkinLabel skinLabel15;
        private CCWin.SkinControl.SkinLabel skinLabel16;
        private CCWin.SkinControl.SkinLabel skinLabel17;
        private System.Windows.Forms.ToolStripMenuItem menuInitTooling;
        private System.Windows.Forms.ToolStripMenuItem menuTareTooling;
        private System.Windows.Forms.TextBox tbRateChannel2;
    }
}

