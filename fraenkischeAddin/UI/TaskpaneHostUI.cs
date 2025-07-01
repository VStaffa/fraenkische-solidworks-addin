using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;
using Fraenkische.SWAddin.Commands;
using System.Drawing;

namespace Fraenkische.SWAddin.UI
{
    [ProgId(SWAddinClass.SWTASKPANE_PROGID)]
    public partial class TaskpaneHostUI : UserControl
    {
        private TabControl tabControl1;
        private TabPage tabPage1;
        private TabPage tabPage2;
        private TabPage tabPage3;
        private TabPage tabPage4;
        private TabPage tabPage5;

        private PictureBox pictureBox1;

        // BUTTONS FOR MANUALS
        private Button btn_man_1;
        private Button btn_man_2;

        // BUTTONS FOR COMMANDS
        private Button btn_cmd_2;
        private Button btn_cmd_7;
        private Button btn_cmd_6;
        private Button btn_cmd_5;
        private Button btn_cmd_4;
        private Button btn_cmd_3;
        private Button btn_cmd_1;
        
        private Label lblActiveDocName;

        public event Action cmd_1_Clicked;
        public event Action cmd_2_Clicked;
        public event Action cmd_3_Clicked;
        public event Action cmd_4_Clicked;
        public event Action cmd_5_Clicked;
        public event Action cmd_6_Clicked;
        public event Action cmd_7_Clicked;

        public TaskpaneHostUI()
        {
            InitializeComponent();
            btn_cmd_1.Click += (s, e) => cmd_1_Clicked?.Invoke();
            btn_cmd_2.Click += (s, e) => cmd_2_Clicked?.Invoke();
            btn_cmd_3.Click += (s, e) => cmd_3_Clicked?.Invoke();
            btn_cmd_4.Click += (s, e) => cmd_4_Clicked?.Invoke();
            btn_cmd_5.Click += (s, e) => cmd_5_Clicked?.Invoke();
            btn_cmd_6.Click += (s, e) => cmd_6_Clicked?.Invoke();
            btn_cmd_7.Click += (s, e) => cmd_7_Clicked?.Invoke();
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TaskpaneHostUI));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.btn_cmd_1 = new System.Windows.Forms.Button();
            this.btn_cmd_2 = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.btn_man_1 = new System.Windows.Forms.Button();
            this.btn_man_2 = new System.Windows.Forms.Button();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.btn_cmd_3 = new System.Windows.Forms.Button();
            this.btn_cmd_4 = new System.Windows.Forms.Button();
            this.btn_cmd_5 = new System.Windows.Forms.Button();
            this.btn_cmd_6 = new System.Windows.Forms.Button();
            this.btn_cmd_7 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lblActiveDocName = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.tabPage5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Controls.Add(this.tabPage5);
            this.tabControl1.Location = new System.Drawing.Point(3, 112);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(337, 402);
            this.tabControl1.TabIndex = 2;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.btn_cmd_1);
            this.tabPage1.Controls.Add(this.btn_cmd_2);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(329, 376);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "PART";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // BTN_CMD_1
            // 
            this.btn_cmd_1.Location = new System.Drawing.Point(3, 45);
            this.btn_cmd_1.Name = "BTN_CMD_1";
            this.btn_cmd_1.Size = new System.Drawing.Size(149, 33);
            this.btn_cmd_1.TabIndex = 4;
            this.btn_cmd_1.Text = "BOM Export";
            this.btn_cmd_1.UseVisualStyleBackColor = true;
            // 
            // BTN_CMD_2
            // 
            this.btn_cmd_2.Location = new System.Drawing.Point(3, 6);
            this.btn_cmd_2.Name = "BTN_CMD_2";
            this.btn_cmd_2.Size = new System.Drawing.Size(149, 33);
            this.btn_cmd_2.TabIndex = 2;
            this.btn_cmd_2.Text = "Bodies To STP Export";
            this.btn_cmd_2.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(329, 399);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "ASSEMBLY";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tabPage3
            // 
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(329, 399);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "DRAWING";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.btn_man_1);
            this.tabPage4.Controls.Add(this.btn_man_2);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(329, 399);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "MANUALS";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // btn_man_1
            // 
            this.btn_man_1.Location = new System.Drawing.Point(161, 6);
            this.btn_man_1.Name = "BTN_MAN_1";
            this.btn_man_1.Size = new System.Drawing.Size(162, 33);
            this.btn_man_1.TabIndex = 2;
            this.btn_man_1.Text = "AUTOKoch MANUAL";
            this.btn_man_1.UseVisualStyleBackColor = true;
            this.btn_man_1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // btn_man_2
            // 
            this.btn_man_2.Location = new System.Drawing.Point(6, 6);
            this.btn_man_2.Name = "BTN_MAN_2";
            this.btn_man_2.Size = new System.Drawing.Size(149, 33);
            this.btn_man_2.TabIndex = 3;
            this.btn_man_2.Text = "BOM Export MANUAL";
            this.btn_man_2.UseVisualStyleBackColor = true;
            this.btn_man_2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.btn_cmd_3);
            this.tabPage5.Controls.Add(this.btn_cmd_4);
            this.tabPage5.Controls.Add(this.btn_cmd_5);
            this.tabPage5.Controls.Add(this.btn_cmd_6);
            this.tabPage5.Controls.Add(this.btn_cmd_7);
            this.tabPage5.Location = new System.Drawing.Point(4, 22);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage5.Size = new System.Drawing.Size(329, 376);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "OTHER";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // BTN_CMD_3
            // 
            this.btn_cmd_3.Location = new System.Drawing.Point(174, 298);
            this.btn_cmd_3.Name = "BTN_CMD_3";
            this.btn_cmd_3.Size = new System.Drawing.Size(149, 33);
            this.btn_cmd_3.TabIndex = 8;
            this.btn_cmd_3.Text = "Load Prices to Excel";
            this.btn_cmd_3.UseVisualStyleBackColor = true;
            // 
            // BTN_CMD_4
            // 
            this.btn_cmd_4.Location = new System.Drawing.Point(6, 337);
            this.btn_cmd_4.Name = "BTN_CMD_4";
            this.btn_cmd_4.Size = new System.Drawing.Size(149, 33);
            this.btn_cmd_4.TabIndex = 7;
            this.btn_cmd_4.Text = "Daily T-Number Update";
            this.btn_cmd_4.UseVisualStyleBackColor = true;
            // 
            // BTN_CMD_5
            // 
            this.btn_cmd_5.Location = new System.Drawing.Point(174, 259);
            this.btn_cmd_5.Name = "BTN_CMD_5";
            this.btn_cmd_5.Size = new System.Drawing.Size(149, 33);
            this.btn_cmd_5.TabIndex = 6;
            this.btn_cmd_5.Text = "Merge Excel BOMs in File";
            this.btn_cmd_5.UseVisualStyleBackColor = true;
            // 
            // BTN_CMD_6
            // 
            this.btn_cmd_6.Location = new System.Drawing.Point(174, 337);
            this.btn_cmd_6.Name = "BTN_CMD_6";
            this.btn_cmd_6.Size = new System.Drawing.Size(149, 33);
            this.btn_cmd_6.TabIndex = 5;
            this.btn_cmd_6.Text = "Update Source Excels";
            this.btn_cmd_6.UseVisualStyleBackColor = true;
            // 
            // btn_cmd_7
            // 
            this.btn_cmd_7.Location = new System.Drawing.Point(6, 6);
            this.btn_cmd_7.Name = "BTN_CMD_7";
            this.btn_cmd_7.Size = new System.Drawing.Size(149, 33);
            this.btn_cmd_7.TabIndex = 4;
            this.btn_cmd_7.Text = "Search T-Number";
            this.btn_cmd_7.UseVisualStyleBackColor = true;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.InitialImage = null;
            this.pictureBox1.Location = new System.Drawing.Point(7, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(185, 80);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // lblActiveDocName
            // 
            this.lblActiveDocName.AutoSize = true;
            this.lblActiveDocName.Location = new System.Drawing.Point(6, 92);
            this.lblActiveDocName.Name = "lblActiveDocName";
            this.lblActiveDocName.Size = new System.Drawing.Size(103, 13);
            this.lblActiveDocName.TabIndex = 4;
            this.lblActiveDocName.Text = "DOCUMENT NAME";
            // 
            // TaskpaneHostUI
            // 
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.lblActiveDocName);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.tabControl1);
            this.Name = "TaskpaneHostUI";
            this.Size = new System.Drawing.Size(343, 517);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.tabPage5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            var basePath = Path.Combine(Path.GetDirectoryName(typeof(SWAddinClass).Assembly.Location), @"Resources\Manuals");
            var manualPath = Path.Combine(basePath, "BOM Export - MANUAL.pdf");

            Process.Start(new ProcessStartInfo(manualPath) { UseShellExecute = true });
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            var basePath = Path.Combine(Path.GetDirectoryName(typeof(SWAddinClass).Assembly.Location), @"Resources\Manuals");
            var manualPath = Path.Combine(basePath, "AUTOKoch - USER MANUAL.pdf");

            Process.Start(new ProcessStartInfo(manualPath) { UseShellExecute = true });
        }

        public void UpdateDocumentName(string docName)
        {
            lblActiveDocName.Text = $"Active Document: {docName}";
        }
    }
}