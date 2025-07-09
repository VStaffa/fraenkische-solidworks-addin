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
        private TabPage tabPage5;

        private PictureBox pictureBox1;

        // BUTTONS FOR COMMANDS
        private Button btn_cmd_1;
        private Button btn_cmd_2;
        private Button btn_cmd_3;
        private Button btn_cmd_4;
        private Button btn_cmd_5;
        private Button btn_cmd_6;
        private Button btn_cmd_7;
        private Button btn_man_2;
        private Button btn_man_7;
        private Button btn_man_6;
        private Button btn_man_3;
        private Button btn_man_5;
        private Button btn_man_1;
        private Button btn_man_4;
        private Button btn_man_8;
        private Button btn_cmd_8;
        private Label lblActiveDocName;

        public event Action cmd_1_Clicked;
        public event Action cmd_2_Clicked;
        public event Action cmd_3_Clicked;
        public event Action cmd_4_Clicked;
        public event Action cmd_5_Clicked;
        public event Action cmd_6_Clicked;
        public event Action cmd_7_Clicked;
        public event Action cmd_8_Clicked;

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
            btn_cmd_8.Click += (s, e) => cmd_8_Clicked?.Invoke();
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TaskpaneHostUI));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.btn_man_7 = new System.Windows.Forms.Button();
            this.btn_man_2 = new System.Windows.Forms.Button();
            this.btn_cmd_7 = new System.Windows.Forms.Button();
            this.btn_cmd_2 = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btn_man_8 = new System.Windows.Forms.Button();
            this.btn_cmd_8 = new System.Windows.Forms.Button();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.btn_man_6 = new System.Windows.Forms.Button();
            this.btn_man_3 = new System.Windows.Forms.Button();
            this.btn_man_5 = new System.Windows.Forms.Button();
            this.btn_man_1 = new System.Windows.Forms.Button();
            this.btn_man_4 = new System.Windows.Forms.Button();
            this.btn_cmd_1 = new System.Windows.Forms.Button();
            this.btn_cmd_3 = new System.Windows.Forms.Button();
            this.btn_cmd_4 = new System.Windows.Forms.Button();
            this.btn_cmd_5 = new System.Windows.Forms.Button();
            this.btn_cmd_6 = new System.Windows.Forms.Button();
            this.lblActiveDocName = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.Buttons;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage5);
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(5, 121);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(290, 576);
            this.tabControl1.TabIndex = 2;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.Transparent;
            this.tabPage1.Controls.Add(this.btn_man_7);
            this.tabPage1.Controls.Add(this.btn_man_2);
            this.tabPage1.Controls.Add(this.btn_cmd_7);
            this.tabPage1.Controls.Add(this.btn_cmd_2);
            this.tabPage1.Location = new System.Drawing.Point(4, 28);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(282, 544);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "PART";
            // 
            // btn_man_7
            // 
            this.btn_man_7.BackColor = System.Drawing.Color.Transparent;
            this.btn_man_7.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_man_7.FlatAppearance.BorderSize = 0;
            this.btn_man_7.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_man_7.Image = global::Fraenkische.SWAddin.Properties.Resources.helpIcon;
            this.btn_man_7.Location = new System.Drawing.Point(192, 70);
            this.btn_man_7.Name = "btn_man_7";
            this.btn_man_7.Size = new System.Drawing.Size(35, 35);
            this.btn_man_7.TabIndex = 7;
            this.btn_man_7.UseVisualStyleBackColor = false;
            this.btn_man_7.Click += new System.EventHandler(this.btn_man_7_Click);
            // 
            // btn_man_2
            // 
            this.btn_man_2.BackColor = System.Drawing.Color.Transparent;
            this.btn_man_2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_man_2.FlatAppearance.BorderSize = 0;
            this.btn_man_2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_man_2.Image = global::Fraenkische.SWAddin.Properties.Resources.helpIcon;
            this.btn_man_2.Location = new System.Drawing.Point(192, 14);
            this.btn_man_2.Name = "btn_man_2";
            this.btn_man_2.Size = new System.Drawing.Size(35, 35);
            this.btn_man_2.TabIndex = 6;
            this.btn_man_2.UseVisualStyleBackColor = false;
            this.btn_man_2.Click += new System.EventHandler(this.btn_man_2_Click_1);
            // 
            // btn_cmd_7
            // 
            this.btn_cmd_7.Location = new System.Drawing.Point(6, 62);
            this.btn_cmd_7.Name = "btn_cmd_7";
            this.btn_cmd_7.Size = new System.Drawing.Size(180, 50);
            this.btn_cmd_7.TabIndex = 5;
            this.btn_cmd_7.Text = "T-Number to PART";
            this.btn_cmd_7.UseVisualStyleBackColor = true;
            // 
            // btn_cmd_2
            // 
            this.btn_cmd_2.Location = new System.Drawing.Point(6, 6);
            this.btn_cmd_2.Name = "btn_cmd_2";
            this.btn_cmd_2.Size = new System.Drawing.Size(180, 50);
            this.btn_cmd_2.TabIndex = 2;
            this.btn_cmd_2.Text = "Bodies to STEP";
            this.btn_cmd_2.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.btn_man_8);
            this.tabPage2.Controls.Add(this.btn_cmd_8);
            this.tabPage2.Location = new System.Drawing.Point(4, 28);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(282, 544);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "ASSEMBLY";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // btn_man_8
            // 
            this.btn_man_8.BackColor = System.Drawing.Color.Transparent;
            this.btn_man_8.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_man_8.FlatAppearance.BorderSize = 0;
            this.btn_man_8.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_man_8.Image = global::Fraenkische.SWAddin.Properties.Resources.helpIcon;
            this.btn_man_8.Location = new System.Drawing.Point(192, 14);
            this.btn_man_8.Name = "btn_man_8";
            this.btn_man_8.Size = new System.Drawing.Size(35, 35);
            this.btn_man_8.TabIndex = 8;
            this.btn_man_8.UseVisualStyleBackColor = false;
            this.btn_man_8.Click += new System.EventHandler(this.btn_man_8_Click);
            // 
            // btn_cmd_8
            // 
            this.btn_cmd_8.Location = new System.Drawing.Point(6, 6);
            this.btn_cmd_8.Name = "btn_cmd_8";
            this.btn_cmd_8.Size = new System.Drawing.Size(180, 50);
            this.btn_cmd_8.TabIndex = 7;
            this.btn_cmd_8.Text = "Create Gauge Drawing";
            this.btn_cmd_8.UseVisualStyleBackColor = true;
            // 
            // tabPage3
            // 
            this.tabPage3.Location = new System.Drawing.Point(4, 28);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(282, 544);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "DRAWING";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.btn_man_6);
            this.tabPage5.Controls.Add(this.btn_man_3);
            this.tabPage5.Controls.Add(this.btn_man_5);
            this.tabPage5.Controls.Add(this.btn_man_1);
            this.tabPage5.Controls.Add(this.btn_man_4);
            this.tabPage5.Controls.Add(this.btn_cmd_1);
            this.tabPage5.Controls.Add(this.btn_cmd_3);
            this.tabPage5.Controls.Add(this.btn_cmd_4);
            this.tabPage5.Controls.Add(this.btn_cmd_5);
            this.tabPage5.Controls.Add(this.btn_cmd_6);
            this.tabPage5.Location = new System.Drawing.Point(4, 28);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage5.Size = new System.Drawing.Size(282, 544);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "OTHER";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // btn_man_6
            // 
            this.btn_man_6.BackColor = System.Drawing.Color.Transparent;
            this.btn_man_6.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_man_6.FlatAppearance.BorderSize = 0;
            this.btn_man_6.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_man_6.Image = global::Fraenkische.SWAddin.Properties.Resources.helpIcon;
            this.btn_man_6.Location = new System.Drawing.Point(192, 420);
            this.btn_man_6.Name = "btn_man_6";
            this.btn_man_6.Size = new System.Drawing.Size(35, 35);
            this.btn_man_6.TabIndex = 14;
            this.btn_man_6.UseVisualStyleBackColor = false;
            this.btn_man_6.Click += new System.EventHandler(this.btn_man_6_Click);
            // 
            // btn_man_3
            // 
            this.btn_man_3.BackColor = System.Drawing.Color.Transparent;
            this.btn_man_3.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_man_3.FlatAppearance.BorderSize = 0;
            this.btn_man_3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_man_3.Image = global::Fraenkische.SWAddin.Properties.Resources.helpIcon;
            this.btn_man_3.Location = new System.Drawing.Point(192, 227);
            this.btn_man_3.Name = "btn_man_3";
            this.btn_man_3.Size = new System.Drawing.Size(35, 35);
            this.btn_man_3.TabIndex = 13;
            this.btn_man_3.UseVisualStyleBackColor = false;
            this.btn_man_3.Click += new System.EventHandler(this.btn_man_3_Click);
            // 
            // btn_man_5
            // 
            this.btn_man_5.BackColor = System.Drawing.Color.Transparent;
            this.btn_man_5.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_man_5.FlatAppearance.BorderSize = 0;
            this.btn_man_5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_man_5.Image = global::Fraenkische.SWAddin.Properties.Resources.helpIcon;
            this.btn_man_5.Location = new System.Drawing.Point(192, 173);
            this.btn_man_5.Name = "btn_man_5";
            this.btn_man_5.Size = new System.Drawing.Size(35, 35);
            this.btn_man_5.TabIndex = 12;
            this.btn_man_5.UseVisualStyleBackColor = false;
            this.btn_man_5.Click += new System.EventHandler(this.btn_man_5_Click);
            // 
            // btn_man_1
            // 
            this.btn_man_1.BackColor = System.Drawing.Color.Transparent;
            this.btn_man_1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_man_1.FlatAppearance.BorderSize = 0;
            this.btn_man_1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_man_1.Image = global::Fraenkische.SWAddin.Properties.Resources.helpIcon;
            this.btn_man_1.Location = new System.Drawing.Point(192, 120);
            this.btn_man_1.Name = "btn_man_1";
            this.btn_man_1.Size = new System.Drawing.Size(35, 35);
            this.btn_man_1.TabIndex = 11;
            this.btn_man_1.UseVisualStyleBackColor = false;
            this.btn_man_1.Click += new System.EventHandler(this.btn_man_1_Click);
            // 
            // btn_man_4
            // 
            this.btn_man_4.BackColor = System.Drawing.Color.Transparent;
            this.btn_man_4.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_man_4.FlatAppearance.BorderSize = 0;
            this.btn_man_4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_man_4.Image = global::Fraenkische.SWAddin.Properties.Resources.helpIcon;
            this.btn_man_4.Location = new System.Drawing.Point(192, 14);
            this.btn_man_4.Name = "btn_man_4";
            this.btn_man_4.Size = new System.Drawing.Size(35, 35);
            this.btn_man_4.TabIndex = 10;
            this.btn_man_4.UseVisualStyleBackColor = false;
            this.btn_man_4.Click += new System.EventHandler(this.btn_man_4_Click);
            // 
            // btn_cmd_1
            // 
            this.btn_cmd_1.Location = new System.Drawing.Point(6, 112);
            this.btn_cmd_1.Name = "btn_cmd_1";
            this.btn_cmd_1.Size = new System.Drawing.Size(180, 50);
            this.btn_cmd_1.TabIndex = 9;
            this.btn_cmd_1.Text = "Batch BOM Export";
            this.btn_cmd_1.UseVisualStyleBackColor = true;
            // 
            // btn_cmd_3
            // 
            this.btn_cmd_3.Location = new System.Drawing.Point(6, 219);
            this.btn_cmd_3.Name = "btn_cmd_3";
            this.btn_cmd_3.Size = new System.Drawing.Size(180, 50);
            this.btn_cmd_3.TabIndex = 8;
            this.btn_cmd_3.Text = "Load Prices to Excel";
            this.btn_cmd_3.UseVisualStyleBackColor = true;
            // 
            // btn_cmd_4
            // 
            this.btn_cmd_4.BackColor = System.Drawing.SystemColors.ControlLight;
            this.btn_cmd_4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_cmd_4.Location = new System.Drawing.Point(6, 6);
            this.btn_cmd_4.Name = "btn_cmd_4";
            this.btn_cmd_4.Size = new System.Drawing.Size(180, 50);
            this.btn_cmd_4.TabIndex = 7;
            this.btn_cmd_4.Text = "Daily T-Number Update";
            this.btn_cmd_4.UseVisualStyleBackColor = false;
            // 
            // btn_cmd_5
            // 
            this.btn_cmd_5.Location = new System.Drawing.Point(6, 165);
            this.btn_cmd_5.Name = "btn_cmd_5";
            this.btn_cmd_5.Size = new System.Drawing.Size(180, 50);
            this.btn_cmd_5.TabIndex = 6;
            this.btn_cmd_5.Text = "Merge Excel BOMs";
            this.btn_cmd_5.UseVisualStyleBackColor = true;
            // 
            // btn_cmd_6
            // 
            this.btn_cmd_6.Enabled = false;
            this.btn_cmd_6.Location = new System.Drawing.Point(6, 412);
            this.btn_cmd_6.Name = "btn_cmd_6";
            this.btn_cmd_6.Size = new System.Drawing.Size(180, 50);
            this.btn_cmd_6.TabIndex = 5;
            this.btn_cmd_6.Text = "Update Source Excels";
            this.btn_cmd_6.UseVisualStyleBackColor = true;
            // 
            // lblActiveDocName
            // 
            this.lblActiveDocName.AutoSize = true;
            this.lblActiveDocName.Location = new System.Drawing.Point(4, 95);
            this.lblActiveDocName.Name = "lblActiveDocName";
            this.lblActiveDocName.Size = new System.Drawing.Size(121, 13);
            this.lblActiveDocName.TabIndex = 4;
            this.lblActiveDocName.Text = "NO DOCUMENT OPEN";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.InitialImage = null;
            this.pictureBox1.Location = new System.Drawing.Point(5, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Padding = new System.Windows.Forms.Padding(3);
            this.pictureBox1.Size = new System.Drawing.Size(190, 80);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // TaskpaneHostUI
            // 
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.lblActiveDocName);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.tabControl1);
            this.Name = "TaskpaneHostUI";
            this.Size = new System.Drawing.Size(300, 700);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        public void UpdateDocumentName(string docName)
        {
            lblActiveDocName.Text = $"Active Document: {docName}";
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btn_man_1_Click(object sender, EventArgs e)
        {
            var basePath = Path.Combine(Path.GetDirectoryName(typeof(SWAddinClass).Assembly.Location), @"Resources\Manuals");
            var manualPath = Path.Combine(basePath, "BOM Export - MANUAL.pdf");

            Process.Start(new ProcessStartInfo(manualPath) { UseShellExecute = true });
        }

        private void btn_man_2_Click_1(object sender, EventArgs e)
        {
            var basePath = Path.Combine(Path.GetDirectoryName(typeof(SWAddinClass).Assembly.Location), @"Resources\Manuals");
            var manualPath = Path.Combine(basePath, "ExportBodiesToSTEP.pdf");

            Process.Start(new ProcessStartInfo(manualPath) { UseShellExecute = true });
        }
        private void btn_man_3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Manual for Load Prices to Excel command is not available yet.", "Manual Not Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btn_man_4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Manual for Daily T-Number Update command is not available yet.", "Manual Not Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btn_man_5_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Manual for Merge Excel BOMs command is not available yet.", "Manual Not Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void btn_man_6_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Manual for Update Source Excels command is not available yet.", "Manual Not Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btn_man_7_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Manual for T-Number to PART command is not available yet.", "Manual Not Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btn_man_8_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Manual for Bodies to STEP command is not available yet.", "Manual Not Available", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}