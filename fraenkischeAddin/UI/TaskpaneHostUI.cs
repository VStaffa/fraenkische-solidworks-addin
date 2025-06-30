using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;

namespace Fraenkische.SWAddin.UI
{
    [ProgId(SWAddinClass.SWTASKPANE_PROGID)]
    public partial class TaskpaneHostUI : UserControl
    {
        private Button button1;

        public TaskpaneHostUI()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(93, 26);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(149, 33);
            this.button1.TabIndex = 0;
            this.button1.Text = "AUTOKoch MANUAL";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // TaskpaneHostUI
            // 
            this.Controls.Add(this.button1);
            this.Name = "TaskpaneHostUI";
            this.Size = new System.Drawing.Size(343, 517);
            this.ResumeLayout(false);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var basePath = Path.Combine(Path.GetDirectoryName(typeof(SWAddinClass).Assembly.Location), @"Resources\Manuals");
            var manualPath = Path.Combine(basePath, "AUTOKoch - USER MANUAL.pdf");
            
            Process.Start(new ProcessStartInfo(manualPath) {UseShellExecute = true });
        }
    }
}