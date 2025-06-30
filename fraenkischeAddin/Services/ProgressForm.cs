using System.Windows.Forms;

namespace Fraenkische.SWAddin.Services
{
    public class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private int max;

        public ProgressForm(int max)
        {
            this.max = max;
            this.Text = "Processing...";
            this.Width = 400;
            this.Height = 80;
            progressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Minimum = 0,
                Maximum = max
            };
            Controls.Add(progressBar);
        }

        public void UpdateProgress(int value)
        {
            if (value > max) value = max;
            progressBar.Value = value;
        }
    }
}
