using System;
using System.Drawing;
using System.Windows.Forms;

namespace DaabNavisExport.Utilities
{
    internal sealed class ProgressDialog : Form
    {
        private readonly ProgressBar _progressBar;
        private readonly Label _statusLabel;
        private int _maximum;

        public ProgressDialog(int maximum)
        {
            _maximum = Math.Max(1, maximum);

            FormBorderStyle = FormBorderStyle.FixedDialog;
            ControlBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Daab Navis Export";
            TopMost = true;
            Width = 420;
            Height = 140;

            _statusLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 50,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(12, 12, 12, 0)
            };

            _progressBar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 24,
                Minimum = 0,
                Maximum = _maximum,
                Step = 1,
                Style = ProgressBarStyle.Continuous,
                Margin = new Padding(12)
            };

            Controls.Add(_progressBar);
            Controls.Add(_statusLabel);

            UpdateProgress(0, "Starting export...");
        }

        public void UpdateProgress(int value, string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => UpdateProgress(value, message)));
                return;
            }

            var bounded = Math.Max(_progressBar.Minimum, Math.Min(value, _maximum));
            _progressBar.Value = bounded;
            _statusLabel.Text = message;
        }

        public void UpdateMaximum(int maximum)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => UpdateMaximum(maximum)));
                return;
            }

            _maximum = Math.Max(1, maximum);
            _progressBar.Maximum = _maximum;
        }
    }
}
