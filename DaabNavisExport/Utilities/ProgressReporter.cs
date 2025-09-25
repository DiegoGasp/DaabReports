using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace DaabNavisExport.Utilities
{
    internal sealed class ProgressReporter : IDisposable
    {
        private readonly ProgressDialog? _dialog;
        private readonly bool _uiAvailable;
        private int _maximum;

        public ProgressReporter(int steps)
        {
            _maximum = Math.Max(1, steps);

            try
            {
                _dialog = new ProgressDialog(_maximum);
                _uiAvailable = true;
                _dialog.Show();
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                _uiAvailable = false;
                Debug.WriteLine($"Progress dialog unavailable: {ex.Message}");
            }
        }

        public void UpdateMaximum(int steps)
        {
            _maximum = Math.Max(1, steps);
            if (_uiAvailable && _dialog != null)
            {
                _dialog.UpdateMaximum(_maximum);
                Application.DoEvents();
            }
        }

        public void Report(int completedSteps, string message)
        {
            if (_uiAvailable && _dialog != null)
            {
                _dialog.UpdateProgress(Math.Max(0, Math.Min(completedSteps, _maximum)), message);
                Application.DoEvents();
            }
            else
            {
                Debug.WriteLine($"[Progress {completedSteps}/{_maximum}] {message}");
            }
        }

        public void Dispose()
        {
            if (_uiAvailable && _dialog != null)
            {
                _dialog.Close();
                _dialog.Dispose();
            }
        }
    }
}
