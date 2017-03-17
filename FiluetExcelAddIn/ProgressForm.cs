using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace FiluetExcelAddIn
{
	public partial class ProgressForm : DevExpress.XtraEditors.XtraForm
	{
		public ProgressForm()
		{
			InitializeComponent();
		}

		public void SetTitle(string text)
		{
			if (title.InvokeRequired)
			{
				this.BeginInvoke(
					new Action(() =>
						{
							title.Text = text;
						}
				));
			}
			else
				title.Text = text;

		}

		public void CloseButtonEnable(bool value)
		{
			if (bOK.InvokeRequired)
			{
				bOK.BeginInvoke(
					new Action(() =>
						{
							bOK.Enabled = value;
						}
				));
			}
			else
				bOK.Enabled = value;
		}

		public void SetLog(string text)
		{
			if (log.InvokeRequired)
			{
				log.BeginInvoke(
					new Action(() =>
						{
							log.Text = text;
							log.SelectionStart = text.Length;
							log.ScrollToCaret();
						}
				));
			}
			else
			{
				log.Text = text;
				log.SelectionStart = text.Length;
				log.ScrollToCaret();
			}
		}

		public void SetProgress(int value, int min, int max)
		{
			if (progressBar.InvokeRequired)
			{
				progressBar.BeginInvoke(
					new Action(() =>
						{
							progressBar.Properties.Minimum = min;
							progressBar.Properties.Maximum = max;
							progressBar.Position = value;
						}
				));
			}
			else
			{
				progressBar.Properties.Minimum = min;
				progressBar.Properties.Maximum = max;
				progressBar.Position = value;
			}

		}

		private void bOK_Click(object sender, EventArgs e)
		{
			this.Close();
		}
	}
}