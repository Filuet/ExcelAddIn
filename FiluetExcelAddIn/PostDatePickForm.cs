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
	public partial class PostDatePickForm : DevExpress.XtraEditors.XtraForm
	{
		public PostDatePickForm()
		{
			InitializeComponent();
			dateEdit1.DateTime = ThisAddIn.PostDatePick.DateStart;
			dateEdit2.DateTime = ThisAddIn.PostDatePick.DateEnd;
		}

		private void bClose_Click(object sender, EventArgs e)
		{
			ThisAddIn.PostDatePick = null;
			Close();
		}

		private void bOK_Click(object sender, EventArgs e)
		{
			ThisAddIn.PostDatePick.DateStart = dateEdit1.DateTime;
			ThisAddIn.PostDatePick.DateEnd = dateEdit2.DateTime;
			Close();
		}

		private void dateEdit1_EditValueChanged(object sender, EventArgs e)
		{
			if (dateEdit1.DateTime > dateEdit2.DateTime)
				dateEdit2.DateTime = dateEdit1.DateTime.AddDays(7);
		}

		private void dateEdit2_EditValueChanged(object sender, EventArgs e)
		{
			if (dateEdit1.DateTime > dateEdit2.DateTime)
				dateEdit1.DateTime = dateEdit2.DateTime;
		}
	}
}