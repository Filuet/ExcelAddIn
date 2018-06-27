using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FiluetExcelAddIn
{
    public partial class frmTextBoxParameter : DevExpress.XtraEditors.XtraForm
    {
        public frmTextBoxParameter(string parameterName)
        {
            if (String.IsNullOrEmpty(parameterName) || String.IsNullOrEmpty(parameterName))
            {
                MessageBox.Show("Ошибка! Обратитесь к разработчику!");
                return;
            }
            InitializeComponent();
            this.lblParameterName.Text = parameterName; 
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            string orderCode = this.teParameterValue.Text.Length < 20 ? this.teParameterValue.Text : this.teParameterValue.Text.Substring(0, 20);
            if (String.IsNullOrWhiteSpace(orderCode))
            {
                MessageBox.Show("Некорректный номер заказа");
                return;
            }

            ThisAddIn.OrderCode = orderCode;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            ThisAddIn.OrderCode = null;
            Close();
        }

        private void rectangleShape1_Click(object sender, EventArgs e)
        {

        }        
    }
}
