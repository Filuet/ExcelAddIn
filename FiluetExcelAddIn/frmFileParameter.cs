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
    public partial class frmFileParameter : DevExpress.XtraEditors.XtraForm
    {
        public frmFileParameter()
        {
            InitializeComponent();
        }

        private void frmPathParameter_Load(object sender, EventArgs e)
        {
            cbRU1B10.Checked = true;
        }

        private void buttonEdit1_Properties_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            OpenFileDialog form = new OpenFileDialog();
            form.Filter = "Текстовые документы(*.txt)|*.txt";
            form.Multiselect = false;
            if (form.ShowDialog() == DialogResult.OK)
            {
                buttonEdit1.Text = form.FileName;
            }
            form.Dispose();
        }

        private void buttonEdit1_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            ThisAddIn.OrderCode = null;
            Close();
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            string fileName = this.buttonEdit1.Text;
            if (String.IsNullOrWhiteSpace(fileName))
            {
                MessageBox.Show("Файл не выбран");
                return;
            }

            ThisAddIn.fileName = fileName;
            ThisAddIn.RU1B10 = cbRU1B10.Checked ? 1: 0;
            ThisAddIn.RU1B80 = cbRU1B80.Checked ? 3: 0;
            ThisAddIn.RUCB20 = cbRUCB20.Checked ? 4: 0;
            ThisAddIn.RUCB80 = cbRUCB80.Checked ? 5: 0;
            Close();        
        }       
    }
}
