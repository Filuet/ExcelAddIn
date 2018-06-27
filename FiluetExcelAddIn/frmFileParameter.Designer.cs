namespace FiluetExcelAddIn
{
    partial class frmFileParameter
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonEdit1 = new DevExpress.XtraEditors.ButtonEdit();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.btnCancel = new DevExpress.XtraEditors.SimpleButton();
            this.btnConfirm = new DevExpress.XtraEditors.SimpleButton();
            this.shapeContainer1 = new Microsoft.VisualBasic.PowerPacks.ShapeContainer();
            this.rectangleShape1 = new Microsoft.VisualBasic.PowerPacks.RectangleShape();
            this.cbRU1B10 = new DevExpress.XtraEditors.CheckEdit();
            this.cbRUCB20 = new DevExpress.XtraEditors.CheckEdit();
            this.cbRUCB80 = new DevExpress.XtraEditors.CheckEdit();
            this.cbRU1B80 = new DevExpress.XtraEditors.CheckEdit();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            ((System.ComponentModel.ISupportInitialize)(this.buttonEdit1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbRU1B10.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbRUCB20.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbRUCB80.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbRU1B80.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonEdit1
            // 
            this.buttonEdit1.Location = new System.Drawing.Point(12, 30);
            this.buttonEdit1.Name = "buttonEdit1";
            this.buttonEdit1.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
            this.buttonEdit1.Properties.ButtonClick += new DevExpress.XtraEditors.Controls.ButtonPressedEventHandler(this.buttonEdit1_Properties_ButtonClick);
            this.buttonEdit1.Size = new System.Drawing.Size(458, 20);
            this.buttonEdit1.TabIndex = 0;
            this.buttonEdit1.EditValueChanged += new System.EventHandler(this.buttonEdit1_EditValueChanged);
            // 
            // labelControl1
            // 
            this.labelControl1.Location = new System.Drawing.Point(12, 11);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(252, 13);
            this.labelControl1.TabIndex = 1;
            this.labelControl1.Text = "Выберите файл HL_Quantity_Onhand_Report*.txt ";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(279, 104);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(72, 23);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnConfirm
            // 
            this.btnConfirm.Location = new System.Drawing.Point(357, 104);
            this.btnConfirm.Name = "btnConfirm";
            this.btnConfirm.Size = new System.Drawing.Size(72, 23);
            this.btnConfirm.TabIndex = 3;
            this.btnConfirm.Text = "OK";
            this.btnConfirm.Click += new System.EventHandler(this.btnConfirm_Click);
            // 
            // shapeContainer1
            // 
            this.shapeContainer1.Location = new System.Drawing.Point(0, 0);
            this.shapeContainer1.Margin = new System.Windows.Forms.Padding(0);
            this.shapeContainer1.Name = "shapeContainer1";
            this.shapeContainer1.Shapes.AddRange(new Microsoft.VisualBasic.PowerPacks.Shape[] {
            this.rectangleShape1});
            this.shapeContainer1.Size = new System.Drawing.Size(496, 170);
            this.shapeContainer1.TabIndex = 4;
            this.shapeContainer1.TabStop = false;
            // 
            // rectangleShape1
            // 
            this.rectangleShape1.Location = new System.Drawing.Point(1, 0);
            this.rectangleShape1.Name = "rectangleShape1";
            this.rectangleShape1.Size = new System.Drawing.Size(494, 169);
            // 
            // cbRU1B10
            // 
            this.cbRU1B10.Location = new System.Drawing.Point(12, 83);
            this.cbRU1B10.Name = "cbRU1B10";
            this.cbRU1B10.Properties.Caption = "RU1B10";
            this.cbRU1B10.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            this.cbRU1B10.Properties.DisplayValueChecked = "2";
            this.cbRU1B10.Properties.DisplayValueUnchecked = "0";
            this.cbRU1B10.Properties.ValueChecked = 1;
            this.cbRU1B10.Properties.ValueUnchecked = 0;
            this.cbRU1B10.Size = new System.Drawing.Size(75, 19);
            this.cbRU1B10.TabIndex = 5;
            this.cbRU1B10.ToolTip = "A, B, IN2";
            // 
            // cbRUCB20
            // 
            this.cbRUCB20.Location = new System.Drawing.Point(12, 108);
            this.cbRUCB20.Name = "cbRUCB20";
            this.cbRUCB20.Properties.Caption = "RUCB20";
            this.cbRUCB20.Properties.ValueChecked = 4;
            this.cbRUCB20.Properties.ValueUnchecked = 0;
            this.cbRUCB20.Size = new System.Drawing.Size(109, 19);
            this.cbRUCB20.TabIndex = 6;
            this.cbRUCB20.ToolTip = "C";
            // 
            // cbRUCB80
            // 
            this.cbRUCB80.Location = new System.Drawing.Point(159, 108);
            this.cbRUCB80.Name = "cbRUCB80";
            this.cbRUCB80.Properties.Caption = "RUCB80";
            this.cbRUCB80.Properties.ValueChecked = 5;
            this.cbRUCB80.Properties.ValueUnchecked = 0;
            this.cbRUCB80.Size = new System.Drawing.Size(75, 19);
            this.cbRUCB80.TabIndex = 7;
            this.cbRUCB80.ToolTip = "BRAK-C";
            // 
            // cbRU1B80
            // 
            this.cbRU1B80.Location = new System.Drawing.Point(159, 83);
            this.cbRU1B80.Name = "cbRU1B80";
            this.cbRU1B80.Properties.Caption = "RU1B80";
            this.cbRU1B80.Properties.ValueChecked = 3;
            this.cbRU1B80.Properties.ValueUnchecked = 0;
            this.cbRU1B80.Size = new System.Drawing.Size(75, 19);
            this.cbRU1B80.TabIndex = 9;
            this.cbRU1B80.ToolTip = "IN1";
            // 
            // labelControl2
            // 
            this.labelControl2.Location = new System.Drawing.Point(12, 64);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(99, 13);
            this.labelControl2.TabIndex = 10;
            this.labelControl2.Text = "Выберите локатор:";
            // 
            // frmFileParameter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(496, 170);
            this.Controls.Add(this.labelControl2);
            this.Controls.Add(this.cbRU1B80);
            this.Controls.Add(this.cbRUCB80);
            this.Controls.Add(this.cbRUCB20);
            this.Controls.Add(this.cbRU1B10);
            this.Controls.Add(this.btnConfirm);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.labelControl1);
            this.Controls.Add(this.buttonEdit1);
            this.Controls.Add(this.shapeContainer1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "frmFileParameter";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmPathParameter";
            this.Load += new System.EventHandler(this.frmPathParameter_Load);
            ((System.ComponentModel.ISupportInitialize)(this.buttonEdit1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbRU1B10.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbRUCB20.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbRUCB80.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbRU1B80.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraEditors.ButtonEdit buttonEdit1;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.SimpleButton btnCancel;
        private DevExpress.XtraEditors.SimpleButton btnConfirm;
        private Microsoft.VisualBasic.PowerPacks.ShapeContainer shapeContainer1;
        private Microsoft.VisualBasic.PowerPacks.RectangleShape rectangleShape1;
        private DevExpress.XtraEditors.CheckEdit cbRU1B10;
        private DevExpress.XtraEditors.CheckEdit cbRUCB20;
        private DevExpress.XtraEditors.CheckEdit cbRUCB80;
        private DevExpress.XtraEditors.CheckEdit cbRU1B80;
        private DevExpress.XtraEditors.LabelControl labelControl2;
    }
}