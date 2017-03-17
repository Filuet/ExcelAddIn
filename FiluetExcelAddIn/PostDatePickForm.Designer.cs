namespace FiluetExcelAddIn
{
	partial class PostDatePickForm
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
			this.dateEdit1 = new DevExpress.XtraEditors.DateEdit();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.dateEdit2 = new DevExpress.XtraEditors.DateEdit();
			this.bClose = new DevExpress.XtraEditors.SimpleButton();
			this.bOK = new DevExpress.XtraEditors.SimpleButton();
			this.shapeContainer1 = new Microsoft.VisualBasic.PowerPacks.ShapeContainer();
			this.rectangleShape1 = new Microsoft.VisualBasic.PowerPacks.RectangleShape();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit1.Properties.VistaTimeProperties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit1.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit2.Properties.VistaTimeProperties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit2.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// dateEdit1
			// 
			this.dateEdit1.EditValue = new System.DateTime(2015, 6, 16, 14, 52, 18, 0);
			this.dateEdit1.Location = new System.Drawing.Point(13, 41);
			this.dateEdit1.Name = "dateEdit1";
			this.dateEdit1.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dateEdit1.Properties.DisplayFormat.FormatString = "dd.MM.yyyy";
			this.dateEdit1.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dateEdit1.Properties.EditFormat.FormatString = "dd.MM.yyyy";
			this.dateEdit1.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dateEdit1.Properties.Mask.EditMask = "dd.MM.yyyy";
			this.dateEdit1.Properties.VistaTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
			this.dateEdit1.Size = new System.Drawing.Size(100, 20);
			this.dateEdit1.TabIndex = 0;
			this.dateEdit1.EditValueChanged += new System.EventHandler(this.dateEdit1_EditValueChanged);
			// 
			// labelControl1
			// 
			this.labelControl1.Location = new System.Drawing.Point(13, 13);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(256, 13);
			this.labelControl1.TabIndex = 1;
			this.labelControl1.Text = "Выберете даты загрузки заказов для обновления";
			// 
			// dateEdit2
			// 
			this.dateEdit2.EditValue = new System.DateTime(2015, 6, 16, 14, 52, 18, 0);
			this.dateEdit2.Location = new System.Drawing.Point(169, 41);
			this.dateEdit2.Name = "dateEdit2";
			this.dateEdit2.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dateEdit2.Properties.DisplayFormat.FormatString = "dd.MM.yyyy";
			this.dateEdit2.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dateEdit2.Properties.EditFormat.FormatString = "dd.MM.yyyy";
			this.dateEdit2.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dateEdit2.Properties.Mask.EditMask = "dd.MM.yyyy";
			this.dateEdit2.Properties.VistaTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
			this.dateEdit2.Size = new System.Drawing.Size(100, 20);
			this.dateEdit2.TabIndex = 2;
			this.dateEdit2.EditValueChanged += new System.EventHandler(this.dateEdit2_EditValueChanged);
			// 
			// bClose
			// 
			this.bClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.bClose.Location = new System.Drawing.Point(113, 82);
			this.bClose.Name = "bClose";
			this.bClose.Size = new System.Drawing.Size(75, 23);
			this.bClose.TabIndex = 3;
			this.bClose.Text = "Отмена";
			this.bClose.Click += new System.EventHandler(this.bClose_Click);
			// 
			// bOK
			// 
			this.bOK.Location = new System.Drawing.Point(194, 82);
			this.bOK.Name = "bOK";
			this.bOK.Size = new System.Drawing.Size(75, 23);
			this.bOK.TabIndex = 4;
			this.bOK.Text = "OK";
			this.bOK.Click += new System.EventHandler(this.bOK_Click);
			// 
			// shapeContainer1
			// 
			this.shapeContainer1.Location = new System.Drawing.Point(0, 0);
			this.shapeContainer1.Margin = new System.Windows.Forms.Padding(0);
			this.shapeContainer1.Name = "shapeContainer1";
			this.shapeContainer1.Shapes.AddRange(new Microsoft.VisualBasic.PowerPacks.Shape[] {
            this.rectangleShape1});
			this.shapeContainer1.Size = new System.Drawing.Size(286, 120);
			this.shapeContainer1.TabIndex = 5;
			this.shapeContainer1.TabStop = false;
			// 
			// rectangleShape1
			// 
			this.rectangleShape1.Location = new System.Drawing.Point(2, 1);
			this.rectangleShape1.Name = "rectangleShape1";
			this.rectangleShape1.Size = new System.Drawing.Size(281, 117);
			// 
			// PostDatePickForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.bClose;
			this.ClientSize = new System.Drawing.Size(286, 120);
			this.ControlBox = false;
			this.Controls.Add(this.bOK);
			this.Controls.Add(this.bClose);
			this.Controls.Add(this.dateEdit2);
			this.Controls.Add(this.labelControl1);
			this.Controls.Add(this.dateEdit1);
			this.Controls.Add(this.shapeContainer1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.Name = "PostDatePickForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "PostDatePickForm";
			((System.ComponentModel.ISupportInitialize)(this.dateEdit1.Properties.VistaTimeProperties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit1.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit2.Properties.VistaTimeProperties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dateEdit2.Properties)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private DevExpress.XtraEditors.DateEdit dateEdit1;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private DevExpress.XtraEditors.DateEdit dateEdit2;
		private DevExpress.XtraEditors.SimpleButton bClose;
		private DevExpress.XtraEditors.SimpleButton bOK;
		private Microsoft.VisualBasic.PowerPacks.ShapeContainer shapeContainer1;
		private Microsoft.VisualBasic.PowerPacks.RectangleShape rectangleShape1;
	}
}