namespace FiluetExcelAddIn
{
	partial class ProgressForm
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
			this.bOK = new DevExpress.XtraEditors.SimpleButton();
			this.progressBar = new DevExpress.XtraEditors.ProgressBarControl();
			this.shapeContainer1 = new Microsoft.VisualBasic.PowerPacks.ShapeContainer();
			this.rectangleShape1 = new Microsoft.VisualBasic.PowerPacks.RectangleShape();
			this.title = new DevExpress.XtraEditors.LabelControl();
			this.log = new DevExpress.XtraEditors.MemoEdit();
			((System.ComponentModel.ISupportInitialize)(this.progressBar.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.log.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// bOK
			// 
			this.bOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.bOK.Enabled = false;
			this.bOK.Location = new System.Drawing.Point(197, 61);
			this.bOK.Name = "bOK";
			this.bOK.Size = new System.Drawing.Size(75, 23);
			this.bOK.TabIndex = 0;
			this.bOK.Text = "ОК";
			this.bOK.Click += new System.EventHandler(this.bOK_Click);
			// 
			// progressBar
			// 
			this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.progressBar.Location = new System.Drawing.Point(12, 37);
			this.progressBar.Name = "progressBar";
			this.progressBar.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
			this.progressBar.Properties.DisplayFormat.FormatString = "0\"%\"";
			this.progressBar.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
			this.progressBar.Properties.Maximum = 203;
			this.progressBar.Properties.ShowTitle = true;
			this.progressBar.Size = new System.Drawing.Size(260, 18);
			this.progressBar.TabIndex = 1;
			// 
			// shapeContainer1
			// 
			this.shapeContainer1.Location = new System.Drawing.Point(0, 0);
			this.shapeContainer1.Margin = new System.Windows.Forms.Padding(0);
			this.shapeContainer1.Name = "shapeContainer1";
			this.shapeContainer1.Shapes.AddRange(new Microsoft.VisualBasic.PowerPacks.Shape[] {
            this.rectangleShape1});
			this.shapeContainer1.Size = new System.Drawing.Size(284, 262);
			this.shapeContainer1.TabIndex = 2;
			this.shapeContainer1.TabStop = false;
			// 
			// rectangleShape1
			// 
			this.rectangleShape1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.rectangleShape1.Location = new System.Drawing.Point(0, 0);
			this.rectangleShape1.Name = "rectangleShape1";
			this.rectangleShape1.Size = new System.Drawing.Size(283, 261);
			// 
			// title
			// 
			this.title.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.title.Appearance.Font = new System.Drawing.Font("Tahoma", 10F);
			this.title.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
			this.title.Location = new System.Drawing.Point(12, 12);
			this.title.Name = "title";
			this.title.Size = new System.Drawing.Size(260, 19);
			this.title.TabIndex = 3;
			this.title.Text = "labelControl1";
			// 
			// log
			// 
			this.log.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.log.Location = new System.Drawing.Point(12, 90);
			this.log.Name = "log";
			this.log.Properties.ReadOnly = true;
			this.log.Size = new System.Drawing.Size(260, 160);
			this.log.TabIndex = 4;
			// 
			// ProgressForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(284, 262);
			this.ControlBox = false;
			this.Controls.Add(this.log);
			this.Controls.Add(this.title);
			this.Controls.Add(this.progressBar);
			this.Controls.Add(this.bOK);
			this.Controls.Add(this.shapeContainer1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.Name = "ProgressForm";
			this.ShowIcon = false;
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "ProgressForm";
			((System.ComponentModel.ISupportInitialize)(this.progressBar.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.log.Properties)).EndInit();
			this.ResumeLayout(false);

		}

		#endregion

		private DevExpress.XtraEditors.SimpleButton bOK;
		private DevExpress.XtraEditors.ProgressBarControl progressBar;
		private Microsoft.VisualBasic.PowerPacks.ShapeContainer shapeContainer1;
		private Microsoft.VisualBasic.PowerPacks.RectangleShape rectangleShape1;
		private DevExpress.XtraEditors.LabelControl title;
		private DevExpress.XtraEditors.MemoEdit log;
	}
}