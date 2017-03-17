namespace FiluetExcelAddIn
{
	partial class RibbonFiluet : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public RibbonFiluet()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

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

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonFiluet));
			this.group1 = this.Factory.CreateRibbonGroup();
			this.tab2 = this.Factory.CreateRibbonTab();
			this.group2 = this.Factory.CreateRibbonGroup();
			this.group3 = this.Factory.CreateRibbonGroup();
			this.button1 = this.Factory.CreateRibbonButton();
			this.button2 = this.Factory.CreateRibbonButton();
			this.button3 = this.Factory.CreateRibbonButton();
			this.button5 = this.Factory.CreateRibbonButton();
			this.button4 = this.Factory.CreateRibbonButton();
			this.tab2.SuspendLayout();
			this.group2.SuspendLayout();
			this.group3.SuspendLayout();
			// 
			// group1
			// 
			this.group1.Label = "group1";
			this.group1.Name = "group1";
			// 
			// tab2
			// 
			this.tab2.Groups.Add(this.group2);
			this.tab2.Groups.Add(this.group3);
			this.tab2.Label = "Filuet";
			this.tab2.Name = "tab2";
			// 
			// group2
			// 
			this.group2.Items.Add(this.button1);
			this.group2.Items.Add(this.button2);
			this.group2.Items.Add(this.button3);
			this.group2.Items.Add(this.button5);
			this.group2.Label = "Обработка Счетов";
			this.group2.Name = "group2";
			// 
			// group3
			// 
			this.group3.Items.Add(this.button4);
			this.group3.Label = "Billing Herbalife";
			this.group3.Name = "group3";
			// 
			// button1
			// 
			this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.button1.Image = global::FiluetExcelAddIn.Properties.Resources.DPD_Logo;
			this.button1.Label = "ДПД";
			this.button1.Name = "button1";
			this.button1.ShowImage = true;
			this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
			// 
			// button2
			// 
			this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.Label = "Пик Поинт";
			this.button2.Name = "button2";
			this.button2.ShowImage = true;
			this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
			// 
			// button3
			// 
			this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.Label = "СПСР";
			this.button3.Name = "button3";
			this.button3.ShowImage = true;
			this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
			// 
			// button5
			// 
			this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.button5.Image = global::FiluetExcelAddIn.Properties.Resources.russian_post;
			this.button5.Label = "Почта России";
			this.button5.Name = "button5";
			this.button5.ShowImage = true;
			this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
			// 
			// button4
			// 
			this.button4.Image = global::FiluetExcelAddIn.Properties.Resources.money;
			this.button4.Label = "ShipTo TransCost";
			this.button4.Name = "button4";
			this.button4.ShowImage = true;
			this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
			// 
			// RibbonFiluet
			// 
			this.Name = "RibbonFiluet";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.tab2);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonFiluet_Load);
			this.tab2.ResumeLayout(false);
			this.tab2.PerformLayout();
			this.group2.ResumeLayout(false);
			this.group2.PerformLayout();
			this.group3.ResumeLayout(false);
			this.group3.PerformLayout();

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
	}

	partial class ThisRibbonCollection
	{
		internal RibbonFiluet RibbonFiluet
		{
			get { return this.GetRibbon<RibbonFiluet>(); }
		}
	}
}
