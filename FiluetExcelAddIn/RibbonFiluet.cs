using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace FiluetExcelAddIn
{
	public partial class RibbonFiluet
	{
		private void RibbonFiluet_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void button1_Click(object sender, RibbonControlEventArgs e)
		{			
			ThisAddIn.ImportDPD();
		}

		private void button2_Click(object sender, RibbonControlEventArgs e)
		{
			ThisAddIn.ImportPKP();
		}

		private void button3_Click(object sender, RibbonControlEventArgs e)
		{
			ThisAddIn.ImportSPSR();
		}

		private void button4_Click(object sender, RibbonControlEventArgs e)
		{
			ThisAddIn.ShipToTransCost();
		}

		private void button5_Click(object sender, RibbonControlEventArgs e)
		{
			ThisAddIn.PostImport();
		}
	}
}
