﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Globalization;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Threading;


namespace FiluetExcelAddIn
{
	public partial class ThisAddIn
	{
		private const string connStr = "Server=RU-LOB-WMS01;Database=ExchangeDB;User Id=ExUser;Password=good4you";
		private static bool isImportPKP_Running = false;
		private static bool isImportSPSR_Running = false;
		private static bool isShipToTransCost_Running = false;
		private static bool isPostImport_Running = false;
		private static ProgressForm formProgress = new ProgressForm();
		public static DatePeriod PostDatePick = new DatePeriod();

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{

		}

		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
		}

		#region Import SPSR

		public static void ImportSPSR()
		{
			if (isImportSPSR_Running)
			{
				MessageBox.Show("Импорт уже запущен!");
				return;
			}

			formProgress.SetTitle("Импорт Счета СПСР");
			formProgress.CloseButtonEnable(false);
			ThreadStart bts = new ThreadStart(ImportSPSR_Step01);
			Thread bt = new Thread(bts);
			bt.Start();
			formProgress.ShowDialog();
			isImportSPSR_Running = false;
		}

		private static void ImportSPSR_Step01()
		{
			Cursor.Current = Cursors.WaitCursor;

			isImportSPSR_Running = true;

			DateTime d = DateTime.Now;
			CultureInfo provider = CultureInfo.InvariantCulture;
			Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
			Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
			Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
			int lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;
			string log = String.Format("Старт...\r\nВсего строк для обработки: {0}\r\n", lastR - 7);
			int err = 0;
			for (int curR = 7; curR <= lastR; curR++)
			{
				try
				{
					Invoice inv = new Invoice();
					inv.BoxQty = 1;
					inv.Date = DateTime.FromOADate(ws.Cells[curR, 1].Value2);
					inv.InvoiceNo = ws.Cells[curR, 2].Text.Trim();
					inv.OrderCode = ws.Cells[curR, 3].Text.Trim();
					inv.CityFrom = ws.Cells[curR, 4].Text.Trim().ToUpper();
					inv.CityTo = ws.Cells[curR, 5].Text.Trim().ToUpper();
					inv.Weight = Convert.ToDecimal(ws.Cells[curR, 6].Value2);
					inv.VolumeWeight = Convert.ToDecimal(ws.Cells[curR, 7].Value2);
					inv.Amount = Convert.ToDecimal(ws.Cells[curR, 9].Value2 * 1.18);
					InsertInvoice(inv, "SPSR");
				}
				catch
				{
					err++;
					log += String.Format("Ошибка импорта строки {0}\r\n", curR);
				}
				if (formProgress.InvokeRequired)
				{
					formProgress.BeginInvoke(
						new System.Action(() =>
						{
							formProgress.SetProgress(curR, 7, lastR);
							formProgress.SetLog(log);
						}
					));
				}

			}
			DateTime d2 = DateTime.Now;
			log += String.Format("Импорт завершен.\r\nВремя обработки {0:0} секунд.\r\nКоличество ошибок: {1}", (d2 - d).TotalSeconds, err);
			if (formProgress.InvokeRequired)
			{
				formProgress.BeginInvoke(
					new System.Action(() =>
					{
						formProgress.CloseButtonEnable(true);
						formProgress.SetLog(log);
					}
							));
			}
			Cursor.Current = Cursors.Default;
		}

		#endregion
		#region Import PKP

		public static void ImportPKP()
		{
			if (isImportPKP_Running)
			{
				MessageBox.Show("Импорт уже запущен!");
				return;
			}

			formProgress.SetTitle("Импорт Счета Пик Поинт");
			formProgress.CloseButtonEnable(false);
			ThreadStart bts = new ThreadStart(ImportPKP_Step01);
			Thread bt = new Thread(bts);
			bt.Start();
			formProgress.ShowDialog();
			isImportPKP_Running = false;
		}

		private static void ImportPKP_Step01()
		{
			Cursor.Current = Cursors.WaitCursor;

			isImportPKP_Running = true;

			DateTime d = DateTime.Now;
			CultureInfo provider = CultureInfo.InvariantCulture;
			Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
			Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
			Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
			int lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;
			string log = String.Format("Старт...\r\nВсего строк для обработки: {0}\r\n", lastR - 7);
			int err = 0;
			for (int curR = 7; curR <= lastR; curR++)
			{
				try
				{
					Invoice inv = new Invoice();
					inv.BoxQty = 1;
					inv.Date = DateTime.FromOADate(ws.Cells[curR, 1].Value2);
					inv.InvoiceNo = ws.Cells[curR, 2].Text.Trim();
					inv.OrderCode = ws.Cells[curR, 3].Text.Trim();
					inv.CityFrom = ws.Cells[curR, 4].Text.Trim().ToUpper();
					inv.CityTo = ws.Cells[curR, 5].Text.Trim().ToUpper();
					inv.Weight = Convert.ToDecimal(ws.Cells[curR, 7].Value2);
					inv.Amount = Convert.ToDecimal(ws.Cells[curR, 9].Value2 * 1.18);
					InsertInvoice(inv, "PKP");
				}
				catch
				{
					err++;
					log += String.Format("Ошибка импорта строки {0}\r\n", curR);
				}
				if (formProgress.InvokeRequired)
				{
					formProgress.BeginInvoke(
						new System.Action(() =>
							{
								formProgress.SetProgress(curR, 7, lastR);
								formProgress.SetLog(log);
							}
					));
				}

			}
			DateTime d2 = DateTime.Now;
			log += String.Format("Импорт завершен.\r\nВремя обработки {0:0} секунд.Количество ошибок: {1}", (d2 - d).TotalSeconds, err);
			if (formProgress.InvokeRequired)
			{
				formProgress.BeginInvoke(
					new System.Action(() =>
						{
							formProgress.CloseButtonEnable(true);
							formProgress.SetLog(log);
						}
							));
			}
			Cursor.Current = Cursors.Default;
		}

		#endregion
		#region Import DPD

		public static void ImportDPD()
		{
			//try
			//{
			CultureInfo provider = CultureInfo.InvariantCulture;
			Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
			Excel.Worksheet ws = wb.Worksheets["attachment"];
			int firstR = ws.UsedRange.Row;
			int lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;

			List<InvoiceDPD> Invoices = new List<InvoiceDPD>();
			for (int curR = ws.UsedRange.Row; curR <= lastR; curR++)
			{
				var cell = ws.Cells[curR, 2];
				if (cell != null)
				{
					string t = cell.Text;
					if (t.Contains("Отправка"))
					{
						InvoiceDPD i = new InvoiceDPD();
						string c8 = ws.Cells[curR, 8].Text;
						string[] t1 = c8.Split('-');
						i.InvoiceNo = t1[0];
						i.OrderNo = t1[1];
						i.OrderDate = DateTime.ParseExact(t1[2], "dd.MM.yyyy", provider);
						string br = t1[4];

						i.Branch = t1[4].Contains("Санкт") ? "Санкт-Петербург" : t1[4];

						string c3 = ws.Cells[curR, 3].Text;
						int.TryParse(c3, out i.BoxQty);

						string c4 = ws.Cells[curR, 4].Text;
						decimal.TryParse(c4, out i.Weight);

						string c5 = ws.Cells[curR, 5].Text;
						decimal.TryParse(c5, out i.DeliveryCost);

						string c7 = ws.Cells[curR, 7].Text;
						decimal.TryParse(c7, out i.DeliveryCostVAT);

						//System.Diagnostics.Debug.WriteLine(i);
						Invoices.Add(i);
					}
				}
			}

			for (int curR = ws.UsedRange.Row; curR <= lastR; curR++)
			{
				var cell = ws.Cells[curR, 2];
				if (cell != null)
				{
					string t = cell.Text;
					if (t.Contains("Прием"))
					{
						decimal cost = 0;
						string c5 = ws.Cells[curR, 5].Text;
						decimal.TryParse(c5, out cost);

						decimal costVat = 0;
						string c7 = ws.Cells[curR, 7].Text;
						decimal.TryParse(c7, out costVat);

						List<string> pick = new List<string>();
						//string c8 = ws.Cells[curR, 8].Text;
						//string[] t1 = c8.Split('-');
						//pick.Add(t1[0].Trim());
						string c82 = "";
						while (string.IsNullOrEmpty(c82))
						{
							curR++;
							string c22 = ws.Cells[curR, 2].Text;
							if (c22.Contains("Прием") | c22.Contains("Отправка"))
								break;
							else
								c82 = ws.Cells[curR, 8].Text;
						}
						System.Diagnostics.Debug.WriteLine(curR.ToString() + " - " + c82);
						string[] t2 = c82.Split(',');
						for (int x = 0; x < t2.Length; x++)
							pick.Add(t2[x].Trim());

						cost = cost / pick.Count;
						costVat = costVat / pick.Count;

						foreach (string p in pick)
						{
							InvoiceDPD i = Invoices.Find(x => x.InvoiceNo.Contains(p));
							if (i != null)
							{
								i.PickCost += cost;
								i.PickCostVAT += costVat;
							}
						}
					}
				}
			}

			CreateWS(wb, Invoices);

			//}
			//catch { }
		}

		private static void CreateWS(Excel.Workbook wb, List<InvoiceDPD> Invoices)
		{
			string tmpFile = Path.GetTempFileName();
			File.WriteAllBytes(tmpFile, Properties.Resources.DPD);
			Excel.Workbook wbT = Globals.ThisAddIn.Application.Workbooks.Add(tmpFile);
			Excel.Worksheet wsT = wbT.Worksheets["Отчет Филуэт"];

			int totalSheets = wb.Worksheets.Count;
			wsT.Copy(After: wb.Worksheets[totalSheets]);
			Excel.Worksheet ws = wb.Worksheets[wsT.Name];

			wbT.Close();
			File.Delete(tmpFile);



			List<InvoiceDPD> _invoices = new List<InvoiceDPD>();
			_invoices = Invoices.OrderBy(i => i.Branch).ThenBy(i => i.OrderDate).ToList();

			int row = 2;
			Excel.Range rng = ws.Rows[row].EntireRow;

			foreach (InvoiceDPD inv in _invoices)
			{
				rng = ws.Rows[row].EntireRow;
				rng.Copy(ws.Rows[row + 1]);
				//rng.Insert(Excel.XlInsertShiftDirection.xlShiftDown,true);
				ws.Cells[row, 1].Value2 = inv.Branch;
				ws.Cells[row, 2].Value2 = inv.InvoiceNo;
				ws.Cells[row, 3].Value2 = inv.OrderNo;
				ws.Cells[row, 4].Value2 = inv.OrderDate;
				ws.Cells[row, 5].Value2 = inv.BoxQty;
				ws.Cells[row, 6].Value2 = inv.Weight;
				ws.Cells[row, 7].Value2 = inv.DeliveryCost;
				ws.Cells[row, 8].Value2 = inv.DeliveryCostVAT;
				ws.Cells[row, 9].Value2 = inv.PickCost;
				ws.Cells[row, 10].Value2 = inv.PickCostVAT;
				ws.Cells[row, 11].Value2 = inv.DeliveryCost + inv.PickCost;
				ws.Cells[row, 12].Value2 = inv.DeliveryCostVAT + inv.PickCostVAT;
				row++;
			}
			rng = ws.Rows[row].EntireRow;
			rng.Delete();
		}

		#endregion
		#region ShipTo Transfer Cost

		public static void ShipToTransCost()
		{
			if (isShipToTransCost_Running)
			{
				MessageBox.Show("Обработка уже работает!");
				return;
			}

			formProgress.SetTitle("Добавить стоимость доставки ShipTo");
			formProgress.CloseButtonEnable(false);
			ThreadStart bts = new ThreadStart(ShipToTransCost_01);
			Thread bt = new Thread(bts);
			bt.Start();
			formProgress.ShowDialog();
			isShipToTransCost_Running = false;
		}

		private static void ShipToTransCost_01()
		{
			Cursor.Current = Cursors.WaitCursor;

			isShipToTransCost_Running = true;

			DateTime d = DateTime.Now;
			CultureInfo provider = CultureInfo.InvariantCulture;
			Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
			Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
			Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
			int lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;
			string log = String.Format("Старт...\r\nВсего строк для обработки: {0}\r\n", lastR - 7);
			int err = 0;
			try
			{
				using (SqlConnection conn = new SqlConnection(connStr))
				{
					conn.Open();
					for (int curR = 10; curR <= lastR; curR++)
					{
						try
						{
							string sql = string.Format("EXEC [dbo].[ShipTo_GetTransportCost] @Code = N'{0}'", ws.Cells[curR, 1].Text.Trim());
							using (SqlCommand cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn, CommandText = sql })
							{
								var r = cmd.ExecuteScalar();
								decimal res = 0;
								if (decimal.TryParse(r.ToString(), out res))
									ws.Cells[curR, 17].Value = res;
								else
									ws.Cells[curR, 17].Value = "";
							}
						}
						catch
						{
							err++;
							log += String.Format("Ошибка импорта строки {0}\r\n", curR);
						}
						if (formProgress.InvokeRequired)
						{
							formProgress.BeginInvoke(
								new System.Action(() =>
								{
									formProgress.SetProgress(curR, 7, lastR);
									formProgress.SetLog(log);
								}
							));
						}
					}
					conn.Close();
				}
			}
			catch (Exception ex)
			{
				err++;
				log += String.Format("Ошибка подключения к БД\r\n {0}\r\n", ex.Message);
			}

			DateTime d2 = DateTime.Now;
			log += String.Format("Импорт завершен.\r\nВремя обработки {0:0} секунд.Количество ошибок: {1}", (d2 - d).TotalSeconds, err);
			if (formProgress.InvokeRequired)
			{
				formProgress.BeginInvoke(
					new System.Action(() =>
					{
						formProgress.CloseButtonEnable(true);
						formProgress.SetLog(log);
					}
							));
			}
			Cursor.Current = Cursors.Default;
		}

		#endregion
		#region Post Import

		public static void PostImport()
		{
			if (isPostImport_Running)
			{
				MessageBox.Show("Обработка уже работает!");
				return;
			}

			isPostImport_Running = true;
			Cursor.Current = Cursors.WaitCursor;

			PostDatePick.DateStart = DateTime.Today.AddDays(-7);
			PostDatePick.DateEnd = DateTime.Today;
			PostDatePickForm form = new PostDatePickForm();
			form.ShowDialog();
			if (PostDatePick == null)
				return;

			Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
			string tmpFile = Path.GetTempFileName();
			File.WriteAllBytes(tmpFile, Properties.Resources.Post);
			Excel.Workbook wbT = Globals.ThisAddIn.Application.Workbooks.Add(tmpFile);
			Excel.Worksheet wsT = wbT.Worksheets["Отчет Почта"];

			wsT.Copy(After: wb.Worksheets[wb.Worksheets.Count]);
			Excel.Worksheet ws = wb.Worksheets[wb.Worksheets.Count];

			Clipboard.Clear();
			wbT.Close();
			File.Delete(tmpFile);

			DataTable dt = PostImport_01();
			PostImport_02(ws, dt);

			formProgress.SetTitle("Импорт данных РПО");
			formProgress.CloseButtonEnable(false);
			ThreadStart bts = new ThreadStart(PostImport_03);
			Thread bt = new Thread(bts);
			bt.Start();
			formProgress.ShowDialog();
			Cursor.Current = Cursors.Default;
			isPostImport_Running = false;
			Clipboard.Clear();
		}

		private static DataTable PostImport_01()
		{
			DataTable dt = new DataTable();
			string sql = string.Format("EXEC [ExchangeDB].[dbo].[Register_GetPost] @Date1 = N'{0:yyyy-MM-dd}T00:00:00', @Date2 = N'{1:yyyy-MM-dd}T23:59:59'", PostDatePick.DateStart, PostDatePick.DateEnd);
			try
			{
				using (SqlConnection conn = new SqlConnection(connStr))
				{
					conn.Open();
					using (SqlCommand cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn, CommandText = sql })
					{
						SqlDataReader dr = cmd.ExecuteReader();
						dt.Load(dr);
						dr.Close();
					}
					conn.Close();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: " + ex.Message);
			}
			DataView dv = dt.DefaultView;
			dv.Sort = "InputDate, OrderCode";
			dt = dv.ToTable();
			return dt;
		}

		private static void PostImport_02(Excel.Worksheet ws, DataTable dt)
		{
			int row = 2;
			Excel.Range rng = ws.Rows[row].EntireRow;

			foreach (DataRow drow in dt.Rows)
			{
				rng = ws.Rows[row].EntireRow;
				rng.Copy(ws.Rows[row + 1]);

				ws.Cells[row, 1].Value2 = drow["InputDate"];
				ws.Cells[row, 2].Value2 = drow["PostDate"];
				ws.Cells[row, 3].Value2 = drow["OrderCode"];
				ws.Cells[row, 4].Value2 = drow["BarcodePost"];
				ws.Cells[row, 5].Value2 = drow["IndexPost"];
				ws.Cells[row, 6].Value2 = drow["RegionPost"];
				ws.Cells[row, 7].Value2 = drow["CityPost"];
				decimal p1 = 0;
				decimal.TryParse(drow["ValueRate"].ToString(), out p1);
				decimal p2 = 0;
				decimal.TryParse(drow["MassRate"].ToString(), out p2);
				decimal p3 = 0;
				decimal.TryParse(drow["AviaRate"].ToString(), out p3);
				ws.Cells[row, 8].Value2 = p1 + p2 + p3;
				row++;
			}
			rng = ws.Rows[row].EntireRow;
			rng.Delete();
		}

		private static void PostImport_03()
		{
			List<Invoice> Invcs = new List<Invoice>();
			DateTime d = DateTime.Now;
			Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
			Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
			Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;

			int lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;
			string log = String.Format("Старт...\r\nВсего строк для обработки: {0}\r\n", lastR);
			int err = 0;

			for (int curR = 2; curR <= lastR; curR++)
			{
				try
				{
					ws.Cells[curR, 1].Select();
					string orderCode = ws.Cells[curR, 3].Text.Trim();
					string bcode = ws.Cells[curR, 4].Text.Trim();
					if (string.IsNullOrEmpty(bcode))
						throw new Exception();
					PostResponse resp = new PostResponse();
					bool x = false;
					int xq = 5;
					do
					{
						app.StatusBar = string.Format("Строка {0}, Заказ {1}, ШПИ {2}, Попытка загрузки данных {3}", curR, orderCode, bcode, 6 - xq);
						x = RussianPost.GetStatus(bcode, ref resp);
						xq--;
					} while (!(x || (xq == 0)));
					if (!x)
						throw new Exception();

					DateTime statusDate = new DateTime();
					DateTime postDate = new DateTime(2100,1,1);
					string status = "";
					decimal massRate = 0;
					decimal insRate = 0;
					decimal mass = 0;
					string addressDest = "";
					string addressOper = "";

					foreach (PostResponse.clBody.clOperationHistoryData.clHistoryRecord rec in resp.Body.OperationHistoryData.historyRecord)
					{
						if (statusDate < rec.OperationParameters.OperDate)
						{
							statusDate = rec.OperationParameters.OperDate;
							status = string.Format("{0} - {1}", rec.OperationParameters.OperType.Name, rec.OperationParameters.OperAttr.Name);
						}
						if (postDate>rec.OperationParameters.OperDate)
						{
							massRate = rec.FinanceParameters.MassRate;
							insRate = rec.FinanceParameters.InsrRate;
							mass = rec.ItemParameters.Mass;
							postDate = rec.OperationParameters.OperDate;
							addressDest = rec.AddressParameters.DestinationAddress.Description;
							addressOper = rec.AddressParameters.OperationAddress.Description;
						}
					}
					ws.Cells[curR, 9].Value2 = (massRate + insRate) / 100;
					ws.Cells[curR, 11].Value2 = statusDate;
					ws.Cells[curR, 12].Value2 = status;
					
					Invoice inv = new Invoice();
					inv.OrderCode = orderCode;
					inv.Amount = (massRate + insRate) / 100;
					inv.InvoiceNo = bcode;
					inv.Date = postDate;
					inv.Weight = mass;
					inv.BoxQty = 1;
					inv.CityFrom = addressDest;
					inv.CityTo = addressOper;
					Invcs.Add(inv);
				}
				catch
				{
					err++;
					log += String.Format("Ошибка импорта строки {0}\r\n", curR);
				}
				if (formProgress.InvokeRequired)
				{
					formProgress.BeginInvoke(
						new System.Action(() =>
						{
							formProgress.SetProgress(curR, 2, lastR);
							formProgress.SetLog(log);
						}
					));
				}
			}

			List<Invoice> invGr = Invcs.GroupBy(i => i.OrderCode).Select(g => new Invoice
			{
				OrderCode = g.Key,
				Amount = g.Sum(i => i.Amount),
				CityFrom = g.Max(i => i.CityFrom),
				BoxQty = g.Sum(i => i.BoxQty),
				CityTo = g.Max(i => i.CityTo),
				Date = g.Max(i => i.Date),
				Weight = g.Max(i => i.Weight),
				InvoiceNo = g.Key,
				VolumeWeight = 0
			}
			).ToList();

			app.StatusBar = "Загрузка данных в ДБ";
			foreach (Invoice item in invGr)
				InsertInvoice(item, "POST");

			DateTime d2 = DateTime.Now;
			log += String.Format("Импорт завершен.\r\nВремя обработки {0:0} секунд.Количество ошибок: {1}", (d2 - d).TotalSeconds, err);
			if (formProgress.InvokeRequired)
			{
				formProgress.BeginInvoke(
					new System.Action(() =>
					{
						formProgress.CloseButtonEnable(true);
						formProgress.SetLog(log);
					}
							));
			}
			Cursor.Current = Cursors.Default;
			app.StatusBar = false;
		}

		#endregion

		private static void InsertInvoice(Invoice inv, string type)
		{
			string sqlD = String.Format("DELETE FROM CourierBills WHERE (Type = '{1}') AND (InvoiceNo ='{0}')", inv.InvoiceNo, type);
			string sql = "INSERT INTO CourierBills (Type, Date, OrderCode, InvoiceNo, CityFrom, CityTo, Weight, VolumeWeight, Amount, BoxQty) VALUES ";
			sql += string.Format("('{9}','{0:yyyy-MM-dd}','{1}','{2}','{3}','{4}',{5},{6},{7},{8})", inv.Date, inv.OrderCode, inv.InvoiceNo, inv.CityFrom, inv.CityTo, inv.Weight, inv.VolumeWeight, inv.Amount, inv.BoxQty, type);
			try
			{
				using (SqlConnection conn = new SqlConnection(connStr))
				{
					conn.Open();
					using (SqlCommand cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn, CommandText = sqlD })
					{
						cmd.ExecuteNonQuery();
					}
					using (SqlCommand cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn, CommandText = sql })
					{
						cmd.ExecuteNonQuery();
					}
					conn.Close();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: " + ex.Message);
			}
		}

		private static int ExecQuery(string sql)
		{
			int res = 0;
			try
			{
				using (SqlConnection conn = new SqlConnection(connStr))
				{
					conn.Open();
					using (SqlCommand cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn, CommandText = sql })
					{
						res = cmd.ExecuteNonQuery();
					}
					conn.Close();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: " + ex.Message);
			}
			return res;
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new EventHandler(ThisAddIn_Startup);
			this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}

	public class InvoiceDPD
	{
		public string InvoiceNo = "";
		public string OrderNo = "";
		public string Branch = "";
		public DateTime OrderDate = new DateTime();
		public int BoxQty = 0;
		public decimal DeliveryCost = 0;
		public decimal DeliveryCostVAT = 0;
		public decimal PickCost = 0;
		public decimal PickCostVAT = 0;
		public decimal Weight = 0;
	}

	public class Invoice
	{
		public DateTime Date = new DateTime();
		public string InvoiceNo = "";
		public string OrderCode = "";
		public string CityFrom = "";
		public string CityTo = "";
		public int BoxQty = 1;
		public decimal Amount = 0;
		public decimal Weight = 0;
		public decimal VolumeWeight = 0;
	}

	public class DatePeriod
	{
		public DateTime DateStart { get; set; }
		public DateTime DateEnd { get; set; }
	}

}