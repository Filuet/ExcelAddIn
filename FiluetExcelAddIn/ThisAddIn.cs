using System;
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
        private static string connStr =  Properties.Resources.ConnectionString; //"Server=RU-LOB-WMS01;Database=ExchangeDB;User Id=ExUser;Password=good4you";
        private static bool isImportPKP_Running = false;
        private static bool isImportSPSR_Running = false;
        private static bool isShipToTransCost_Running = false;
        private static bool isPostImport_Running = false;
        private static bool isOrderBulk_LvOrclLotMismatch_Running = false;
        private static bool isReceiptBulk_LvOrclLotMismatch_Running = false;
        private static bool isStockBulk_LvOrclLotMismatch_Running = false;

        public static ProgressForm formProgress = new ProgressForm();
        public static DatePeriod PostDatePick = new DatePeriod();
        public static Excel.Worksheet ws;

        public static string OrderCode = null;
        public static string fileName = null;
        public static int RU1B10 = 0;
        public static int RU1B68 = 0;
        public static int RU1B80 = 0;
        public static int RUCB20 = 0;
        public static int RUCB80 = 0;



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
            ws = wb.Worksheets[wb.Worksheets.Count];

            Clipboard.Clear();
            wbT.Close();
            File.Delete(tmpFile);

            ThreadStart bts = new ThreadStart(StartImport);
            Thread bt = new Thread(bts);
            bt.Start();
            formProgress.ShowDialog();

            Cursor.Current = Cursors.Default;
            isPostImport_Running = false;
            Clipboard.Clear();
        }

        private static void StartImport()
        {
            formProgress.CloseButtonEnable(false);
            bool r;
            DataTable dt = new DataTable();
            r = PostImport_01(ref dt);
            if (r)
            {
                formProgress.SetLog(string.Format("Всего отправлений: {0}", dt.Rows.Count));
                r = PostImport_02(ws, dt);
                if (r)
                {
                    formProgress.SetLog("Done");
                }
            }

            formProgress.CloseButtonEnable(true);
        }

        private static bool PostImport_01(ref DataTable dt)
        {
            formProgress.SetTitle("Импорт данных из ЛВижн");
            bool res = true;

            dt = new DataTable();
            string sql = string.Format("EXEC FiluetWH.RuPost.GetPost_XL @StartDate = N'{0:yyyy-MM-dd}', @EndDate = N'{1:yyyy-MM-dd}'", PostDatePick.DateStart, PostDatePick.DateEnd);
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
                DataView dv = dt.DefaultView;
                dv.Sort = "InputDate, OrderCode";
                dt = dv.ToTable();
                formProgress.SetProgress(100, 0, 100);
            }
            catch (Exception ex)
            {
                formProgress.SetLog("Error: " + ex.Message);
                res = false;
            }
            return res;
        }

        private static bool PostImport_02(Excel.Worksheet ws, DataTable dt)
        {
            int row = 2;
            bool res = true;
            RuPostQueryAPI api = new RuPostQueryAPI();
            try
            {
                Excel.Range rng = ws.Rows[row].EntireRow;

                foreach (DataRow drow in dt.Rows)
                {
                    rng = ws.Rows[row].EntireRow;
                    rng.Copy(ws.Rows[row + 1]);

                    ws.Cells[row, 1].Value2 = drow["InputDate"];
                    ws.Cells[row, 2].Value2 = drow["ExecuteDate"];
                    ws.Cells[row, 3].Value2 = drow["OrderCode"];
                    ws.Cells[row, 5].Value2 = drow["Post_index"];
                    ws.Cells[row, 6].Value2 = drow["Post_region"];
                    ws.Cells[row, 7].Value2 = drow["Post_place"];
                    int boxid = 0;
                    bool rr = int.TryParse(drow["boxPostID"].ToString(), out boxid);
                    if (boxid!=0)
                    {
                        PostSearch ps = api.SearchRPO(drow["boxPostID"].ToString());
                        if (ps != null)
                        {
                            ws.Cells[row, 4].Value2 = ps.Barcode;
                            ws.Cells[row, 12].Value2 = ps.HumanOperationName;
                            ws.Cells[row, 11].Value2 = ps.LastOperDate;
                            ws.Cells[row, 9].Value2 = ps.TotalRateWoVat / 100;
                        }
                    }
                    row++;
                    formProgress.SetProgress(row - 2, 0, dt.Rows.Count);
                }
                rng = ws.Rows[row].EntireRow;
                rng.Delete();
            }
            catch (Exception ex)
            {
                formProgress.SetLog("Error: " + ex.Message);
                res = false;
            }
            return res;
        }


        #endregion

        #region LvOrclStock_ReceiptBulkMismatch
        //Сравнение запасов по партии, товару и количеству в LV и Oracle
        public static void LvOrclStock_ReceiptBulkMismatch_Report()
        {
            string ErrMsg = null;

            if (isReceiptBulk_LvOrclLotMismatch_Running)
            {
                MessageBox.Show("Обработка уже идет!");
                return;
            }

            isReceiptBulk_LvOrclLotMismatch_Running = true;
            Cursor.Current = Cursors.WaitCursor;

            frmTextBoxParameter form = new frmTextBoxParameter("Приход:");
            form.ShowDialog();
            if (OrderCode == null)
            {
                isReceiptBulk_LvOrclLotMismatch_Running = false;
                return;
            }

            //Проверка условий выпуска отчета
            DataTable dt = new DataTable();
            string sql = string.Format("select [ExchangeDB].[dbo].[FIL_HRBL_LvOrclStock_ReceiptBulkCompleteCheck]('{0}')", OrderCode);
            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn, CommandText = sql, CommandTimeout = 120 })
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
                isReceiptBulk_LvOrclLotMismatch_Running = false;
                return;
            }

            if (dt.Rows.Count != 0)
            {
                ErrMsg = dt.Rows[0][0].ToString();
                if (!String.IsNullOrEmpty(ErrMsg))
                {
                    MessageBox.Show(ErrMsg);
                    isReceiptBulk_LvOrclLotMismatch_Running = false;
                    return;
                }
            }

            //Вставка в докмуент листа Excel из ресурсов
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string tmpFile = Path.GetTempFileName();
            File.WriteAllBytes(tmpFile, Properties.Resources.LvOrclStock);
            Excel.Workbook wbT = Globals.ThisAddIn.Application.Workbooks.Add(tmpFile);
            Excel.Worksheet wsT = wbT.Worksheets["Контроль прихода"];

            wsT.Copy(After: wb.Worksheets[wb.Worksheets.Count]);
            Excel.Worksheet ws = wb.Worksheets[wb.Worksheets.Count];

            Clipboard.Clear();
            wbT.Close();
            File.Delete(tmpFile);

            //Получаем данные для отчета
            DataSet ds = null;
            ds = LvOrclStock_ReceiptBulkMismatch_GetData(OrderCode);
            //Грузим их в Excel
            LvOrclStock_ReceiptBulkMismatch_FillExcel(ws, ds);

            Cursor.Current = Cursors.Default;
            isReceiptBulk_LvOrclLotMismatch_Running = false;
            Clipboard.Clear();
        }

        private static DataSet LvOrclStock_ReceiptBulkMismatch_GetData(string OrderCode)
        {
            DataSet ds = new DataSet();
            string sql = string.Format("exec [dbo].[FIL_HRBL_LvOrclStock_ReceiptBulkMismatch] @ReceiptCode = '{0}'", OrderCode);
            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn, CommandText = sql, CommandTimeout = 120 })
                    {
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(ds);
                        }
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
                isReceiptBulk_LvOrclLotMismatch_Running = false;
                return null;
            }
            return ds;
        }

        private static void LvOrclStock_ReceiptBulkMismatch_FillExcel(Excel.Worksheet ws, DataSet ds)
        {
            if (ds == null || ds.Tables.Count != 2)
            {
                MessageBox.Show("Нет данных");
                return;
            }

            DataTable dtHeader = ds.Tables[0];
            DataTable dtDetail = ds.Tables[1];

            ws.Cells[1, 2].Value2 = dtHeader.Rows[0]["OrderNo"];
            ws.Cells[2, 2].Value2 = dtHeader.Rows[0]["InputDate"];
            ws.Cells[3, 2].Value2 = dtHeader.Rows[0]["ReceiptDate"];

            if (ds.Tables[1].Rows.Count == 0)
            {
                MessageBox.Show("Расхождений не обнаружено");
                return;
            }

            int row = 6;
            Excel.Range rngDetail = ws.Rows[row].EntireRow;

            foreach (DataRow drow in dtDetail.Rows)
            {
                rngDetail = ws.Rows[row].EntireRow;
                rngDetail.Copy(ws.Rows[row + 1]);

                ws.Cells[row, 1].Value2 = drow["StockCode"];
                ws.Cells[row, 2].Value2 = drow["SellCode"];
                ws.Cells[row, 3].Value2 = drow["Lot"];
                ws.Cells[row, 4].Value2 = drow["Qty"];
                ws.Cells[row, 5].Value2 = drow["System"];
                if (drow["System"].ToString().Equals("Oracle"))
                    ws.Rows[row].EntireRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.MintCream);
                row++;
            }
            rngDetail = ws.Rows[row].EntireRow;
            rngDetail.Delete();
        }
        #endregion
        #region LvOrclStock_OrderBulkMismatch
        //Сравнение запасов по партии, товару и количеству в LV и Oracle
        public static void LvOrclStock_OrderBulkMismatch_Report()
        {
            string ErrMsg = null;

            if (isOrderBulk_LvOrclLotMismatch_Running)
            {
                MessageBox.Show("Обработка уже идет!");
                return;
            }

            isOrderBulk_LvOrclLotMismatch_Running = true;
            Cursor.Current = Cursors.WaitCursor;

            frmTextBoxParameter form = new frmTextBoxParameter("Расход:");
            form.ShowDialog();
            if (OrderCode == null)
            {
                isOrderBulk_LvOrclLotMismatch_Running = false;
                return;
            }

            //Проверка условий выпуска отчета
            DataTable dt = new DataTable();
            string sql = string.Format("select [FiluetWH].[dbo].[FIL_HRBL_LvOrclStock_OrderBulkAllocatedCheck]('{0}')", OrderCode);
            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn, CommandText = sql, CommandTimeout = 120 })
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
                isOrderBulk_LvOrclLotMismatch_Running = false;
                return;
            }

            if (dt.Rows.Count != 0)
            {
                ErrMsg = dt.Rows[0][0].ToString();
                if (!String.IsNullOrEmpty(ErrMsg))
                {
                    MessageBox.Show(ErrMsg);
                    isOrderBulk_LvOrclLotMismatch_Running = false;
                    return;
                }
            }

            //Вставка в докмуент листа Excel из ресурсов
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string tmpFile = Path.GetTempFileName();
            File.WriteAllBytes(tmpFile, Properties.Resources.LvOrclStock);
            Excel.Workbook wbT = Globals.ThisAddIn.Application.Workbooks.Add(tmpFile);
            Excel.Worksheet wsT = wbT.Worksheets["Контроль расхода"];

            wsT.Copy(After: wb.Worksheets[wb.Worksheets.Count]);
            Excel.Worksheet ws = wb.Worksheets[wb.Worksheets.Count];

            Clipboard.Clear();
            wbT.Close();
            File.Delete(tmpFile);

            //Получаем данные для отчета
            DataSet ds = null;
            ds = LvOrclStock_BulkOrderMismatch_GetData(OrderCode);
            //Грузим их в Excel
            LvOrclStock_BulkOrderMismatch_FillExcel(ws, ds);

            Cursor.Current = Cursors.Default;
            isOrderBulk_LvOrclLotMismatch_Running = false;
            Clipboard.Clear();
        }

        private static DataSet LvOrclStock_BulkOrderMismatch_GetData(string OrderCode)
        {
            DataSet ds = new DataSet();
            string sql = string.Format("exec [dbo].[FIL_HRBL_LvOrclStock_OrderBulkMismatch] @OrderCode = '{0}'", OrderCode);
            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn, CommandText = sql, CommandTimeout = 120 })
                    {
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(ds);
                        }
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
                isOrderBulk_LvOrclLotMismatch_Running = false;
                return null;
            }
            return ds;
        }

        private static void LvOrclStock_BulkOrderMismatch_FillExcel(Excel.Worksheet ws, DataSet ds)
        {
            if (ds == null || ds.Tables.Count != 2)
            {
                MessageBox.Show("Нет данных");
                return;
            }

            DataTable dtHeader = ds.Tables[0];
            DataTable dtDetail = ds.Tables[1];

            ws.Cells[1, 2].Value2 = dtHeader.Rows[0]["OrderNo"];
            ws.Cells[2, 2].Value2 = dtHeader.Rows[0]["InputDate"];
            ws.Cells[3, 2].Value2 = dtHeader.Rows[0]["ShipDate"];

            if (ds.Tables[1].Rows.Count == 0)
            {
                MessageBox.Show("Расхождений не обнаружено");
                return;
            }

            int row = 6;
            Excel.Range rngDetail = ws.Rows[row].EntireRow;

            foreach (DataRow drow in dtDetail.Rows)
            {
                rngDetail = ws.Rows[row].EntireRow;
                rngDetail.Copy(ws.Rows[row + 1]);

                ws.Cells[row, 1].Value2 = drow["StockCode"];
                ws.Cells[row, 2].Value2 = drow["SellCode"];
                ws.Cells[row, 3].Value2 = drow["Lot"];
                ws.Cells[row, 4].Value2 = drow["Qty"];
                ws.Cells[row, 5].Value2 = drow["System"];
                if (drow["System"].ToString().Equals("Oracle"))
                    ws.Rows[row].EntireRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.MintCream);
                row++;
            }
            rngDetail = ws.Rows[row].EntireRow;
            rngDetail.Delete();
        }
        #endregion LvOrclStock
        #region LvOrclStock_StockMismatch
        /// <summary>
        /// Производит запрос параметров 
        /// Вывод в Excel
        /// отчета "Сверка остатков" 
        /// </summary>
        public static void LvOrclStock_StockBulkMismatch_Report()
        {
            //1.Check if process is already started
            string ErrMsg = null;

            if (isStockBulk_LvOrclLotMismatch_Running)
            {
                MessageBox.Show("Обработка уже идет!");
                return;
            }

            isStockBulk_LvOrclLotMismatch_Running = true;
            Cursor.Current = Cursors.WaitCursor;

            frmFileParameter form = new frmFileParameter();
            form.ShowDialog();
            if (fileName == null)
            {
                isStockBulk_LvOrclLotMismatch_Running = false;
                return;
            }

            //Чтение разбор файла, загрузка данных в базу
            LvOrclStock_StockBulkMismatch_ParseFile();

            //Вставка в докмуент листа Excel из ресурсов
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string tmpFile = Path.GetTempFileName();
            File.WriteAllBytes(tmpFile, Properties.Resources.LvOrclStock);
            Excel.Workbook wbT = Globals.ThisAddIn.Application.Workbooks.Add(tmpFile);
            Excel.Worksheet wsT = wbT.Worksheets["Сверка остатков"];

            wsT.Copy(After: wb.Worksheets[wb.Worksheets.Count]);
            Excel.Worksheet ws = wb.Worksheets[wb.Worksheets.Count];

            Clipboard.Clear();
            wbT.Close();
            File.Delete(tmpFile);

            //Получаем данные для отчета
            DataSet ds = null;
            ds = LvOrclStock_StockBulkMismatch_GetData();
            //Грузим их в Excel
            LvOrclStock_StockBulkMismatch_FillExcel(ws, ds);

            Cursor.Current = Cursors.Default;
            isStockBulk_LvOrclLotMismatch_Running = false;
            Clipboard.Clear();
        }

        /// <summary>
        /// Чтение и разбор текстового файла и загрузка данных в базу
        /// </summary>
        private static bool LvOrclStock_StockBulkMismatch_ParseFile()
        {
            string s = null;
            string sExpDate = null;
            string sQty = null;
            string Sku = null;
            string SkuNew = null;
            string Lot = null;
            string sql = null;
            string Locator = null;
            string LocatorNew = null;

            int ReportPage = 0;
            int res = 0; //Номер страницы отчеты

            DateTime ReportDateTime = DateTime.MinValue;
            DateTime ExpDate = DateTime.MinValue; //Дата формирования отчета

            float Qty = 0;

            //Очистка данных в базе
            sql = "EXEC [FiluetWH].[dbo].[FIL_HRBL_LvOrclStock_StockOrclDelete]";

            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn, CommandText = sql, CommandTimeout = 120 })
                    {
                        res = cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
                return false;
            }


            //Читаем файл отчета
            StreamReader sr = new StreamReader(fileName);

            s = sr.ReadLine();

            while (!sr.EndOfStream)
            {
                //Если с текущей строки начинается страница
                if (s.Contains("HL Quantity Onhand Report"))
                {
                    if (!DateTime.TryParse(s.Substring(212).Trim(), CultureInfo.InvariantCulture.DateTimeFormat, System.Globalization.DateTimeStyles.None, out ReportDateTime))
                        return false;
                }

                //Если строка содержит номер страницы
                else if (s.Contains("Page:"))
                {
                    try
                    {
                        ReportPage = int.Parse(s.Substring(212).Trim());
                    }
                    catch//System.IndexOutOfRangeException, System.FormatException
                    {
                        return false;
                    }

                    //Пропускаем строки до таблицы
                    for (int i = 1; i <= 9; i++)
                        s = sr.ReadLine();

                }  //Исключаем дополнительные строки с переносом данных
                else if (s.Length > 193 && !String.IsNullOrWhiteSpace(sQty = s.Substring(182, 10).Replace(",", "")))
                {
                    //Получаем количество
                    if (!float.TryParse(sQty, System.Globalization.NumberStyles.Float, CultureInfo.InvariantCulture.NumberFormat, out Qty))
                        return false;

                    //...Артикул
                    SkuNew = s.Substring(4, 12).Trim();
                    Sku = String.IsNullOrWhiteSpace(SkuNew) ? Sku : SkuNew;
                    if (String.IsNullOrWhiteSpace(Sku))
                        return false;

                    //...Locator
                    LocatorNew = s.Substring(94, 6).Trim();
                    Locator = String.IsNullOrWhiteSpace(LocatorNew) ? Locator : LocatorNew;
                    if (String.IsNullOrWhiteSpace(Locator))
                        return false;

                    //...Lot
                    Lot = s.Substring(149, 10).Split('-')[0].Trim();
                    Lot = String.IsNullOrEmpty(Lot) ? "NULL" : "'" + Lot + "'";

                    //...ExpDate

                    if (DateTime.TryParse(s.Substring(161, 9).Trim(), CultureInfo.InvariantCulture.DateTimeFormat, System.Globalization.DateTimeStyles.None, out ExpDate))
                        sExpDate = "'" + ExpDate.ToString("yyyyMMdd HH:mm:ss.000") + "'";
                    else
                        sExpDate = "NULL";

                    //Запись в базу
                    sql = string.Format("EXEC [FiluetWH].[dbo].[FIL_HRBL_LvOrclStock_StockOrclItemInsert] @Sku = '{0}', @Lot = {1}, @ExpDate = {2}, @Qty = {3}, @ReportDateTime = '{4}', @ReportPage = {5}, @Locator = '{6}'",
                    Sku, Lot, sExpDate, Qty, ReportDateTime.ToString("yyyyMMdd HH:mm:ss.000"), ReportPage, Locator);

                    try
                    {
                        using (SqlConnection conn = new SqlConnection(connStr))
                        {
                            conn.Open();
                            using (SqlCommand cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn, CommandText = sql, CommandTimeout = 120 })
                            {
                                res = cmd.ExecuteNonQuery();
                            }
                            conn.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                        return false;
                    }
                }

                //Считываем следующую строку
                s = sr.ReadLine();

            };
            return true;
        }

        private static DataSet LvOrclStock_StockBulkMismatch_GetData()
        {
            DataSet ds = new DataSet();
            string sql = string.Format("exec [FiluetWH].[dbo].[FIL_HRBL_LvOrclStock_StockBulkMismatch] @LocatorList = '{0}'", "#" + RU1B10.ToString() + "#" + RU1B80.ToString() + "#" + RUCB20.ToString() + "#" + RUCB80.ToString() + "#");

            try
            {
                using (SqlConnection conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn, CommandText = sql, CommandTimeout = 120 })
                    {
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(ds);
                        }
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
                isReceiptBulk_LvOrclLotMismatch_Running = false;
                return null;
            }
            return ds;
        }

        private static void LvOrclStock_StockBulkMismatch_FillExcel(Excel.Worksheet ws, DataSet ds)
        {


            if (ds == null || ds.Tables.Count != 2)
            {
                MessageBox.Show("Нет данных");
                return;
            }

            DataTable dtHeader = ds.Tables[0];
            DataTable dtDetail = ds.Tables[1];

            ws.Cells[1, 2].Value2 = dtHeader.Rows[0]["StockOrclDateTime"];
            ws.Cells[2, 2].Value2 = dtHeader.Rows[0]["Today"];

            if (ds.Tables[1].Rows.Count == 0)
            {
                MessageBox.Show("Расхождений не обнаружено");
                return;
            }

            int row = 5;
            Excel.Range rngDetail = ws.Rows[row].EntireRow;

            foreach (DataRow drow in dtDetail.Rows)
            {
                rngDetail = ws.Rows[row].EntireRow;
                rngDetail.Copy(ws.Rows[row + 1]);

                ws.Cells[row, 1].Value2 = drow["LVLocations"];
                ws.Cells[row, 2].Value2 = drow["OrclLocator"];
                ws.Cells[row, 3].Value2 = drow["StockCode"];
                ws.Cells[row, 4].Value2 = drow["SellCode"];
                ws.Cells[row, 5].Value2 = drow["Lot"];
                ws.Cells[row, 6].Value2 = drow["ExpDate"];
                ws.Cells[row, 7].Value2 = drow["OrclQty"];
                ws.Cells[row, 8].Value2 = drow["LvQty"];

                row++;
            }
            rngDetail = ws.Rows[row].EntireRow;
            rngDetail.Delete();
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
