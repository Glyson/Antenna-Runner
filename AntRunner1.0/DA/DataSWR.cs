using AntRunner.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace AntRunner
{
    public class DataSWR : DataBase
    {
        //SWR Report
        public override void Report()
        {
            Progress = 0;
            int count1, count2, count3, count4;
            int pass1, pass2, pass3, pass4;
            List<SingleData> list1, list2, list3, list4;
            GetSingleData_LOG(Settings.Default.OutputDir, Settings.Default.Para1.Trace,
                out count1, out pass1, out list1);
            GetSingleData_LOG(Settings.Default.OutputDir, Settings.Default.Para2.Trace,
                out count2, out pass2, out list2);
            GetSingleData_LOG(Settings.Default.OutputDir, Settings.Default.Para3.Trace,
                out count3, out pass3, out list3);
            GetSingleData_LOG(Settings.Default.OutputDir, Settings.Default.Para4.Trace,
                out count4, out pass4, out list4);

            StreamWriter writer = null;
            try
            {
                Excel.Application app = new Excel.Application();
                Excel.Workbooks bks = app.Workbooks;
                Excel.Workbook bk = bks.Add(true);
                Excel.Worksheet sh = (Excel.Worksheet)bk.Sheets[1];
                sh.Name = "Report Data";
                ((Excel.Range)sh.Columns[2, Type.Missing]).NumberFormat = "@";
                sh.Columns.ColumnWidth = 12;
                sh.Columns[1, Type.Missing].ColumnWidth = 28;
                int r = 0;
                r++;
                ((Excel.Range)sh.Rows[r, Type.Missing]).Interior.ColorIndex = 37;
                ((Excel.Range)sh.Rows[r, Type.Missing]).Font.Bold = true;
                sh.Cells[r, 1] = "Summary";
                sh.Cells[r, 2] = "Pass/Sum";
                sh.Cells[r, 3] = "Pass Rate";

                sh.Cells[++r, 1] = "Port1(S11)";
                sh.Cells[r, 2] = pass1 + "/" +count1;
                sh.Cells[r, 3] = (count1 == 0 ? 0 : Math.Round(pass1 / (double)count1 * 100, 4)) + "%";

                sh.Cells[++r, 1] = "Port2(S22)";
                sh.Cells[r, 2] = pass2 + "/" + count2;
                sh.Cells[r, 3] = (count2 == 0 ? 0 : Math.Round(pass2 / (double)count2 * 100, 4)) + "%";

                sh.Cells[++r, 1] = "Port3(S33)";
                sh.Cells[r, 2] = pass3 + "/" + count3;
                sh.Cells[r, 3] = (count3 == 0 ? 0 : Math.Round(pass3 / (double)count3 * 100, 4)) + "%";

                sh.Cells[++r, 1] = "Port4(S44)";
                sh.Cells[r, 2] = pass4 + "/" + count4;
                sh.Cells[r, 3] = (count4 == 0 ? 0 : Math.Round(pass4 / (double)count4 * 100, 4)) + "%";

                int pass = pass1 + pass2 + pass3 + pass4;
                int count = count1 + count2 + count3 + count4;
                sh.Cells[++r, 1] = "Total";
                sh.Cells[r, 2] = pass + "/" + count;
                sh.Cells[r, 3] = (count == 0 ? 0 : Math.Round(pass / (double)count * 100, 4)) + "%";

                //DUT information
                r++;
                r++;
                ((Excel.Range)sh.Rows[r, Type.Missing]).Interior.ColorIndex = 37;
                ((Excel.Range)sh.Rows[r, Type.Missing]).Font.Bold = true;
                sh.Cells[r, 1] = "DUT Information";
                r++;
                sh.Cells[r, 1] = "Code";
                sh.Cells[r, 2] = Settings.Default.Code;
                r++;
                sh.Cells[r, 1] = "Manufacturer";
                sh.Cells[r, 2] = Settings.Default.Manufacture;
                r++;
                sh.Cells[r, 1] = "Memo";
                sh.Cells[r, 2] = Settings.Default.Memo;

                //Parameters
                r++;
                r++;
                ((Excel.Range)sh.Rows[r, Type.Missing]).Interior.ColorIndex = 37;
                ((Excel.Range)sh.Rows[r, Type.Missing]).Font.Bold = true;
                sh.Cells[r, 1] = "Parameters";
                sh.Cells[r, 2] = "S11";
                sh.Cells[r, 3] = "S22";
                sh.Cells[r, 4] = "S33";
                sh.Cells[r, 5] = "S44";
                r++;
                sh.Cells[r, 1] = "Cut Power";
                sh.Cells[r, 2] = string.Format("{0} MHz", Settings.Default.Para1.CutBW);
                sh.Cells[r, 3] = string.Format("{0} MHz", Settings.Default.Para2.CutBW);
                sh.Cells[r, 4] = string.Format("{0} MHz", Settings.Default.Para3.CutBW);
                sh.Cells[r, 5] = string.Format("{0} MHz", Settings.Default.Para4.CutBW);
                r++;
                sh.Cells[r, 1] = "Frequency Difference";
                sh.Cells[r, 2] = string.Format("{0} MHz", Settings.Default.Para1.DiffFreq);
                sh.Cells[r, 3] = string.Format("{0} MHz", Settings.Default.Para2.DiffFreq);
                sh.Cells[r, 4] = string.Format("{0} MHz", Settings.Default.Para3.DiffFreq);
                sh.Cells[r, 5] = string.Format("{0} MHz", Settings.Default.Para4.DiffFreq);
                r++;
                sh.Cells[r, 1] = "Power Reference";
                sh.Cells[r, 2] = string.Format("{0} dBm", Settings.Default.Para1.CutPow);
                sh.Cells[r, 3] = string.Format("{0} dBm", Settings.Default.Para2.CutPow);
                sh.Cells[r, 4] = string.Format("{0} dBm", Settings.Default.Para3.CutPow);
                sh.Cells[r, 5] = string.Format("{0} dBm", Settings.Default.Para4.CutPow);
                r++;
                sh.Cells[r, 1] = "Power Difference";
                sh.Cells[r, 2] = string.Format("{0} dB", Settings.Default.Para1.DiffPower);
                sh.Cells[r, 3] = string.Format("{0} dB", Settings.Default.Para2.DiffPower);
                sh.Cells[r, 4] = string.Format("{0} dB", Settings.Default.Para3.DiffPower);
                sh.Cells[r, 5] = string.Format("{0} dB", Settings.Default.Para4.DiffPower);

                //raw data
                if (list1 != null && list1.Count > 0)
                {
                    InsertData(sh, ref r, list1, Settings.Default.Para1);
                }
                if (list2 != null && list2.Count > 0)
                {
                    InsertData(sh, ref r, list2, Settings.Default.Para1);
                }
                if (list3 != null && list3.Count > 0)
                {
                    InsertData(sh, ref r, list3, Settings.Default.Para1);
                }
                if (list4 != null && list4.Count > 0)
                {
                    InsertData(sh, ref r, list4, Settings.Default.Para1);
                }



                string root = string.Format("{0}\\{1}", Settings.Default.OutputDir, "Report");
                if (!Directory.Exists(root))
                    Directory.CreateDirectory(root);
                string path = string.Format("{0}\\Report_{1}.xlsx", root, DateTime.Now.ToString("MMddHHmmss"));
                app.AlertBeforeOverwriting = false;
                bk.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                app.Quit();

                System.Diagnostics.Process.Start(path);
            }
            catch (Exception ex)
            {
                if (writer != null)
                {
                    writer.Close();
                    writer = null;
                }
                AppLog.Error("Report2 has error.", ex);
                MessageBox.Show("Report error! \n\n\n" + ex.Message);
            }
            finally
            {
            }
        }
        private void InsertData(Excel.Worksheet sh, ref int r, List<SingleData> list1, ParaObject para)
        {
            r++;
            int c = 0;
            if (list1 != null && list1.Count > 0)
            {
                r++;
                ((Excel.Range)sh.Rows[r, Type.Missing]).Interior.ColorIndex = 37;
                ((Excel.Range)sh.Rows[r, Type.Missing]).Font.Bold = true;
                c = 0;
                sh.Cells[r, ++c] = string.Format("Data File ( {0} )", list1[0].TraceType);
                sh.Cells[r, ++c] = "Code";
                sh.Cells[r, ++c] = "Result";
                c++;
                foreach (KeyValuePair<double, double> item in list1[0].ReferData)
                {
                    sh.Cells[r, ++c] = Math.Round(item.Key, 2) + "(MHz)";
                    sh.Cells[r + 1, c] = Math.Round(item.Value, 2);
                }
                foreach (SingleData data in list1)
                {
                    r++;
                    c = 0;
                    sh.Cells[r, ++c] = data.Filename;
                    sh.Cells[r, ++c] = data.Code;
                    sh.Cells[r, ++c] = data.Result;

                    c++;
                    foreach (KeyValuePair<double, double> item in data.ListData)
                    {
                        sh.Cells[r, ++c] = Math.Round(item.Value, 2);
                    }
                    if (data.Result == "Fail")
                    {
                        ((Excel.Range)sh.Rows[r, Type.Missing]).Font.ColorIndex = 3;
                    }
                }
            }
        }
        private void GetSingleData_LOG(string path, string traceType, out int count, out int pass, out List<SingleData> list)
        {
            count = 0;
            pass = 0;
            list = new List<SingleData>();
            string dir = string.Format("{0}\\{1}", path, traceType);
            if (!Directory.Exists(dir))
                return;
            string[] files = Directory.GetFiles(dir);
            string str;
            SingleData single = null;
            SortedList<double, double> caliData = null;
            SortedList<double, double> testData = null;
            StreamReader sr = null;
            try
            {
                foreach (string file in files)
                {
                    Progress++;
                    using (sr = new StreamReader(file))
                    {
                        str = sr.ReadLine();
                        if (!str.Contains("Result"))
                        {
                            sr.Close();
                            continue;
                        }
                        single = new SingleData();
                        single.Filename = System.IO.Path.GetFileName(file);
                        single.Result = str.Split(',')[1];
                        single.Errors = sr.ReadLine().Split(',')[1];
                        single.Code = sr.ReadLine().Split(',')[1];
                        single.TraceType = sr.ReadLine().Split(',')[1];

                        while ((str = sr.ReadLine()) != null)
                        {
                            if (str.Contains("Calibration"))
                                break;
                        }
                        caliData = new SortedList<double, double>();
                        while ((str = sr.ReadLine()) != null)
                        {
                            if (str.Trim() == "") break;
                            caliData.Add(double.Parse(str.Split(',')[1]), double.Parse(str.Split(',')[2]));
                        }
                        single.ReferData = caliData;

                        while ((str = sr.ReadLine()) != null)
                        {
                            if (str.Contains("Test"))
                                break;
                        }
                        testData = new SortedList<double, double>();
                        while ((str = sr.ReadLine()) != null)
                        {
                            if (str.Trim() == "") break;
                            testData.Add(double.Parse(str.Split(',')[1]), double.Parse(str.Split(',')[2]));
                        }
                        single.ListData = testData;

                        sr.Close();
                        count++;
                        if (single.Result == "Pass")
                            pass++;
                        list.Add(single);
                    }
                }
            }
            catch (Exception ex)
            {
                AppLog.Error("GetSingleData has error.", ex);
                if (sr != null)
                {
                    sr.Close();
                    sr = null;
                }
            }
        }
    }
}
