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
    public class DataLOG : DataBase
    {
        #region single test output
        public override string Output(ParaObject para, SortedList<double, double> list, List<ErrorCode> errors)
        {
            StreamWriter writer = null;
            string path = null;
            try
            {
                bool pass = !(errors != null && errors.Count > 0);
                path = GetFilePath(para, pass);
                using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    writer = new StreamWriter(fs);
                    writer.WriteLine("Result," + (pass ? "Pass" : "Fail"));
                    writer.WriteLine("Error Code," + string.Join("|", errors));
                    writer.WriteLine("DUT Code," + Settings.Default.Code);
                    writer.WriteLine("Trace Type," + para.Trace);
                    writer.WriteLine("Cut Power," + para.CutPow);
                    writer.WriteLine("Frequency Width ," + para.CutBW);
                    writer.WriteLine("Frequency Width Difference," + para.DiffBW);
                    writer.WriteLine("Frequency Difference," + para.DiffFreq);
                    writer.WriteLine("Power Difference," + para.DiffPower);
                    writer.WriteLine("Memo," + Settings.Default.Memo);

                    writer.WriteLine();
                    writer.WriteLine("Calibration,Frequency,Data");
                    foreach (KeyValuePair<double, double> item in (Dictionary<double, double>)para.MarkersCal)
                    {
                        writer.WriteLine(" ,{0},{1}", item.Key, item.Value);
                    }

                    writer.WriteLine();
                    writer.WriteLine("Test,Frequency,Data");
                    foreach (KeyValuePair<double, double> item in (Dictionary<double, double>)para.Markers)
                    {
                        writer.WriteLine(" ,{0},{1}", item.Key, item.Value);
                    }
                    writer.Flush();
                    writer.Close();
                    writer = null;
                    fs.Close();
                    return path;
                }
            }
            catch (Exception ex)
            {
                File.Delete(path);
                AppLog.Error("Output has error.", ex);
                return null;
            }
        }
        #endregion

        #region Report
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
                sh.Cells[r, 2] = pass1 + "/" + count1;
                sh.Cells[r, 3] = (count1 == 0 ? 0 : Math.Round(pass1 / (double)count1 * 100, 4)) + "%";
                sh.Cells[r, 6] = "偏低：";
                ((Excel.Range)sh.Cells[r, 7]).Interior.ColorIndex = 6;

                sh.Cells[++r, 1] = "Port2(S22)";
                sh.Cells[r, 2] = pass2 + "/" + count2;
                sh.Cells[r, 3] = (count2 == 0 ? 0 : Math.Round(pass2 / (double)count2 * 100, 4)) + "%";
                sh.Cells[r, 6] = "偏高：";
                ((Excel.Range)sh.Cells[r, 7]).Interior.ColorIndex = 3;

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
                for (int i = 0; i < 4; i++)
                {
                    ParaObject tempPara = null;
                    List<SingleData> listData = null;
                    if (i == 0)
                    {
                        tempPara = Settings.Default.Para1;
                        listData = list1;
                    }
                    else if (i == 1)
                    {
                        tempPara = Settings.Default.Para2;
                        listData = list2;
                    }
                    else if (i == 2)
                    {
                        tempPara = Settings.Default.Para3;
                        listData = list3;
                    }
                    else if (i == 3)
                    {
                        tempPara = Settings.Default.Para4;
                        listData = list4;
                    }

                    r++;
                    r++;
                    ((Excel.Range)sh.Rows[r, Type.Missing]).Interior.ColorIndex = 37;
                    ((Excel.Range)sh.Rows[r, Type.Missing]).Font.Bold = true;
                    sh.Cells[r, 1] = string.Format("Parameters({0})", tempPara.Trace);
                    r++;
                    sh.Cells[r, 1] = "Cut Power";
                    sh.Cells[r, 2] = string.Format("{0} dBm", tempPara.CutPow);
                    r++;
                    sh.Cells[r, 1] = "Frequency Width";
                    sh.Cells[r, 2] = string.Format("{0:N2} MHz", tempPara.CutBW);
                    r++;
                    sh.Cells[r, 1] = "Frequency Width Difference";
                    sh.Cells[r, 2] = string.Format("{0} MHz", tempPara.DiffBW);
                    r++;
                    sh.Cells[r, 1] = "Frequency Difference";
                    sh.Cells[r, 2] = string.Format("{0} MHz", tempPara.DiffFreq);
                    r++;
                    sh.Cells[r, 1] = "Power Difference";
                    sh.Cells[r, 2] = string.Format("{0} dB", tempPara.DiffPower);

                    if (listData.Count > 0)
                    {
                        int j = 0;
                        string tag = "Left";
                        foreach (KeyValuePair<double, double> item in listData[0].ReferData)
                        {
                            r++;
                            j++;
                            if (j == 2)
                            {
                                tag = "Low";
                            }
                            else if (j == 3)
                            {
                                tag = "Right";
                            }

                            sh.Cells[r, 1] = string.Format("Marker ({0})", tag);
                            sh.Cells[r, 2] = string.Format("{0,4} MHz", Math.Round(item.Key, 2));
                            sh.Cells[r, 3] = string.Format("{0} dBm", Math.Round(item.Value, 2));
                        }
                    }
                }

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
                AppLog.Error("Report_LOG has error.", ex);
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (writer != null)
                {
                    writer.Close();
                    writer = null;
                }
            }
        }
        private void InsertData(Excel.Worksheet sh, ref int r, List<SingleData> listData, ParaObject para)
        {
            r++;
            int c = 0;
            if (listData != null && listData.Count > 0)
            {
                r++;
                ((Excel.Range)sh.Rows[r, Type.Missing]).Interior.ColorIndex = 37;
                ((Excel.Range)sh.Rows[r, Type.Missing]).Font.Bold = true;
                c = 0;
                sh.Cells[r, ++c] = string.Format("Data File ( {0} )", listData[0].TraceType);
                sh.Cells[r, ++c] = "Code";
                sh.Cells[r, ++c] = "Result";
                //c++;
                sh.Cells[r, ++c] = "频率偏差(MHz)";
                ((Excel.Range)sh.Cells[r, c]).ColumnWidth = 18;
                sh.Cells[r, ++c] = "功率偏差(dBm)";
                ((Excel.Range)sh.Cells[r, c]).ColumnWidth = 18;
                sh.Cells[r, ++c] = "Marker1(MHz)";
                ((Excel.Range)sh.Cells[r, c]).ColumnWidth = 18;
                sh.Cells[r, ++c] = "Marker2(MHz)";
                ((Excel.Range)sh.Cells[r, c]).ColumnWidth = 18;
                sh.Cells[r, ++c] = "频宽偏差[Marker2-Marker1](MHz)";
                ((Excel.Range)sh.Cells[r, c]).ColumnWidth = 36;
                sh.Cells[r, ++c] = "短路(是/否)";
                foreach (SingleData data in listData)
                {
                    Progress++;
                    r++;
                    c = 0;
                    sh.Cells[r, ++c] = data.Filename;
                    sh.Cells[r, ++c] = data.Code;
                    sh.Cells[r, ++c] = data.Result;

                    if (data.Result == "Fail")
                    {
                        ((Excel.Range)sh.Rows[r, Type.Missing]).Font.ColorIndex = 3;
                    }
                    else
                    {
                        ((Excel.Range)sh.Rows[r, Type.Missing]).Font.ColorIndex = 1;
                    }

                    //c++;
                    string tag = "正常";
                    ((Excel.Range)sh.Cells[r, 4]).Font.ColorIndex = 1;
                    if (data.Errors.Contains(ErrorCode.FreqL.ToString()))
                    {
                        tag = "偏低";
                        ((Excel.Range)sh.Cells[r, 4]).Interior.ColorIndex = 6;
                    }
                    else if (data.Errors.Contains(ErrorCode.FreqH.ToString()))
                    {
                        tag = "偏高";
                        ((Excel.Range)sh.Cells[r, 4]).Interior.ColorIndex = 3;
                    }
                    string str = string.Format("{0:N2}", data.ListData.Keys[1]);
                    sh.Cells[r, ++c] = str;

                    ++c;
                    tag = "正常";
                    ((Excel.Range)sh.Cells[r, c]).Font.ColorIndex = 1;
                    if (data.Errors.Contains(ErrorCode.PowL.ToString()))
                    {
                        tag = "偏低";
                        ((Excel.Range)sh.Cells[r, c]).Interior.ColorIndex = 6;
                    }
                    else if (data.Errors.Contains(ErrorCode.PowH.ToString()))
                    {
                        tag = "偏高";
                        ((Excel.Range)sh.Cells[r, c]).Interior.ColorIndex = 3;
                    }
                    str = string.Format("{0:N2}", data.ListData.Values[1]);
                    sh.Cells[r, c] = str;

                    ++c;
                    sh.Cells[r, c] = string.Format("{0:N2}", data.ListData.Keys.First());
                    ((Excel.Range)sh.Cells[r, c]).Font.ColorIndex = 1;

                    ++c;
                    sh.Cells[r, c] = string.Format("{0:N2}", data.ListData.Keys.Last());
                    ((Excel.Range)sh.Cells[r, c]).Font.ColorIndex = 1;

                    ++c;
                    tag = "正常";
                    ((Excel.Range)sh.Cells[r, c]).Font.ColorIndex = 1;
                    if (data.Errors.Contains(ErrorCode.FreqBandWidthL.ToString()))
                    {
                        tag = "偏低";
                        ((Excel.Range)sh.Cells[r, c]).Interior.ColorIndex = 6;
                    }
                    else if (data.Errors.Contains(ErrorCode.FreqBandWidthH.ToString()))
                    {
                        tag = "偏高";
                        ((Excel.Range)sh.Cells[r, c]).Interior.ColorIndex = 3;
                    }
                    str = string.Format("{0:N2}", data.ListData.Keys.Last() - data.ListData.Keys.First());
                    sh.Cells[r, c] = str;

                    tag = "否";
                    ++c;
                    ((Excel.Range)sh.Cells[r, c]).Font.ColorIndex = 1;
                    if (data.Errors.Contains(ErrorCode.Bad.ToString()))
                    {
                        tag = "是";
                        ((Excel.Range)sh.Cells[r, c]).Interior.ColorIndex = 3;
                    }
                    sh.Cells[r, c] = tag;


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
        #endregion

    }
}
