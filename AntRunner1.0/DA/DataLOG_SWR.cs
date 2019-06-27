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
    public class DataLOG_SWR : DataBase
    {
        #region single test output
        public override string Output(ParaObject para, List<ErrorCode> errors)
        {
            StreamWriter writer = null;
            try
            {
                bool pass = !(errors != null && errors.Count > 0);
                string path = GetFilePath(para, pass);
                using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    writer = new StreamWriter(fs);
                    writer.WriteLine("Result," + (pass ? "Pass" : "Fail"));
                    writer.WriteLine("Error Code," + string.Join("|", errors));
                    writer.WriteLine("DUT Code," + Settings.Default.Code);
                    writer.WriteLine("Memo," + Settings.Default.Memo);
                    writer.WriteLine("S22 Type," + para.S22TraceFormat);
                    writer.WriteLine("S21 Min," + para.S21Min);
                    writer.WriteLine("S21 Max ," + para.S21Max);
                    writer.WriteLine("S22 Max," + para.S22Max);


                    writer.WriteLine();
                    writer.WriteLine("Test,Frequency,S21,S22");
                    foreach (KeyValuePair<double, string> item in (Dictionary<double, string>)para.Markers)
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
                AppLog.Error("Output has error.", ex);
                return null;
            }
        }
        #endregion

        #region Report
        public override void Report()
        {
            Progress = 0;
            int count1;
            int pass1;
            List<SingleData> list1;
            GetSingleData_LOG8SWR(Settings.Default.OutputDir, Settings.Default.Para1.Trace,
                out count1, out pass1, out list1);
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

                sh.Cells[++r, 1] = "Port1";
                sh.Cells[r, 2] = pass1 + "/" + count1;
                sh.Cells[r, 3] = (count1 == 0 ? 0 : Math.Round(pass1 / (double)count1 * 100, 4)) + "%";
                sh.Cells[r, 6] = "偏低：";
                ((Excel.Range)sh.Cells[r, 7]).Interior.ColorIndex = 6;

                int pass = pass1;
                int count = count1;
                sh.Cells[++r, 1] = "Total";
                sh.Cells[r, 2] = pass + "/" + count;
                sh.Cells[r, 3] = (count == 0 ? 0 : Math.Round(pass / (double)count * 100, 4)) + "%";
                sh.Cells[r, 6] = "偏高：";
                ((Excel.Range)sh.Cells[r, 7]).Interior.ColorIndex = 3;


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
                for (int i = 0; i < 1; i++)
                {
                    ParaObject tempPara = null;
                    List<SingleData> listData = null;
                    if (i == 0)
                    {
                        tempPara = Settings.Default.Para1;
                        listData = list1;
                    }

                    r++;
                    r++;
                    ((Excel.Range)sh.Rows[r, Type.Missing]).Interior.ColorIndex = 37;
                    ((Excel.Range)sh.Rows[r, Type.Missing]).Font.Bold = true;
                    sh.Cells[r, 1] = string.Format("Parameters");
                    r++;
                    sh.Cells[r, 1] = "S22 Type";
                    sh.Cells[r, 2] = string.Format("{0}", tempPara.S22TraceFormat);
                    r++;
                    sh.Cells[r, 1] = "S21 Min";
                    sh.Cells[r, 2] = string.Format("{0} dBm", tempPara.S21Min);
                    r++;
                    sh.Cells[r, 1] = "S21 Max";
                    sh.Cells[r, 2] = string.Format("{0} dBm", tempPara.S21Max);
                    r++;
                    sh.Cells[r, 1] = "S22 Max";
                    sh.Cells[r, 2] = string.Format("{0} dBm", tempPara.S22Max);

                    //    if (listData.Count > 0)
                    //    {
                    //        int j = 0;
                    //        string tag = "Left";
                    //        foreach (KeyValuePair<double, double> item in listData[0].ReferData)
                    //        {
                    //            r++;
                    //            j++;
                    //            if (j == 2)
                    //            {
                    //                tag = "Low";
                    //            }
                    //            else if (j == 3)
                    //            {
                    //                tag = "Right";
                    //            }

                    //            sh.Cells[r, 1] = string.Format("Marker ({0})", tag);
                    //            sh.Cells[r, 2] = string.Format("{0,4} MHz", Math.Round(item.Key, 2));
                    //            sh.Cells[r, 3] = string.Format("{0} dBm", Math.Round(item.Value, 2));
                    //        }
                    //    }
                }
                //raw data
                if (list1 != null && list1.Count > 0)
                {
                    InsertData_LOG8SWR(sh, ref r, list1, Settings.Default.Para1);
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
                AppLog.Error("Report_LOG8SWR has error.", ex);
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
        private void GetSingleData_LOG8SWR(string path, string traceType, out int count, out int pass, out List<SingleData> list)
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
            StreamReader sr = null;

            foreach (string file in files)
            {
                try
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
                        single.Memo = sr.ReadLine().Split(',')[1];

                        while ((str = sr.ReadLine()) != null)
                        {
                            if (str.Contains("Test"))
                                break;
                        }

                        single.ListData = new SortedList<double, double>();
                        single.ListData2 = new SortedList<double, double>();
                        while ((str = sr.ReadLine()) != null)
                        {
                            if (str.Trim() == "") break;
                            single.ListData.Add(double.Parse(str.Split(',')[1]), double.Parse(str.Split(',')[2]));
                            single.ListData2.Add(double.Parse(str.Split(',')[1]), double.Parse(str.Split(',')[3]));
                        }
                        list.Add(single);
                        count++;
                        if (single.Result == "Pass")
                            pass++;

                        sr.Close();
                    }
                }
                catch (Exception ex)
                {
                    AppLog.Error("GetSingleData_LOG8SWR has error.", ex);
                    if (sr != null)
                    {
                        sr.Close();
                        sr = null;
                    }
                }
            }
        }
        private void InsertData_LOG8SWR(Excel.Worksheet sh, ref int r, List<SingleData> listData, ParaObject para)
        {
            r++;
            int c = 0;
            if (listData != null && listData.Count > 0)
            {
                r++;
                ((Excel.Range)sh.Rows[r, Type.Missing]).Interior.ColorIndex = 37;
                ((Excel.Range)sh.Rows[r, Type.Missing]).Font.Bold = true;
                c = 0;
                sh.Cells[r, ++c] = string.Format("Data File ( {0} )", listData.Count);
                sh.Cells[r, ++c] = "Code";
                sh.Cells[r, ++c] = "Result";
                //c++;
                sh.Cells[r, ++c] = "S21";
                ((Excel.Range)sh.Cells[r, c]).ColumnWidth = 18;
                sh.Cells[r, ++c] = "S22";
                ((Excel.Range)sh.Cells[r, c]).ColumnWidth = 18;
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
                    if (data.Errors.Contains(ErrorCode.PowS21L.ToString()))
                    {
                        tag = "偏低";
                        ((Excel.Range)sh.Cells[r, 4]).Interior.ColorIndex = 6;
                    }
                    else if (data.Errors.Contains(ErrorCode.PowS21H.ToString()))
                    {
                        tag = "偏高";
                        ((Excel.Range)sh.Cells[r, 4]).Interior.ColorIndex = 3;
                    }
                    //string str = string.Format("{0:N2}", data.ListData.Keys[1]);
                    sh.Cells[r, ++c] = tag;

                    ++c;
                    tag = "正常";
                    ((Excel.Range)sh.Cells[r, c]).Font.ColorIndex = 1;
                    if (data.Errors.Contains(ErrorCode.PowS22H.ToString()))
                    {
                        tag = "偏高";
                        ((Excel.Range)sh.Cells[r, c]).Interior.ColorIndex = 3;
                    }
                    //str = string.Format("{0:N2}", data.ListData.Values[1]);
                    sh.Cells[r, c] = tag;

                    //++c;
                    //sh.Cells[r, c] = string.Format("{0:N2}", data.ListData.Keys.First());
                    //((Excel.Range)sh.Cells[r, c]).Font.ColorIndex = 1;

                    //++c;
                    //sh.Cells[r, c] = string.Format("{0:N2}", data.ListData.Keys.Last());
                    //((Excel.Range)sh.Cells[r, c]).Font.ColorIndex = 1;

                    //++c;
                    //tag = "正常";
                    //((Excel.Range)sh.Cells[r, c]).Font.ColorIndex = 1;
                    //if (data.Errors.Contains(ErrorCode.FreqBandWidthL.ToString()))
                    //{
                    //    tag = "偏低";
                    //    ((Excel.Range)sh.Cells[r, c]).Interior.ColorIndex = 6;
                    //}
                    //else if (data.Errors.Contains(ErrorCode.FreqBandWidthH.ToString()))
                    //{
                    //    tag = "偏高";
                    //    ((Excel.Range)sh.Cells[r, c]).Interior.ColorIndex = 3;
                    //}
                    //str = string.Format("{0:N2}", data.ListData.Keys.Last() - data.ListData.Keys.First());
                    //sh.Cells[r, c] = str;

                    //tag = "否";
                    //++c;
                    //((Excel.Range)sh.Cells[r, c]).Font.ColorIndex = 1;
                    //if (data.Errors.Contains(ErrorCode.Bad.ToString()))
                    //{
                    //    tag = "是";
                    //    ((Excel.Range)sh.Cells[r, c]).Interior.ColorIndex = 3;
                    //}
                    //sh.Cells[r, c] = tag;


                }
            }
        }
        #endregion
    }
}
