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
    public class DataBase
    {
        public static DataBase self;
        public static TraceFormat Format;
        public static DataBase GetDataHandler(TraceFormat format)
        {
            if (Format == format && self != null)
            {
                return self;
            }
            else
            {
                Format = format;
                if (format == TraceFormat.SWR)
                {
                    return new DataSWR();
                }
                else if (format == TraceFormat.LOG)
                {
                    return new DataLOG();
                }
                else if (format == TraceFormat.LOG_SWR)
                {
                    return new DataLOG_SWR();
                }
                else
                {
                    return null;
                }
            }
        }
        public virtual void Report() { }
        public virtual string Output(ParaObject para, List<ErrorCode> errors) { return null; }
        public virtual string Output(ParaObject para, SortedList<double, double> list, List<ErrorCode> errors) { return null; }

        public static double Progress = 0;

        protected string GetFilePath(ParaObject para, bool pass)
        {
            string root = string.Format("{0}\\{1}",
                    Settings.Default.OutputDir,
                    para.Trace);
            if (!Directory.Exists(root))
                Directory.CreateDirectory(root);
            string path = string.Format("{0}\\{1}_{2}_{3}_{4}.csv",
                    root,
                    para.Trace,
                    Settings.Default.Code,
                    DateTime.Now.ToString("MMddHHmmss"),
                    pass ? "Pass" : "Fail");
            return path;
        }
        public void GetCount8Pass(string path, string traceType, out int count, out int pass)
        {
            count = 0;
            pass = 0;
            string dir = string.Format("{0}\\{1}", path, traceType);
            if (!Directory.Exists(dir))
                return;
            string[] files = Directory.GetFiles(dir);
            try
            {
                foreach (string file in files)
                {
                    if (file.Contains("_Pass") || file.Contains("_Fail"))
                    {
                        if (file.Contains("_Pass"))
                        {
                            pass++;
                        }
                        count++;
                    }
                }
            }
            catch (Exception ex)
            {
                AppLog.Error("GetCount8Pass has error.", ex);
            }
        }
    }
}
