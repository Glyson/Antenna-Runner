using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AntRunner.Properties;
using NationalInstruments.VisaNS;
using System.Windows;
using System.Threading;

namespace AntRunner
{
    public class VNA_AT8753ES : VNA
    {
        public override void Config()
        {
            //Write("SYST:PRES");
            GetPorts();
            //string chCmd = "DISP:SPL D1";
            //int portCnt = portList.Count;
            //if (portCnt == 2)
            //    chCmd = "DISP:SPL D1_2";
            //if (portCnt == 3)
            //    chCmd = "DISP:SPL D1_2_3";
            //if (portCnt == 4)
            //    chCmd = "DISP:SPL D1_2_3_4";

            //Write(chCmd);
            ////Write("TRIG:SCOP ALL");// no in 5071B
            //Write("TRIG:SOUR INT");
            //Write("FORM:DATA ASC");

            Write("DUAC {0}", "ON");

        }
        //S21:LOG,  S22:LOG/SWR   
        public override void Setup(ParaObject para)
        {
            if (MainWindow.IsSkip) return;
            Setup(para, 1);
            Setup(para, 2);
        }

        private void Setup(ParaObject para, int ch)
        {
            Write("CHAN{0}", ch);// ChAN1
            if (ch == 1)
            {
                Write("S21");
                Write("LOGM");
            }
            else if (ch == 2)
            {
                Write("S22");
                if (para.S22TraceFormat == S22TraceFormat.SWR)
                {
                    Write("SWR");//LOGM/SWR
                }
                else
                {
                    Write("LOGM");
                }
            }
            Write("STAR {0}", para.FreqStart * 1E6);
            Write("STOP {0}", para.FreqStop * 1E6);
            Write("POIN {0}", para.Points);
            Write("PWRR {0}", "PAUTO");//PAUTO, PMAN
            Write("POWE {0:0.##}", para.Power);
            Write("IFBW {0}", para.Bandwidth);//10,30,100,300,1000,3000,3700,6000
            Write("AUTO");
        }
        //private int AnalyzeIFBW(double bw)
        //{
        //    if (bw <= 10)
        //        return 10;
        //    else if (bw <= 30)
        //        return 30;
        //    else if (bw <= 100)
        //        return 100;
        //    else if (bw <= 300)
        //        return 300;
        //    else if (bw <= 1000)
        //        return 1000;
        //    else if (bw <= 3000)
        //        return 3000;
        //    else if (bw <= 3700)
        //        return 3700;

        //    return 6000;
        //}
        public override SortedList<double, double> ReadTrace(ParaObject para, int traceNum = 1)
        {
            if (MainWindow.IsSkip)
            {
                SortedList<double, double> list2 = new SortedList<double, double>();
                double freq = para.FreqStart;
                double step = (para.FreqStop - para.FreqStart) / (para.Points - 1);
                for (int i = 0; i < para.Points; i++, freq += step)
                    list2.Add(freq, 10);
                return list2;
            }
            int dimCnt = 2;
            string[] arr = ReadInsTrace(para, traceNum, dimCnt);
            SortedList<double, double> list = FixTrace(para, arr, dimCnt, 1);
            return list;
        }
        private string[] ReadInsTrace(ParaObject para, int ch = 1, int dimCnt = 2)
        {
            Write("CHAN{0}", ch);
            Write("FORM4");
            Write("OUTPFORM");

            Write("AUTO");
            string cur = string.Empty;
            string[] arrCur;
            do
            {
                cur += ses.ReadString().Trim();
                arrCur = cur.Split('\n', ',');
            } while (arrCur.Length < para.Points * dimCnt);
            return arrCur;
        }
    }
}
