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
        public void Setup()
        {
            //Write("SYST:PRES");
            portList.Clear();
            if (Settings.Default.Para1.Enable)
                portList.Add(1);
            if (Settings.Default.Para2.Enable)
                portList.Add(2);
            if (Settings.Default.Para3.Enable)
                portList.Add(3);
            if (Settings.Default.Para4.Enable)
                portList.Add(4);

            string chCmd = "DISP:SPL D1";
            int portCnt = portList.Count;
            if (portCnt == 2)
                chCmd = "DISP:SPL D1_2";
            if (portCnt == 3)
                chCmd = "DISP:SPL D1_2_3";
            if (portCnt == 4)
                chCmd = "DISP:SPL D1_2_3_4";

            Write(chCmd);
            //Write("TRIG:SCOP ALL");// no in 5071B
            Write("TRIG:SOUR INT");
            Write("FORM:DATA ASC");
        }
        public void Setup(ParaObject para)
        {
            if (MainWindow.IsSkip) return;

            int ch = int.Parse(para.Trace.Last().ToString());
            ch = portList.IndexOf(ch) + 1;
            Write("CHAN{0}", ch);
            Write(para.Trace);
            //Write("CALC{0}:PAR:COUN 1", ch);
            if (Settings.Default.TraceFormat == TraceFormat.LOG.ToString())
            {
                Write("CALC{0}:FORM LOGM", ch);
            }
            else
            {
                Write("CALC{0}:FORM SWR", ch);
            }
            Write("STAR {0}", para.FreqStart * 1E6);
            Write("STOP {0}", para.FreqStop * 1E6);
            Write("POIN {0}", para.Points);
            Write("SENS{0}:SWE:TYPE LIN", ch);

            //Write("SENS{0}:BWID {1}", ch, Settings.Default.Para1.Bandwidth);
            //Write("SOUR{0}:POW {1}", ch, para.Power);
            //Write("SENS{0}:aver on", ch);
            //Write("SENS{0}:aver:coun {1}", ch, 1);
            //Write("INIT{0}:CONT ON", ch);
            Write("INIT{0}:CONT ON", ch);


            Write("DISP:WIND{0}:TRAC1:Y:AUTO", ch);
        }

        public override SortedList<double, double> ReadTrace(ParaObject para)
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
            string[] arr = ReadTrace(para, dimCnt);
            SortedList<double, double> list = GetTrace(para, arr, dimCnt, 1);
            return list;
        }
        private string[] ReadTrace(ParaObject para, int dimCnt = 1)
        {
            int ch = int.Parse(para.Trace.Last().ToString());

            Write("CHAN{0}", 1);
            Write("FORM4");
            Write("OUTPFORM");

            Write(string.Format("DISP:WIND{0}:TRAC1:Y:AUTO", ch));
            Write(string.Format("CALC{0}:DATA:FDAT?", ch));//CAL
            string cur = string.Empty;
            string[] arrCur;
            do
            {
                cur += ses.ReadString();
                arrCur = cur.Split(',');
            } while(arrCur.Length < para.Points * dimCnt);
            return arrCur;
        }
        private SortedList<double, double> GetTrace(ParaObject para, string[] arrCur, int dimCnt = 1, int dim = 1)
        {
            SortedList<double, double> list = new SortedList<double, double>();
            double freq = para.FreqStart;
            double step = (para.FreqStop - para.FreqStart) / (para.Points - 1);
            for (int i = 0; i < para.Points; i++, freq += step)
                list.Add(freq, double.Parse(arrCur[i * dimCnt + dim - 1]));
            return list;
        }
    }
}
