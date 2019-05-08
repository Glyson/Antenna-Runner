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
    public class VNA_AT5071C
    {
        ResourceManager rm = null;
        IMessageBasedSession ses = null;
        public bool IsOK = false;
        private static VNA_AT5071C instance;
        private VNA_AT5071C() { }
        public static VNA_AT5071C GetInstance()
        {
            if (instance == null)
            {
                instance = new VNA_AT5071C();
            }
            return instance;
        }

        ~VNA_AT5071C()
        {
            if (ses != null)
            {
                (ses as Session).Dispose();
            }
        }

        public bool Init(string gpib)
        {
            try
            {
                rm = ResourceManager.GetLocalManager();
                Session s = rm.Open(gpib);
                s.Timeout = 60000;//1 mins
                ses = s as IMessageBasedSession;
                IsOK = true;
                return true;
            }
            catch (Exception ex)
            {
                MainWindow.Self.Dispatcher.Invoke(new Action(delegate
                {
                    MessageBox.Show(MainWindow.Self, ex.Message, "Error");
                }));
                IsOK = false;
                return false;
            }
        }

        private void Write(string cmd, params object[] paras)
        {
            if (paras != null && paras.Length > 0)
            {
                cmd = string.Format(cmd, paras);
            }
            ses.Write(cmd);
        }
        public void Setup()
        {
            int chCnt = 0;
            if (Settings.Default.Para1.Enable)
                chCnt++;
            if (Settings.Default.Para2.Enable)
                chCnt++;
            if (Settings.Default.Para3.Enable)
                chCnt++;
            if (Settings.Default.Para4.Enable)
                chCnt++;

            string chCmd = "DISP:SPL D1";
            if (chCnt == 2)
                chCmd = "DISP:SPL D1_2";
            if (chCnt == 3)
                chCmd = "DISP:SPL D1_2_3";
            if (chCnt == 4)
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
            Write("DISP:WIND{0}:ACT", ch);
            Write("CALC{0}:PAR:DEF {1}", ch, para.Trace);
            //Write("CALC{0}:PAR:COUN 1", ch);
            Write("CALC{0}:FORM SWR", ch);
            Write("SENS{0}:FREQ:STAR {1}", ch, para.FreqStart * 1E6);
            Write("SENS{0}:FREQ:STOP {1}", ch, para.FreqStop * 1E6);
            Write("SENS{0}:SWE:POIN {1}", ch, para.Points);
            Write("SENS{0}:SWE:TYPE LIN", ch);

            //Write("SENS{0}:BWID {1}", ch, Settings.Default.Para1.Bandwidth);
            //Write("SOUR{0}:POW {1}", ch, para.Power);
            //Write("SENS{0}:aver on", ch);
            //Write("SENS{0}:aver:coun {1}", ch, 1);
            //Write("INIT{0}:CONT ON", ch);
            //Write("INIT{0}:CONT ON", ch);


            Write("DISP:WIND{0}:TRAC1:Y:AUTO", ch);
        }

        public void SetupBack()
        {

        }
        //public void Trigger(ParaObject para)
        //{
        //    //int ch = int.Parse(para.Trace.Last().ToString());
        //    //Write("DISP:WIND{0}:ACT", ch);

        //    //int swpTmt = 60000;
        //    //int swpPll = 200;
        //    //int counter = 0, ub = swpTmt / swpPll;
        //    //string status = "";
        //    //Write("SENS{0}:AVER ON", ch);
        //    //Write("SENS{0}:AVER:CLE", ch);
        //    //for (int i = 0; i < 1; i++)
        //    //{
        //    //    counter = 0;
        //    //    Write("TRIG:SING");
        //    //    while (counter < ub && (status != "+1" && status != "1" && status != "+1\n"))
        //    //    {
        //    //        Thread.Sleep(swpPll);
        //    //        status = ses.Query("*OPC?");
        //    //        counter++;
        //    //    }
        //    //    if (counter >= ub)
        //    //        MessageBox.Show(MainWindow.Self, "Sweep timeout, result may be incorrect.", "Error");
        //    //}
        //    //Write(string.Format("DISP:WIND{0}:TRAC1:Y:AUTO", ch));
        //    //Write(string.Format("CALC{0}:DATA:FDAT?", ch));//CAL
        //}
        public SortedList<double, double> ReadSWR(ParaObject para)
        {
            SortedList<double, double> raw = ReadSWRByTrace(para);
            SortedList<double, double> list = new SortedList<double, double>();
            if (para.MarkerType == MarkerType.Points.ToString())
            {
                int cnt = para.Points;
                int cell = para.MarkerPoints;
                int step = (int)((cnt - 1) / (cell - 1));
                int mark;
                for (int i = 0; i < cell; i++)
                {
                    mark = i * step;
                    list.Add(raw.Keys[mark], raw.Values[mark]);
                }
            }
            else
            {
                List<double> markers = GetMarker(para.MarkerText);
                foreach (double marker in markers)
                {
                    list.Add(marker, GetRefer(raw, marker));
                }

            }
            return list;
        }
        private double GetRefer(SortedList<double, double> trace, double freq)
        {
            if (trace.ContainsKey(freq))
                return trace[freq];
            double[] keys = trace.Keys.ToArray<double>();
            Array.Sort<double>(keys);
            if (freq < keys[0])
                return trace[keys[0]];
            if (freq > keys[keys.Length - 1])
                return trace[keys[keys.Length - 1]];
            int index = ~Array.BinarySearch(keys, freq);
            double gt = keys[index];
            double lt = keys[index - 1];
            return trace[gt] - (gt - freq) * (trace[gt] - trace[lt]) / (gt - lt);
        }
        private List<double> GetMarker(string text)
        {
            List<double> list = new List<double>();
            string[] arr = text.Split('\r', '\n');
            double fq;
            foreach (string str in arr)
            {
                if (double.TryParse(str, out fq))
                {
                    if (!list.Contains(fq))
                        list.Add(fq);
                }
            }
            return list;
        }

        public SortedList<double, double> ReadSWRByTrace(ParaObject para)
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
            Write(string.Format("DISP:WIND{0}:TRAC1:Y:AUTO", ch));
            Write(string.Format("CALC{0}:DATA:FDAT?", ch));//CAL
            string cur = string.Empty;
            string[] arrCur;
            do
            {
                cur += ses.ReadString();
                arrCur = cur.Split(',');
            } while (arrCur.Length < para.Points * dimCnt);
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
