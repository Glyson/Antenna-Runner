using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Xceed.Wpf.AvalonDock.Layout;
using System.Threading;
using AntRunner.Properties;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO.Ports;
using System.Globalization;
using System.ComponentModel;

namespace AntRunner
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public const bool IsSkip = true;
        volatile int Count1, Count2, Count3, Count4;
        volatile int Pass1, Pass2, Pass3, Pass4;
        string code1, code2, code3, code4;
        Thread t1, t2, t3, t4;
        bool manual1, manual2, manual3, manual4;
        SortedList<double, double> refer1, refer2, refer3, refer4;
        Dictionary<double, double> markersCal1, markersCal2, markersCal3, markersCal4;
        Dictionary<double, double> markers1, markers2, markers3, markers4;
        public static MainWindow Self;
        public VNA vna = null;
        Brush bshPass = Brushes.SkyBlue;
        Brush bshFail = Brushes.Red;
        Brush bshTgr = Brushes.ForestGreen;
        private State state = State.Stoped;

        public State State
        {
            get { return state; }
            set { state = value; }
        }

        public MainWindow()
        {
            Self = this;
            InitializeComponent();
            BrushConverter brushConverter = new BrushConverter();
            bshPass = (Brush)brushConverter.ConvertFromString("#FF3092FC");
            bshTgr = (Brush)brushConverter.ConvertFromString("#FF28F31B");

            //string unique = MyMD5.GetComputerId();

            if (HasExpired())
            {
                List<string> list = new List<string>();
                list.Add("System.Net.dll has expired.");
                list.Add("System.Data.framework.dll has expired.");
                list.Add("System.Data.Core.dll has expired.");
                list.Add("Microsoft.CSharp.dll has expired.");
                list.Add("NationalInstruments.Common.dll has expired.");
                list.Add("System.Windows.Forms.dll has expired.");
                list.Add("System.Xaml.dll has expired.");
                list.Add("System.WindowsBase.dll has expired.");
                list.Add("Xceed.Wpf.Linq.dll has expired.");
                list.Add("Xceed.Wpf.Data.dll has expired.");
                list.Add("Xceed.Wpf.Core.dll has expired.");
                list.Add("Xceed.Wpf.AvalonDock.dll has expired.");
                MessageBox.Show(list[DateTime.Now.Month - 1]);
                Application.Current.Shutdown();
            }
            if (Settings.Default.Para1 == null)
                Settings.Default.Para1 = new ParaObject();
            if (Settings.Default.Para2 == null)
                Settings.Default.Para2 = new ParaObject();
            if (Settings.Default.Para3 == null)
                Settings.Default.Para3 = new ParaObject();
            if (Settings.Default.Para4 == null)
                Settings.Default.Para4 = new ParaObject();
            Settings.Default.Para1.Trace = Trace.S11.ToString();
            Settings.Default.Para2.Trace = Trace.S22.ToString();
            Settings.Default.Para3.Trace = Trace.S33.ToString();
            Settings.Default.Para4.Trace = Trace.S44.ToString();

            btnStart1.IsEnabled = false;
            btnStart2.IsEnabled = false;
            btnStart3.IsEnabled = false;
            btnStart4.IsEnabled = false;

            InitCount();
        }

        private bool HasExpired()
        {
            FileStream fs = new FileStream("license.config", FileMode.Open, FileAccess.Read);
            StreamReader sr = new StreamReader(fs);
            string str = sr.ReadLine();
            str = Helper.Decode(str);

            DateTime time = DateTime.Parse(str);
            return DateTime.Now > time;
        }
        public void InitCount()
        {
            DataAccess.GetCount8Pass(Settings.Default.OutputDir, Settings.Default.Para1.Trace, out Count1, out Pass1);
            DataAccess.GetCount8Pass(Settings.Default.OutputDir, Settings.Default.Para2.Trace, out Count2, out Pass2);
            DataAccess.GetCount8Pass(Settings.Default.OutputDir, Settings.Default.Para3.Trace, out Count3, out Pass3);
            DataAccess.GetCount8Pass(Settings.Default.OutputDir, Settings.Default.Para4.Trace, out Count4, out Pass4);
            tb1.Text = GetPercentStr(Count1, Pass1);
            tb2.Text = GetPercentStr(Count2, Pass2);
            tb3.Text = GetPercentStr(Count3, Pass3);
            tb4.Text = GetPercentStr(Count4, Pass4);
            tbAll.Text = GetPercentStr(Count1 + Count2 + Count3 + Count4, Pass1 + Pass2 + Pass3 + Pass4);
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (State == State.Running)
            {
                MessageBox.Show("正在测试...");
                return;
            }
            SetupWin win = new SetupWin();
            win.Owner = this;
            win.ShowDialog();
        }
        private void MenuItem_Click_3(object sender, RoutedEventArgs e)
        {
            if (State == State.Running)
            {
                MessageBox.Show(this, "正在测试...", "Tips", MessageBoxButton.OK);
                return;
            }
            if (!InitVNA())
                return;

            if (Settings.Default.MatchCnt < 2)
                Settings.Default.MatchCnt = 2;

            Start(Settings.Default.Para1, ref refer1, ref ellipse1, ref btnStart1, ref t1);
            Start(Settings.Default.Para2, ref refer2, ref ellipse2, ref btnStart2, ref t2);
            Start(Settings.Default.Para3, ref refer3, ref ellipse3, ref btnStart3, ref t3);
            Start(Settings.Default.Para4, ref refer4, ref ellipse4, ref btnStart4, ref t4);
        }
        private void StartCOM1(string name)
        {
            SerialPort com = new SerialPort(name);
            com.DataReceived += new SerialDataReceivedEventHandler(com_DataReceived1);
            com.Open();
        }

        void com_DataReceived1(object sender, SerialDataReceivedEventArgs e)
        {

        }
        private bool InitVNA()
        {

            if (vna == null || !vna.IsOK)
            {
                vna = VNA.CreateVNA();

                if (IsSkip) return true;
                if (!vna.Init(Settings.Default.GPIB))
                {
                    return false;
                }
            }
            vna.Config();
            return true;
        }
        private void Start(ParaObject para, ref SortedList<double, double> refer, ref Ellipse elli, ref Button btn, ref Thread td)
        {
            if (para.Enable)
            {
                if (para.Trace == "S11")
                {
                    refer1 = GetRefer(Settings.Default.Para1);
                    if (refer1 == null) return;
                }
                else if (para.Trace == "S22")
                {
                    refer2 = GetRefer(Settings.Default.Para2);
                    if (refer2 == null) return;
                }
                else if (para.Trace == "S33")
                {
                    refer3 = GetRefer(Settings.Default.Para3);
                    if (refer3 == null) return;
                }
                else if (para.Trace == "S44")
                {
                    refer4 = GetRefer(Settings.Default.Para4);
                    if (refer4 == null) return;
                }

                lock (vna)
                {
                    if (vna != null)
                    {
                        vna.Setup(para);
                    }
                }
                elli.Fill = bshPass;
                btn.IsEnabled = true;
                State = State.Running;
                if (para.Trace == "S11")
                {
                    td = new Thread(ThreadStart1);
                }
                else if (para.Trace == "S22")
                {
                    td = new Thread(ThreadStart2);
                }
                else if (para.Trace == "S33")
                {
                    td = new Thread(ThreadStart3);
                }
                else if (para.Trace == "S44")
                {
                    td = new Thread(ThreadStart4);
                }
                td.Start();
            }
        }
        private void ThreadStart1()
        {
            while (true)
            {
                if (Settings.Default.TriggerType == TriggerType.Auto.ToString())
                {
                    AutoTriger(Settings.Default.Para1);
                }
                else if (Settings.Default.TriggerType == TriggerType.Scanner.ToString())
                {
                    return;
                }
                Run1();
                manual1 = false;
            }
        }
        private void ThreadStart2()
        {
            while (true)
            {
                if (Settings.Default.TriggerType == TriggerType.Auto.ToString())
                {
                    AutoTriger(Settings.Default.Para2);
                }
                else if (Settings.Default.TriggerType == TriggerType.Scanner.ToString())
                {
                    return;
                }
                Run2();
                manual2 = false;
            }
        }
        private void ThreadStart3()
        {
            while (true)
            {
                if (Settings.Default.TriggerType == TriggerType.Auto.ToString())
                {
                    AutoTriger(Settings.Default.Para3);
                }
                else if (Settings.Default.TriggerType == TriggerType.Scanner.ToString())
                {
                    return;
                }
                Run3();
                manual3 = false;
            }
        }
        private void ThreadStart4()
        {
            while (true)
            {
                if (Settings.Default.TriggerType == TriggerType.Auto.ToString())
                {
                    AutoTriger(Settings.Default.Para4);
                }
                else if (Settings.Default.TriggerType == TriggerType.Scanner.ToString())
                {
                    return;
                }
                Run4();
                manual4 = false;
            }
        }
        private void Run1()
        {
            bool pass;
            List<int> errors;
            SortedList<double, double> list = null;
            btnStart1.Dispatcher.BeginInvoke(new Action(delegate
            {
                btnStart1.IsEnabled = false;
                ellipse1.Fill = bshTgr;
            }));
            lock (vna)
            {
                pass = CheckPass(ref list, Settings.Default.Para1, out errors);
                string path = DataAccess.Output(Settings.Default.Para1, list, errors);
                LogMsg(Settings.Default.Para1.Trace, pass, (pass ? "Pass" : "Fail") + " | Stored in : ", path);
                Count1++;
            }
            this.Dispatcher.BeginInvoke(new Action(delegate
            {
                if (pass) Pass1++;
                ellipse1.Fill = pass ? bshPass : bshFail;
                blk1.Text = pass ? "合格" : "不合格";
                tb1.Text = GetPercentStr(Count1, Pass1);
                UpdateAll();
                btnStart1.IsEnabled = true;
            }));
        }
        private void Run2()
        {
            bool pass;
            List<int> errors;
            SortedList<double, double> list = null;
            btnStart2.Dispatcher.BeginInvoke(new Action(delegate
            {
                btnStart2.IsEnabled = false;
                ellipse2.Fill = bshTgr;
            }));
            lock (vna)
            {
                pass = CheckPass(ref list, Settings.Default.Para2, out errors);
                string path = DataAccess.Output(Settings.Default.Para2, list, errors);
                LogMsg(Settings.Default.Para2.Trace, pass, (pass ? "Pass" : "Fail") + " | Stored in : ", path);
                Count2++;
            }
            this.Dispatcher.BeginInvoke(new Action(delegate
            {
                if (pass) Pass2++;
                ellipse2.Fill = pass ? bshPass : bshFail;
                blk2.Text = pass ? "合格" : "不合格";
                tb2.Text = GetPercentStr(Count2, Pass2);
                UpdateAll();
                btnStart2.IsEnabled = true;
            }));
        }
        private void Run3()
        {
            bool pass;
            List<int> errors;
            SortedList<double, double> list = null;
            btnStart3.Dispatcher.BeginInvoke(new Action(delegate
            {
                btnStart3.IsEnabled = false;
                ellipse3.Fill = bshTgr;
            }));
            lock (vna)
            {
                pass = CheckPass(ref list, Settings.Default.Para3, out errors);
                string path = DataAccess.Output(Settings.Default.Para3, list, errors);
                LogMsg(Settings.Default.Para3.Trace, pass, (pass ? "Pass" : "Fail") + " | Stored in : ", path);
                Count3++;
            }
            this.Dispatcher.BeginInvoke(new Action(delegate
            {
                if (pass) Pass3++;
                ellipse3.Fill = pass ? bshPass : bshFail;
                blk3.Text = pass ? "合格" : "不合格";
                tb3.Text = GetPercentStr(Count3, Pass3);
                UpdateAll();
                btnStart3.IsEnabled = true;
            }));
        }
        private void Run4()
        {
            bool pass;
            List<int> errors;
            SortedList<double, double> list = null;
            btnStart4.Dispatcher.BeginInvoke(new Action(delegate
            {
                btnStart4.IsEnabled = false;
                ellipse4.Fill = bshTgr;
            }));
            lock (vna)
            {
                pass = CheckPass(ref list, Settings.Default.Para4, out errors);
                string path = DataAccess.Output(Settings.Default.Para4, list, errors);
                LogMsg(Settings.Default.Para4.Trace, pass, (pass ? "Pass" : "Fail") + " | Stored in : ", path);
                Count4++;
            }
            this.Dispatcher.BeginInvoke(new Action(delegate
            {
                if (pass) Pass4++;
                ellipse4.Fill = pass ? bshPass : bshFail;
                blk4.Text = pass ? "合格" : "不合格";
                tb4.Text = GetPercentStr(Count4, Pass4);
                UpdateAll();
                btnStart4.IsEnabled = true;
            }));
        }
        private bool AutoTriger(ParaObject para)
        {
            Trace type = (Trace)Enum.Parse(typeof(Trace), para.Trace);
            SortedList<double, double> list = null;
            List<SortedList<double, double>> all = new List<SortedList<double, double>>();
            bool triggerDeep = false;
            while (true)
            {
                if (type == Trace.S11 && manual1)
                    return true;
                else if (type == Trace.S22 && manual2)
                    return true;
                if (type == Trace.S33 && manual3)
                    return true;
                if (type == Trace.S44 && manual4)
                    return true;
                Thread.Sleep(Settings.Default.AutoDelay);
                lock (vna)
                {
                    list = vna.ReadTrace(para);//读曲线
                }
                if (IsDeep(list, Settings.Default.Deep))//如果是底噪，即触发过底噪，之前数据清空，路过
                {
                    triggerDeep = true;
                    all.Clear();
                    continue;
                }

                if (!triggerDeep && !IsSkip)//如果没有触发过底噪，跳过
                    continue;

                //如果触发过底噪（说明重新拎过天线），且数据不是底噪，这个数据才是可用的
                btnStart1.Dispatcher.BeginInvoke(new Action(delegate
                {
                    switch (type)
                    {
                        case Trace.S11:
                            ellipse1.Fill = bshTgr;
                            blk1.Text = "正在触发 . . .";
                            break;
                        case Trace.S22:
                            ellipse2.Fill = bshTgr;
                            blk2.Text = "正在触发 . . .";
                            break;
                        case Trace.S33:
                            ellipse3.Fill = bshTgr;
                            blk3.Text = "正在触发 . . .";
                            break;
                        case Trace.S44:
                            ellipse4.Fill = bshTgr;
                            blk4.Text = "正在触发 . . .";
                            break;
                    }
                }));
                all.Add(list);
                if (all.Count >= Settings.Default.MatchCnt)
                {
                    //如果所有曲线全匹配，ok，否则把最旧的数据删掉，再测一条曲线
                    if (MatchTrace(all, Settings.Default.AutoDiff))
                    {
                        return true;
                    }
                    else
                    {
                        all.RemoveAt(0);
                    }
                }
            }
        }
        //private bool MatchTrace(SortedList<double, double> trace1, SortedList<double, double> trace2, double diff)
        //{
        //    if (trace1.Count() != trace2.Count())
        //        return false;
        //    else
        //    {
        //        diff = Math.Abs(diff);
        //        for (int i = 0; i < trace1.Count(); i++)
        //        {
        //            if (Math.Abs(trace1.Values[i] - trace2.Values[i]) > diff)
        //                return false;
        //        }
        //        return true;
        //    }
        //}
        private bool MatchTrace(List<SortedList<double, double>> lists, double diff)
        {
            if (lists.Count() <= 1)
                return false;

            diff = Math.Abs(diff);
            int traceCnt = lists.Count();
            for (int i = 0; i < lists[traceCnt - 1].Count(); i++)
            {
                for (int j = 0; j < traceCnt - 1; j++)
                {
                    if (Math.Abs(lists[traceCnt - 1].Values[i] - lists[j].Values[i]) > diff)
                        return false;
                }
            }
            return true;
        }

        /// <summary>
        /// 判断是否是底噪，（曲线全是deep值之上）
        /// </summary>
        /// <param name="trace"></param>
        /// <param name="deep"></param>
        /// <returns></returns>
        private bool IsDeep(SortedList<double, double> trace, double deep)
        {
            foreach (KeyValuePair<double, double> item in trace)
            {
                if (item.Value < deep)
                    return false;
            }
            return true;
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
        private bool CheckPass(ref SortedList<double, double> list, ParaObject para, out List<int> errors)
        {
            if (Settings.Default.TraceFormat == TraceFormat.LOG.ToString())
            {
                list = vna.ReadTrace(para);
                return CheckPass2(list, para, out errors);
            }
            else
            {
                SortedList<double, double> raw = vna.ReadTrace(para);
                list = new SortedList<double, double>();
                List<double> markers = GetMarker(para.MarkerText);
                foreach (double marker in markers)
                {
                    list.Add(marker, GetPointByTrace(raw, marker));
                }
                errors = new List<int>();
                return CheckPass1(list, para);
            }
        }

        private bool CheckPass1(SortedList<double, double> list, ParaObject para)
        {
            SortedList<double, double> referTrace = GetReferTrace(para);
            double refer, min, max;
            foreach (KeyValuePair<double, double> item in list)
            {
                refer = GetPointByTrace(referTrace, item.Key);
                min = refer - para.ReferDiff;
                max = refer + para.ReferDiff;
                if (item.Value < min || item.Value > max)
                    return false;
            }
            return true;
        }

        private bool CheckPass2(SortedList<double, double> trace, ParaObject para, out List<int> errorCode)
        {
            errorCode = new List<int>();
            double powRef, fqRef, powMin, fqMin;
            SortedList<double, double> referTrace = GetReferTrace(para);
            GetTraceMin(referTrace, out fqRef, out powRef);
            GetTraceMin(trace, out fqMin, out powMin);
            double diffFreq = Math.Abs(para.DiffFreq);
            double diffPower = Math.Abs(para.DiffPower);
            //检查最低点的频率偏差(横向比较)
            if (fqMin < fqRef - diffFreq || fqMin > fqRef + diffFreq)
            {
                errorCode.Add(1);
            }

            //检查最低点的功能偏差（纵向比较）
            if (powMin < powRef - diffPower || powMin > powRef + diffPower)
            {
                errorCode.Add(2);
            }

            //检查功率切线的频宽（频宽比较）
            ParaObject para2 = new ParaObject();
            para2.CutPow = para.CutPow;
            para.Markers = GetParaDataInTrace(trace, para2);
            double diffBW = Math.Abs(para.DiffBW);
            if (para2.CutBW < para.CutBW - diffBW || para2.CutBW > para.CutBW + diffBW)
            {
                errorCode.Add(3);
            }

            return errorCode.Count == 0;
        }

        private Dictionary<double, double> GetParaDataInTrace(SortedList<double, double> trace, ParaObject para)
        {
            double minFreq, minPower;
            GetTraceMin(trace, out minFreq, out minPower);
            SortedList<double, double> trace1 = new SortedList<double, double>();
            SortedList<double, double> trace2 = new SortedList<double, double>();

            foreach (KeyValuePair<double, double> item in trace)
            {
                if (item.Key <= minFreq)
                {
                    trace1.Add(item.Key, item.Value);
                }
                else
                {
                    trace2.Add(item.Key, item.Value);
                }
            }
            double cutPower = para.CutPow;
            double cutF1, cutF2;
            double[] values = trace1.Values.ToArray<double>();
            int index = LocateTrace(values, cutPower);
            if (index <= 0)
            {
                cutF1 = trace1.Keys[index];
            }
            else
            {
                double gt = values[index];
                double lt = values[index - 1];
                double gtF = trace1.Keys[index];
                double ltF = trace1.Keys[index - 1];
                cutF1 = gtF - (gtF - ltF) * (gt - cutPower) / (gt - lt);
            }

            values = trace2.Values.ToArray<double>();
            index = LocateTrace(values, cutPower);
            if (index <= 0)
            {
                cutF2 = trace2.Keys[index];
            }
            else
            {
                double gt = values[index];
                double lt = values[index - 1];
                double gtF = trace2.Keys[index];
                double ltF = trace2.Keys[index - 1];
                cutF2 = gtF - (gtF - ltF) * (gt - cutPower) / (gt - lt);
            }
            if (cutF2 < cutF1)
            {
                para.CutLeftFreq = cutF2;
                para.CutRightFreq = cutF1;
            }
            else
            {
                para.CutLeftFreq = cutF1;
                para.CutRightFreq = cutF2;
            }
            para.CutBW = para.CutRightFreq - para.CutLeftFreq;

            Dictionary<double, double> markers = new Dictionary<double, double>();
            markers.Add(para.CutLeftFreq, para.CutPow);
            if (!markers.ContainsKey(minFreq))
            {
                markers.Add(minFreq, minPower);
            }
            if (!markers.ContainsKey(para.CutRightFreq))
            {
                markers.Add(para.CutRightFreq, para.CutPow);
            }
            return markers;
        }

        private int LocateTrace(double[] arr, double value)
        {
            for (int i = 0; i < arr.Length - 1; i++)
            {
                if (arr[i] < value && arr[i + 1] > value)
                {
                    return i;
                }
                else if (arr[i] > value && arr[i + 1] < value)
                {
                    return i;
                }
                else if (arr[i] == value)
                {
                    return i;
                }
            }
            return 0;
        }
        private SortedList<double, double> GetReferTrace(ParaObject para)
        {
            SortedList<double, double> trace = null;
            Trace type = (Trace)Enum.Parse(typeof(Trace), para.Trace);
            switch (type)
            {
                case Trace.S11:
                    trace = refer1;
                    break;
                case Trace.S22:
                    trace = refer2;
                    break;
                case Trace.S33:
                    trace = refer3;
                    break;
                case Trace.S44:
                    trace = refer4;
                    break;
            }
            return trace;
        }

        private void GetTraceMin(SortedList<double, double> trace, out double freq, out double pow)
        {
            //KeyValuePair<double, double> min = trace.Min();
            pow = trace.Values[0]; ;
            freq = trace.Keys[0];
            foreach (KeyValuePair<double, double> item in trace)
            {
                if (item.Value < pow)
                {
                    pow = item.Value;
                    freq = item.Key;
                }
            }
        }

        private double GetPointByTrace(SortedList<double, double> trace, double freq)
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


        private void MenuItem_Click_6(object sender, RoutedEventArgs e)
        {
            MenuItem menu = sender as MenuItem;
            if (menu.Tag.ToString() == "1")
            {
                pgp1.Inlines.Clear();
            }
            else if (menu.Tag.ToString() == "2")
            {
                pgp2.Inlines.Clear();
            }
            else if (menu.Tag.ToString() == "3")
            {
                pgp3.Inlines.Clear();
            }
            else if (menu.Tag.ToString() == "4")
            {
                pgp4.Inlines.Clear();
            }
        }

        private SortedList<double, double> GetRefer(ParaObject para)
        {
            int port = int.Parse(para.Trace.ToString().Last().ToString());
            string path = para.ReferTracePath;
            if (!File.Exists(path))
            {
                LogMsg(port, Brushes.Red, string.Format("No found Cal File({0})", path));
                return null;
            }
            else
            {
                SortedList<double, double> list = new SortedList<double, double>();
                StreamReader sr = null;
                string str;
                string[] arr;

                try
                {
                    using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
                    {
                        sr = new StreamReader(fs);
                        while ((str = sr.ReadLine()) != null)
                        {
                            arr = str.Split(',');
                            list.Add(double.Parse(arr[0]), double.Parse(arr[1]));
                        }
                        sr.Close();
                        fs.Close();
                    }
                    para.MarkersCal = GetParaDataInTrace(list, para);
                    return list;
                }
                catch (Exception ex)
                {
                    if (sr != null)
                    {
                        sr.Close();
                        sr = null;
                    }
                    LogMsg(port, Brushes.Red, string.Format("Failed reading Cal File({0}), caused by: {1}", path, ex.Message));
                    return null;
                }
            }
        }
        private void UpdateAll()
        {
            tbAll.Text = GetPercentStr(Count1 + Count2 + Count3 + Count4, Pass1 + Pass2 + Pass3 + Pass4);
        }

        public string GetPercentStr(int count, int pass)
        {
            return string.Format("{0}/{1} | {2}%", pass, count, Math.Round(pass / (double)count * 100, 4));
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            StackPanel stack = Helper.GetVisualParent<StackPanel>(btn);
            int port = int.Parse(stack.Tag.ToString());
            switch (port)
            {
                case 1:
                    manual1 = true;
                    break;
                case 2:
                    manual2 = true;
                    break;
                case 3:
                    manual3 = true;
                    break;
                case 4:
                    manual4 = true;
                    break;
            }

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            Button btn = sender as Button;
            StackPanel stack = Helper.GetVisualParent<StackPanel>(btn);
            int port = int.Parse(stack.Tag.ToString());
            Stop(port);
            this.Cursor = Cursors.Arrow;
        }

        private void Stop(int port)
        {
            switch (port)
            {
                case 1:
                    if (t1 != null)
                    {
                        t1.Abort();
                        t1 = null;
                        ellipse1.Fill = Brushes.Gray;
                        blk1.Text = "停止";
                        btnStart1.IsEnabled = false;
                    }
                    break;
                case 2:
                    if (t2 != null)
                    {
                        t2.Abort();
                        t2 = null;
                        ellipse2.Fill = Brushes.Gray;
                        blk2.Text = "停止";
                        btnStart2.IsEnabled = false;
                    }
                    break;
                case 3:
                    if (t3 != null)
                    {
                        t3.Abort();
                        t3 = null;
                        ellipse3.Fill = Brushes.Gray;
                        blk3.Text = "停止";
                        btnStart3.IsEnabled = false;
                    }
                    break;
                case 4:
                    if (t4 != null)
                    {
                        t4.Abort();
                        t4 = null;
                        ellipse4.Fill = Brushes.Gray;
                        blk4.Text = "停止";
                        btnStart4.IsEnabled = false;
                    }
                    break;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Stop();
            Settings.Default.Save();
            Application.Current.Shutdown();
        }

        private void MenuItem_Click_2(object sender, RoutedEventArgs e)
        {
            Stop();
        }
        private void Stop()
        {
            Stop(1);
            Stop(2);
            Stop(3);
            Stop(4);
            State = State.Stoped;
        }

        #region LOG
        void link_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Uri.LocalPath);
        }
        private void LogMsg(int port, Brush bsh, string msg, string path = null)
        {
            Paragraph pgp = pgp1;
            RichTextBox rtb = rtb1;
            switch (port)
            {
                case 1:
                    pgp = pgp1;
                    rtb = rtb1;
                    break;
                case 2:
                    pgp = pgp2;
                    rtb = rtb2;
                    break;
                case 3:
                    pgp = pgp3;
                    rtb = rtb3;
                    break;
                case 4:
                    pgp = pgp4;
                    rtb = rtb4;
                    break;
            }
            if (pgp.Inlines.Count > 1000)
            {
                pgp.Inlines.Clear();
            }
            Run run = new Run(DateTime.Now.ToString("HH:mm:ss >> "));
            run.Foreground = Brushes.Purple;
            run.FontFamily = new FontFamily("Batang,Arial");
            pgp.Inlines.Add(run);

            run = new Run(msg);
            run.Foreground = bsh;
            pgp.Inlines.Add(run);

            Uri u;
            if (path != null && Uri.TryCreate(path, UriKind.Absolute, out u))
            {
                run = new Run(System.IO.Path.GetFileName(path));
                run.Foreground = bsh;
                Hyperlink link = new Hyperlink(run);
                link.ToolTip = u.LocalPath;
                link.NavigateUri = u;
                link.RequestNavigate += new System.Windows.Navigation.RequestNavigateEventHandler(link_RequestNavigate);
                ToolTipService.SetInitialShowDelay(link, 2000);
                pgp.Inlines.Add(link);
            }

            LineBreak br = new LineBreak();
            pgp.Inlines.Add(br);
            rtb.ScrollToEnd();
        }
        private void LogMsg(int port, bool isOK, string msg, string path = null)
        {
            Brush bsh = isOK ? Brushes.Blue : Brushes.Red;
            LogMsg(port, bsh, msg, path);
        }
        private void LogMsg(string trace, bool isOK, string msg, string path = null)
        {
            Brush bsh = isOK ? Brushes.Blue : Brushes.Red;
            int port = int.Parse(trace.Last().ToString());
            btnStart1.Dispatcher.BeginInvoke(new Action(delegate
            {
                LogMsg(port, bsh, msg, path);
            }));
        }
        #endregion

        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            if (Settings.Default.TraceFormat == TraceFormat.LOG.ToString())
            {
                DataAccess.Report_LOG();
            }
            else
            {
                Report2();
            }

            this.Cursor = Cursors.Arrow;
        }

        //LOG Report

        //SWR Report
        private void Report2()
        {
            List<SingleData> list1, list2, list3, list4;
            DataAccess.GetSingleData_LOG(Settings.Default.OutputDir, Settings.Default.Para1.Trace, out Count1, out Pass1, out list1);
            DataAccess.GetSingleData_LOG(Settings.Default.OutputDir, Settings.Default.Para2.Trace, out Count2, out Pass2, out list2);
            DataAccess.GetSingleData_LOG(Settings.Default.OutputDir, Settings.Default.Para3.Trace, out Count3, out Pass3, out list3);
            DataAccess.GetSingleData_LOG(Settings.Default.OutputDir, Settings.Default.Para4.Trace, out Count4, out Pass4, out list4);

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
                sh.Cells[r, 2] = Pass1 + "/" + Count1;
                sh.Cells[r, 3] = (Count1 == 0 ? 0 : Math.Round(Pass1 / (double)Count1 * 100, 4)) + "%";

                sh.Cells[++r, 1] = "Port2(S22)";
                sh.Cells[r, 2] = Pass2 + "/" + Count2;
                sh.Cells[r, 3] = (Count2 == 0 ? 0 : Math.Round(Pass2 / (double)Count2 * 100, 4)) + "%";

                sh.Cells[++r, 1] = "Port3(S33)";
                sh.Cells[r, 2] = Pass3 + "/" + Count3;
                sh.Cells[r, 3] = (Count3 == 0 ? 0 : Math.Round(Pass3 / (double)Count3 * 100, 4)) + "%";

                sh.Cells[++r, 1] = "Port4(S44)";
                sh.Cells[r, 2] = Pass4 + "/" + Count4;
                sh.Cells[r, 3] = (Count4 == 0 ? 0 : Math.Round(Pass4 / (double)Count4 * 100, 4)) + "%";

                int pass = Pass1 + Pass2 + Pass3 + Pass4;
                int count = Count1 + Count2 + Count3 + Count4;
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
        private void MenuItem_Click_4(object sender, RoutedEventArgs e)
        {
            if (State == State.Running)
            {
                MessageBox.Show("正在测试...", "Tips", MessageBoxButton.OK);
            }
            else
            {
                if (MessageBox.Show("Are you sure to reset?", "Tips", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    Settings.Default.Reset();
                    Settings.Default.Para1 = new ParaObject();
                    Settings.Default.Para2 = new ParaObject();
                    Settings.Default.Para3 = new ParaObject();
                    Settings.Default.Para4 = new ParaObject();
                    Settings.Default.Para1.Trace = Trace.S11.ToString();
                    Settings.Default.Para2.Trace = Trace.S22.ToString();
                    Settings.Default.Para3.Trace = Trace.S33.ToString();
                    Settings.Default.Para4.Trace = Trace.S44.ToString();

                    Settings.Default.Save();
                }
            }

        }

        private void MenuItem_Click_5(object sender, RoutedEventArgs e)
        {
            AboutWin about = new AboutWin();
            about.Owner = this;
            about.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
            about.ShowDialog();
        }

    }
}
