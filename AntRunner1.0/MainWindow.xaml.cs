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
        public static MainWindow Self;
        public VNA_AT5071C VNA = null;
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
            List<SingleData> temp;
            GetCount(Settings.Default.OutputDir, Settings.Default.Para1?.Trace, out Count1, out Pass1, out temp);
            GetCount(Settings.Default.OutputDir, Settings.Default.Para2?.Trace, out Count2, out Pass2, out temp);
            GetCount(Settings.Default.OutputDir, Settings.Default.Para3?.Trace, out Count3, out Pass3, out temp);
            GetCount(Settings.Default.OutputDir, Settings.Default.Para4?.Trace, out Count4, out Pass4, out temp);
            tb1.Text = GetPercentStr(Count1, Pass1);
            tb2.Text = GetPercentStr(Count2, Pass2);
            tb3.Text = GetPercentStr(Count3, Pass3);
            tb4.Text = GetPercentStr(Count4, Pass4);
            tbAll.Text = GetPercentStr(Count1 + Count2 + Count3 + Count4, Pass1 + Pass2 + Pass3 + Pass4);
        }
        private void GetCount(string path, string trace, out int count, out int pass, out List<SingleData> list)
        {
            count = 0;
            pass = 0;
            list = new List<SingleData>();
            string dir = string.Format("{0}\\{1}", path, trace);
            if (!Directory.Exists(dir))
                return;
            string[] files = Directory.GetFiles(dir);
            string str;
            SingleData single = null;
            SortedList<double, double> listData = null;
            SortedList<double, double> referData = null;
            StreamReader sr = null;
            try
            {
                foreach (string file in files)
                {
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
                        single.Code = sr.ReadLine().Split(',')[1];
                        single.TraceType = sr.ReadLine().Split(',')[1];
                        sr.ReadLine();
                        single.Memo = sr.ReadLine().Split(',')[1];
                        sr.ReadLine();
                        sr.ReadLine();
                        listData = new SortedList<double, double>();
                        referData = new SortedList<double, double>();
                        while ((str = sr.ReadLine()) != null)
                        {
                            listData.Add(double.Parse(str.Split(',')[0]), double.Parse(str.Split(',')[1]));
                            referData.Add(double.Parse(str.Split(',')[0]), double.Parse(str.Split(',')[2]));
                        }
                        single.ListData = listData;
                        single.ReferData = referData;
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
                if (sr != null)
                {
                    sr.Close();
                    sr = null;
                }
            }
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            SetupWin win = new SetupWin();
            win.Owner = this;
            win.ShowDialog(); 
        }
        private void MenuItem_Click_3(object sender, RoutedEventArgs e)
        {
            if (State == State.Running)
            {
                MessageBox.Show(this, "Testing...", "Tips", MessageBoxButton.OK);
                return;
            }
            if (!SetupVNA())
                return;

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
        private bool SetupVNA()
        {

            if (VNA == null || !VNA.IsOK)
            {
                VNA = VNA_AT5071C.GetInstance();

                if (IsSkip) return true;
                if (!VNA.Init(Settings.Default.GPIB))
                {
                    return false;
                }
            }
            VNA.Setup();
            return true;

        }
        private void Start(ParaObject para, ref SortedList<double, double> refer, ref Ellipse elli, ref Button btn, ref Thread td)
        {
            if (para.Enable)
            {
                if (VNA != null)
                {
                    VNA.Setup(para);
                }
                elli.Fill = bshPass; 
                btn.IsEnabled = true;
                State = State.Running;
                if (para.Trace == "S11")
                    td = new Thread(ThreadStart1);
                else if (para.Trace == "S22")
                    td = new Thread(ThreadStart2);
                else if (para.Trace == "S33")
                    td = new Thread(ThreadStart3);
                else if (para.Trace == "S44")
                    td = new Thread(ThreadStart4);

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
            bool ok;
            SortedList<double, double> list = null;
            btnStart1.Dispatcher.BeginInvoke(new Action(delegate
            {
                btnStart1.IsEnabled = false;
                ellipse1.Fill = bshTgr;
            }));
            lock (VNA)
            {
                list = VNA.ReadSWR(Settings.Default.Para1);
                ok = CheckPass(list, Settings.Default.Para1);
                Output(Settings.Default.Para1, list, ok);
                Count1++;
            }
            this.Dispatcher.BeginInvoke(new Action(delegate
            {
                if (ok) Pass1++;
                ellipse1.Fill = ok ? bshPass : bshFail;
                blk1.Text = ok ? "合格" : "不合格";
                tb1.Text = GetPercentStr(Count1, Pass1);
                UpdateAll();
                btnStart1.IsEnabled = true;
            }));
        }
        private void Run2()
        {
            bool ok;
            SortedList<double, double> list = null;
            btnStart2.Dispatcher.BeginInvoke(new Action(delegate
            {
                btnStart2.IsEnabled = false;
                ellipse2.Fill = bshTgr;
            }));
            lock (VNA)
            {
                list = VNA.ReadSWR(Settings.Default.Para2);
                ok = CheckPass(list, Settings.Default.Para2);
                Output(Settings.Default.Para2, list, ok);
                Count2++;
            }
            this.Dispatcher.BeginInvoke(new Action(delegate
            {
                if (ok) Pass2++;
                ellipse2.Fill = ok ? bshPass : bshFail;
                blk2.Text = ok ? "合格" : "不合格";
                tb2.Text = GetPercentStr(Count2, Pass2);
                UpdateAll();
                btnStart2.IsEnabled = true;
            }));
        }
        private void Run3()
        {
            bool ok;
            SortedList<double, double> list = null;
            btnStart3.Dispatcher.BeginInvoke(new Action(delegate
            {
                btnStart3.IsEnabled = false;
                ellipse3.Fill = bshTgr;
            }));
            lock (VNA)
            {
                list = VNA.ReadSWR(Settings.Default.Para3);
                ok = CheckPass(list, Settings.Default.Para3);
                Output(Settings.Default.Para3, list, ok);
                Count3++;
            }
            this.Dispatcher.BeginInvoke(new Action(delegate
            {
                if (ok) Pass3++;
                ellipse3.Fill = ok ? bshPass : bshFail;
                blk3.Text = ok ? "合格" : "不合格";
                tb3.Text = GetPercentStr(Count3, Pass3);
                UpdateAll();
                btnStart3.IsEnabled = true;
            }));
        }
        private void Run4()
        {
            bool ok;
            SortedList<double, double> list = null;
            btnStart4.Dispatcher.BeginInvoke(new Action(delegate
            {
                btnStart4.IsEnabled = false;
                ellipse4.Fill = bshTgr;
            }));
            lock (VNA)
            {
                list = VNA.ReadSWR(Settings.Default.Para4);
                ok = CheckPass(list, Settings.Default.Para4);
                Output(Settings.Default.Para4, list, ok);
                Count4++;
            }
            this.Dispatcher.BeginInvoke(new Action(delegate
            {
                if (ok) Pass4++;
                ellipse4.Fill = ok ? bshPass : bshFail;
                blk4.Text = ok ? "合格" : "不合格";
                tb4.Text = GetPercentStr(Count4, Pass4);
                UpdateAll();
                btnStart4.IsEnabled = true;
            }));
        }
        private bool AutoTriger(ParaObject para)
        {
            SortedList<double, double> refer = null;
            Trace type = (Trace)Enum.Parse(typeof(Trace), para.Trace);
            switch (type)
            {
                case Trace.S11:
                    refer = refer1;
                    break;
                case Trace.S22:
                    refer = refer2;
                    break;
                case Trace.S33:
                    refer = refer3;
                    break;
                case Trace.S44:
                    refer = refer4;
                    break;
            }

            SortedList<double, double> listPre = null;
            SortedList<double, double> list = null;
            List<SortedList<double, double>> all = new List<SortedList<double, double>>();
            bool firstDeep = false;
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
                lock (VNA)
                {
                    list = VNA.ReadSWRByTrace(para);
                }
                if (listPre == null)
                {
                    listPre = list;
                    continue;
                }
                else
                {
                    if (IsDeep(list, Settings.Default.Deep))
                    {
                        firstDeep = true;
                        continue;
                    }

                    if (!firstDeep && !IsSkip)
                        continue;

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
                    if (Settings.Default.MatchCnt < 2)
                        Settings.Default.MatchCnt = 2;
                    if (all.Count >= Settings.Default.MatchCnt)
                    {
                        for (int i = Settings.Default.MatchCnt - 1; i >= 0; i--)
                        {
                            if (!MatchTrace(list, all[i], Settings.Default.AutoDiff))
                                continue;
                        }
                        return true;
                    }
                    all.Add(list);
                }
                listPre = list;
            }
        }
        private bool MatchTrace(SortedList<double, double> trace1, SortedList<double, double> trace2, double prec)
        {
            if (trace1.Count() != trace2.Count())
                return false;
            else
            {
                for (int i = 0; i < trace1.Count(); i++)
                {
                    if (trace1.Values[i] - trace2.Values[i] > prec)
                        return false;
                }
                return true;
            }
        }

        private bool IsDeep(SortedList<double, double> trace, double deep)
        {
            foreach (KeyValuePair<double, double> item in trace)
            {
                if (item.Value < deep)
                    return false;
            }
            return true;
        }

        private bool CheckPass(SortedList<double, double> list, ParaObject para)
        {
            return CheckPass2(list, para);
        }

        private bool CheckPass1(SortedList<double, double> list, ParaObject para)
        {
            double refer, min, max;
            foreach (KeyValuePair<double, double> item in list)
            {
                refer = GetRefer(para, item.Key);
                min = refer - para.ReferDiff;
                max = refer + para.ReferDiff;
                if (item.Value < min || item.Value > max)
                    return false;
            }
            return true;
        }

        private bool CheckPass2(SortedList<double, double> list, ParaObject para)
        {
            double minVal = double.NaN;
            double minValFreq = double.NaN;
            foreach (KeyValuePair<double, double> item in list)
            {
                if (minVal == double.NaN)
                {
                    minVal = item.Value;
                    minValFreq = item.Key;
                }
                else
                {
                    if (item.Value < minVal)
                    {
                        minVal = item.Value;
                        minValFreq = item.Key;
                    }
                }
            }
            if (Math.Abs(minValFreq - para.ReferFreq) > Math.Abs(para.DiffFreq))
                return false;

            if (Math.Abs(minVal - para.ReferPower) > Math.Abs(para.DiffPower))
                return false;

            return true;
        }

        private double GetRefer(ParaObject para, double freq)
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
            double refer = GetRefer(trace, freq);
            return refer;
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
        private string Output(ParaObject para, SortedList<double, double> list, bool isOK)
        {
            StreamWriter writer = null;
            try
            {
                string path = GetFilePath(para, isOK);
                using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    writer = new StreamWriter(fs);
                    writer.WriteLine("Result," + (isOK ? "Pass" : "Fail"));
                    writer.WriteLine("Code," + Settings.Default.Code);
                    writer.WriteLine("Trace," + para.Trace);
                    writer.WriteLine("Refer Span," + para.ReferDiff);
                    writer.WriteLine("Memo," + Settings.Default.Memo);
                    writer.WriteLine();
                    writer.WriteLine("Frequency,Data,Cal");

                    double cal;
                    double freq, val;
                    for (int i = 0; i < list.Count; i++)
                    {
                        freq = list.Keys[i];
                        val = list.Values[i];
                        cal = 0;// GetRefer(para, freq);
                        writer.WriteLine("{0},{1},{2}", freq, val, cal);
                    }
                    writer.Flush();
                    writer.Close();
                    writer = null;
                    fs.Close();
                    LogMsg(para.Trace, isOK, (isOK ? "Pass" : "Fail") + " | Stored in : ", path);
                    return path;
                }
            }
            catch (Exception ex)
            {
                if (writer != null)
                {
                    writer.Flush();
                    writer.Close();
                }
                return null;
            }
        }
        private string GetFilePath(ParaObject para, bool pass)
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

        private void MenuItem_Click_6(object sender, RoutedEventArgs e)
        {
            MenuItem menu = sender as MenuItem;
            if(menu.Tag.ToString()=="1")
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
                    }
                    ellipse1.Fill = Brushes.Gray;
                    blk1.Text = "停止";
                    btnStart1.IsEnabled = false;
                    break;
                case 2:
                    if (t2 != null)
                    {
                        t2.Abort();
                        t2 = null;
                    }
                    ellipse2.Fill = Brushes.Gray;
                    blk2.Text = "停止";
                    btnStart2.IsEnabled = false;
                    break;
                case 3:
                    if (t3 != null)
                    {
                        t3.Abort();
                        t3 = null;
                    }
                    ellipse3.Fill = Brushes.Gray;
                    blk3.Text = "停止";
                    btnStart3.IsEnabled = false;
                    break;
                case 4:
                    if (t4 != null)
                    {
                        t4.Abort();
                        t4 = null;
                    }
                    ellipse4.Fill = Brushes.Gray;
                    blk4.Text = "停止";
                    btnStart4.IsEnabled = false;
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
            Report();
        }

        private void Report()
        {
            List<SingleData> list1, list2, list3, list4;
            GetCount(Settings.Default.OutputDir, Settings.Default.Para1.Trace, out Count1, out Pass1, out list1);
            GetCount(Settings.Default.OutputDir, Settings.Default.Para2.Trace, out Count2, out Pass2, out list2);
            GetCount(Settings.Default.OutputDir, Settings.Default.Para3.Trace, out Count3, out Pass3, out list3);
            GetCount(Settings.Default.OutputDir, Settings.Default.Para4.Trace, out Count4, out Pass4, out list4);

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
                sh.Cells[r, 1] = "Frequency Reference";
                sh.Cells[r, 2] = string.Format("{0} MHz", Settings.Default.Para1.ReferFreq);
                sh.Cells[r, 3] = string.Format("{0} MHz", Settings.Default.Para2.ReferFreq);
                sh.Cells[r, 4] = string.Format("{0} MHz", Settings.Default.Para3.ReferFreq);
                sh.Cells[r, 5] = string.Format("{0} MHz", Settings.Default.Para4.ReferFreq);
                r++;
                sh.Cells[r, 1] = "Frequency Difference";
                sh.Cells[r, 2] = string.Format("{0} MHz", Settings.Default.Para1.DiffFreq);
                sh.Cells[r, 3] = string.Format("{0} MHz", Settings.Default.Para2.DiffFreq);
                sh.Cells[r, 4] = string.Format("{0} MHz", Settings.Default.Para3.DiffFreq);
                sh.Cells[r, 5] = string.Format("{0} MHz", Settings.Default.Para4.DiffFreq);
                r++;
                sh.Cells[r, 1] = "Power Reference";
                sh.Cells[r, 2] = string.Format("{0} dBm", Settings.Default.Para1.ReferPower);
                sh.Cells[r, 3] = string.Format("{0} dBm", Settings.Default.Para2.ReferPower);
                sh.Cells[r, 4] = string.Format("{0} dBm", Settings.Default.Para3.ReferPower);
                sh.Cells[r, 5] = string.Format("{0} dBm", Settings.Default.Para4.ReferPower);
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
                MessageBox.Show("In testing, can not to reset.", "Tips", MessageBoxButton.OK);
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
