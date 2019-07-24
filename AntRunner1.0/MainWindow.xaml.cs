#define FUN1//结果图标闪动提示

using AntRunner.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using Form = System.Windows.Forms;

namespace AntRunner
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public static bool IsSkip = false;
        volatile int Count1, Count2, Count3, Count4;
        volatile int Pass1, Pass2, Pass3, Pass4;
        string code1, code2, code3, code4;
        Thread t1, t2, t3, t4;
        Thread ts1;
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
        public DateTime expireTime;
        Storyboard sb1, sb2, sb3, sb4;
        ReportWaitWin wReport;
        Thread tReport;
        DataBase da;
        KeyboardHook k_hook;
        StringBuilder readKeyText = new StringBuilder();

        public State State
        {
            get { return state; }
            set { state = value; }
        }

        public MainWindow()
        {
            AppLog.Error("Start");
            Self = this;
            InitializeComponent();
            BrushConverter brushConverter = new BrushConverter();
            bshPass = (Brush)brushConverter.ConvertFromString("#FF3092FC");
            bshTgr = (Brush)brushConverter.ConvertFromString("#FF28F31B");

            sb1 = CreateStoryboard(ellipse1, vb1);
            sb2 = CreateStoryboard(ellipse2, vb2);
            sb3 = CreateStoryboard(ellipse3, vb3);
            sb4 = CreateStoryboard(ellipse4, vb4);

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

            k_hook = new KeyboardHook();
            k_hook.KeyPressEvent += K_hook_KeyPressEvent;
            k_hook.Start();
        }

        void K_hook_KeyPressEvent(object sender, Form.KeyPressEventArgs e)
        {
            if (!this.IsActive)
            {
                return;
            }
            char key = e.KeyChar;
            if (key == '\r')
            {
                Settings.Default.Para1.Code = readKeyText.ToString();
                readKeyText.Clear();
            }
            else
            {
                readKeyText.Append(e.KeyChar);
            }
        }
        
        private Storyboard CreateStoryboard(Ellipse ele, Viewbox vb)
        {
            DoubleAnimation ani = new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(300)));
            //ani.AutoReverse = true;
            ani.RepeatBehavior = new RepeatBehavior(TimeSpan.FromSeconds(5));
            Storyboard.SetTarget(ani, ele);
            Storyboard.SetTargetProperty(ani, new PropertyPath(Ellipse.OpacityProperty.ToString()));

            DoubleAnimation ani2 = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(3000)));
            Storyboard.SetTarget(ani2, vb);
            Storyboard.SetTargetProperty(ani2, new PropertyPath(Ellipse.OpacityProperty.ToString()));

            Storyboard sb = new Storyboard();
#if FUN1
            sb.Children.Add(ani);
            sb.Children.Add(ani2);
#endif
            return sb;
        }
        private bool HasExpired()
        {
            FileStream fs = new FileStream("license.config", FileMode.Open, FileAccess.Read);
            StreamReader sr = new StreamReader(fs);
            string str = sr.ReadLine();
            str = Helper.Decode(str);
            expireTime = DateTime.Parse(str);
            return DateTime.Now > expireTime;
        }
        public void InitCount()
        {
            da = DataBase.GetDataHandler(Helper.String2Enum<TraceFormat>(Settings.Default.TraceFormat));
            da.GetCount8Pass(Settings.Default.OutputDir, Settings.Default.Para1.Trace, out Count1, out Pass1);
            da.GetCount8Pass(Settings.Default.OutputDir, Settings.Default.Para2.Trace, out Count2, out Pass2);
            da.GetCount8Pass(Settings.Default.OutputDir, Settings.Default.Para3.Trace, out Count3, out Pass3);
            da.GetCount8Pass(Settings.Default.OutputDir, Settings.Default.Para4.Trace, out Count4, out Pass4);
            tb1.Text = GetPercentStr(Count1, Pass1);
            tb2.Text = GetPercentStr(Count2, Pass2);
            tb3.Text = GetPercentStr(Count3, Pass3);
            tb4.Text = GetPercentStr(Count4, Pass4);
            tbAll.Text = GetPercentStr(Count1 + Count2 + Count3 + Count4, Pass1 + Pass2 + Pass3 + Pass4);
        }
        private bool InitVNA()
        {
            if (vna == null || !vna.IsOK)
            {
                vna = VNA.CreateVNA();

                try
                {
                    VNA.ScanGPIB();
                    MainWindow.IsSkip = false;
                }
                catch (Exception ex)
                {
                    AppLog.Warn("ScanGPIB has error.", ex);
                    MainWindow.IsSkip = true;
                }

                if (IsSkip) return true;
                if (!vna.Init(Settings.Default.GPIB))
                {
                    return false;
                }
            }
            vna.Config();
            return true;
        }
        private void Start()
        {
            try
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
                da = DataBase.GetDataHandler(Helper.String2Enum<TraceFormat>(Settings.Default.TraceFormat));
                Start(Settings.Default.Para1, ref refer1, ref ellipse1, ref btnStart1, ref t1);
                if (Settings.Default.TraceFormat == TraceFormat.LOG_SWR.ToString())
                    return;
                Start(Settings.Default.Para2, ref refer2, ref ellipse2, ref btnStart2, ref t2);
                Start(Settings.Default.Para3, ref refer3, ref ellipse3, ref btnStart3, ref t3);
                Start(Settings.Default.Para4, ref refer4, ref ellipse4, ref btnStart4, ref t4);
            }
            catch (Exception ex)
            {
                AppLog.Error("Test has error.", ex);
            }
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

                if (Settings.Default.UserScanner)
                {
                    readKeyText.Clear();
                    k_hook.Start();//安装键盘钩子
                }
            }
        }

        private void ThreadStart1()
        {
            blk1.Dispatcher.BeginInvoke(new Action(delegate
            {
                blk1.Text = "开始";
            }));
            while (true)
            {
                SortedList<double, double> list1 = null;
                SortedList<double, double> list2 = null;
                if (Settings.Default.TriggerType == TriggerType.Auto.ToString())
                {
                    AutoTriger(Settings.Default.Para1, out list1, out list2);
                }
                else if (Settings.Default.TriggerType == TriggerType.Scanner.ToString())
                {
                    return;
                }
                Run1(list1, list2);
                manual1 = false;
            }
        }
        private void ThreadStart2()
        {
            blk2.Dispatcher.BeginInvoke(new Action(delegate
            {
                blk2.Text = "开始";
            }));
            while (true)
            {
                SortedList<double, double> list1 = null;
                SortedList<double, double> list2 = null;
                if (Settings.Default.TriggerType == TriggerType.Auto.ToString())
                {
                    AutoTriger(Settings.Default.Para2, out list1, out list2);
                }
                else if (Settings.Default.TriggerType == TriggerType.Scanner.ToString())
                {
                    return;
                }
                Run2(list1, list2);
                manual2 = false;
            }
        }
        private void ThreadStart3()
        {
            blk3.Dispatcher.BeginInvoke(new Action(delegate
            {
                blk3.Text = "开始";
            }));
            while (true)
            {
                SortedList<double, double> list1 = null;
                SortedList<double, double> list2 = null;
                if (Settings.Default.TriggerType == TriggerType.Auto.ToString())
                {
                    AutoTriger(Settings.Default.Para3, out list1, out list2);
                }
                else if (Settings.Default.TriggerType == TriggerType.Scanner.ToString())
                {
                    return;
                }
                Run3(list1, list2);
                manual3 = false;
            }
        }
        private void ThreadStart4()
        {
            blk4.Dispatcher.BeginInvoke(new Action(delegate
            {
                blk4.Text = "开始";
            }));
            while (true)
            {
                SortedList<double, double> list1 = null;
                SortedList<double, double> list2 = null;
                if (Settings.Default.TriggerType == TriggerType.Auto.ToString())
                {
                    AutoTriger(Settings.Default.Para4, out list1, out list2);
                }
                else if (Settings.Default.TriggerType == TriggerType.Scanner.ToString())
                {
                    return;
                }
                Run4(list1, list2);
                manual4 = false;
            }
        }
        private void Run1(SortedList<double, double> list1, SortedList<double, double> list2)
        {
            bool pass;
            string path;
            List<ErrorCode> errors;
            btnStart1.Dispatcher.Invoke(new Action(delegate
            {
                btnStart1.IsEnabled = false;
                ellipse1.Fill = bshTgr;
            }));
            lock (vna)
            {
                path = CheckPass(Settings.Default.Para1, list1, list2, out errors);
                pass = errors.Count == 0;
            }
            this.Dispatcher.Invoke(new Action(delegate
            {
                if (path == null)
                {
                    LogMsg(1, bshFail, "报错");
                    txt1.Text = blk1.Text = "报错";
                    txt1.Foreground = ellipse1.Fill = bshFail;
                    sb1.Begin();
                    btnStart1.IsEnabled = true;
                }
                else
                {
                    LogMsg(Settings.Default.Para1.Trace, path, errors);
                    if (pass) Pass1++;
                    Count1++;
                    txt1.Text = blk1.Text = pass ? "合格" : "淘汰";
                    txt1.Foreground = ellipse1.Fill = pass ? bshPass : bshFail;
                    tb1.Text = GetPercentStr(Count1, Pass1);
                    sb1.Begin();
                    UpdateAll();
                    btnStart1.IsEnabled = true;
                }
            }));
        }
        private void Run2(SortedList<double, double> list1, SortedList<double, double> list2)
        {
            bool pass;
            string path;
            List<ErrorCode> errors;
            btnStart2.Dispatcher.Invoke(new Action(delegate
            {
                btnStart2.IsEnabled = false;
                ellipse2.Fill = bshTgr;
            }));
            lock (vna)
            {
                path = CheckPass(Settings.Default.Para2, list1, list2, out errors);
                pass = errors.Count == 0;
            }
            this.Dispatcher.Invoke(new Action(delegate
            {
                if (path == null)
                {
                    LogMsg(2, bshFail, "报错");
                    txt2.Text = blk2.Text = "报错";
                    txt2.Foreground = ellipse2.Fill = bshFail;
                    sb2.Begin();
                    btnStart2.IsEnabled = true;
                }
                else
                {
                    LogMsg(Settings.Default.Para2.Trace, path, errors);
                    if (pass) Pass2++;
                    Count2++;
                    txt2.Text = blk2.Text = pass ? "合格" : "淘汰";
                    txt2.Foreground = ellipse2.Fill = pass ? bshPass : bshFail;
                    tb2.Text = GetPercentStr(Count2, Pass2);
                    sb2.Begin();
                    UpdateAll();
                    btnStart2.IsEnabled = true;
                }
            }));
        }
        private void Run3(SortedList<double, double> list1, SortedList<double, double> list2)
        {
            bool pass;
            string path;
            List<ErrorCode> errors;
            btnStart3.Dispatcher.Invoke(new Action(delegate
            {
                btnStart3.IsEnabled = false;
                ellipse3.Fill = bshTgr;
            }));
            lock (vna)
            {
                path = CheckPass(Settings.Default.Para3, list1, list2, out errors);
                pass = errors.Count == 0;
            }
            this.Dispatcher.Invoke(new Action(delegate
            {
                if (path == null)
                {
                    LogMsg(3, bshFail, "报错");
                    txt3.Text = blk3.Text = "报错";
                    txt3.Foreground = ellipse3.Fill = bshFail;
                    sb3.Begin();
                    btnStart3.IsEnabled = true;
                }
                else
                {
                    LogMsg(Settings.Default.Para3.Trace, path, errors);
                    if (pass) Pass3++;
                    Count3++;
                    txt3.Text = blk3.Text = pass ? "合格" : "淘汰";
                    txt3.Foreground = ellipse3.Fill = pass ? bshPass : bshFail;
                    tb3.Text = GetPercentStr(Count3, Pass3);
                    sb3.Begin();
                    UpdateAll();
                    btnStart3.IsEnabled = true;
                }
            }));
        }
        private void Run4(SortedList<double, double> list1, SortedList<double, double> list2)
        {
            bool pass;
            string path;
            List<ErrorCode> errors;
            btnStart4.Dispatcher.Invoke(new Action(delegate
            {
                btnStart4.IsEnabled = false;
                ellipse4.Fill = bshTgr;
            }));
            lock (vna)
            {
                path = CheckPass(Settings.Default.Para4, list1, list2, out errors);
                pass = errors.Count == 0;
            }
            this.Dispatcher.Invoke(new Action(delegate
            {
                if (path == null)
                {
                    LogMsg(4, bshFail, "报错");
                    txt4.Text = blk4.Text = "报错";
                    txt4.Foreground = ellipse4.Fill = bshFail;
                    sb4.Begin();
                    btnStart4.IsEnabled = true;
                }
                else
                {
                    LogMsg(Settings.Default.Para4.Trace, path, errors);
                    if (pass) Pass4++;
                    Count4++;
                    txt4.Text = blk4.Text = pass ? "合格" : "淘汰";
                    txt4.Foreground = ellipse4.Fill = pass ? bshPass : bshFail;
                    tb4.Text = GetPercentStr(Count4, Pass4);
                    sb4.Begin();
                    UpdateAll();
                    btnStart4.IsEnabled = true;
                }
            }));
        }
        private bool AutoTriger(ParaObject para, out SortedList<double, double> list1, out SortedList<double, double> list2)
        {
            Trace type = (Trace)Enum.Parse(typeof(Trace), para.Trace);
            SortedList<double, double> list = null;
            list1 = null;
            list2 = null;
            List<SortedList<double, double>> all = new List<SortedList<double, double>>();
            bool triggerDeep = false;
            while (true)
            {
                if (type == Trace.S11 && manual1)
                    goto AA;
                else if (type == Trace.S22 && manual2)
                    goto AA;
                if (type == Trace.S33 && manual3)
                    goto AA;
                if (type == Trace.S44 && manual4)
                    goto AA;
                Thread.Sleep(Settings.Default.AutoDelay);
                lock (vna)
                {
                    if (Settings.Default.TraceFormat == TraceFormat.LOG_SWR.ToString())
                    {
                        list = vna.ReadTrace(para, 2);//读曲线
                    }
                    else
                    {
                        list = vna.ReadTrace(para);//读曲线
                    }

                }
                if (da.IsDeep(list, Settings.Default.Deep))//如果是底噪，即触发过底噪，之前数据清空，路过
                {
                    triggerDeep = true;
                    all.Clear();

                    //如果触发过底噪（说明重新拎过天线），且数据不是底噪，这个数据才是可用的
                    btnStart1.Dispatcher.BeginInvoke(new Action(delegate
                    {
                        switch (type)
                        {
                            case Trace.S11:
                                ellipse1.Fill = bshTgr;
                                sb1.Pause(ellipse1);
                                blk1.Text = "正在触发 . . .";
                                break;
                            case Trace.S22:
                                ellipse2.Fill = bshTgr;
                                sb2.Pause(ellipse2);
                                blk2.Text = "正在触发 . . .";
                                break;
                            case Trace.S33:
                                ellipse3.Fill = bshTgr;
                                sb3.Pause(ellipse3);
                                blk3.Text = "正在触发 . . .";
                                break;
                            case Trace.S44:
                                ellipse4.Fill = bshTgr;
                                sb4.Pause(ellipse4);
                                blk4.Text = "正在触发 . . .";
                                break;
                        }
                    }));
                    continue;
                }

                if (!triggerDeep && !IsSkip)//如果没有触发过底噪，跳过
                {
                    continue;
                }
                all.Add(list);
                if (all.Count >= Settings.Default.MatchCnt)
                {
                    //如果所有曲线全匹配，ok，否则把最旧的数据删掉，再测一条曲线
                    if (MatchTrace(all, Settings.Default.AutoDiff))
                    {
                        if (Settings.Default.TraceFormat == TraceFormat.LOG_SWR.ToString())
                        {
                            list2 = list;
                            list1 = vna.ReadTrace(para, 1);//读曲线
                        }
                        else
                        {
                            list1 = list;
                        }
                        return true;
                    }
                    else
                    {
                        all.RemoveAt(0);
                    }
                }
            }
        AA:
            if (Settings.Default.TraceFormat == TraceFormat.LOG_SWR.ToString())
            {
                list2 = list;
                list1 = vna.ReadTrace(para, 1);//读曲线
            }
            else
            {
                list1 = list;
            }
            return true;
        }
        /// <summary>
        /// 多条曲线偏差比较，偏差必须在指定范围内
        /// </summary>
        /// <param name="lists">曲线列表</param>
        /// <param name="diff">偏差值</param>
        /// <returns></returns>
        private bool MatchTrace(List<SortedList<double, double>> lists, double diff)
        {
            if (lists.Count() <= 1)
                return true;

            diff = Math.Abs(diff);
            int traceCnt = lists.Count();
            List<double> tempList = new List<double>();
            for (int i = 0; i < lists[0].Count(); i++)
            {
                tempList.Clear();
                for (int j = 0; j < traceCnt - 1; j++)
                {
                    tempList.Add(lists[j].Values[i]);
                }
                if (tempList.Max() - tempList.Min() > diff)
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
        private void Setup()
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
        private void Report()
        {
            da = DataBase.GetDataHandler(Helper.String2Enum<TraceFormat>(Settings.Default.TraceFormat));
            da.Report();
            wReport.Dispatcher.Invoke(new Action(delegate
            {
                wReport.Close();
            }));
        }
        private void About()
        {
            AboutWin about = new AboutWin();
            about.Owner = this;
            about.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            about.ShowDialog();
        }
        private void StackPanel_Click_1(object sender, RoutedEventArgs e)
        {
            MenuItem mi = e.OriginalSource as MenuItem;
            XmlElement xe = mi.Header as XmlElement;
            string name = xe.Attributes["Name"].Value;

            if (name == "设置")
            {
                Setup();
            }
            else if (name == "导出报告")
            {
                tReport = new Thread(new ThreadStart(new Action(delegate
                {
                    Report();
                })));
                tReport.Start();

                if (Settings.Default.TraceFormat == TraceFormat.LOG_SWR.ToString())
                    wReport = new ReportWaitWin(Count1);
                else
                    wReport = new ReportWaitWin(Count1 + Count2 + Count3 + Count4);
                wReport.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                wReport.Closing += Win_Closing;
                wReport.Owner = this;
                wReport.ShowDialog();
            }
            else if (name == "开始测试")
            {
                Start();
            }
            else if (name == "停止测试")
            {
                Stop();
            }
            else if (name == "关于")
            {
                About();
            }
        }

        private void Win_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            tReport.Abort();
        }
        #region Check
        private string CheckPass(ParaObject para, SortedList<double, double> list1, SortedList<double, double> list2, out List<ErrorCode> errors)
        {
            bool pass;
            string path = null;
            try
            {
                if (Settings.Default.TraceFormat == TraceFormat.SWR.ToString())
                {
                    SortedList<double, double> list = new SortedList<double, double>();
                    List<double> markers = GetMarker(para.MarkerText);
                    foreach (double marker in markers)
                    {
                        list.Add(marker, GetPointByTrace(list1, marker));
                    }
                    pass = CheckPass_SWR(list, para, out errors);
                    path = da.Output(para, list, errors);
                    return path;
                }
                else if (Settings.Default.TraceFormat == TraceFormat.LOG.ToString())
                {
                    pass = CheckPass_LOG(list1, para, out errors);
                    path = da.Output(para, list1, errors);
                    return path;
                }
                else if (Settings.Default.TraceFormat == TraceFormat.LOG_SWR.ToString())
                {
                    pass = CheckPass_LOG8SWR(list1, list2, para, out errors);
                    path = da.Output(para, errors);
                    return path;
                }
                else
                {
                    errors = new List<ErrorCode>();
                    return null;
                }
            }
            catch (Exception ex)
            {
                AppLog.Error("CheckPass has error.", ex);
                errors = new List<ErrorCode>();
                return null;
            }
        }

        private bool CheckPass_SWR(SortedList<double, double> list, ParaObject para, out List<ErrorCode> errorCode)
        {
            errorCode = new List<ErrorCode>();
            SortedList<double, double> referTrace = GetReferTrace(para);
            double refer, min, max;
            foreach (KeyValuePair<double, double> item in list)
            {
                refer = GetPointByTrace(referTrace, item.Key);
                min = refer - para.ReferDiff;
                max = refer + para.ReferDiff;
                if (item.Value < min)
                {
                    errorCode.Add(ErrorCode.PowL);
                }
                if (item.Value > max)
                {
                    errorCode.Add(ErrorCode.PowH);
                }
            }
            return errorCode.Count == 0;
        }

        private bool CheckPass_LOG(SortedList<double, double> trace, ParaObject para, out List<ErrorCode> errorCode)
        {
            errorCode = new List<ErrorCode>();
            double powRef, fqRef, powMin, fqMin;
            SortedList<double, double> referTrace = GetReferTrace(para);
            GetTraceMin(referTrace, out fqRef, out powRef);
            GetTraceMin(trace, out fqMin, out powMin);
            double diffFreq = Math.Abs(para.DiffFreq);
            double diffPower = Math.Abs(para.DiffPower);
            double diffFreqBad = Math.Abs(para.DiffFreq_Bad);
            double diffPowerBad = Math.Abs(para.DiffPower_Bad);
            //errorCode.Add(ErrorCode.FreqL);
            //errorCode.Add(ErrorCode.PowH);
            //检查最低点的频率偏差(横向比较)
            if (fqMin < fqRef - diffFreq)
            {
                errorCode.Add(ErrorCode.FreqL);
            }
            if (fqMin > fqRef + diffFreq)
            {
                errorCode.Add(ErrorCode.FreqH);
            }

            //检查最低点的功能偏差（纵向比较）
            if (powMin < powRef - diffPower)
            {
                errorCode.Add(ErrorCode.PowL);
            }
            if (powMin > powRef + diffPower)
            {
                errorCode.Add(ErrorCode.PowH);
            }

            //检查功率切线的频宽（频宽比较）
            ParaObject para2 = new ParaObject();
            para2.CutPow = para.CutPow;
            para.Markers = GetMarkersInTrace(trace, para2);
            double diffBW = Math.Abs(para.DiffBW);
            if (para2.CutBW < para.CutBW - diffBW)
            {
                errorCode.Add(ErrorCode.FreqBandWidthL);
            }
            if (para2.CutBW > para.CutBW + diffBW)
            {
                errorCode.Add(ErrorCode.FreqBandWidthH);
            }

            //检查短路
            if (powMin < powRef - diffPowerBad || powMin > powRef + diffPowerBad
                || fqMin < fqRef - diffFreqBad || fqMin > fqRef + diffFreqBad)
            {
                errorCode.Add(ErrorCode.Bad);
            }
            return errorCode.Count == 0;
        }

        private bool CheckPass_LOG8SWR(SortedList<double, double> s21, SortedList<double, double> s22, ParaObject para, out List<ErrorCode> errorCode)
        {
            errorCode = new List<ErrorCode>();

            if (para.MarkerType == MarkerType.Points.ToString())
            {
                string[] arr = para.MarkerText.Trim().Split('\n');
                arr = arr.Distinct().ToArray();
                Dictionary<double, string> markers = new Dictionary<double, string>();
                para.Markers = markers;
                foreach (string str in arr)
                {
                    double freq = double.Parse(str);
                    double p21 = GetPointByTrace(s21, freq);
                    double p22 = GetPointByTrace(s22, freq);
                    markers.Add(freq, string.Format("{0},{1}", p21, p22));
                    if (p21 < para.S21Min)
                    {
                        if (!errorCode.Contains(ErrorCode.PowS21L))
                            errorCode.Add(ErrorCode.PowS21L);
                    }
                    if (p21 > para.S21Max)
                    {
                        if (!errorCode.Contains(ErrorCode.PowS21H))
                            errorCode.Add(ErrorCode.PowS21H);
                    }
                    if (p22 > para.S22Max)
                    {
                        if (!errorCode.Contains(ErrorCode.StandingWaveS22H))
                            errorCode.Add(ErrorCode.StandingWaveS22H);
                    }
                }
            }
            else
            {
                Dictionary<double, double> s21List = new Dictionary<double, double>();
                Dictionary<double, double> s22List = new Dictionary<double, double>();
                Dictionary<double, string> markers = new Dictionary<double, string>();
                para.Markers = markers;
                for (int i = 0; i < s21.Count; i++)
                {
                    double freq = s21.Keys[i];
                    if (freq >= para.MarkerStart && freq <= para.MarkerStop)
                    {
                        s21List.Add(freq, s21.Values[i]);
                        s22List.Add(freq, s22.Values[i]);
                        markers.Add(freq, string.Format("{0},{1}", s21.Values[i], s22.Values[i]));
                    }
                }
                if (s21List.Where((i) => i.Value < para.S21Min).Count() > 0)
                {
                    errorCode.Add(ErrorCode.PowS21L);
                }
                if (s21List.Where((i) => i.Value > para.S21Max).Count() > 0)
                {
                    errorCode.Add(ErrorCode.PowS21H);
                }
                if (s22List.Where((i) => i.Value > para.S22Max).Count() > 0)
                {
                    errorCode.Add(ErrorCode.StandingWaveS22H);
                }
            }
            return errorCode.Count == 0;
        }

        #endregion

        /// <summary>
        /// 读曲线中的Marker,(共3个Marker：最低点、功能切线左右两点)
        /// </summary>
        /// <param name="trace"></param>
        /// <param name="para"></param>
        /// <returns></returns>
        private Dictionary<double, double> GetMarkersInTrace(SortedList<double, double> trace, ParaObject para)
        {
            double minFreq, minPower;
            GetTraceMin(trace, out minFreq, out minPower);
            Dictionary<double, double> markers = new Dictionary<double, double>();
            markers.Add(minFreq, minPower);
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
            double cutF1 = 0;
            double cutF2 = 0;
            double[] values = trace1.Values.ToArray<double>();
            int index = LocateTrace(values, cutPower);
            if (values.Count() > 0)
            {
                if (index == 0)
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
            }

            values = trace2.Values.ToArray<double>();
            index = LocateTrace(values, cutPower);
            if (values.Count() > 0)
            {
                if (index == 0)
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

            if (!markers.ContainsKey(para.CutLeftFreq))
            {
                markers.Add(para.CutLeftFreq, para.CutPow);
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
            if (Settings.Default.TraceFormat == TraceFormat.LOG_SWR.ToString())
            {
                return new SortedList<double, double>();
            }
            if (IsSkip)
            {
                SortedList<double, double> list2 = new SortedList<double, double>();
                double freq = para.FreqStart;
                double step = (para.FreqStop - para.FreqStart) / (para.Points - 1);
                for (int i = 0; i < para.Points; i++, freq += step)
                    list2.Add(freq, 10);
                return list2;
            }
            int port = int.Parse(para.Trace.ToString().Last().ToString());
            string path = para.ReferTracePath;
            if (!File.Exists(path))
            {
                LogMsg(port, Brushes.Red, string.Format("没有校准文件：“{0}”", path));
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
                    para.MarkersCal = GetMarkersInTrace(list, para);
                    return list;
                }
                catch (Exception ex)
                {
                    if (sr != null)
                    {
                        sr.Close();
                        sr = null;
                    }
                    AppLog.Error("GetRefer has error.", ex);
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
                        sb1.Stop();
                        t1.Abort();
                        t1 = null;
                        ellipse1.Fill = Brushes.Gray;
                        blk1.Text = "停止";
                        btnStart1.IsEnabled = false;
                    }
                    if (ts1 != null)
                    {
                        ts1.Abort();
                        ts1 = null;
                    }
                    break;
                case 2:
                    if (t2 != null)
                    {
                        sb2.Stop();
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
                        sb3.Stop();
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
                        sb4.Stop();
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

        private void Stop()
        {
            k_hook.Stop();
            Stop(1);
            Stop(2);
            Stop(3);
            Stop(4);
            State = State.Stoped;
        }

        #region LOG
        void link_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Uri.LocalPath);
        }
        private void LogMsg(int port, Brush bsh, string msg)
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

            LineBreak br = new LineBreak();
            pgp.Inlines.Add(br);
            rtb.ScrollToEnd();
        }
        private void LogMsg(string trace, string path, List<ErrorCode> errors)
        {
            int port = int.Parse(trace.Last().ToString());
            bool isOK = errors == null || errors.Count == 0;
            Brush bsh = isOK ? Brushes.Blue : Brushes.Red;
            string pass = isOK ? "合格" : "淘汰";
            btnStart1.Dispatcher.BeginInvoke(new Action(delegate
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

                run = new Run(pass);
                run.Foreground = bsh;

                if (path != null)
                {
                    Uri u;
                    if (Uri.TryCreate(path, UriKind.Absolute, out u))
                    {
                        Hyperlink link = new Hyperlink(run);
                        link.ToolTip = u.LocalPath;
                        link.NavigateUri = u;
                        link.RequestNavigate += new System.Windows.Navigation.RequestNavigateEventHandler(link_RequestNavigate);
                        ToolTipService.SetInitialShowDelay(link, 2000);
                        pgp.Inlines.Add(link);
                    }
                    else
                    {
                        pgp.Inlines.Add(run);
                    }
                }

                StringBuilder errStr = new StringBuilder();
                foreach (ErrorCode e in errors)
                {
                    errStr.Append(" | " + Helper.GetEnumDescription(e));
                }
                run = new Run(errStr.ToString());
                run.Foreground = bsh;
                pgp.Inlines.Add(run);

                LineBreak br = new LineBreak();
                pgp.Inlines.Add(br);
                rtb.ScrollToEnd();
            }));
        }
        #endregion


        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            this.WindowState = this.WindowState == System.Windows.WindowState.Maximized ? System.Windows.WindowState.Normal : System.Windows.WindowState.Maximized;
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            this.WindowState = System.Windows.WindowState.Minimized;
        }
        //private void MenuItem_Click_4(object sender, RoutedEventArgs e)
        //{
        //    if (State == State.Running)
        //    {
        //        MessageBox.Show("正在测试...", "Tips", MessageBoxButton.OK);
        //    }
        //    else
        //    {
        //        if (MessageBox.Show("Are you sure to reset?", "Tips", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
        //        {
        //            Settings.Default.Reset();
        //            Settings.Default.Para1 = new ParaObject();
        //            Settings.Default.Para2 = new ParaObject();
        //            Settings.Default.Para3 = new ParaObject();
        //            Settings.Default.Para4 = new ParaObject();
        //            Settings.Default.Para1.Trace = Trace.S11.ToString();
        //            Settings.Default.Para2.Trace = Trace.S22.ToString();
        //            Settings.Default.Para3.Trace = Trace.S33.ToString();
        //            Settings.Default.Para4.Trace = Trace.S44.ToString();

        //            Settings.Default.Save();
        //        }
        //    }

        //} 

    }
}
