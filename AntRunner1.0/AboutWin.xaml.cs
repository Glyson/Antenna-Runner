using System;
using System.Collections.Generic;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;

namespace AntRunner
{
    /// <summary>
    /// AboutWin.xaml 的交互逻辑
    /// </summary>
    public partial class AboutWin : Window
    {
        SortedList<DateTime, Popup> list = new SortedList<DateTime, Popup>();
        List<string> listTips = new List<string>();
        public AboutWin()
        {
            InitializeComponent();
            tipTime.ToolTip = DateTime.Now.ToString("yyyy/MM/dd") + " - " + MainWindow.Self.expireTime.ToString("yyyy/MM/dd");
            ToolTipService.SetShowDuration(tipTime, 2000);
            ToolTipService.SetInitialShowDelay(tipTime, 10000);

            DispatcherTimer t = new DispatcherTimer();
            t.Interval = TimeSpan.FromMilliseconds(200);
            t.Tick += T_Tick;
            t.Start();

            listTips.Add("点我干嘛！");
            listTips.Add("你好调皮！");
            listTips.Add("点我干嘛！");
            listTips.Add("你好调皮！");
            listTips.Add("点我干嘛！");
            listTips.Add("你好调皮！");
            listTips.Add("点我没用，扫我！！");
            listTips.Add("点我没用，扫我！！");
            listTips.Add("你想干嘛!");
            listTips.Add("撩我？");
            listTips.Add("一边去！");
            listTips.Add("不要乱点哦！");
            listTips.Add("小傻瓜！");
            listTips.Add("你好坏！");
            listTips.Add("坏！");
            listTips.Add("扫我吗！");
            listTips.Add("扫我，一切皆有可能！");
            listTips.Add("不扫我何以扫天下！");
            listTips.Add("不扫我何以扫天下！");
            listTips.Add("快去工作！");
            listTips.Add("别玩了！");
            listTips.Add("你是帅哥还是美女？");
            listTips.Add("不懂扫我！");
            listTips.Add("小生恭候多时了！");
            listTips.Add("感觉人生到达了高潮！");
            listTips.Add("你很有探索精神哦！");
            listTips.Add("这都被你发现了！");

        }

        private void T_Tick(object sender, EventArgs e)
        {
            if (list.Count == 0) return;
            DateTime time;
            DateTime now = DateTime.Now;
            for (int i = list.Count - 1; i >= 0; i--)
            {
                time = list.Keys[i];
                double span = now.Subtract(time).TotalMilliseconds;
                if (span >= 2000)
                {
                    list.Values[i].IsOpen = false;
                    list.RemoveAt(i);
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void image2_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Random r = new Random();
            Popup pop = new Popup();
            pop.Placement = PlacementMode.Relative;
            pop.HorizontalOffset = r.Next(400);
            pop.VerticalOffset = r.Next(30, 280);
            pop.PlacementTarget = this;

            Border b = new Border();
            b.Background = Brushes.AliceBlue;
            b.BorderBrush = Brushes.SteelBlue;
            b.BorderThickness = new Thickness(1);
            TextBlock blk = new TextBlock();
            b.Child = blk;
            blk.Text = listTips[r.Next(listTips.Count)];
            pop.Child = b;
            pop.IsOpen = true;
            list.Add(DateTime.Now, pop);
        }

    }
}
