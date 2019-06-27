using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Threading.Tasks;

namespace AntRunner
{
    /// <summary>
    /// ReportWaitWin.xaml 的交互逻辑
    /// </summary>
    public partial class ReportWaitWin : Window
    {
        int count;
        public ReportWaitWin(int count)
        {
            InitializeComponent();
            this.count = count;
            Task.Factory.StartNew(new Action(delegate
            {
                while (true)
                {
                    Thread.Sleep(300);
                    Action act = new Action(UpdateProgress);
                    blk.Dispatcher.Invoke(act);
                }
            }));
        }
        private void UpdateProgress()
        {
            blk.Text = string.Format("{0} %", (int)(DataBase.Progress / (count * 2) * 100));
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
