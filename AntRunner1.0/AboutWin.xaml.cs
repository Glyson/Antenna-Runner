using System;
using System.Windows;
using System.Windows.Controls;

namespace AntRunner
{
    /// <summary>
    /// AboutWin.xaml 的交互逻辑
    /// </summary>
    public partial class AboutWin : Window
    {
        public AboutWin()
        {
            InitializeComponent();
            tipTime.ToolTip = DateTime.Now.ToString("yyyy/MM/dd") + " - " + MainWindow.Self.expireTime.ToString("yyyy/MM/dd");
            ToolTipService.SetShowDuration(tipTime, 2000);
            ToolTipService.SetInitialShowDelay(tipTime, 10000);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
