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
using System.Windows.Shapes;
using System.Windows.Forms;
using AntRunner.Properties;
using NationalInstruments.VisaNS;
using System.ComponentModel;
using System.Windows.Markup;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Ports;

namespace AntRunner
{
    /// <summary>
    /// SetupWin.xaml 的交互逻辑
    /// </summary>
    public partial class SetupWin : Window
    {

        public SetupWin()
        {
            InitializeComponent();
            Settings.Default.Para1.MarkerType = MarkerType.Markers.ToString();
            Settings.Default.Para2.MarkerType = MarkerType.Markers.ToString();
            Settings.Default.Para3.MarkerType = MarkerType.Markers.ToString();
            Settings.Default.Para4.MarkerType = MarkerType.Markers.ToString();
            RefreshGPIB();
            //InitCOM();
            if (MainWindow.Self.State == State.Running)
            {
                this.IsEnabled = false;
            }
        }

        private void btnSelDir_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Settings.Default.OutputDir = fbd.SelectedPath;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Settings.Default.Save();
        }

        private void btnIns_Click(object sender, RoutedEventArgs e)
        {
            RefreshGPIB();
        }

        private void RefreshGPIB()
        {
            try
            {
                cbGPIB.ItemsSource = null;
                cbGPIB.Items.Clear();
                cbGPIB.Foreground = Brushes.Black;
                string pre = cbGPIB.Text;
                ResourceManager manager = ResourceManager.GetLocalManager();
                string[] listGPIB = manager.FindResources("GPIB?*INSTR");
                cbGPIB.ItemsSource = listGPIB;
                if (listGPIB != null && listGPIB.Length > 0)
                {
                    if (cbGPIB.Items.Contains(pre))
                        cbGPIB.Text = pre;
                    else
                        cbGPIB.Text = cbGPIB.Items[0].ToString();
                }
                else
                {
                    cbGPIB.Items.Add("No Instrument");
                    cbGPIB.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                cbGPIB.Foreground = Brushes.Red;
                cbGPIB.Items.Add("No Instrument");
                cbGPIB.SelectedIndex = 0;
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            Settings.Default.Para2.Power = Settings.Default.Para1.Power;
            //Settings.Default.Para2.Enable = Settings.Default.Para1.Enable;
            Settings.Default.Para2.Bandwidth = Settings.Default.Para1.Bandwidth;
            Settings.Default.Para2.FreqStart = Settings.Default.Para1.FreqStart;
            Settings.Default.Para2.FreqStop = Settings.Default.Para1.FreqStop;
            Settings.Default.Para2.Points = Settings.Default.Para1.Points;
            Settings.Default.Para2.MarkerPoints = Settings.Default.Para1.MarkerPoints;
            Settings.Default.Para2.MarkerText = Settings.Default.Para1.MarkerText;
            Settings.Default.Para2.MarkerType = Settings.Default.Para1.MarkerType;
            Settings.Default.Para2.ReferDiff = Settings.Default.Para1.ReferDiff;
            Settings.Default.Para2.ReferTracePath = Settings.Default.Para1.ReferTracePath;
            Settings.Default.Para2.CutPow = Settings.Default.Para1.CutPow;
            Settings.Default.Para2.DiffBW = Settings.Default.Para1.DiffBW;
            Settings.Default.Para2.DiffFreq = Settings.Default.Para1.DiffFreq;
            Settings.Default.Para2.DiffPower = Settings.Default.Para1.DiffPower;

            Settings.Default.Para3.Power = Settings.Default.Para1.Power;
            //Settings.Default.Para3.Enable = Settings.Default.Para1.Enable;
            Settings.Default.Para3.Bandwidth = Settings.Default.Para1.Bandwidth;
            Settings.Default.Para3.FreqStart = Settings.Default.Para1.FreqStart;
            Settings.Default.Para3.FreqStop = Settings.Default.Para1.FreqStop;
            Settings.Default.Para3.Points = Settings.Default.Para1.Points;
            Settings.Default.Para3.MarkerPoints = Settings.Default.Para1.MarkerPoints;
            Settings.Default.Para3.MarkerText = Settings.Default.Para1.MarkerText;
            Settings.Default.Para3.MarkerType = Settings.Default.Para1.MarkerType;
            Settings.Default.Para3.ReferDiff = Settings.Default.Para1.ReferDiff;
            Settings.Default.Para3.ReferTracePath = Settings.Default.Para1.ReferTracePath;
            Settings.Default.Para3.CutPow = Settings.Default.Para1.CutPow;
            Settings.Default.Para3.DiffBW = Settings.Default.Para1.DiffBW;
            Settings.Default.Para3.DiffFreq = Settings.Default.Para1.DiffFreq;
            Settings.Default.Para3.DiffPower = Settings.Default.Para1.DiffPower;

            Settings.Default.Para4.Power = Settings.Default.Para1.Power;
            //Settings.Default.Para4.Enable = Settings.Default.Para1.Enable;
            Settings.Default.Para4.Bandwidth = Settings.Default.Para1.Bandwidth;
            Settings.Default.Para4.FreqStart = Settings.Default.Para1.FreqStart;
            Settings.Default.Para4.FreqStop = Settings.Default.Para1.FreqStop;
            Settings.Default.Para4.Points = Settings.Default.Para1.Points;
            Settings.Default.Para4.MarkerPoints = Settings.Default.Para1.MarkerPoints;
            Settings.Default.Para4.MarkerText = Settings.Default.Para1.MarkerText;
            Settings.Default.Para4.MarkerType = Settings.Default.Para1.MarkerType;
            Settings.Default.Para4.ReferDiff = Settings.Default.Para1.ReferDiff;
            Settings.Default.Para4.ReferTracePath = Settings.Default.Para1.ReferTracePath;
            Settings.Default.Para4.CutPow = Settings.Default.Para1.CutPow;
            Settings.Default.Para4.DiffBW = Settings.Default.Para1.DiffBW;
            Settings.Default.Para4.DiffFreq = Settings.Default.Para1.DiffFreq;
            Settings.Default.Para4.DiffPower = Settings.Default.Para1.DiffPower;

            System.Windows.MessageBox.Show(this, "Mapping OK.", "Tips");
        }

        private bool OutputRefer(SortedList<double, double> list, string path)
        {
            StreamWriter writer = null; try
            {
                using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                {
                    writer = new StreamWriter(fs);
                    for (int i = 0; i < list.Count; i++)
                    {
                        writer.WriteLine(list.Keys[i] + "," + list.Values[i]);
                    }
                    writer.Flush();
                    writer.Close();
                    fs.Close();
                }
                return true;
            }
            catch
            {
                if (writer != null)
                {
                    writer.Flush();
                    writer.Close();
                }
                return false;
            }
        }
        private void btnCali1_Click(object sender, RoutedEventArgs e)
        {
            if (MainWindow.Self.State == State.Running)
            {
                System.Windows.MessageBox.Show(this, "正在测试...", "Tips");
                return;
            }
            if (MainWindow.Self.vna == null || !MainWindow.Self.vna.IsOK)
            {
                MainWindow.Self.vna = VNA.CreateVNA();
                if (!MainWindow.IsSkip)
                {
                    if (!MainWindow.Self.vna.Init(Settings.Default.GPIB))
                    {
                        return;
                    }
                }
            }
            System.Windows.Controls.Button btn = sender as System.Windows.Controls.Button;
            this.Cursor = System.Windows.Input.Cursors.Wait;
            string trace = btn.Tag.ToString();
            SortedList<double, double> list;
            string root = string.Format("{0}\\{1}", Settings.Default.OutputDir, "Calibration");
            if (!Directory.Exists(root))
                Directory.CreateDirectory(root);
            string path;
            string dateStr = DateTime.Now.ToString("yyyy.MM.dd HH.mm");
            MainWindow.Self.vna.Config();
            switch (trace)
            {
                case "S11":
                    MainWindow.Self.vna.Setup(Settings.Default.Para1);
                    list = MainWindow.Self.vna.ReadTrace(Settings.Default.Para1);
                    path = string.Format("{0}\\REFER_{1}_{2}.csv", root, Settings.Default.Para1.Trace, dateStr);
                    if (OutputRefer(list, path))
                    {
                        Settings.Default.Para1.ReferTracePath = path;
                    }
                    break;
                case "S22":
                    MainWindow.Self.vna.Setup(Settings.Default.Para2);
                    list = MainWindow.Self.vna.ReadTrace(Settings.Default.Para2);
                    path = string.Format("{0}\\REFER_{1}_{2}.csv", root, Settings.Default.Para2.Trace, dateStr);
                    if (OutputRefer(list, path))
                    {
                        Settings.Default.Para2.ReferTracePath = path;
                    }
                    break;
                case "S33":
                    MainWindow.Self.vna.Setup(Settings.Default.Para3);
                    list = MainWindow.Self.vna.ReadTrace(Settings.Default.Para3);
                    path = string.Format("{0}\\REFER_{1}_{2}.csv", root, Settings.Default.Para3.Trace, dateStr);
                    if (OutputRefer(list, path))
                    {
                        Settings.Default.Para3.ReferTracePath = path;
                    }
                    break;
                case "S44":
                    MainWindow.Self.vna.Setup(Settings.Default.Para4);
                    list = MainWindow.Self.vna.ReadTrace(Settings.Default.Para4);
                    path = string.Format("{0}\\REFER_{1}_{2}.csv", root, Settings.Default.Para4.Trace, dateStr);
                    if (OutputRefer(list, path))
                    {
                        Settings.Default.Para4.ReferTracePath = path;
                    }
                    break;
            }
            this.Cursor = System.Windows.Input.Cursors.Arrow;
            System.Windows.MessageBox.Show(this, "Calibrate completed.", "Tips", MessageBoxButton.OK);
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (MainWindow.Self.State == State.Running)
                return;

            if (this.IsLoaded)
                MainWindow.Self.InitCount();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            //dialog.InitialDirectory = Settings.Default.OutputDir;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string path = dialog.FileName;
                System.Windows.Controls.Button btn = sender as System.Windows.Controls.Button;
                string trace = btn.Tag.ToString();
                switch (trace)
                {
                    case "S11":
                        Settings.Default.Para1.ReferTracePath = path;
                        break;
                    case "S22":
                        Settings.Default.Para2.ReferTracePath = path;
                        break;
                    case "S33":
                        Settings.Default.Para3.ReferTracePath = path;
                        break;
                    case "S44":
                        Settings.Default.Para4.ReferTracePath = path;
                        break;
                }
            }
        }

            //private void InitCOM()
            //{
            //    string[] ports = SerialPort.GetPortNames();
            //    cbPort1.ItemsSource = ports;
            //    cbPort2.ItemsSource = ports;
            //    cbPort3.ItemsSource = ports;
            //    cbPort4.ItemsSource = ports;

            //}
            //private void Button_Click(object sender, RoutedEventArgs e)
            //{
            //    string[] ports = SerialPort.GetPortNames();
            //    string pre;
            //    int tag = int.Parse((sender as System.Windows.Controls.Button).Tag.ToString());
            //    switch (tag)
            //    {
            //        case 1:
            //            pre = cbPort1.Text;
            //            cbPort1.ItemsSource = null;
            //            cbPort1.Items.Clear();
            //            cbPort1.ItemsSource = ports;
            //            cbPort1.Text = pre;
            //            break;
            //        case 2:
            //            pre = cbPort2.Text;
            //            cbPort2.ItemsSource = null;
            //            cbPort2.Items.Clear();
            //            cbPort2.ItemsSource = ports;
            //            cbPort2.Text = pre;
            //            break;
            //        case 3:
            //            pre = cbPort3.Text;
            //            cbPort3.ItemsSource = null;
            //            cbPort3.Items.Clear();
            //            cbPort3.ItemsSource = ports;
            //            cbPort3.Text = pre;
            //            break;
            //        case 4:
            //            pre = cbPort4.Text;
            //            cbPort4.ItemsSource = null;
            //            cbPort4.Items.Clear();
            //            cbPort4.ItemsSource = ports;
            //            cbPort4.Text = pre;
            //            break;
            //    }
            //}
        }
    }
