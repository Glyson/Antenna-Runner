using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Globalization;
using System.Windows.Data;
using AntRunner.Properties;

namespace AntRunner
{
    public class TestType2Visibility : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (parameter == null)
            {
                return Visibility.Collapsed;
            }
            else
            {
                string[] arr = parameter.ToString().Split(',');
                if (arr.Contains(Settings.Default.TraceFormat))
                    return Visibility.Visible;
                return Visibility.Collapsed;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}
