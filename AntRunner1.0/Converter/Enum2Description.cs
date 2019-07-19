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
    public class Enum2Description : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Helper.GetEnumDescription((Enum)value);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Helper.DescToEnum(value.ToString(), targetType);
        }
    }
}
