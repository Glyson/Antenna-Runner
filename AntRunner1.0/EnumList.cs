using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Windows.Markup;
using System.Reflection;
using System.Windows.Data;

namespace AntRunner
{
    public enum CompareType
    {
        [Description("全部高于")]
        AllUp,
        [Description("全部低于")]
        AllDown,
        [Description("高于")]
        Up,
        [Description("低于")]
        Down,
    }
    public enum ErrorCode
    {
        [Description("功率偏低")]
        PowL = 1,
        [Description("功率偏高")]
        PowH,
        [Description("频率偏低")]
        FreqL,
        [Description("频率偏高")]
        FreqH,
        [Description("频宽偏低")]
        FreqBandWidthL,
        [Description("频宽偏高")]
        FreqBandWidthH,
        [Description("短路")]
        Bad,
        [Description("功率(S21)偏低")]
        PowS21L,
        [Description("功率(S21)偏高")]
        PowS21H,
        [Description("驻波(S22)偏高")]
        StandingWaveS22H,
    }
    public enum Instrument
    {
        Agilent_5071C,
        Agilent_8753ES,
    }
    public enum TraceFormat
    {
        SWR,
        LOG,
        LOG_SWR,
    }
    public enum S22TraceFormat
    {
        SWR,
        LOG,
    }
    public enum TriggerType
    {
        Auto,
        //Manual,
        Scanner,
    }
    public enum MarkerType
    {
        Range,
        Points,
    }
    public enum Trace
    {
        S11,
        S22,
        S33,
        S44,
    }
    public enum State
    {
        Stoped,
        Trigger,
        ReadTrace,
        Setting,
        Running,
        Pause,
    }

    #region enum binding source operations
    public class EnumerationExtension : MarkupExtension
    {
        private Type _enumType;

        public Type EnumType
        {
            get { return _enumType; }
            private set
            {
                if (_enumType == value)
                    return;
                var enumType = Nullable.GetUnderlyingType(value) ?? value;
                if (enumType.IsEnum == false)
                    throw new ArgumentException("Type must be an Enum");
                _enumType = value;
            }
        }

        public EnumerationExtension(Type enumType)
        {
            if (enumType == null)
                throw new ArgumentNullException("enumType");
            EnumType = enumType;
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            var enumValues = Enum.GetValues(EnumType);
            return (
                from object enumValue in enumValues
                select new EnumerationMember
                {
                    Value = enumValue,
                    Description = GetDescription(enumValue)
                }).ToArray();
        }

        public static string GetDescription(object value)
        {
            FieldInfo field = value.GetType().GetField(value.ToString());
            DescriptionAttribute[] attributes = (DescriptionAttribute[])field.GetCustomAttributes(typeof(DescriptionAttribute), false);
            return (attributes.Length > 0) ? attributes[0].Description : value.ToString();
        }

        public class EnumerationMember
        {
            public string Description { get; set; }
            public object Value { get; set; }
        }
    }
    public class BoolToString : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value.ToString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return bool.Parse(value.ToString());
        }
    }
    public class StringToBool : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value.ToString() != parameter.ToString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value.ToString();
        }
    }
    public class BoolToVisibility : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return ((bool)value) ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (value.ToString() == System.Windows.Visibility.Visible.ToString()) ? true : false;
        }
    }
    public class StringToVisibility : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value.ToString() == parameter.ToString())
                return System.Windows.Visibility.Visible;
            else
                return System.Windows.Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value.ToString();
        }
    }
    #endregion
}
