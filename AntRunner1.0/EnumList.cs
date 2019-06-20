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
    public enum ErrorCode
    {
        [Description("Power Low")]
        PowL=1,
        [Description("Power High")]
        PowH,
        [Description("Frequency Low")]
        FreqL,
        [Description("Frequency High")]
        FreqH,
        [Description("Frequency Width Low")]
        FreqBandWidthL,
        [Description("Frequency Width High")]
        FreqBandWidthH,
        [Description("Short Circuit")]
        Bad,
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
    public enum TriggerType
    {
        Auto,
        //Manual,
        Scanner,
    }
    public enum MarkerType
    {
        Points,
        Markers,
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
