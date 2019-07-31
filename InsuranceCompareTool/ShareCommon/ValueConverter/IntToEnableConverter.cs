using System;
using System.Globalization;
using System.Windows.Data;
namespace InsuranceCompareTool.ShareCommon.ValueConverter {
    public class IntToEnableConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var v = System.Convert.ToInt32(value);
            var s = System.Convert.ToInt32(parameter);
            return v == s;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }
}