using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace ParametricBuilder.Helpers
{
    public class CommaSeparatedValuesConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var selectedValues = value as IEnumerable<string>;
            if (selectedValues != null)
            {
                return string.Join(", ", selectedValues);
            }
            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Convert back logic (if needed)
            return value.ToString().Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries).ToList();
        }
    }
}
