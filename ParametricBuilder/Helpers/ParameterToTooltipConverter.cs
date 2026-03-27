using ParametricBuilder.Models;
using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace ParametricBuilder.Helpers
{
    public class ParameterToTooltipConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ParameterModel param)
            {
                // Show validation errors if any exist
                if (param.HasErrors)
                {
                    var errors = param.GetErrors(null)?.Cast<string>();
                    return errors != null ? string.Join(Environment.NewLine, errors) : null;
                }

                // Show validation rules as hints
                if (param.AllowedValues?.Count > 0)
                {
                    return $"Allowed values: {string.Join(", ", param.AllowedValues)}";
                }
                else if (param.MinValue.HasValue && param.MaxValue.HasValue)
                {
                    return $"Range: {param.MinValue} to {param.MaxValue}";
                }
                else if (param.MinValue.HasValue)
                {
                    return $"Min: {param.MinValue}";
                }
                else if (param.MaxValue.HasValue)
                {
                    return $"Max: {param.MaxValue}";
                }
            }
            return null; // No tooltip if no validation rules
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}