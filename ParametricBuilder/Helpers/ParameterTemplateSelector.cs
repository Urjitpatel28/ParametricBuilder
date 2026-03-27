using ParametricBuilder.Models;
using System.Windows;
using System.Windows.Controls;

namespace ParametricBuilder.Helpers
{
    public class ParameterTemplateSelector : DataTemplateSelector
    {
        public DataTemplate TextBoxTemplate { get; set; }
        public DataTemplate ComboBoxTemplate { get; set; }
        public DataTemplate MultiSelectTemplate { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            var parameter = item as ParameterModel;
            switch (parameter.ControlType)
            {
                case "ComboBox":
                    return ComboBoxTemplate;
                case "MultiSelect":
                    return MultiSelectTemplate;
                case "TextBox":
                default:
                    return TextBoxTemplate;
            }
        }
    }
}