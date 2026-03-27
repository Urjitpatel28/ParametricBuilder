using Microsoft.Xaml.Behaviors;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;

namespace ParametricBuilder.Behavior
{
    public class ListBoxSelectedItemsBehavior : Behavior<ListBox>
    {
        public static readonly DependencyProperty SelectedItemsProperty =
            DependencyProperty.Register("SelectedItems", typeof(IList<string>), typeof(ListBoxSelectedItemsBehavior), new PropertyMetadata(null));

        public ObservableCollection<string> SelectedItems
        {
            get { return (ObservableCollection<string>)GetValue(SelectedItemsProperty); }
            set { SetValue(SelectedItemsProperty, value); }
        }

        protected override void OnAttached()
        {
            base.OnAttached();
            AssociatedObject.SelectionChanged += AssociatedObject_SelectionChanged;
        }

        private void AssociatedObject_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedValues = new ObservableCollection<string>();

            // Add each selected item to the collection
            foreach (var item in AssociatedObject.SelectedItems)
            {
                selectedValues.Add(item.ToString());  // Add the item (as string or other desired format)
            }

            // Now set the selected values to the SelectedItems
            SelectedItems = selectedValues;
        }

        protected override void OnDetaching()
        {
            AssociatedObject.SelectionChanged -= AssociatedObject_SelectionChanged;
            base.OnDetaching();
        }
    }
}