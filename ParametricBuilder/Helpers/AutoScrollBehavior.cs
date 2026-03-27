using Microsoft.Xaml.Behaviors;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;

namespace ParametricBuilder.Helpers
{
    public class AutoScrollBehavior : Behavior<ListBox>
    {
        protected override void OnAttached()
        {
            base.OnAttached();
            if (AssociatedObject.Items is INotifyCollectionChanged notifyCollection)
            {
                notifyCollection.CollectionChanged += OnCollectionChanged;
            }

            // Scroll to the bottom initially if there are items
            //if (AssociatedObject.Items.Count > 0)
            //{
            //    AssociatedObject.Loaded += (s, e) => ScrollToBottom();
            //}
        }

        private void OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            //if (e.Action == NotifyCollectionChangedAction.Add && AssociatedObject.Items.Count > 0)
            //{
            //    ScrollToBottom();
            //}
        }

        private void ScrollToBottom()
        {
            if (AssociatedObject.Items.Count > 0)
            {
                var border = (Border)VisualTreeHelper.GetChild(AssociatedObject, 0);
                var scrollViewer = (ScrollViewer)VisualTreeHelper.GetChild(border, 0);
                scrollViewer.ScrollToBottom();
            }
        }

        protected override void OnDetaching()
        {
            base.OnDetaching();
            if (AssociatedObject.Items is INotifyCollectionChanged notifyCollection)
            {
                notifyCollection.CollectionChanged -= OnCollectionChanged;
            }
        }
    }
}
