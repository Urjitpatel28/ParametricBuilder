using System.ComponentModel;

namespace ParametricBuilder.Models
{
    public class ModelData : INotifyPropertyChanged
    {
        //private bool _isChecked;
        private string _value;

        public string Name { get; set; }

        public string Value
        {
            get => _value;
            set
            {
                _value = value;
                OnPropertyChanged(nameof(Value));
            }
        }

        //public bool IsChecked
        //{
        //    get => _isChecked;
        //    set
        //    {
        //        _isChecked = value;
        //        OnPropertyChanged(nameof(IsChecked));
        //    }
        //}


        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}