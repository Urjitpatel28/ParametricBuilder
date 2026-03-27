using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;

namespace ParametricBuilder.Models
{
    public class ParameterModel : INotifyPropertyChanged, INotifyDataErrorInfo
    {
        private string _value;
        private double? _minValue;
        private double? _maxValue;
        private List<string> _allowedValues;
        private readonly Dictionary<string, List<string>> _errors = new Dictionary<string, List<string>>();

        public string DependsOn { get; set; } // Name of the controlling parameter
        public string TriggerValue { get; set; } // Value that enables this parameter


        public string Name { get; set; }

        public string Value
        {
            get => _value;
            set
            {
                if (_value != value)
                {
                    string oldValue = _value;
                    _value = value;
                    Validate(value);
                    OnPropertyChanged(nameof(Value));

                    // Notify dependents when value changes
                    if (HasDependents)
                    {
                        DependencyManager.NotifyDependents(Name);
                    }
                }
            }
        }

        private bool _hasDependents;
        public bool HasDependents
        {
            get => _hasDependents;
            set
            {
                _hasDependents = value;
                OnPropertyChanged(nameof(HasDependents));
            }
        }
        private ParameterModel _controller;
        public ParameterModel Controller
        {
            get => _controller;
            set
            {
                _controller = value;
                OnPropertyChanged(nameof(Controller));
                UpdateState(); // Update state when controller is set
            }
        }

        public void UpdateState()
        {
            if (Controller == null || string.IsNullOrEmpty(TriggerValue))
                return;

            bool shouldBeActive = !string.IsNullOrEmpty(Controller.Value) &&
                                 Controller.Value.Equals(TriggerValue, StringComparison.OrdinalIgnoreCase);

            // Remember previous state
            bool wasActive = IsVisible && IsEnabled;

            // Update state
            IsVisible = shouldBeActive;
            IsEnabled = shouldBeActive;

            // Clear value when becoming inactive
            if (wasActive && !shouldBeActive)
            {
                ClearValue();
            }
        }
        public void ClearValue()
        {
            Value = null;

            // Clear CheckBoxList selections
            if (ControlType == "CheckBoxList" && SelectedValues != null)
            {
                SelectedValues.Clear();
            }
        }
        public string ControlType { get; set; } = "TextBox";

        // Modify Value property to support multiple selections
        private ObservableCollection<string> _selectedValues = new ObservableCollection<string>();
        public ObservableCollection<string> SelectedValues
        {
            get => _selectedValues;
            set
            {
                if (_selectedValues != value)
                {
                    _selectedValues = value;
                    OnPropertyChanged(nameof(SelectedValues));
                    Value = string.Join(", ", _selectedValues);
                }
            }
        }

        private bool _isVisible = true;
        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                if (_isVisible != value)
                {
                    _isVisible = value;
                    OnPropertyChanged(nameof(IsVisible));
                }
            }
        }

        private bool _isEnabled = true;
        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                if (_isEnabled != value)
                {
                    _isEnabled = value;
                    OnPropertyChanged(nameof(IsEnabled));
                }
            }
        }

        public double? MinValue
        {
            get => _minValue;
            set => _minValue = value;
        }

        public double? MaxValue
        {
            get => _maxValue;
            set => _maxValue = value;
        }
        
        public List<string> AllowedValues
        {
            get => _allowedValues;
            set => _allowedValues = value;
        }

        private void Validate(string value)
        {
            // If not visible or not enabled, skip validation and clear errors
            if (!IsVisible || !IsEnabled)
            {
                ClearErrors(nameof(Value));
                return;
            }

            ClearErrors(nameof(Value));

            if (AllowedValues != null && AllowedValues.Count > 0)
            {
                if (!AllowedValues.Contains(value, StringComparer.OrdinalIgnoreCase))
                {
                    AddError(nameof(Value), $"Value must be one of: {string.Join(", ", AllowedValues)}");
                }
            }
            else if (MinValue.HasValue || MaxValue.HasValue)
            {
                if (!double.TryParse(value, out double numericValue))
                {
                    AddError(nameof(Value), "Value must be a number");
                }
                else
                {
                    if (MinValue.HasValue && numericValue < MinValue.Value)
                    {
                        AddError(nameof(Value), $"Value must be ≥ {MinValue}");
                    }
                    if (MaxValue.HasValue && numericValue > MaxValue.Value)
                    {
                        AddError(nameof(Value), $"Value must be ≤ {MaxValue}");
                    }
                }
            }

            // Add any other validation rules here
        }

        #region INotifyDataErrorInfo Implementation
        public event EventHandler<DataErrorsChangedEventArgs> ErrorsChanged;

        public System.Collections.IEnumerable GetErrors(string propertyName)
        {
            if (string.IsNullOrEmpty(propertyName))
            {
                // Handle the case when propertyName is null or empty (return an empty collection or null)
                return null;
            }

            return _errors.ContainsKey(propertyName) ? _errors[propertyName] : null;
        }

        public bool HasErrors => _errors.Any();

        private void AddError(string propertyName, string error)
        {
            if (!_errors.ContainsKey(propertyName))
                _errors[propertyName] = new List<string>();

            if (!_errors[propertyName].Contains(error))
            {
                _errors[propertyName].Add(error);
                OnErrorsChanged(propertyName);
            }
        }

        private void ClearErrors(string propertyName)
        {
            if (_errors.ContainsKey(propertyName))
            {
                _errors.Remove(propertyName);
                OnErrorsChanged(propertyName);
            }
        }


        private void OnErrorsChanged(string propertyName)
        {
            ErrorsChanged?.Invoke(this, new DataErrorsChangedEventArgs(propertyName));
        }

        #endregion

        // Existing INotifyPropertyChanged implementation...
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        internal object GetErrors()
        {
            throw new NotImplementedException();
        }


        public void UpdateDependencyState(ParameterModel controller)
        {
            if (controller.Name != DependsOn) return;

            bool shouldBeActive = controller.Value == TriggerValue;
            IsVisible = shouldBeActive;
            IsEnabled = shouldBeActive;
        }
    }
}