using DocumentFormat.OpenXml.Spreadsheet;
using NLog;
using ParametricBuilder.Commands;
using ParametricBuilder.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Input;

namespace ParametricBuilder.ViewModels
{
    public class MainWindowViewModel : INotifyPropertyChanged, IFileHandler
    {
        private Logger _logger = CadAiLogger.ConfigureLogger();
        public ObservableCollection<ExcelFileEntry> AvailableFiles { get; set; } = new ObservableCollection<ExcelFileEntry>();

        //Properties
        public string ApplicationVersion
        {
            get
            {
                string titlePrefix = "Parametric Builder";
                var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                return $"{titlePrefix} V{version.Major}.{version.Minor}";
            }
        }

        private string _filePath;
        public string FilePath
        {
            get => _filePath;
            set
            {
                if (_filePath != value)
                {
                    _filePath = value;
                    OnPropertyChanged(nameof(FilePath));
                }
            }
        }

        private string _selectedFileName;
        public string SelectedFileName
        {
            get => _selectedFileName;
            set
            {
                if (_selectedFileName != value)
                {
                    _selectedFileName = value;
                    OnPropertyChanged(nameof(SelectedFileName));
                }
            }
        }

        private ObservableCollection<string> _logMessages;
        public ObservableCollection<string> LogMessages
        {
            get => _logMessages;
            set
            {

                _logMessages = value;
                OnPropertyChanged(nameof(LogMessages));
            }
        }

        private Visibility _loadingWheelVisibility = Visibility.Hidden;
        public Visibility LoadingWheelVisibility
        {
            get => _loadingWheelVisibility;
            set
            {
                _loadingWheelVisibility = value;
                OnPropertyChanged(nameof(LoadingWheelVisibility));
            }
        }

        //update status
        private string _statusMessage = "Welcome! Please select your Excel file to get started.";  // Default message or initialize in constructor

        public string StatusMessage
        {
            get { return _statusMessage; }
            set
            {
                if (_statusMessage != value)
                {
                    _statusMessage = value;
                    OnPropertyChanged(nameof(StatusMessage));  // Notify the UI of property change
                }
            }
        }

        private string _outputPath;
        public string OutputPath
        {
            get => _outputPath;
            set
            {
                if (_outputPath != value)
                {
                    _outputPath = value;
                    OnPropertyChanged(nameof(OutputPath));
                }
            }
        }

        public ObservableCollection<ParameterModel> Parameters { get; } = new ObservableCollection<ParameterModel>();
        public Dictionary<string, List<ModelData>> ModelParameterMap { get; } = new Dictionary<string, List<ModelData>>(StringComparer.OrdinalIgnoreCase);
        public ObservableCollection<ParameterModel> ModelData { get; } = new ObservableCollection<ParameterModel>();
        public ObservableCollection<string> Models { get; } = new ObservableCollection<string>();

        private string _selectedModel;
        public string SelectedModel
        {
            get => _selectedModel;
            set
            {
                if (_selectedModel != value)
                {
                    _selectedModel = value;
                    OnPropertyChanged(nameof(SelectedModel));

                    if (string.IsNullOrWhiteSpace(_selectedModel))
                    {
                        // Handle case where model is null or empty
                        StatusMessage = "No model selected.";
                        OnPropertyChanged(nameof(StatusMessage));
                        return;
                    }

                    // Check if the selected model exists in the dictionary
                    if (ModelParameterMap.ContainsKey(_selectedModel))
                    {
                        var modelParameters = ModelParameterMap[_selectedModel];

                        // Match parameters by name and update values
                        foreach (var uiParam in Parameters)
                        {
                            // Find matching parameter from model data
                            var modelParam = modelParameters.FirstOrDefault(p =>
                                p.Name.Equals(uiParam.Name, StringComparison.OrdinalIgnoreCase));

                            if (modelParam != null)
                            {
                                uiParam.Value = modelParam.Value;
                                // Raise property changed to update UI binding
                                OnPropertyChanged(nameof(uiParam.Value));
                            }
                        }

                        StatusMessage = $"Loaded {_selectedModel} with {modelParameters.Count} parameters.";
                    }
                    else
                    {
                        // Handle case where the model is not found
                        StatusMessage = $"Model '{_selectedModel}' not found in data.";
                    }
                    OnPropertyChanged(nameof(StatusMessage));
                }
            }
        }


        private Configurator _selectedConfigurator;
        public ObservableCollection<Configurator> Configurators { get; }
        public Configurator SelectedConfigurator
        {
            get => _selectedConfigurator;
            set
            {
                if (_selectedConfigurator != value)
                {            
                    ModelParameterMap.Clear();
                    Parameters.Clear();
                    Models.Clear();
                    _selectedConfigurator = value;

                    //// Temp code ....Null check before accessing DisplayName
                    //if (value != null && (value.DisplayName == "GPHE" || value.DisplayName == "PLATE"))
                    //{
                    //    _selectedConfigurator = null;
                    //    MessageBox.Show("Configurator not available!!", "Under Progress!", MessageBoxButton.OK, MessageBoxImage.Information);
                    //    OnPropertyChanged(nameof(SelectedConfigurator));
                    //    return;
                    //}

                    OnPropertyChanged(nameof(SelectedConfigurator));
                    // Access backend paths through SelectedConfigurator when needed
                }
            }
        }


        public MainWindowViewModel()
        {

            var dataService = new ConfiguratorDataService();
            Configurators = dataService.LoadConfiguratorData(ConfigurationManager.AppSettings["MappingExcel"]);

            //LoadExcelFilesFromConfiguredDirectory();
            _outputPath = ConfigurationManager.AppSettings["OutPutDirectory"] ?? string.Empty;
            _logMessages = new ObservableCollection<string>();
            LogMessages.CollectionChanged += LogMessages_CollectionChanged; // Add this line
            LogMessages.Clear();

            Parameters.CollectionChanged += Parameters_CollectionChanged;

            CadAiLogger.LogMessageWritten -= OnLogMessageWritten; // Detach previous handler if any
            CadAiLogger.LogMessageWritten += OnLogMessageWritten; // Attach new handler

            //Command
            BrowseCommand = new BrowseCommand(this);
            GetExcelParametersCommand = new GetExcelParametersCommand(this);
            RunCadCasterCommand = new RunCadCasterCommand(this);
            //GetValuesCommand = new GetValuesCommand(this);
            ChangeOutputPathCommand = new ChangeOutputPathCommand(this);
            UpdateCommand = new UpdateCommand(this);
            AddCommand = new AddCommand(this);
            DeleteCommand = new DeleteCommand(this);

        }

        private void Parameters_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.NewItems != null)
            {
                foreach (ParameterModel item in e.NewItems)
                {
                    item.ErrorsChanged += Parameter_ErrorsChanged;
                }
            }
            if (e.OldItems != null)
            {
                foreach (ParameterModel item in e.OldItems)
                {
                    item.ErrorsChanged -= Parameter_ErrorsChanged;
                }
            }
        }

        private void Parameter_ErrorsChanged(object sender, DataErrorsChangedEventArgs e)
        {
            OnPropertyChanged(nameof(HasValidationErrors));
        }


        //Commands
        public ICommand BrowseCommand { get; set; }
        public ICommand GetExcelParametersCommand { get; set; }
        public ICommand RunCadCasterCommand { get; set; }
        public ICommand GetValuesCommand { get; set; }
        public ICommand ChangeOutputPathCommand { get; set; }
        public ICommand UpdateCommand { get; set; }
        public ICommand AddCommand { get; set; }
        public ICommand DeleteCommand { get; set; }


        //Methods
        private void LogMessages_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Add)
            {
                // Scroll to the last item (code depends on UI control)
            }
        }

        private void OnLogMessageWritten(object sender, string logMessage)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                if (LogMessages.Count > 1000)
                {
                    LogMessages.RemoveAt(0);
                }

                if (!LogMessages.Contains(logMessage))
                {
                    LogMessages.Add(logMessage);
                }

            });
        }

        public void HandleDroppedFiles(string[] filePaths)
        {
            var validFile = filePaths.FirstOrDefault(IsExcelFile);
            if (validFile != null)
            {
                FilePath = validFile;
            }
        }

        private bool IsExcelFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            var ext = Path.GetExtension(path).ToLower();
            return ext == ".xls" || ext == ".xlsx";
        }

        private void LoadExcelFilesFromConfiguredDirectory()
        {
            try
            {
                string relativeFolder = Path.Combine(ConfigurationManager.AppSettings["ExcelFileDirectory"], "ConfigExcelTemplates");
                if (string.IsNullOrEmpty(relativeFolder)) return;

                string exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string folderPath = Path.Combine(exePath, relativeFolder);

                if (Directory.Exists(folderPath))
                {
                    var files = Directory.GetFiles(folderPath, "*.xlsx");
                    foreach (var fullPath in files)
                    {
                        var displayName = Path.GetFileNameWithoutExtension(fullPath);
                        AvailableFiles.Add(new ExcelFileEntry
                        {
                            DisplayName = displayName,
                            FullPath = fullPath
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error loading files: {ex.Message}";
                _logger.Error(ex, "Failed to load Excel files");
            }
        }

        public void HandleFileSelected(string filePath)
        {
            FilePath = filePath;
            Parameters.Clear();
            StatusMessage = "Click The 'Get Parameters' Button To Proceed.";
        }


        public bool HasValidationErrors => Parameters.Any(p => p.HasErrors);

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private Dictionary<string, List<ParameterModel>> _dependentsMap = new Dictionary<string, List<ParameterModel>>();

        public void SetupDependencies()
        {
            _dependentsMap.Clear();

            // Map controllers to their dependents
            foreach (var param in Parameters)
            {
                if (!string.IsNullOrEmpty(param.DependsOn))
                {
                    if (!_dependentsMap.ContainsKey(param.DependsOn))
                        _dependentsMap[param.DependsOn] = new List<ParameterModel>();

                    _dependentsMap[param.DependsOn].Add(param);
                }
            }

            // Set up value change handlers
            foreach (var param in Parameters.Where(p => _dependentsMap.ContainsKey(p.Name)))
            {
                param.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(ParameterModel.Value))
                    {
                        UpdateDependents(param);
                    }
                };
                // Initialize state
                UpdateDependents(param);
            }
        }

        private void UpdateDependents(ParameterModel controller)
        {
            if (!_dependentsMap.TryGetValue(controller.Name, out var dependents))
                return;

            foreach (var dependent in dependents)
            {
                dependent.UpdateDependencyState(controller);
            }
        }
    }
}