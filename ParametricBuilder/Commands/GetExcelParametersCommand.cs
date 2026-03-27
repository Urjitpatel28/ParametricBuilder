using ClosedXML.Excel;
using NLog;
using ParametricBuilder.Models;
using ParametricBuilder.ViewModels;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace ParametricBuilder.Commands
{
    public class GetExcelParametersCommand : BaseCommand
    {
        private readonly MainWindowViewModel _viewModel;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public GetExcelParametersCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
            _logger.Debug("GetExcelParameters command initialized.");
        }

        public override bool CanExecute(object parameter) => true;

        public override async void Execute(object parameter)
        {
            try
            {
                BeginLoading();

                var parameters = await ReadParametersFromExcelAsync();
                SetupDependencies(parameters);
                UpdateViewModelParameters(parameters);

                await LoadModelsFromFileAsync();

                _logger.Info("Excel Parameters Successfully Loaded.");
            }
            catch (IOException ioEx)
            {
                HandleFileAccessError(ioEx);
            }
            catch (Exception ex)
            {
                HandleUnexpectedError(ex);
            }
            finally
            {
                EndLoading();
            }
        }

        // Show loading UI and status
        private void BeginLoading()
        {
            Mouse.OverrideCursor = Cursors.Wait;
            _viewModel.LoadingWheelVisibility = Visibility.Visible;
            _viewModel.StatusMessage = "Reading Parameters From Excel...";
            _viewModel.Parameters.Clear();
        }

        // Hide loading UI and update status
        private void EndLoading()
        {
            Mouse.OverrideCursor = null;
            _viewModel.LoadingWheelVisibility = Visibility.Hidden;
            _viewModel.StatusMessage = "Parameters Successfully Loaded. Please Input Values And Configure Your Model!";
        }

        // Read and parse parameters from Excel
        private async Task<List<ParameterModel>> ReadParametersFromExcelAsync()
        {
            var excelPath = _viewModel.SelectedConfigurator.ConfigExcelTemplatePath;
            _logger.Info($"Reading Parameters From Excel File: {excelPath}");

            return await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook(excelPath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var paramList = new List<ParameterModel>();
                    var dependencyMap = new Dictionary<string, ParameterModel>();

                    foreach (var row in worksheet.RowsUsed())
                    {
                        if (row.Cell(3).GetString().Trim().Equals("INPUT", StringComparison.OrdinalIgnoreCase))
                        {
                            var param = CreateParameterModelFromRow(row);
                            ParseValidationRule(row.Cell(5).GetString().Trim(), param);
                            paramList.Add(param);
                            dependencyMap[param.Name] = param;
                        }
                    }

                    // Setup controller references after all params created
                    foreach (var param in paramList.Where(p => !string.IsNullOrEmpty(p.DependsOn)))
                    {
                        if (dependencyMap.TryGetValue(param.DependsOn, out var controller))
                        {
                            param.Controller = controller;
                        }
                    }

                    return paramList;
                }
            });
        }

        // Setup dependencies after all parameters created
        private void SetupDependencies(List<ParameterModel> paramList)
        {
            foreach (var param in paramList.Where(p => p.Controller != null))
            {
                DependencyManager.RegisterDependency(param.Controller.Name, param);
                param.UpdateState();
            }
        }

        // Push the new parameters to the ViewModel on the UI thread
        private void UpdateViewModelParameters(List<ParameterModel> parameters)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                _viewModel.Parameters.Clear();
                foreach (var param in parameters)
                    _viewModel.Parameters.Add(param);
            });
        }

        // Create ParameterModel from an Excel row
        private ParameterModel CreateParameterModelFromRow(IXLRow row)
        {
            var controlType = row.Cell(4).GetString().Trim();

            var parameterModel = new ParameterModel
            {
                Name = row.Cell(1).GetString().Trim(),
                Value = string.Empty,
                ControlType = controlType,
                IsVisible = row.Cell(7).GetString().Trim().ToUpper() != "FALSE",
                IsEnabled = row.Cell(8).GetString().Trim().ToUpper() != "FALSE",
                DependsOn = row.Cell(9).GetString().Trim(),
                TriggerValue = row.Cell(10).GetString().Trim()
            };

            if (controlType == "CheckBoxList")
                parameterModel.SelectedValues = new ObservableCollection<string>();

            return parameterModel;
        }

        // Parse allowed values or range rules
        private void ParseValidationRule(string rule, ParameterModel parameter)
        {
            if (string.IsNullOrWhiteSpace(rule)) return;

            if (rule.Contains("-"))
            {
                var parts = rule.Split('-');
                if (parts.Length == 2 &&
                    double.TryParse(parts[0], out double min) &&
                    double.TryParse(parts[1], out double max))
                {
                    if (min <= max)
                    {
                        parameter.MinValue = min;
                        parameter.MaxValue = max;
                    }
                    else
                    {
                        _logger.Warn($"Invalid range format in Excel for {parameter.Name}: {rule} (min > max)");
                    }
                }
                else
                {
                    _logger.Warn($"Invalid range format in Excel for {parameter.Name}: {rule}");
                }
            }
            else if (rule.Contains(","))
            {
                var values = rule.Split(',')
                    .Select(x => x.Trim())
                    .Where(x => !string.IsNullOrEmpty(x))
                    .ToList();

                if (values.Any())
                {
                    parameter.AllowedValues = values;
                }
                else
                {
                    _logger.Warn($"Empty allowed values list in Excel for {parameter.Name}: {rule}");
                }
            }
            // Additional custom validation logic can go here if needed
        }

        // Handle file access exceptions
        private void HandleFileAccessError(IOException ioEx)
        {
            _logger.Error(ioEx, "File Access Error While Reading The Excel File.");
            MessageBox.Show($"Error Accessing The Excel File. Ensure It's Not Open In Another Application.\n\n{ioEx.Message}",
                "File Access Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        // Handle other exceptions
        private void HandleUnexpectedError(Exception ex)
        {
            _logger.Error(ex, "Unexpected Error While Reading Excel Parameters.");
            MessageBox.Show($"An Error Occurred While Reading The Excel File.\n\n{ex.Message}",
                "Excel Read Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        // ----- Model Data Loading Section -----

        public async Task LoadModelsFromFileAsync()
        {
            var excelDir = ConfigurationManager.AppSettings["ExcelFileDirectory"];
            if (string.IsNullOrEmpty(excelDir))
                throw new ConfigurationErrorsException("ExcelFileDirectory is missing in config.");

            var dataExcelPath = _viewModel.SelectedConfigurator.ModelDataExcelPath;

            try
            {
                var (models, parameterMap) = await ReadExcelDataAsync(dataExcelPath);

                Application.Current.Dispatcher.Invoke(() =>
                {
                    UpdateViewModelCollections(models, parameterMap);
                    _viewModel.SelectedModel = _viewModel.Models.FirstOrDefault();
                });
            }
            catch (Exception ex)
            {
                HandleLoadError(ex);
            }
        }

        private async Task<(List<string>, Dictionary<string, List<ModelData>>)> ReadExcelDataAsync(string filePath)
        {
            return await Task.Run(() =>
            {
                var models = new List<string>();
                var parameterMap = new Dictionary<string, List<ModelData>>(StringComparer.OrdinalIgnoreCase);
                var existingModels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                try
                {
                    _logger.Info($"Attempting to read Excel file: {filePath}");

                    using (var workbook = new XLWorkbook(filePath))
                    {
                        var worksheet = workbook.Worksheet(1);
                        ValidateWorksheetStructure(worksheet);

                        var headerRow = worksheet.Row(1);
                        var parameterNames = GetParameterNames(headerRow).ToList();

                        foreach (var row in worksheet.RowsUsed().Skip(1))
                        {
                            var (modelName, modelData) = ProcessDataRow(row, parameterNames);
                            if (string.IsNullOrWhiteSpace(modelName)) continue;

                            if (existingModels.Add(modelName))
                            {
                                models.Add(modelName);
                                _logger.Debug($"Model '{modelName}' added.");
                            }

                            parameterMap[modelName] = modelData;
                        }
                        _logger.Info("Successfully read and parsed Excel data.");
                    }
                }
                catch (IOException ioEx)
                {
                    _logger.Error(ioEx, $"File in use: {filePath}");
                    MessageBox.Show($"The Excel file is currently in use. Please close it and try again.\n\nDetails: {ioEx.Message}",
                                    "File Access Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "An unexpected error occurred while reading the Excel file.");
                    MessageBox.Show($"An unexpected error occurred while reading the Excel file:\n\n{ex.Message}",
                                    "Unexpected Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                return (models, parameterMap);
            });
        }

        private void ValidateWorksheetStructure(IXLWorksheet worksheet)
        {
            var headerRow = worksheet.Row(1);
            if (!headerRow.Cell(1).GetString().Trim().Equals("Model", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("First column header must be 'Model'");
            }
        }

        private IEnumerable<string> GetParameterNames(IXLRow headerRow)
        {
            int lastColumn = headerRow.LastCellUsed().Address.ColumnNumber;
            for (int col = 1; col <= lastColumn; col++)
            {
                yield return headerRow.Cell(col).GetString().Trim();
            }
        }

        private (string ModelName, List<ModelData> Data) ProcessDataRow(IXLRow row, List<string> parameterNames)
        {
            var modelName = row.Cell(1).GetString().Trim();
            var modelData = new List<ModelData>();

            int columnCount = parameterNames.Count;
            for (int col = 1; col <= columnCount; col++)
            {
                var cellValue = row.Cell(col).GetString().Trim();
                modelData.Add(new ModelData
                {
                    Name = parameterNames[col - 1],
                    Value = cellValue
                });
            }
            return (modelName, modelData);
        }

        private void UpdateViewModelCollections(List<string> newModels, Dictionary<string, List<ModelData>> newParameterMap)
        {
            var modelsToRemove = _viewModel.Models
                .Where(m => !m.Equals("(New Model)", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var model in modelsToRemove)
            {
                _viewModel.Models.Remove(model);
                _viewModel.ModelParameterMap.Remove(model);
            }

            foreach (var model in newModels)
            {
                if (!_viewModel.Models.Any(m => m.Equals(model, StringComparison.OrdinalIgnoreCase)))
                {
                    _viewModel.Models.Add(model);
                    _viewModel.ModelParameterMap[model] = newParameterMap[model];
                }
            }

            // Ensure "(New Model)" exists and is first
            if (!_viewModel.Models.Contains("(New Model)"))
            {
                _viewModel.Models.Insert(0, "(New Model)");
            }
            else
            {
                var currentIndex = _viewModel.Models.IndexOf("(New Model)");
                if (currentIndex > 0)
                    _viewModel.Models.Move(currentIndex, 0);
            }
        }

        private void HandleLoadError(Exception ex)
        {
            _logger.Error(ex, "Model loading error");

            string message = ex is InvalidDataException ide
                ? $"Invalid Excel format: {ide.Message}"
                : ex is IOException ioe
                ? $"File access error: {ioe.Message}"
                : $"Error loading models: {ex.Message}";

            Application.Current.Dispatcher.Invoke(() =>
            {
                MessageBox.Show(message, "Loading Error", MessageBoxButton.OK, MessageBoxImage.Error);
                _viewModel.StatusMessage = "Error loading model data";
            });
        }
    }
}
