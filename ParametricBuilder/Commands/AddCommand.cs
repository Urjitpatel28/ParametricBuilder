using ClosedXML.Excel;
using NLog;
using ParametricBuilder.Models;
using ParametricBuilder.ViewModels;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Windows;

namespace ParametricBuilder.Commands
{
    public class AddCommand : BaseCommand
    {
        private readonly MainWindowViewModel _viewModel;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public AddCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        public override bool CanExecute(object parameter)
        {
            return !string.IsNullOrWhiteSpace(_viewModel.SelectedModel);
        }

        public override void Execute(object parameter)
        {
            try
            {
                // Get the first parameter's value to be used as the model name
                var newModelName = _viewModel.Parameters.FirstOrDefault(p => !string.IsNullOrWhiteSpace(p.Value))?.Value;
                _logger.Info($"Attempting to add new model with name: {newModelName}");

                if (string.IsNullOrEmpty(newModelName))
                {
                    _logger.Warn("Model name is empty.");
                    MessageBox.Show("Model name cannot be empty.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Check if the model already exists in the map
                if (_viewModel.ModelParameterMap.ContainsKey(newModelName))
                {
                    _logger.Warn($"Model '{newModelName}' already exists.");
                    MessageBox.Show($"Model '{newModelName}' already exists. \nPlease update the model instead.", "Model Exists!", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return; // Prevent adding the model if it already exists
                }

                // Prepare the new model's data from parameters
                var newModelData = _viewModel.Parameters.Select(param => new ModelData
                {
                    Name = param.Name,
                    Value = param.Value
                }).ToList();

                // Add the new model data to the parameter map
                _viewModel.ModelParameterMap[newModelName] = newModelData;
                _logger.Info($"Added new model '{newModelName}' to ModelParameterMap.");

                // Add model name to Models collection
                _viewModel.Models.Add(newModelName);
                _logger.Info($"Added model '{newModelName}' to Models collection.");

                // Save the new model to Excel
                AddNewModelToExcel(newModelName, newModelData);

                // Set the newly added model as the selected model
                _viewModel.SelectedModel = newModelName;
                _logger.Info($"Selected model set to: {newModelName}");

                // Update status message
                _viewModel.StatusMessage = $"Added new model: {newModelName}.";
                _logger.Info($"New model '{newModelName}' added successfully.");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error adding new model.");
                MessageBox.Show("Error adding new model.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddNewModelToExcel(string modelName, List<ModelData> modelData)
        {
            var excelPath = _viewModel.SelectedConfigurator.ModelDataExcelPath;
            _logger.Debug($"Saving new model '{modelName}' to Excel file at path: {excelPath}");

            try
            {
                using (var workbook = new XLWorkbook(excelPath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                    var newRow = worksheet.Row(lastRow + 1); // Add new row after the last used row

                    // Set the model name in the first column
                    newRow.Cell(1).Value = modelName;
                    _logger.Debug($"Set model name '{modelName}' in row {lastRow + 1}");

                    // Set the parameter values for the new model (start from second column)
                    for (int i = 1; i < modelData.Count; i++)
                    {
                        newRow.Cell(i + 1).Value = modelData[i].Value; // Starting from column 2 (index 1)
                        _logger.Debug($"Set parameter '{modelData[i].Name}' to value '{modelData[i].Value}' in row {lastRow + 1}, column {i + 2}");
                    }

                    workbook.Save();
                    _logger.Info($"Excel file saved successfully with new model '{modelName}'.");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error saving new model data to Excel.");
                MessageBox.Show("Error saving model data to Excel.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}