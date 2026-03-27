using ClosedXML.Excel;
using NLog;
using ParametricBuilder.Models;
using ParametricBuilder.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace ParametricBuilder.Commands
{
    public class UpdateCommand : BaseCommand
    {
        private readonly MainWindowViewModel _viewModel;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public UpdateCommand(MainWindowViewModel viewModel)
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
                var modelNameToUpdate = _viewModel.SelectedModel;
                _logger.Info($"Attempting to update model: {modelNameToUpdate}");

                if (modelNameToUpdate == "(New Model)")
                {
                    _logger.Warn("Attempted to update a new model.");
                    MessageBox.Show("You cannot update a new model.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return; // Exit the method if it's a new model
                }

                // Check if the model exists in the ModelParameterMap
                if (!_viewModel.ModelParameterMap.ContainsKey(modelNameToUpdate))
                {
                    _logger.Warn($"Model '{modelNameToUpdate}' not found in ModelParameterMap.");
                    MessageBox.Show($"Model '{modelNameToUpdate}' not found.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var modelParameters = _viewModel.ModelParameterMap[modelNameToUpdate];
                var modelParamDictionary = modelParameters.ToDictionary(p => p.Name, StringComparer.OrdinalIgnoreCase);

                _logger.Debug($"Found model '{modelNameToUpdate}' with {modelParameters.Count} parameters to update.");

                // Update parameters based on UI changes (bound to `Parameters`)
                foreach (var param in _viewModel.Parameters)
                {
                    // Find matching parameter in the model data using the dictionary for fast lookups
                    if (modelParamDictionary.TryGetValue(param.Name, out var modelParam))
                    {
                        modelParam.Value = param.Value; // Update the model parameter value
                        _logger.Debug($"Updated parameter '{param.Name}' to value '{param.Value}' in model '{modelNameToUpdate}'.");
                    }
                    else
                    {
                        _logger.Warn($"Parameter '{param.Name}' not found in model '{modelNameToUpdate}'.");
                    }
                }

                // Save the changes to the Excel file
                SaveModelDataToExcel(modelNameToUpdate, modelParameters);

                // Update the status message
                _viewModel.StatusMessage = $"Updated {modelNameToUpdate} with new parameters.";
                _logger.Info($"Model '{modelNameToUpdate}' updated successfully.");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error updating model data.");
                MessageBox.Show("Error updating model data.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveModelDataToExcel(string modelName, List<ModelData> modelParameters)
        {
            var excelPath = _viewModel.SelectedConfigurator.ModelDataExcelPath;
            _logger.Info($"Saving model '{modelName}' data to Excel file at path: {excelPath}");

            try
            {
                using (var workbook = new XLWorkbook(excelPath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var modelRow = worksheet.RowsUsed().FirstOrDefault(row => row.Cell(1).GetString().Trim() == modelName);

                    if (modelRow != null)
                    {
                        _logger.Debug($"Found model row for '{modelName}'. Updating values...");
                        // Update the row with new values for each parameter
                        for (int i = 1; i < modelParameters.Count; i++)
                        {
                            modelRow.Cell(i + 1).Value = modelParameters[i].Value; // Starting from 2nd column (index 1)
                            _logger.Debug($"Set value for parameter '{modelParameters[i].Name}' to '{modelParameters[i].Value}'.");
                        }
                        workbook.Save();
                        _logger.Info($"Model '{modelName}' data saved successfully to Excel.");
                    }
                    else
                    {
                        _logger.Warn($"Model '{modelName}' not found in Excel file.");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to save model data to Excel.");
                MessageBox.Show("Error saving model data to Excel.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}