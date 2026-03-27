using ClosedXML.Excel;
using NLog;
using ParametricBuilder.ViewModels;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Windows;

namespace ParametricBuilder.Commands
{
    public class DeleteCommand : BaseCommand
    {
        private readonly MainWindowViewModel _viewModel;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();


        public DeleteCommand(MainWindowViewModel viewModel)
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
                var modelNameToDelete = _viewModel.SelectedModel; // Use selected model for deletion
                _logger.Info($"Attempting to delete model: {modelNameToDelete}");

                // Check if the selected model is "(New Model)"
                if (modelNameToDelete == "(New Model)")
                {
                    _logger.Warn("Attempted to delete a new model.");
                    MessageBox.Show("You cannot delete a new model.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return; // Exit the method if it's a new model
                }

                if (!_viewModel.ModelParameterMap.ContainsKey(modelNameToDelete))
                {
                    _logger.Warn($"Model '{modelNameToDelete}' not found in data.");
                    MessageBox.Show("Model Not Found in Data!", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return; // Exit the method if the model doesn't exist in the map
                }

                // Confirm with the user if they really want to delete the model
                var confirmationResult = MessageBox.Show(
                    $"Are you sure you want to delete the model: '{modelNameToDelete}'?",
                    "Confirm Deletion",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning
                );

                // If user clicks 'Yes', proceed with deletion
                if (confirmationResult == MessageBoxResult.Yes)
                {
                    // Remove from ModelParameterMap
                    if (_viewModel.ModelParameterMap.ContainsKey(modelNameToDelete))
                    {
                        _viewModel.ModelParameterMap.Remove(modelNameToDelete);
                        _logger.Info($"Model '{modelNameToDelete}' removed from ModelParameterMap.");
                    }

                    // Remove from Models collection
                    _viewModel.Models.Remove(modelNameToDelete);
                    _logger.Info($"Model '{modelNameToDelete}' removed from Models collection.");

                    // Clear the Values in Parameters collection since the model was deleted (retain other properties)
                    foreach (var param in _viewModel.Parameters)
                    {
                        param.Value = string.Empty; // Only clear the value, other properties remain intact
                        _logger.Debug($"Cleared value for parameter: {param.Name}");
                    }
                    _logger.Info("Cleared Parameters collection after deletion.");

                    // Delete the model from Excel
                    DeleteModelFromExcel(modelNameToDelete);

                    // Update the StatusMessage after deletion
                    _viewModel.StatusMessage = $"Deleted model: {modelNameToDelete}.";
                    _logger.Info($"Successfully deleted model '{modelNameToDelete}'.");

                    // Set SelectedModel to "(New Model)" and reset parameters
                    _viewModel.SelectedModel = "(New Model)";
                    _logger.Info("SelectedModel set to '(New Model)' after deletion.");

                    // Optionally, reset the Parameter collection to clear the values
                    //foreach (var param in _viewModel.Parameters)
                    //{
                    //    param.Value = string.Empty;
                    //    _logger.Debug($"Cleared value for parameter: {param.Name}");
                    //}
                }
                else
                {
                    _viewModel.StatusMessage = "Model deletion canceled.";
                    _logger.Info("Model deletion canceled by user.");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error deleting model.");
                MessageBox.Show("Error deleting model.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DeleteModelFromExcel(string modelName)
        {
            var excelPath = _viewModel.SelectedConfigurator.ModelDataExcelPath;
            _logger.Info($"Attempting to delete model '{modelName}' from Excel file at path: {excelPath}");

            try
            {
                using (var workbook = new XLWorkbook(excelPath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var modelRow = worksheet.RowsUsed().FirstOrDefault(row => row.Cell(1).GetString().Trim() == modelName);

                    if (modelRow != null)
                    {
                        modelRow.Delete();  // Delete the row
                        workbook.Save();    // Save the workbook
                        _logger.Info($"Model '{modelName}' deleted from Excel file.");
                    }
                    else
                    {
                        _logger.Warn($"Model '{modelName}' not found in Excel file.");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"Failed to delete model '{modelName}' from Excel.");
                MessageBox.Show("Error deleting model from Excel.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}