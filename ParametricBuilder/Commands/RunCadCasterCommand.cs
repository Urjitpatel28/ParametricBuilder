using ClosedXML.Excel;
using NLog;
using ParametricBuilder.Helpers;
using ParametricBuilder.ViewModels;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace ParametricBuilder.Commands
{
    public class RunCadCasterCommand : BaseCommand
    {
        private readonly MainWindowViewModel _viewModel;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public RunCadCasterCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        public override bool CanExecute(object parameter)
        {
            return _viewModel.Parameters
                .Where(p => p.IsVisible && p.IsEnabled)
                .All(p => !p.HasErrors && !string.IsNullOrWhiteSpace(p.Value))
                && !string.IsNullOrWhiteSpace(_viewModel.OutputPath);
        }

        public override async void Execute(object parameter)
        {
            try
            {
                BeginLoading();
                ExcelHelper.CloseOrphanedExcelProcesses();

                await Task.Run(async () =>
                {
                    using (var workbook = new XLWorkbook(_viewModel.SelectedConfigurator.ConfigExcelTemplatePath))
                    {
                        UpdateParameterValuesInWorkbook(workbook);
                        UpdateOutputPathInWorkbook(workbook);

                        workbook.Save();
                        _logger.Info("Workbook Saved Successfully.");

                        string cadCasterPath = ConfigurationManager.AppSettings["CadCasterPath"];
                        if (!ValidateCadCasterPath(cadCasterPath))
                            return;

                        string configFilePath = Path.Combine(cadCasterPath, "SolidworksConfigurator.exe.config");
                        string exePath = Path.Combine(cadCasterPath, "SolidworksConfigurator.exe");
                        string excelFullPath = GetExcelFullPath(_viewModel.SelectedConfigurator.ConfigExcelTemplatePath);

                        if (!UpdateCadCasterConfig(configFilePath, excelFullPath))
                            return;

                        await RunCadCasterExeAsync(exePath);
                    }
                });

                _viewModel.StatusMessage = "Done!";
            }
            catch (Exception ex)
            {
                HandleError(ex);
            }
            finally
            {
                ClearValueColumn(_viewModel.SelectedConfigurator.ConfigExcelTemplatePath);
                EndLoading();
                ExcelHelper.CloseOrphanedExcelProcesses();
            }
        }

        // ----- UI/Loading Helpers -----

        private void BeginLoading()
        {
            Mouse.OverrideCursor = Cursors.Wait;
            _viewModel.LoadingWheelVisibility = Visibility.Visible;
            _viewModel.StatusMessage = "Updating Excel Parameters....";
        }

        private void EndLoading()
        {
            Mouse.OverrideCursor = null;
            _viewModel.LoadingWheelVisibility = Visibility.Hidden;
            _logger.Info("Excel Update Process Completed.");
            _logger.Info("CAD Model Update Process Completed.");
        }

        // ----- Workbook/Parameter Helpers -----

        private void UpdateParameterValuesInWorkbook(XLWorkbook workbook)
        {
            var worksheet = workbook.Worksheet(1);
            int rowIndex = 2;
            foreach (var param in _viewModel.Parameters)
            {
                var cell = worksheet.Cell(rowIndex, 2);
                if (double.TryParse(param.Value, out double numericValue))
                {
                    cell.Value = numericValue;
                    _logger.Debug($"Parameter '{param.Name}' Set To Numeric Value: {numericValue} At Row {rowIndex}.");
                }
                else
                {
                    cell.Value = param.Value;
                    _logger.Debug($"Parameter '{param.Name}' Set To String Value: {param.Value} At Row {rowIndex}.");
                }
                rowIndex++;
            }
        }

        private void UpdateOutputPathInWorkbook(XLWorkbook workbook)
        {
            var secondWorksheet = workbook.Worksheet(2);
            secondWorksheet.Cell("B2").Value = _viewModel.OutputPath;
        }

        // ----- CadCaster Path/Config/Execution -----

        private bool ValidateCadCasterPath(string cadCasterPath)
        {
            if (string.IsNullOrWhiteSpace(cadCasterPath) || !Directory.Exists(cadCasterPath))
            {
                string errorMessage = "CadCaster folder not found! Please check the CadCaster location and update the config file accordingly.";
                MessageBox.Show(errorMessage, "Check CadCaster Location", MessageBoxButton.OK, MessageBoxImage.Error);
                _logger.Warn(errorMessage);
                return false;
            }
            return true;
        }

        private string GetExcelFullPath(string relativePath)
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            return Path.Combine(baseDirectory, relativePath);
        }

        private bool UpdateCadCasterConfig(string configFilePath, string excelFullPath)
        {
            bool result = UpdateExcelPathInConfig.UpdateExcelPathInConfigSafe(configFilePath, excelFullPath);

            if (!result)
            {
                string updateError = "Failed to update the Excel path in the config file.";
                MessageBox.Show(updateError, "Update Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                _logger.Error(updateError);
                return false;
            }

            _viewModel.StatusMessage = "Updating CAD Model";
            return true;
        }

        private async Task RunCadCasterExeAsync(string exePath)
        {
            try
            {
                await exeHelper.RunAsAdministrator(exePath);
            }
            catch (Exception ex)
            {
                string executionError = $"Failed to execute '{exePath}'. Error: {ex.Message}";
                MessageBox.Show(executionError, "Execution Error", MessageBoxButton.OK, MessageBoxImage.Error);
                _logger.Error(executionError);
            }
        }

        // ----- Error Handling -----

        private void HandleError(Exception ex)
        {
            _logger.Error(ex, "Failed To Update The Workbook.");
            MessageBox.Show($"An Error Occurred While Updating The Excel File.\n\n{ex.Message}", "Excel Update Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void ClearValueColumn(string excelPath)
        {
            try
            {
                using (var workbook = new XLWorkbook(excelPath))
                {
                    var worksheet = workbook.Worksheet(1);
                    // Assuming first row is header, data starts from row 2
                    int lastRow = worksheet.LastRowUsed().RowNumber();

                    for (int row = 2; row <= lastRow; row++)
                    {
                        worksheet.Cell(row, 2).Clear();
                    }

                    workbook.Save();
                    _logger.Info("Cleared all values from column 2 and saved workbook.");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to clear second column values.");
                MessageBox.Show($"Failed to clear second column: {ex.Message}");
            }
        }
    }
}