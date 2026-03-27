using NLog;
using ParametricBuilder.Models;
using ParametricBuilder.ViewModels;
using System;

namespace ParametricBuilder.Commands
{
    public class BrowseCommand : BaseCommand
    {
        private readonly IFileHandler _viewModel;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public BrowseCommand(IFileHandler viewModel)
        {
            _viewModel = viewModel;
        }

        public override bool CanExecute(object parameter) => true;

        public override void Execute(object parameter)
        {
            try
            {
                _logger.Info("Executing BrowseCommand.");
                var openFileDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    _viewModel.HandleFileSelected(openFileDialog.FileName);
                    _logger.Info($"File selected: {openFileDialog.FileName}");
                }
                else
                {
                    _logger.Warn("No File Was Selected.");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed To Execute BrowseCommand.");
            }
        }
    }
}