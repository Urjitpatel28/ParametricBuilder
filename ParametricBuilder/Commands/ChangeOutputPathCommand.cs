using ParametricBuilder.ViewModels;
using System.Configuration;
using System.Windows.Forms;

namespace ParametricBuilder.Commands
{
    public class ChangeOutputPathCommand : BaseCommand
    {
        private readonly MainWindowViewModel _viewModel;

        public ChangeOutputPathCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        public override bool CanExecute(object parameter)
        {
            // You can add logic here to disable/enable the button if needed
            return true;
        }

        public override void Execute(object parameter)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select a folder for the output path";

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string newPath = folderDialog.SelectedPath;

                    // Update the OutputPath in the ViewModel
                    _viewModel.OutputPath = newPath;

                    // Ask the user whether they want to save this change permanently
                    var result = MessageBox.Show(
                        "Do you want to change the output directory permanently?",
                        "Save Output Directory",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        // If yes, update the app.config permanently
                        UpdateAppConfig(_viewModel.OutputPath);
                    }
                    else
                    {
                        // If no, just update the OutputPath for this session
                        // No need to do anything else, the change will be temporary
                    }
                }
            }
        }

        private void UpdateAppConfig(string newPath)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["OutPutDirectory"].Value = newPath;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }       
    }
}
