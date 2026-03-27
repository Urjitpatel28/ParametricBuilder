using ClosedXML.Excel;
using ParametricBuilder.Models;
using System.Collections.ObjectModel;
using System.IO;

namespace ParametricBuilder.ViewModels
{
    public class ConfiguratorDataService
    {
        public ObservableCollection<Configurator> LoadConfiguratorData(string filePath)
        {
            var configurators = new ObservableCollection<Configurator>();

            // Open the Excel file using ClosedXML
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);  // Assumes data is in the first worksheet

                // Get the last row with data
                var lastRow = worksheet.LastRowUsed().RowNumber();

                // Loop through the rows to fetch configurator data (starting from row 2, assuming row 1 has headers)
                for (int row = 2; row <= lastRow; row++)
                {
                    var configurator = new Configurator
                    {
                        DisplayName =  worksheet.Cell(row, 1).GetValue<string>(), // Configurator Name
                        ConfigExcelTemplatePath = Path.Combine("Config.Data\\ConfigExcelTemplates", worksheet.Cell(row, 2).GetValue<string>()), // Config Excel Template Path
                        ModelDataExcelPath = Path.Combine("Config.Data\\ModelDataExcel", worksheet.Cell(row, 3).GetValue<string>()), // Model Data Excel Path
                        MasterCardFolderPath = Path.Combine("Config.Data\\MasterCadFiles", worksheet.Cell(row, 4).GetValue<string>()),// Master CAD Folder Path
                        GeometryImagePath = Path.Combine("Config.Data\\GeometryImages", worksheet.Cell(row, 5).GetValue<string>())
                    };

                    configurators.Add(configurator);
                }
            }

            return configurators;
        }
    }
}