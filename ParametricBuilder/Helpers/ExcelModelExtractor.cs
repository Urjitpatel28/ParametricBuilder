using ClosedXML.Excel;
using System.Collections.ObjectModel;
using System.IO;

namespace ParametricBuilder.Helpers
{
    public class ExcelModelExtractor
    {
        public static ObservableCollection<string> GetModelList(string filePath)
        {
            var models = new ObservableCollection<string>();

            if (!File.Exists(filePath))
                return models;

            models.Add("(New Model)");
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // assuming data is in the first sheet
                int row = 2; // start from row 2 to skip header

                while (true)
                {
                    var modelCell = worksheet.Cell(row, 1); // Column A
                    var modelValue = modelCell.GetString();

                    if (string.IsNullOrWhiteSpace(modelValue))
                        break;

                    models.Add(modelValue);
                    row++;
                }
            }

            return models;
        }
    }
}
