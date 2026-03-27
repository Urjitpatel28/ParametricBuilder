using System;
using System.IO;

namespace ParametricBuilder.Models
{
    public class Configurator
    {
        // UI-facing property
        public string DisplayName { get; set; }

        // Backend configuration paths
        public string ConfigExcelTemplatePath { get; set; }
        public string ModelDataExcelPath { get; set; }
        public string MasterCardFolderPath { get; set; }
        public string GeometryImagePath { get; set; }
        public string GeometryImagePathAbsolute
        {
            get
            {
                if (string.IsNullOrWhiteSpace(GeometryImagePath))
                    return null;

                // If it's already rooted (absolute), return directly
                if (Path.IsPathRooted(GeometryImagePath))
                    return GeometryImagePath;

                // Otherwise, base it on app directory
                var exeDir = AppDomain.CurrentDomain.BaseDirectory;
                return Path.Combine(exeDir, GeometryImagePath);
            }
        }
    }
}
