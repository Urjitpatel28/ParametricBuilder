using System;
using System.Configuration;
using System.Xml;

namespace ParametricBuilder.Helpers
{
    public static class UpdateExcelPathInConfig
    {
        public static bool UpdateExcelPathInConfigSafe(string configFilePath, string excelPath)
        {
            try
            {
                XmlDocument configDoc = new XmlDocument();
                configDoc.Load(configFilePath);

                XmlNode excelNode = configDoc.SelectSingleNode(
                    "/configuration/appSettings/add[@key='CADExcelSheet']");

                if (excelNode?.Attributes?["value"] == null)
                {
                    throw new ConfigurationErrorsException("Invalid configuration format");
                }

                excelNode.Attributes["value"].Value = excelPath;
                configDoc.Save(configFilePath);
                return true;
            }
            catch (Exception ex)
            {
                // Handle specific exceptions here
                throw new ApplicationException(
                    $"Failed to update configuration file: {ex.Message}", ex);
            }
        }
    }
}
