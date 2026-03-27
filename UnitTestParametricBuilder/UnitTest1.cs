using Microsoft.VisualStudio.TestTools.UnitTesting;
using ParametricBuilder.Helpers;
using System;

namespace UnitTestParametricBuilder
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            string configFilePath = @"C:\Users\upatel\source\repos\ParametricBuilder\ParametricBuilder\Resource\CadCasterSWX\SolidworksConfigurator.exe.config";
            string excelFilePath = @"C:\Users\upatel\Desktop\ConfigData.xlsx";
            UpdateExcelPathInConfig.UpdateExcelPathInConfigSafe(configFilePath, excelFilePath);
        }
    }
}
