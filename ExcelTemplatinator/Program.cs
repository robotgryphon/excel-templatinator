using HandlebarsDotNet;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToDataTable;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace ScarfPupperBestPupper
{
    class Program
    {
        static DirectoryInfo OUTPUT_ROOT;
        static readonly Regex NON_ALPHANUMERIC = new Regex(@"[^\w]");

        static void Main()
        {
            TemplateDataConfig templateConfig = LoadTemplateConfigFile();

            if (templateConfig.Areas.Length == 0)
            {
                Console.WriteLine("No template areas defined. Define areas and re-run the program.");
                return;
            }

            if (string.IsNullOrEmpty(templateConfig.Input) || string.IsNullOrEmpty(templateConfig.OutputTemplate))
            {
                Console.WriteLine("No input or output file defined. Check the config and re-run the program.");
                return;
            }

            FileInfo INPUT_FILE = new FileInfo(templateConfig.Input);
            OUTPUT_ROOT = new DirectoryInfo(templateConfig.OutputDir);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(INPUT_FILE))
            {
                var template = package.Workbook.Worksheets[0];
                Dictionary<string, TemplateCompiled> templateLookup = BuildCompiledTemplateLookup(templateConfig, template);

                var data = LoadDataFromFile(templateConfig);

                PopulateOutputFiles(templateConfig, package, template, templateLookup, data);
            }

            var holdColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Press enter to exit.");
            Console.ReadLine();
            Console.ForegroundColor = holdColor;
        }

        private static Dictionary<string, TemplateCompiled> BuildCompiledTemplateLookup(TemplateDataConfig templateConfig, ExcelWorksheet template)
        {
            var templateLookup = new Dictionary<string, TemplateCompiled>();
            foreach (var area in templateConfig.Areas)
            {
                var range = template.Cells[area.Range];
                foreach (var cell in range)
                {
                    Console.WriteLine(cell.Address);

                    if (!templateLookup.ContainsKey(cell.Address))
                    {
                        string cellValueString = cell.GetValue<string>();

                        var cellTemplate = Handlebars.Compile(cellValueString);
                        templateLookup.Add(cell.Address, new TemplateCompiled
                        {
                            Template = cellTemplate
                        });
                    }
                }
            }

            return templateLookup;
        }

        private static void PopulateOutputFiles(TemplateDataConfig templateConfig, ExcelPackage package, ExcelWorksheet template, Dictionary<string, TemplateCompiled> templateLookup, DataTable dt)
        {
            List<string> columns = new List<string>();
            Dictionary<string, string> dtMapping = new Dictionary<string, string>();

            for (int coli = 0; coli < dt.Columns.Count; coli++)
            {
                // Replace all non-alphanumeric and underscores with underscores (template simplicity)
                string originalColName = dt.Columns[coli].ColumnName;
                string colNameFixed = NON_ALPHANUMERIC.Replace(originalColName, "_");

                columns.Add(colNameFixed);
                dtMapping.Add(colNameFixed, originalColName);
            }

            for (int rowi = 0; rowi < dt.Rows.Count; rowi++)
            {
                var row = dt.Rows[rowi];

                Dictionary<string, object> rowData = new Dictionary<string, object>();

                foreach (var col in columns)
                {
                    object colData = row[dtMapping[col]];
                    rowData.Add(col, colData);

                }

                FillTemplateForEntry(templateConfig, package, template, templateLookup, rowData);
            }
        }

        private static DataTable? LoadDataFromFile(TemplateDataConfig templateConfig)
        {
            using ExcelPackage templateDataPkg = new ExcelPackage(new FileInfo(templateConfig.Data.File));

            var data = templateDataPkg.Workbook.Worksheets[0];
            var dataSel = data.Cells[templateConfig.Data.Range];

            if (dataSel.Rows <= 1)
            {
                Console.WriteLine("No header or not enough data to populate template.");
                return null;
            }

            var opts = ToDataTableOptions.Create();
            opts.ExcelErrorParsingStrategy = ExcelErrorParsingStrategy.HandleExcelErrorsAsBlankCells;
            opts.ColumnNameParsingStrategy = NameParsingStrategy.SpaceToUnderscore;
            opts.EmptyRowStrategy = EmptyRowsStrategy.Ignore;

            return dataSel.ToDataTable(opts);
        }

        private static void FillTemplateForEntry(TemplateDataConfig templateConfig, ExcelPackage package, ExcelWorksheet template, Dictionary<string, TemplateCompiled> templateLookup, Dictionary<string, object> data)
        {

            // Template registration complete; now loop back through and set ALL the data
            foreach (var temp in templateLookup)
            {
                string finalData = temp.Value.Template(data);

                template.Cells[temp.Key].Value = finalData;
            }

            var outputFilenameTemp = Handlebars.Compile(templateConfig.OutputTemplate);
            string outputFilenameReal = outputFilenameTemp(data);

            FileInfo OUTPUT_FILE = new FileInfo(Path.Combine(OUTPUT_ROOT.FullName, outputFilenameReal));
            if (!OUTPUT_ROOT.Exists)
                OUTPUT_ROOT.Create();

            // Delete output if it exists already
            if (OUTPUT_FILE.Exists)
                OUTPUT_FILE.Delete();

            package.SaveAs(OUTPUT_FILE);
        }

        private static TemplateDataConfig LoadTemplateConfigFile()
        {
            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("template-data.json", false, false)
                .Build();

            var templateConfig = config.Get<TemplateDataConfig>();
            return templateConfig;
        }
    }
}
