using ClosedXML.Excel;
using CommandLine;
using CommandLine.Text;
using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;

namespace DjangoTranslationHelper
{
    class Program
    {

        class Options
        {
            [Option('p', "archivopo", Required = true, HelpText = "PO Input file to be processed.")]
            public string POInputFile { get; set; }

            [Option('l', "lenguajes", Required = true, HelpText = "A comma separated string with the languages to proc")]
            public string Languages { get; set; }

            [Option('e', "excelfile", Required = false, HelpText = "A xlsx file name")]
            public string FileName { get; set; }

            [Option('m', "modotrabajo", Required = true, HelpText = "A xlsx file name")]
            public string TypeOfWork { get; set; }

            [ParserState]
            public IParserState LastParserState { get; set; }

            [HelpOption]
            public string GetUsage()
            {
                return HelpText.AutoBuild(this, (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
            }
        }

        // Consume them
        static void Main(string[] args)
        {
            var options = new Options();
            if (!CommandLine.Parser.Default.ParseArguments(args, options))
            {
                Console.Write(options.GetUsage());
                return;
            }

            if (!File.Exists(options.POInputFile))
            {
                Console.Write(options.GetUsage());
                return;
            }

            string[] languages = options.Languages.Split(',');
            if (languages.Length == 0)
            {
                Console.Write(options.GetUsage());
                return;
            }

            if (options.TypeOfWork == "exportar")
            {
                ExportFile(options);
            }

            if (options.TypeOfWork == "importar")
            {
                ImportFile(options);
            }


        }

        private static void ImportFile(Options options)
        {
            var dictionaries = GetLanguageDictinaries(options.FileName);
            int currentLanguage = 0;

            string[] templateFile = File.ReadAllLines(options.POInputFile);

            foreach (string language in options.Languages.Split(','))
            {
                string folder = Path.GetDirectoryName(options.FileName);
                string newPoFileName = Path.GetFileNameWithoutExtension(options.FileName);
                newPoFileName = folder + "/" +newPoFileName + "-" +language +".po";
                CreatePoFromTemplate(templateFile, dictionaries[currentLanguage], newPoFileName);
                currentLanguage++;
            }
        }

        private static void CreatePoFromTemplate(string[] templateFile, Dictionary<string,string> currentLanguage, string newPoFileName)
        {
            string[] result = templateFile;
            foreach (string key in currentLanguage.Keys)
            {
                string regExp = "msgid \"" + key + "\"";
                string replace = "msgstr \""+ currentLanguage[key]+ "\"";

                for (int i = 0; i < templateFile.Length; i++)
                {
                    if (templateFile[i] == regExp)
                    {
                        result[i + 1] = replace;
                    }
                }
            }

            File.WriteAllLines(newPoFileName, result);
        }

        private static List<Dictionary<string,string>> GetLanguageDictinaries(string fileName)
        {
            List<Dictionary<string, string>> result = new List<Dictionary<string, string>>();
            XLWorkbook data = new XLWorkbook(fileName);

            var traducciones = data.Worksheets.First();
            int numberOfLanguages = traducciones.Row(1).CellsUsed().Count() - 1;
            for (int j = 0; j < numberOfLanguages; j++)
            {
                result.Add(new Dictionary<string, string>());
            }

            var rows  = traducciones.Rows();
            int maximunRows = rows.Count();

            for (int i = 2; i < maximunRows; i++)
            {
               var values = traducciones.Row(i).Cells().Select(s => s.Value.ToString()).ToArray();

                for (int k = 0; k < numberOfLanguages; k++)
                {
                    if (!result[k].ContainsKey(values[0]))
                    {
                        result[k].Add(values[0], values[k+1]);
                    }
                    
                }

            }

            return result;


        }

        private static void ExportFile(Options options)
        {

            DataTable data = CreateDataTable(options.POInputFile, options.Languages.Split(','));
            XLWorkbook workbook = new XLWorkbook();
            workbook.Worksheets.Add(data);

            if (File.Exists(options.FileName))
            {
                File.Delete(options.FileName);
            }

            workbook.SaveAs(options.FileName);
        }

        private static DataTable CreateDataTable(string pOInputFile, string[] languages)
        {
            DataTable result = new DataTable("Traducciones");
            result.Columns.Add("Identificador");
            foreach (string language in languages)
            {
                result.Columns.Add(language);
            }

            foreach (string identifier in GetIdentifiers(pOInputFile))
            {
                result.Rows.Add(new object[] { identifier });
            }

            return result;
        }

        private static IEnumerable<string> GetIdentifiers(string pOInputFile)
        {
            Regex regExpress = new Regex("msgid \"(.*)\"");
            string text = File.ReadAllText(pOInputFile);

            List<string> results = new List<string>();

            foreach (Match match in regExpress.Matches(text))
            {
                var identifier = match.Groups[1].Value.ToString().Trim();
                if (!string.IsNullOrEmpty(identifier))
                {
                    results.Add(identifier);
                }
            }

            return results;
        }
    }
}
