using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;

namespace Xls2Cql
{
    class Program
    {

        private const int CodeColumn = 2;
        private const int NameColumn = 3;
        private const int DiscussionColumn = 4;
        private const int NumeratorDefinitionColumn = 5;
        private const int NumeratorComputationColumn = 6;
        private const int DenominatorDefinitionColumn = 7;
        private const int DenominatorComputationColumn = 8;
        private const int DisaggregationColumn = 9;
        private const int ReferenceColumn = 10;

        /// <summary>
        /// Process spreadsheet into CQL files
        /// </summary>
        static void Main(string[] args)
        {
            
            if(args.Length != 2)
            {
                Console.WriteLine("Use: dotnet xls2cql [XLSX] [OUTPUT]");
                Environment.Exit(1);
                return;
            }
            else if(!File.Exists(args[0]))
            {
                Console.WriteLine("File {0} does not exist", args[0]);
                Environment.Exit(2);
                return;
            }

            // Process file
            try
            {
                var output = args[1];
                if (!Path.IsPathRooted(output))
                {
                    output = Path.Combine(Path.GetDirectoryName(typeof(Program).GetType().Assembly.Location), output);
                }
                if (!Directory.Exists(output))
                {
                    Directory.CreateDirectory(output);
                }

                using(var excelStream = File.OpenRead(args[0]))
                {
                    using(var wkb = new XLWorkbook(excelStream))
                    {
                        var wksht = wkb.Worksheets.FirstOrDefault(o => o.Name.Trim().Equals("Indicator Table", StringComparison.OrdinalIgnoreCase));
                        if(wksht == null)
                        {
                            Console.WriteLine("No worksheet named 'Indicator Table' found");
                        }

                        bool isReading = false;
                        foreach(var row in wksht.Rows())
                        {
                            if(row.Cell(CodeColumn).GetValue<String>()?.Trim().Equals("Indicator Code", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                isReading = true;
                            }
                            else if(isReading && !row.Cell(CodeColumn).IsEmpty())
                            {
                                GenerateCQLFile(row, output);
                                GenerateMeasureFile(row, output);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Fatal Error: {0}", e);
                Environment.Exit(911);
            }
        }

        /// <summary>
        /// Generate measure definition file
        /// </summary>
        private static void GenerateMeasureFile(IXLRow row, String outputDirectory)
        {

        }

        /// <summary>
        /// Generate CQL file
        /// </summary>
        private static void GenerateCQLFile(IXLRow row,String outputDirectory)
        {
            var indicatorName = row.Cell(CodeColumn).GetValue<String>().Replace(".", "");
            var fileName = Path.ChangeExtension(Path.Combine(outputDirectory, indicatorName), ".cql");
            Console.WriteLine("Creating {0}", fileName);
            using (var tw = File.CreateText(fileName))
            {
                tw.WriteLine("/*");
                tw.WriteLine(" * Library: {0}", row.Cell(CodeColumn).GetValue<String>());
                tw.WriteLine(" * {0} \r\n * {1}", row.Cell(NameColumn).GetValue<String>(), row.Cell(DiscussionColumn).GetValue<String>());
                tw.WriteLine(" * ");
                tw.WriteLine(" * Numerator: {0} \r\n * Numerator Computation: {1}\r\n * Denominator: {2}\r\n * Denominator Computation: {3}",
                    row.Cell(NumeratorDefinitionColumn).GetValue<String>(),
                    row.Cell(NumeratorComputationColumn).GetValue<String>(),
                    row.Cell(DenominatorDefinitionColumn).GetValue<String>(),
                    row.Cell(DenominatorComputationColumn).GetValue<String>());
                tw.WriteLine(" * ");
                tw.WriteLine(" * Disaggregation:");
                foreach (var d in row.Cell(DisaggregationColumn).GetValue<String>().Split('\r', '\n'))
                {
                    tw.WriteLine(" * {0}", d);
                }
                tw.WriteLine(" * See: {0}", row.Cell(ReferenceColumn).GetValue<String>());
                tw.WriteLine(" */\r\n");
                tw.WriteLine("library {0}\r\n", indicatorName);
                tw.WriteLine("using FHIR version '4.0.1'\r\n");
                tw.WriteLine("include FHIRHelpers version '4.0.1'");
                tw.WriteLine("include IMMZCommon called IMMZCom");
                tw.WriteLine("include IMMZConcepts called IMMZc");
                tw.WriteLine("include IMMZStratifiers called IMMZStratifiers");
                tw.WriteLine("include IMMZVaccineLibrary called IMMZvl\r\n");
                tw.WriteLine("parameter \"Measurement Period\" Interval<Date>\r\n");
                tw.WriteLine("context Patient\r\n");

                tw.WriteLine("/*\r\n * Numerator: {0}\r\n* Numerator Computation: {1}\r\n */", row.Cell(NumeratorDefinitionColumn).GetValue<String>(), row.Cell(NumeratorComputationColumn).GetValue<String>());
                tw.WriteLine("define \"numerator\":\r\n\ttrue // TODO: Write logic here \r\n");
                tw.WriteLine("/*\r\n * Denominator: {0}\r\n* Denominator Computation: {1}\r\n */", row.Cell(DenominatorDefinitionColumn).GetValue<String>(), row.Cell(DenominatorComputationColumn).GetValue<String>());
                tw.WriteLine("define \"denominator\":\r\n\ttrue // TODO: Write logic here \r\n");

                foreach (var d in row.Cell(DisaggregationColumn).GetValue<String>().Split('\r', '\n'))
                {
                    tw.WriteLine("/*\r\n * Disaggregator: {0}\r\n */", d);

                    var dn = d;
                    if(dn.Contains("("))
                    {
                        dn = dn.Substring(0, dn.IndexOf("("));
                    }

                    tw.WriteLine("define \"{0} Stratifier\":\r\n\ttrue // todo: fill in logic\r\n", dn);
                }
            }
        }
    }
}
