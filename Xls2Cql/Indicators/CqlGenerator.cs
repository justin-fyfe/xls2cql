using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Xls2Cql.Indicators
{
    /// <summary>
    /// Generator for indicator CQL
    /// </summary>
    public class CqlGenerator : IGenerator
    {
        /// <inheritdoc/>
        public string Name => "who.dak.l2.ind.cql";

        /// <inheritdoc/>
        public string Description => "WHO DAK L2 Indicator Table to CQL";

        /// <inheritdoc/>
        public void Generate(IXLWorkbook workbook, string rootPath, bool replaceExisting, string skelFile)
        {
            var idRegex = new Regex(@"^([^\d]*?)(\d*)$"); // regex to extract ID from Excel
            var defineRegex = new Regex(@"(define\s?\""(.*?)\""[\S\s]*?)\/\*", RegexOptions.Multiline | RegexOptions.IgnoreCase); // Regex to extract DEFINE statements from existing CQL file
            var parameterRegex = new Regex(@"^parameter.*?$", RegexOptions.Multiline | RegexOptions.IgnoreCase); // Regex to extract parameter definitions from existing CQL file


            var sheet = workbook.Worksheets.FirstOrDefault(o => o.Name.Equals("Indicator table", StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
            {
                throw new InvalidOperationException("Cannot find a worksheet named 'indicator table'");
            }

            // skel file
            var skelContents = String.Empty;
            if (!String.IsNullOrEmpty(skelFile) && File.Exists(skelFile))
            {
                skelContents = File.ReadAllText(skelFile);
            }
            else
            {
                skelContents = File.ReadAllText("skel.cql");
            }

            foreach (var row in sheet?.Rows())
            {
                var codeCell = row.Cell(IndicatorConstants.CodeColumn).GetValue<String>().Trim();
                if (String.IsNullOrEmpty(codeCell) || codeCell.Equals("Indicator Code", StringComparison.OrdinalIgnoreCase) || 
                    row.Cell(IndicatorConstants.CodeColumn).IsMerged())
                {
                    continue;
                }

                var code = idRegex.Replace(codeCell, o => $"{o.Groups[1].Value}{Int32.Parse(o.Groups[2].Value).ToString("00")}"); // Code for the indicator
                var indicatorName = codeCell.Replace(".", "").Trim(); // Gets the name of the indiactor for the current row
               
                indicatorName = idRegex.Replace(indicatorName, o => $"{o.Groups[1].Value}{Int32.Parse(o.Groups[2].Value).ToString("00")}"); // Format and pad the ID
                var fileName = Path.ChangeExtension(Path.Combine(rootPath, "input", "cql", indicatorName), ".cql");
                Console.WriteLine("Creating {0}", fileName);

                // existing statements (DEFINE)
                var existingStatements = new Dictionary<String, String>();
                // existing parameters 
                var parameters = new List<String>();

                
                if(!Directory.Exists(Path.GetDirectoryName(fileName)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(fileName));
                }

                // File exists ? Load the contents so we don't nuke any of the current logic
                if (File.Exists(fileName))
                {
                    if (!replaceExisting)
                    {
                        Console.WriteLine("File {0} already exists - skipping", fileName);
                        continue;
                    }

                    // Select contents and extract existing definitions and parameters
                    var contents = File.ReadAllText(fileName);
                    foreach (Match m in defineRegex.Matches(contents + "/*"))
                    {
                        existingStatements.Add(m.Groups[2].Value, m.Groups[1].Value.Trim());
                    }
                    parameters = parameterRegex.Matches(contents).Select(o => o.Value.Trim()).ToList();
                }

                using (var tw = File.CreateText(fileName))
                {
                    // Emit friendly header with the documentation for the file
                    tw.WriteLine("/*");
                    tw.WriteLine(" * Library: {0}", code);
                    tw.WriteLine(" * {0} \r\n * {1}", row.Cell(IndicatorConstants.NameColumn).GetValue<String>(), row.Cell(IndicatorConstants.DiscussionColumn).GetValue<String>());
                    tw.WriteLine(" * ");
                    tw.WriteLine(" * Numerator: {0} \r\n * Numerator Computation: {1}\r\n * Denominator: {2}\r\n * Denominator Computation: {3}",
                        row.Cell(IndicatorConstants.NumeratorDefinitionColumn).GetValue<String>(),
                        row.Cell(IndicatorConstants.NumeratorComputationColumn).GetValue<String>(),
                        row.Cell(IndicatorConstants.DenominatorDefinitionColumn).GetValue<String>(),
                        row.Cell(IndicatorConstants.DenominatorComputationColumn).GetValue<String>());
                    tw.WriteLine(" * ");
                    tw.WriteLine(" * Disaggregation:");
                    foreach (var d in row.Cell(IndicatorConstants.DisaggregationColumn).GetValue<String>().Split('\r', '\n').Where(o => !String.IsNullOrEmpty(o)))
                    {
                        tw.WriteLine(" *   - {0}", d);
                    }
                    tw.WriteLine(" * ");
                    tw.WriteLine(" * References: {0}", String.Join(", ", row.Cell(IndicatorConstants.ReferenceColumn).GetValue<String>().Split('\n')));
                    tw.WriteLine(" */\r\n");

                    // Define the library
                    tw.WriteLine("library {0}\r\n", indicatorName);

                    // Standard headers
                    tw.WriteLine(skelContents);

                    if (parameters.Any())
                    {
                        foreach (var p in parameters)
                        {
                            tw.WriteLine(p);
                        }
                    }
                    else
                    {
                        tw.WriteLine("parameter \"Measurement Period\" Interval<Date>\r\n");
                    }
                    tw.WriteLine("context Patient\r\n");

                    tw.WriteLine("/*\r\n * Numerator: {0}\r\n * Numerator Computation: {1}\r\n */", row.Cell(IndicatorConstants.NumeratorDefinitionColumn).GetValue<String>(), row.Cell(IndicatorConstants.NumeratorComputationColumn).GetValue<String>());

                    if (existingStatements.TryGetValue("numerator", out var numerator))
                    {
                        tw.WriteLine(numerator);
                    }
                    else
                    {
                        tw.WriteLine("define \"numerator\":\r\n\ttrue // TODO: Write logic here \r\n");
                    }
                    tw.WriteLine("/*\r\n * Denominator: {0}\r\n * Denominator Computation: {1}\r\n */", row.Cell(IndicatorConstants.DenominatorDefinitionColumn).GetValue<String>(), row.Cell(IndicatorConstants.DenominatorComputationColumn).GetValue<String>());

                    if (existingStatements.TryGetValue("denominator", out var denom))
                    {
                        tw.WriteLine(denom);
                    }
                    else
                    {
                        tw.WriteLine("define \"denominator\":\r\n\ttrue // TODO: Write logic here \r\n");
                    }

                    foreach (var d in row.Cell(IndicatorConstants.DisaggregationColumn).GetValue<String>().Split('\r', '\n'))
                    {
                        tw.WriteLine("/*\r\n * Disaggregator: {0}\r\n */", d);

                        var dn = d;
                        if (dn.Contains("("))
                        {
                            dn = dn.Substring(0, dn.IndexOf("("));
                        }

                        if (existingStatements.TryGetValue($"{dn} Stratifier", out var strat))
                        {
                            tw.WriteLine(strat);
                        }
                        else
                        {
                            tw.WriteLine("define \"{0} Stratifier\":\r\n\ttrue // todo: fill in logic\r\n", dn);
                        }
                    }

                    tw.WriteLine("/* End of {0} */", code);
                }
            }
        }
    }
}
