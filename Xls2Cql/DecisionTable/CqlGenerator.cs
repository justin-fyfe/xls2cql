/*
 * Licensed under the Apache License, Version 2.0 (the "License"); you 
 * may not use this file except in compliance with the License. You may 
 * obtain a copy of the License at 
 * 
 * http://www.apache.org/licenses/LICENSE-2.0 
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the 
 * License for the specific language governing permissions and limitations under 
 * the License.
 * 
 * User: fyfej
 * Date: 2022-3-4
 */
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Xls2Cql.DecisionTable
{
    /// <summary>
    /// CQL Generator for the decision logic tables
    /// </summary>
    public class CqlGenerator : IGenerator
    {

        /// <summary>
        /// Decision rule group
        /// </summary>
        private class DecisionRuleGroup
        {

            public DecisionRuleGroup()
            {
                this.Reference = new HashSet<string>();
                this.Output = new HashSet<string>();
                this.Annotation = new HashSet<string>();
            }

            /// <summary>
            /// Gets or sets the action
            /// </summary>
            public string Action { get; set; }

            /// <summary>
            /// Gets or sets the annotation
            /// </summary>
            public HashSet<String> Annotation { get; set; }

            /// <summary>
            /// Gets or sets the referene
            /// </summary>
            public HashSet<String> Reference { get; set; }

            /// <summary>
            /// Gets or sets the output
            /// </summary>
            public HashSet<String> Output { get; set; }

            /// <summary>
            /// Gets or sets the inputs
            /// </summary>
            public CqlExpression Expression { get; set; }

        }

        // Common sheets which don't need to be processed 
        private readonly string[] ignoreSheets =
        {
            "readme",
            "cover",
            "references"
        };

        /// <summary>
        /// Gets the name of the generator
        /// </summary>
        public string Name => "who.dak.l2.dt.cql";

        /// <summary>
        /// Gets the description of the generator
        /// </summary>
        public string Description => "Decision Tables to CQL";

        /// <inheritdoc/>
        public void Generate(IXLWorkbook workbook, string rootPath, string skelFile, IDictionary<String, Object> arguments)
        {

            // Regex for decision ID column
            var decisionIdRegex = new Regex(@"(\w*?\.\w*?\.\d*)(.*)");
            // Regex for existing DEFINE statements
            var defineRegex = new Regex(@"(define\s?\""(.*?)\""[\S\s]*?)\/\*", RegexOptions.Multiline | RegexOptions.IgnoreCase); // Regex to extract DEFINE statements from existing CQL file
            // Regex for existing parameters
            var parameterRegex = new Regex(@"^parameter.*?$", RegexOptions.Multiline | RegexOptions.IgnoreCase); // Regex to extract parameter definitions from existing CQL file

            var skelContents = String.Empty;
            if (!String.IsNullOrEmpty(skelFile) && File.Exists(skelFile))
            {
                skelContents = File.ReadAllText(skelFile);
            }
            else
            {
                skelContents = File.ReadAllText("skel.cql");
            }

            rootPath = Path.Combine(rootPath, "input", "cql");
            if (!Directory.Exists(rootPath))
            {
                Directory.CreateDirectory(rootPath);
            }

            // Worksheets processing
            foreach (var worksheet in workbook.Worksheets)
            {
                if (ignoreSheets.Any(i => i.Equals(worksheet.Name, StringComparison.OrdinalIgnoreCase)) || worksheet.Visibility == XLWorksheetVisibility.Hidden)
                {
                    continue;
                }

                // Find the root cell 
                foreach (var labelCell in worksheet.CellsUsed(o =>
                 {
                     try
                     {
                         return !o.IsEmpty() && !o.IsMerged() && o.Value.ToString().Trim().Equals("Decision ID", StringComparison.OrdinalIgnoreCase);
                     }
                     catch
                     {
                         return false;
                     }
                 }))
                {
                    var decisionIdCellValue = labelCell.CellRight().GetValue<String>().Trim();
                    var decisionIdMatch = decisionIdRegex.Match(decisionIdCellValue);
                    if (!decisionIdMatch.Success)
                    {
                        throw new InvalidOperationException($"Cannot parse {decisionIdCellValue} should be in format AAA.BBB.##.XXXXXX for example IMMZ.DT.01.Some-Description-Of-The-Decision");
                    }

                    var code = decisionIdMatch.Groups[1].Value.Trim();
                    var libraryName = code.Replace(".", "");

                    var mnemonic = decisionIdMatch.Groups[2].Value.Trim();

                    var description = labelCell.CellBelow().CellRight().GetValue<String>().Trim();
                    var triggerDescription = labelCell.CellBelow(2).CellRight().GetValue<String>().Trim();

                    var inputCell = labelCell.CellBelow(3); // This should be "INPUT"
                    if (!inputCell.GetValue<String>().Trim().Equals("inputs", StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException($"Expected the value 'input' in cell {inputCell.FormulaR1C1}");
                    }

                    int inputStartCol = inputCell.WorksheetColumn().ColumnNumber(),
                        outputStartCol = inputStartCol;

                    // Process until we hit OUTPUT
                    while ( !inputCell.GetValue<String>().Trim().Equals("Output", StringComparison.OrdinalIgnoreCase))
                    {
                        inputCell = inputCell.CellRight();
                        outputStartCol++;
                    }

                    int actionStartCol = outputStartCol + 1, // action can be 1..n so we need to count
                        annotationStartCol = outputStartCol;
                    // Process until we HIT ANNOTATION column
                    while( !inputCell.GetValue<String>().Trim().Equals("Annotations", StringComparison.OrdinalIgnoreCase))
                    {
                        inputCell = inputCell.CellRight();
                        annotationStartCol++;
                    }

                    int referenceCol = annotationStartCol;
                    // Process until we HIT REFERENCE(S) column
                    while ( !inputCell.GetValue<String>().Trim().Equals("Reference(s)", StringComparison.OrdinalIgnoreCase))
                    {
                        inputCell = inputCell.CellRight();
                        referenceCol++;
                    }


                    // Now we process
                    var fileName = Path.ChangeExtension(Path.Combine(rootPath, libraryName), "cql");
                    Console.WriteLine("Generating {0}...", fileName);

                    var existingStatements = new Dictionary<String, String>();
                    var parameters = new List<String>();

                    if (File.Exists(fileName))
                    {
                        if (!arguments.TryGetValue("replace", out _))
                        {
                            Console.WriteLine("File {0} already exists - skipping (use --replace)", fileName);
                            continue;
                        }


                        // Select contents and extract existing definitions and parameters
                        var contents = File.ReadAllText(fileName);
                        foreach (Match m in defineRegex.Matches(contents + "/*"))
                        {
                            existingStatements.Add(m.Groups[2].Value, m.Groups[1].Value.Trim());
                        }
                        parameters.AddRange(parameterRegex.Matches(contents).Select(o => o.Value.Trim()));
                    }

                    using (var tw = File.CreateText(fileName))
                    {
                        // Emit friendly header with the documentation for the file
                        tw.WriteLine("/*");
                        tw.WriteLine(" * Library: {0} ({1}{2})", libraryName, code, mnemonic);
                        tw.WriteLine(" * Rule: {0} \r\n * Trigger: {1}", description, triggerDescription);
                        tw.WriteLine(" */");

                        tw.WriteLine("library {0}", libraryName);

                        // Standard headers
                        tw.WriteLine(skelContents);

                        if (parameters.Any())
                        {
                            foreach (var p in parameters)
                            {
                                tw.WriteLine(p);
                            }
                        }

                        tw.WriteLine("context Patient\r\n");

                        // Parse the rules into a rule structure
                        List<DecisionRuleGroup> rules = new List<DecisionRuleGroup>();

                        // Collect the actual logic and group it into rule groups 
                        var row = inputCell.WorksheetRow().RowBelow();
                        while(!row.Cell(inputStartCol).IsEmpty() || (row.Cell(inputStartCol).IsMerged() && !row.Cell(inputStartCol).MergedRange().FirstCell().IsEmpty()))
                        {

                            var action = String.Empty;
                            // Process actions
                            for (var c = actionStartCol; c < annotationStartCol; c++)
                            {
                                var value = row.Cell(c).GetValue<String>().Trim();

                                if (row.Cell(c).IsMerged())
                                {
                                    value =row.Cell(c).MergedRange().FirstCell().GetValue<String>();
                                }
                                if (!String.IsNullOrEmpty(value))
                                {
                                    action += $"{value} then ";
                                }
                            }

                            if (action.Length > 0)
                            {
                                action = action.Substring(0, action.Length - 6);
                            }

                            string annotation = row.Cell(annotationStartCol).GetValue<String>().Trim(),
                                output = row.Cell(outputStartCol).GetValue<String>().Trim(),
                                reference = row.Cell(referenceCol).GetValue<String>().Trim();

                            // First, try to find the existing rule
                            var rule = rules.Find(o => o.Action == action);
                            if (rule == null)
                            {
                                rule = new DecisionRuleGroup()
                                {
                                    Action = action
                                };
                                rules.Add(rule);
                            }

                            if(!rule.Annotation.Contains(annotation))
                            {
                                rule.Annotation.Add(annotation);
                            }

                            // Process the INPUT columns
                            CqlExpression rowExpression = null;
                            for (int c = inputStartCol; c < outputStartCol; c++)
                            {
                                try
                                {
                                    String value = String.Empty;
                                    if (row.Cell(c).IsMerged())
                                    {
                                        value = row.Cell(c).MergedRange().FirstCell().GetValue<String>();
                                    }
                                    else
                                    {
                                        value = row.Cell(c).GetValue<String>();
                                    }

                                    if(String.IsNullOrEmpty(value))
                                    {
                                        continue;
                                    }

                                    var expr = CqlExpression.Parse(value);
                                    if (rowExpression == null)
                                    {
                                        rowExpression = expr;
                                    }
                                    else
                                    {
                                        rowExpression = new CqlBinaryExpression(CqlBinaryOperator.And, rowExpression, expr);
                                    }
                                }
                                catch(Exception e)
                                {
                                    Console.WriteLine("WARNING: Cell {0} - {1}", row.Cell(c).ToString(), e.Message);
                                }
                            }

                            if (rule.Expression == null)
                            {
                                rule.Expression = rowExpression;
                            }
                            else
                            {
                                rule.Expression = new CqlBinaryExpression(CqlBinaryOperator.Or, rule.Expression, rowExpression);
                            }

                            // Process output column
                            if (!rule.Output.Contains(output)) {
                                rule.Output.Add(output);
                            }
                            if (!rule.Reference.Contains(reference))
                            {
                                rule.Reference.Add(reference);
                            }

                        row = row.RowBelow(); // iterate row

                        }

                        HashSet<String> existingElements = new HashSet<string>();
                        // Now generate the CQL
                        foreach (var r in rules)
                        {

                            // emit DEFINE statements for each term that is applicable
                            if (!arguments.TryGetValue("rulesonly", out _))
                            {
                                foreach (var itm in this.GetDataElements(r.Expression))
                                {
                                    var defineTerm = itm.Replace("\r", "").Replace("\n", "").Replace("\"", "");

                                    if (!existingElements.Contains(defineTerm))
                                    {
                                        tw.WriteLine("/* \r\n * @dataElement {0}\r\n */", defineTerm);
                                        if (existingStatements.TryGetValue(defineTerm.Trim(), out var existingStmt) && !arguments.TryGetValue("refresh", out _))
                                        {
                                            tw.WriteLine("{0}\r\n", existingStmt);
                                        }
                                        else
                                        {
                                            tw.WriteLine("define \"{0}\":\r\n\t0 // TODO: Define this\r\n", defineTerm);
                                        }
                                        existingElements.Add(defineTerm);
                                    }
                                }
                            }

                            tw.WriteLine("/*");
                            tw.WriteLine(" * Rule: {0}", r.Action);
                            tw.WriteLine(" * Annotations:");
                            foreach (var itm in r.Annotation)
                            {
                                tw.WriteLine(" * \t - {0}", itm);
                            }
                            tw.WriteLine(" * Outputs:");
                            foreach(var itm in r.Output)
                            {
                                tw.WriteLine(" * \t - {0}", itm);
                            }
                            tw.WriteLine(" * References:");
                            foreach(var itm in r.Reference)
                            {
                                tw.WriteLine(" * \t- {0}", itm);
                            }
                            tw.WriteLine(" * Logic:\r\n *\t {0}", r.Expression);
                            tw.WriteLine(" */");

                            var name = r.Action;
                          
                            if (existingStatements.TryGetValue(name.Trim(), out string existing) && !arguments.TryGetValue("refresh", out _)) {
                                tw.WriteLine("{0}", existing);
                            }
                            else 
                            {
                                tw.WriteLine("define \"{0}\":\r\n\t{1}\r\n", name, r.Expression);
                            }
                        }

                    }
                }
            }
        }

        /// <summary>
        /// Get the data elements for the expression
        /// </summary>
        private IEnumerable<String> GetDataElements(CqlExpression expression)
        {
            switch(expression)
            {
                case CqlIdentifier ci:
                    if (ci.Identifier.StartsWith("\""))
                    {
                        yield return ci.Identifier;
                    }
                    break;
                case CqlBinaryExpression cb:
                    foreach (var itm in this.GetDataElements(cb.Left).Union(this.GetDataElements(cb.Right)))
                        yield return itm;
                    break;
            }
        }
    }
}
