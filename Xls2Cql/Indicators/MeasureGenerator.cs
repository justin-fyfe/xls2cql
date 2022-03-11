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
using Hl7.Fhir.Model;
using Hl7.Fhir.Serialization;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Xls2Cql.Indicators
{
    /// <summary>
    /// Generates the measure file
    /// </summary>
    public class MeasureGenerator : IGenerator
    {
        ///<inheritdoc/>
        public string Name => "who.dak.l2.ind.measure";

        /// <inheritdoc/>
        public String Description => "WHO DAK L2 Indicator to Measure JSON Resources";

        ///<inheritdoc/>
        public void Generate(IXLWorkbook workbook, string rootPath, string skelFile, IDictionary<String, Object> arguments )
        {

            var replaceRegex = new Regex(@"[\s\-\(\)]");
            var disaggregatorRegex = new Regex(@"^([^\(]*)\s?\(?.*$");
            var idRegex = new Regex(@"^([^\d]*?)(\d*)$");

            var sheet = workbook.Worksheets.FirstOrDefault(o => o.Name.Equals("Indicator table", StringComparison.OrdinalIgnoreCase));
            if (sheet == null)
            {
                throw new InvalidOperationException("Cannot find a worksheet named 'indicator table'");
            }
            foreach (var row in sheet?.Rows())
            {
                var codeCell = row.Cell(IndicatorConstants.CodeColumn).GetValue<String>().Trim();
                if (String.IsNullOrEmpty(codeCell) || codeCell.Equals("Indicator Code", StringComparison.OrdinalIgnoreCase) ||
                    row.Cell(IndicatorConstants.CodeColumn).IsMerged())
                {
                    continue;
                }

                var indicatorName = codeCell.Replace(".", "");
                var code = idRegex.Replace(codeCell, o => $"{o.Groups[1].Value}{Int32.Parse(o.Groups[2].Value).ToString("00")}");

                indicatorName = idRegex.Replace(indicatorName, o => $"{o.Groups[1].Value}{Int32.Parse(o.Groups[2].Value).ToString("00")}");
                var measure = new Measure()
                {
                    Id = indicatorName,
                    Name = indicatorName,
                    Title = $"{code} {row.Cell(IndicatorConstants.NameColumn).GetValue<String>()}",
                    Url = $"http://fhir.org/guides/who/Immz/Measure/{indicatorName}",
                    Date = DateTime.Now.ToString("o"),
                    Description = new Markdown(row.Cell(IndicatorConstants.DiscussionColumn).GetValue<String>()),
                    Scoring = new CodeableConcept("http://terminology.hl7.org/CodeSystem/measure-scoring", "proportion"),
                    Type = new System.Collections.Generic.List<CodeableConcept>()
                {
                    new CodeableConcept("http://terminology.hl7.org/CodeSystem/measure-type", "process")
                },
                    ImprovementNotation = new CodeableConcept("http://terminology.hl7.org/CodeSystem/measure-improvement-notation", "increase"),
                    Group = new System.Collections.Generic.List<Measure.GroupComponent>()
                {
                    new Measure.GroupComponent()
                    {
                        ElementId = indicatorName,
                        Population = new System.Collections.Generic.List<Measure.PopulationComponent>()
                        {
                            new Measure.PopulationComponent()
                            {
                                Description = row.Cell(IndicatorConstants.NumeratorDefinitionColumn).GetValue<String>(),
                                ElementId = "numerator",
                                Criteria = new Expression()
                                {
                                    Expression_ = "numerator",
                                    Language = "text/cql"
                                },
                                Code = new CodeableConcept("http://terminology.hl7.org/CodeSystem/measure-population", "numerator"),
                            },
                            new Measure.PopulationComponent()
                            {
                                Description = row.Cell(IndicatorConstants.DenominatorDefinitionColumn).GetValue<String>(),
                                Criteria = new Expression()
                                {
                                    Expression_ = "numerator",
                                    Language = "text/cql"
                                },
                                ElementId = "denominator",
                                Code = new CodeableConcept("http://terminology.hl7.org/CodeSystem/measure-population", "denominator"),
                            }
                        },
                        Stratifier = row.Cell(IndicatorConstants.DisaggregationColumn).GetValue<String>().Split('\r', '\n').Select(o=> new Measure.StratifierComponent()
                        {
                            Criteria = new Expression()
                            {
                                Expression_ = $"{disaggregatorRegex.Match(o).Groups[1].Value} Stratifier".Replace("  ", " "),  // HACK, 
                                Language = "text/cql"
                            },
                            ElementId = replaceRegex.Replace(disaggregatorRegex.Match(o).Groups[1].Value.Trim(), x=>"-").ToLowerInvariant() + "-stratifier"
                        }).ToList()
                    }
                }
                };

                var serializer = new FhirJsonSerializer(new SerializerSettings()
                {
                    Pretty = true
                });
                var fileName = Path.ChangeExtension(Path.Combine(rootPath, "input", "resources", "measure", $"measure-{indicatorName}"), ".json");


                if (!Directory.Exists(Path.GetDirectoryName(fileName)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(fileName));
                }

                Console.WriteLine("Generating {0}...", fileName);
                // File exists? Want to make sure we actually want to replace it
                if (File.Exists(fileName) && !arguments.TryGetValue("refresh", out _))
                {
                    Console.WriteLine("File {0} already exists - skipping", fileName);
                    continue;
                }
                else
                {
                    using (var tw = File.CreateText(fileName))
                    {
                        using (var jw = new JsonTextWriter(tw)
                        {
                            Formatting = Formatting.Indented
                        })
                        {
                            serializer.Serialize(measure, jw);
                        }
                    }
                }
            }
        }
    }
}