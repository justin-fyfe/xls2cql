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
 * User: Nityan Khanna
 * Date: 2022-3-25
 */

using ClosedXML.Excel;
using Hl7.Fhir.Model;
using Hl7.Fhir.Serialization;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Xls2Cql.DecisionTable
{
    /// <summary>
    /// Represents a plan definition generator.
    /// </summary>
    public class PlanDefinitionGenerator : IGenerator
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PlanDefinitionGenerator"/> class.
        /// </summary>
        public PlanDefinitionGenerator()
        {

        }

        /// <summary>
        /// Gets the name of the generator.
        /// </summary>
        public string Name => "who.dak.l2.dt.pd";

        /// <summary>
        /// Gets the description of the generator.
        /// </summary>
        public string Description => "Decision Tables to Plan Definition Resources";

        /// <summary>
        /// Generate the resource or CQL file
        /// </summary>
        /// <param name="workbook">The workbook to be processed</param>
        /// <param name="rootPath">The path to generate the resource files in</param>
        /// <param name="parameters">Other parameters passed on the command line</param>
        /// <param name="skelFile">Skeleton file which should be used</param>
        /// <returns>The generated file</returns>
        public void Generate(IXLWorkbook workbook, string rootPath, string skelFile, IDictionary<string, object> parameters)
        {
            //var replaceRegex = new Regex(@"[\s\-\(\)]");
            //var disaggregatorRegex = new Regex(@"^([^\(]*)\s?\(?.*$");
            //var idRegex = new Regex(@"^([^\d]*?)(\d*)$");

            foreach (var sheet in workbook.Worksheets.Where(c => c.Name.StartsWith("IMMZ.DT.") && c.Name != "IMMZ.DT.00.Common"))
            {
                var planDefinition = new PlanDefinition();

                IXLCell outputHeaderCell = null;
                IXLCell actionHeaderCell = null;
                IXLCell annotationHeaderCell = null;
                IXLCell actionCell = null;

                foreach (var row in sheet.Rows())
                {
                    
                    // we are done processing the final input on the final row
                    if (row.RowNumber() > PlanDefinitionConstants.InputsRowStart && row.Cell(PlanDefinitionConstants.InputsColumnStart)?.Value?.ToString() == string.Empty)
                    {
                        break;
                    }

                    switch (row.RowNumber())
                    {
                        case 4:
                            planDefinition.Id = $"{PlanDefinitionConstants.PlanDefinitionBaseUrl}{row.Cell(3).Value}";
                            planDefinition.Name = row.Cell(3).Value?.ToString();
                            continue;
                        case 5:
                            planDefinition.Description = new Markdown(row.Cell(3).Value?.ToString());
                            continue;
                        case 7:
                            outputHeaderCell = row.Cells(c => c.Value?.ToString() == PlanDefinitionConstants.OutputLabel).Single();
                            actionHeaderCell = row.Cells(c => c.Value?.ToString() == PlanDefinitionConstants.ActionLabel).Single();
                            annotationHeaderCell = row.Cells(c => c.Value?.ToString() == PlanDefinitionConstants.AnnotationsLabel).Single();
                            continue;
                            //case 6:
                            //    planDefinition.tr= row.Cell(2).Value?.ToString();
                            //    continue;
                    }

                    if (row.RowNumber() >= PlanDefinitionConstants.InputsRowStart)
                    {
                        // TODO: build actions
                        // keep getting input cells until we reach the output cell

                        //var activityDefinition = new ActivityDefinition
                        //{
                        //    Title = row.Cell(outputCell.Address.ColumnNumber)?.Value?.ToString(),
                        //    Description = new Markdown(row.Cell(annotationCell.Address.ColumnNumber)?.Value?.ToString())
                        //};

                        var action = new PlanDefinition.ActionComponent
                        {
                            Title = row.Cell(outputHeaderCell.Address.ColumnNumber)?.Value?.ToString(),
                            Description = row.Cell(annotationHeaderCell.Address.ColumnNumber)?.Value?.ToString()
                        };

                        if (actionCell == null)
                        {
                            actionCell = row.Cell(actionHeaderCell.Address.ColumnNumber);
                        }

                        foreach (var cellEntry in row.Cells(c => c.Address.ColumnNumber < outputHeaderCell.Address.ColumnNumber).Where(c => c.Value?.ToString() != string.Empty))
                        {
                            ;
                            if (!actionCell.IsMerged())
                            {
                                actionCell = row.Cell(actionHeaderCell.Address.ColumnNumber);
                            }

                            action.Condition.Add(new PlanDefinition.ConditionComponent
                            {
                                Expression = new Expression
                                {
                                    Description = cellEntry.Value?.ToString(),
                                    Language = "text/cql",
                                    Expression_ = actionCell.Value?.ToString()
                                },
                                Kind = ActionConditionKind.Applicability,
                            });
                        }

                        planDefinition.Action.Add(action);



                        //activityDefinition
                    }
                }

                var serializer = new FhirJsonSerializer(new SerializerSettings()
                {
                    Pretty = true
                });

                var fileName = Path.ChangeExtension(Path.Combine(rootPath, "input", "resources", "plandefinition", $"{sheet.Name}"), ".json");

                if (!Directory.Exists(Path.GetDirectoryName(fileName)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(fileName));
                }

                Console.WriteLine("Generating {0}...", fileName);

                // File exists? Want to make sure we actually want to replace it
                if (File.Exists(fileName) && !parameters.TryGetValue("refresh", out _))
                {
                    Console.WriteLine("File {0} already exists - skipping", fileName);
                    continue;
                }
                else
                {
                    using var tw = File.CreateText(fileName);
                    using var jw = new JsonTextWriter(tw)
                    {
                        Formatting = Formatting.Indented
                    };

                    serializer.Serialize(planDefinition, jw);
                }
            }
        }
    }
}
