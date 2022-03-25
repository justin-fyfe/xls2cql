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
            foreach (var sheet in workbook.Worksheets.Where(c => c.Name.StartsWith("IMMZ.DT.") && c.Name != "IMMZ.DT.00.Common"))
            {
                var resources = new List<Resource>();
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
                    }


                    // move to the next iteration if we have not yet reached the input row
                    if (row.RowNumber() < PlanDefinitionConstants.InputsRowStart)
                    {
                        continue;
                    }

                    var activityDefinitionId = row.Cell(outputHeaderCell.Address.ColumnNumber)?.Value?.ToString()?.Replace(" ", "-").Replace("---", "-");

                    // remove the trailing dash from the id if neecessary
                    activityDefinitionId = activityDefinitionId.EndsWith("-") ? activityDefinitionId.Substring(0, activityDefinitionId.Length - 1) : activityDefinitionId;

                    var action = new PlanDefinition.ActionComponent
                    {
                        Title = row.Cell(outputHeaderCell.Address.ColumnNumber)?.Value?.ToString(),
                        Description = row.Cell(annotationHeaderCell.Address.ColumnNumber)?.Value?.ToString(),
                        Definition = new Canonical($"{PlanDefinitionConstants.ActivityDefinitionCanonicalBaseUrl}{activityDefinitionId}")
                    };

                    // we need to maintain the original reference to the Action
                    // in the case the action cell is a merged cell
                    // meaning that multiple logical groupings of inputs in the sheet
                    // all result in the same action
                    actionCell ??= row.Cell(actionHeaderCell.Address.ColumnNumber);

                    foreach (var cellEntry in row.Cells(c => c.Address.ColumnNumber < outputHeaderCell.Address.ColumnNumber).Where(c => c.Value?.ToString() != string.Empty))
                    {
                        // if the action cell is not merged, then get a reference to the correct action cell
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
                            Kind = ActionConditionKind.Applicability
                        });
                    }

                    planDefinition.Action.Add(action);

                    // add the plan definition and activity definition to the list of resources to be written to the file system
                    resources.Add(new ActivityDefinition
                    {
                        Code = new CodeableConcept(PlanDefinitionConstants.SnomedCtUrl, PlanDefinitionConstants.SnomedCtVaccinationCode, PlanDefinitionConstants.SnomedCtDescription, null),
                        DoNotPerform = false,
                        Id = activityDefinitionId,
                        Intent = RequestIntent.Proposal,
                        Meta = new Meta
                        {
                            Profile = new List<string>
                            {
                                PlanDefinitionConstants.CpgImmunizationActivityDefinitionProfileUrl
                            }
                        },
                        Status = PublicationStatus.Draft,
                        Publisher = PlanDefinitionConstants.ActivityDefinitionPublisher,
                        Description = new Markdown(action.Description)
                    });
                    resources.Add(planDefinition);
                }

                foreach (var resource in resources)
                {
                    var name = resource switch
                    {
                        PlanDefinition _ => sheet.Name,
                        ActivityDefinition _ => $"{resource.Id.Replace("/", "-").Replace("--", "-").Replace("\\", "-")}",
                        _ => throw new InvalidOperationException($"Unknown resource type: {resource?.GetType().Name}")
                    };

                    this.WriteToFile(rootPath, name, parameters, resource);
                }
            }
        }

        /// <summary>
        /// Writes a FHIR resource to a file.
        /// </summary>
        /// <param name="rootPath">The root path.</param>
        /// <param name="sheetName">The sheet name.</param>
        /// <param name="parameters">The parameters.</param>
        /// <param name="resource">The resource.</param>
        /// <exception cref="InvalidOperationException"></exception>
        private void WriteToFile(string rootPath, string sheetName, IDictionary<string, object> parameters, Resource resource)
        {
            var serializer = new FhirJsonSerializer(new SerializerSettings
            {
                Pretty = true
            });

            var path = resource switch
            {
                PlanDefinition _ => "plandefinition",
                ActivityDefinition _ => "activitydefinition",
                _ => throw new InvalidOperationException($"Unknown resource type: {resource?.GetType().Name}")
            };

            var fileName = Path.ChangeExtension(Path.Combine(rootPath, "input", "resources", path, $"{sheetName}"), ".json");

            if (!Directory.Exists(Path.GetDirectoryName(fileName)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(fileName));
            }

            Console.WriteLine("Generating {0}...", fileName);

            // File exists? Want to make sure we actually want to replace it
            if (File.Exists(fileName) && !parameters.TryGetValue("replace", out _))
            {
                Console.WriteLine("File {0} already exists - skipping", fileName);
            }
            else
            {
                using var streamWriter = File.CreateText(fileName);
                using var jsonTextWriter = new JsonTextWriter(streamWriter)
                {
                    Formatting = Formatting.Indented
                };

                serializer.Serialize(resource, jsonTextWriter);
            }
        }
    }
}
