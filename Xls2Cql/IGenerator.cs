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

namespace Xls2Cql
{
    /// <summary>
    /// Base level FHIR generator interface
    /// </summary>
    public interface IGenerator
    {

        /// <summary>
        /// Gets the name of the XLS <> CQL Generator
        /// </summary>
        String Name { get; }

        /// <summary>
        /// Gets the description of the generator
        /// </summary>
        String Description { get; }

        /// <summary>
        /// Generate the resource or CQL file
        /// </summary>
        /// <param name="workbook">The workbook to be processed</param>
        /// <param name="rootPath">The path to generate the resource files in</param>
        /// <param name="parameters">Other parameters passed on the command line</param>
        /// <param name="skelFile">Skeleton file which should be used</param>
        /// <returns>The generated file</returns>
        void Generate(IXLWorkbook workbook, string rootPath, string skelFile, IDictionary<String, Object> parameters);
    }
}