﻿using ClosedXML.Excel;
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