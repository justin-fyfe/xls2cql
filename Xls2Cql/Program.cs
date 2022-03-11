using ClosedXML.Excel;
using Hl7.Fhir.Model;
using Hl7.Fhir.Serialization;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Xls2Cql
{
    class Program
    {

        private const string PARM_HELP = "help";
        private const string PARM_GENERATE = "generate";
        private const string PARM_INPUT = "input";
        private const string PARM_OUTPUT = "output";
        private const string PARM_SKEL = "skel";
        private const string PARM_REPLACE = "replace";
        private const string PARM_REFRESH = "refresh";
        

        private static readonly Regex parmExtract = new Regex(@"^--(\w*?)(?:\=(.*?))?$");

        /// <summary>
        /// Parse parameters into key-value pairs
        /// </summary>
        static IEnumerable<KeyValuePair<String, String>> ParseParameters(string[] args)
        {

            foreach (var itm in args)
            {
                var matches = parmExtract.Match(itm);
                if (matches.Success)
                {
                    if (matches.Groups.Count == 1)
                    {
                        yield return new KeyValuePair<String, String>(matches.Groups[1].Value, "true"); // handles --help or --flag
                    }
                    else
                    {
                        yield return new KeyValuePair<String, String>(matches.Groups[1].Value, matches.Groups[2].Value); // handles --file=foo.bar 
                    }
                }
                else
                {
                    throw new InvalidOperationException($"Can't parse {itm} - use --parameter=value");
                }
            }
        }

        /// <summary>
        /// Process spreadsheet into CQL files
        /// </summary>
        static void Main(string[] args)
        {

            // Process file
            try
            {
                var settings = ParseParameters(args).GroupBy(o => o.Key).ToDictionary(o => o.Key, o => o.Select(v => v.Value).ToList());
                var generators = typeof(Program).Assembly
                    .ExportedTypes
                    .Where(t => !t.IsAbstract && !t.IsInterface && typeof(IGenerator).IsAssignableFrom(t))
                    .Select(t => Activator.CreateInstance(t))
                    .OfType<IGenerator>()
                    .ToDictionary(o => o.Name, o => o);

                if (settings.TryGetValue(PARM_HELP, out _))
                {
                    ShowHelp(generators);
                }
                else if (settings.TryGetValue(PARM_GENERATE, out var generate))
                {
                    string inputFile = String.Empty, outputDirectory = String.Empty;

                    if(settings.TryGetValue(PARM_INPUT, out var inputList))
                    {
                        inputFile = inputList.First();
                    } 
                    else
                    {
                        throw new InvalidOperationException("Must pass --input parameter");
                    }

                    if(settings.TryGetValue(PARM_OUTPUT, out var outputList))
                    {
                        outputDirectory = outputList.First(); 
                    }
                    else
                    {
                        outputDirectory = Path.GetDirectoryName(typeof(Program).Assembly.Location);
                    }

                    using (var excelStream = File.OpenRead(inputFile))
                    {
                        using (var wkb = new XLWorkbook(excelStream))
                        {
                            foreach (string itm in generate)
                            {
                                if (generators.TryGetValue(itm, out var generator))
                                {
                                    generator.Generate(wkb, outputDirectory, settings.TryGetValue(PARM_SKEL, out var skel) ? skel.First() : null, settings.ToDictionary(o=>o.Key, o=>(object)o.Value));
                                }
                                else
                                {
                                    throw new InvalidOperationException($"Don't have a generator for {itm}");
                                }
                            }
                        }
                    }
    
                }
                else {
                    ShowHelp(generators);
                }
               
            }
            catch (Exception e)
            {
                Console.WriteLine("Fatal Error: {0}", e);
                Environment.Exit(911);
            }
        }

        /// <summary>
        /// Show help contents
        /// </summary>
        private static void ShowHelp(IDictionary<String, IGenerator> generators)
        {
            Console.WriteLine("Use: xls2cql [options] where options are one of:");
            Console.WriteLine($"--{PARM_GENERATE}=generatorName\tGenerate Output Type");
            Console.WriteLine($"--{PARM_HELP}\t\t\t\tShow this help and exit");
            Console.WriteLine($"--{PARM_INPUT}=input.xlsx\t\tInput Excel spread sheet");
            Console.WriteLine($"--{PARM_OUTPUT}=directory\t\tThe output directory (the tool will create input\\cql\\XXXX.cql)");
            Console.WriteLine($"--{PARM_REPLACE}\t\t\tReplace/overwrite existing files");
            Console.WriteLine($"--{PARM_REFRESH}\t\t\tRefresh the contents of the define statements");
            Console.WriteLine($"--{PARM_SKEL}=fileName.cql\t\tThe skeleton file to use (for your includes and any header contents)");
            Console.WriteLine("\r\nWhere generatorName is one of:");
            foreach(var itm in generators)
            {
                Console.WriteLine("\t{0} - {1}", itm.Key, itm.Value.Description);
            }
        }
    }
}
