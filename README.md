# WHO L2 DAK Excel File to CQL Scaffolder

## Use

To use the tool: `dotnet Xls2Cql.dll [arguments]` where arguments are:

```
--generate=generatorName      Use the specified generator
--help                          Show this help and exit
--input=input.xlsx            Input Excel spread sheet in XLS or XLSX format
--output=directory				The root directory of your IG (the tool will create input\cql\XXXX.cql and input\resources\XXXXX\*.json)
--replace                       Replace/overwrite existing files (this will keep any modified definitions and just update comments)
--refresh                       Refresh the contents of the define statements (resets the contents of the define to the defaults)
--skel=fileName.cql             The skeleton file to use (for your includes and any header contents)
```

## Skeleton File

The skeleton file is used to place a common header (after `library`) in the CQL output, and is a simple CQL file. For example, for the Immunization DAK:

```
using FHIR version '4.0.1'
include FHIRHelpers version '4.0.1'
include IMMZCommon called IMMZCom
include IMMZConcepts called IMMZc
include IMMZStratifiers called IMMZStratifiers
include IMMZVaccineLibrary called IMMZvl
```

## Generating Decision Table CQL

To generate decision table CQL you should use the `who.dak.l2.dt.cql` generator. For example, to take the contents of dt.xlsx and generate the CQL in a project:

```
dotnet Xls2Cql.dll --generate=who.dak.l2.dt.cql --input="dt.xlsx" --output="C:\Users\myuser\source\repos\my-ig" --skel=skel.cql --replace
```

## Generating Plan Definitions and Activity Definitions

To generate decision table `PlanDefinition` and `ActivityDefinition` resources you should use the `who.dak.l2.dt.pd` generator. For example, to take the contents of dt.xlsx and generate the CQL in a project:

```
dotnet Xls2Cql.dll --generate=who.dak.l2.dt.pd --input="dt.xlsx" --output="C:\Users\myuser\source\repos\my-ig" --replace
```

## Generating Indicator CQL

To generate indicator CQL you should use the `who.dak.l2.ind.cql` generator. For example, to take the contents of ind.xlsx and generate the CQL in a project:

```
dotnet Xls2Cql.dll --generate=who.dak.l2.ind.cql --input="ind.xlsx" --output="C:\Users\myuser\source\repos\my-ig" --skel=skel.cql --replace
```

You can also generate the indicator measure resources off the indicator table by adding the `who.dak.l2.ind.measure` generator, such as:

```
dotnet Xls2Cql.dll --generate=who.dak.l2.ind.cql --generate=who.dak.l2.ind.measure --input="ind.xlsx" --output="C:\Users\myuser\source\repos\my-ig" --skel=skel.cql --replace
```
