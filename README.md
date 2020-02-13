# Pega PowerPoint Generator

## Table of Contents

* [Introduction](#Introduction)
* [Installation](#Installation)
* [Sample Application](#Sample-Application)
* [UI Components](#UI-Components)
* [Supported Property References](#Supported-Property-References)
* [Supported Formats](#Supported-Formats)
* [Supported Actions](#Supported-Actions)
* [Pega Mail Merge Syntax](#Pega-Mail-Merge-Syntax)
* [Limitations](#Limitations)
* [Issues](#Issues)

## Introduction

Pega PowerPoint Generator is an Exchange component that enables Pega users to generate PowerPoint presentations from Pega cases. It uses the case data and a template to generate the presentation. It allows Pega users to:
* Reference Pega properties from their slides
* Generate native PowerPoint tables
* Generate native PowerPoint bullets
* Generate native PowerPoint pie, bar and line graphs
* Include images from Pega in their slides

The template is a PowerPoint file. it can include any PowerPoint components as well as references to Pega properties and actions. References and actions can be included using Pega Mail Merge expressions.

## Installation

1. Download `pegaPowerPointGenerator_pkg.zip` from Pega Exchange.
2. Unzip `pegaPowerPointGenerator_pkg.zip`, to create `pegaPowerPointGenerator_pkg` directory that contains the following three sub-directories: 
	 - `archives`: contains `pegaPowerPointGenerator_v1.zip` installable RAP.
	 - `examples`: contains `pptxdemo.zip` sample application and `ppt-demo.pptx` sample  template.
	 -  `documents`: contains `README.md` file 
3. Import `pegaPowerPointGenerator_v1.zip` RAP from the `archives` directory
4. Restart your Pega application server.
5. Add `PegaPowerPointGenerator`  ruleset to your application rulesets.
6. Add `GeneratePowerPoint` Flow Action to your process.
7. Create a PowerPoint template
8. Create a `Rule-Binary-File` and upload your template:
	 - The binary file rule must use pptx extension.
	 - The binary file rule must use ppt directory.

> The sample application contains all the required rulesets. If you are planning to install the sample application, please skip this section and follow the steps in [Sample Application](#Sample-Application).
 
## Sample Application

Pega PowerPoint Generator ships with the PptxDemo sample application. Follow these steps to install the and run application:

* Import the `pptxdemo.zip` RAP from examples directory.
* Restart your application server.
* Login using the following credentials:
	* User: admin@pptx
	* Password: ppt4Pega$
* Create a Demo case and run the Generate Presentation flow.
* Review the template: `ppt.demo.pptx` and the generated presentation

> The demo application uses `ppt.demo.pptx` binary file as a template. `ppt.demo.pptx`  contains expressions for all supported references and actions.

## UI Components

#### GeneratePowerPoint Flow Action

This is the main user interface for the component. It allows the user to:
- Select the presentation template
- Specify the presentation output file name
- Configure the presentation output method:
	 - Attach it to the case
	 - Generate a downloadable link
- Display the settings Flow Action
- Review errors and warnings

#### PowerPointConfiguration Flow Action

This is the main user interface for the component configuration. It allows the user to:
- Select a different locale. The default locale is the requestor locale.
- Select the date/datetime format:

| Format | Date                       | DateTime                           
|--------|----------------------------|------------------------------------|
| Short  | 1/8/20                     | 1/8/20 8:02 PM                     |
| Medium | Jan 8, 2020                | Jan 8, 2020 8:02 PM                |
| Long   | January 8, 2020            | January 8, 2020 8:02 PM            |
| Full   | Wednesday, January 8, 2020 | Wednesday, January 8, 2020 8:02 PM |

- Create a link to the case and add it to the first slide.
- Configure the number of rows per slide for tables.
- Configure the maximum number of slides per table.
- Configure font size for table cells.

> All rules for this components are included in PegaPowerPointGenerator ruleset.

## Supported Property References

- Case reference `{{.MyProperty}}`
- Page reference `{{MyPage.MyProperty}}`
- PageList reference:
	- `{{.MyPageList(1).pyFullName}}`
	- `{{MyPageList(1).pyFullName}}`
- DataPage reference `{{D_PegaPowerPointTestList.pxResults(13).ID}}`
- DataPage reference with parameters `{{D_PegaPowerPointTestList.pxResults(13).Name ID=.ID}}`

> Users can add format to any of the references above using Format parameter. For example, use  `{{pyWorkPage.MyNumberProperty Format=C}}` to format `MyNumberProperty` in currency.

## Supported Formats

| Symbol | Name|Example | Result                 |                                   
|--------|--------------|------------------------|----------------------------------|
| I     | Integer      | `{{.MyProp Format=I}}`  | 12,345                           |
| N     | No format    | `{{.MyProp Format=N}}`  | The property clipboard value.    |
| C     | Currency     | `{{.MyProp Format=C}}`  | $12,345.00 (based on the locale) |
| K     | Thousands    | `{{.MyProp Format=K}}`  | 12.34K                           |
| M     | Millions     | `{{.MyProp Format=M}}`  | 1.23M                            |

> By default, the component uses the requestor locale to generate currency symbol. If a different locale is selected in the settings screen, the component uses the settings locale.

### Default formats

| Property Type   | Example                        
|-----------------|--------------------------------|
| Boolean         | true or false                  |
| Date            | Jan 8, 2020 (Medium)           |
| DateTime        | Jan 8, 2020 8:02 PM (Medium)   |
| Decimal/Double  | 12,345.67                      |
| Integer         | 12345                          |

>Default format is be based on the requestor or the selected locale. The example above uses en_US locale.

## Supported Actions

### Table Action

$Table action generates PowerPoint tables from a Report Definition, a Data Page or a PageList. Generated tables can span multiple pages. Users can use the [settings screen](PowerPointConfiguration-Flow-Action) to configure the table output.

|Data Source|Syntax|Description
|--|--|--|
|Report Definition|`{{$Table Data-PegaPowerPointTest.PegaPowerPointTableTest}}`|Generates a table from `PegaPowerPointTableTest` Report Definition
|Report Definition with Parameters|`{{$Table Data-PegaPowerPointTest.PegaPowerPointTableWithParamsTest ID=.ID Quarter=.Quarter}}`|Generates a table from `PegaPowerPointTableWithParamsTest` Report Definition with parameters Quarter equals .Quarter property.
|Data Page|`{{$Table D_PegaPowerPointTestList.pxResults}}`|Generates a table from a Data Table
|Data Page with Parameters|`{{$Table D_PegaPowerPointTestList.pxResults ID=.ID Quarter=.Quarter}}`|Generate table from a Data Table with parameters Quarter equals .Quarter property.
|PageList|`{{$Table pyWorkPage.Clients}}`|Generates a table from a PageList

>  Users can select the columns to use in the generated table by adding `Columns=[X,Y,Z]` parameter. For example: `{{$Table pyWorkPage.Clients Columns=[AcctNum, Name]}}`  will generate a table with two columns: AcctNum and Name.
>
>  Users can change the table headers by adding `Headers=[X,Y,Z]` parameter. For example: `{{$Table pyWorkPage.Clients Columns=[AcctNum, Name] Headers=[Account Number, Full Name]}}`  will generate a table with two columns: AcctNum and Name and add "Account Number" and "Full Name" headers.
>
>  Users can format the table cell by adding `Format=[X,Y,Z]` parameter. For example: `{{$Table pyWorkPage.QuarterResults Columns=[Name,Qtr,Amt] Headers=[Name,Quarter, Amount] Format=[N,I,C}}`  will generate a table with two columns: AcctNum and Name and add "Account Number" and "Full Name" headers. The Name column will not be formatted, the Qtr column will be formatted as an integer and the Amt column will be formatted as currency.

### Bullets Action

$Bullet action generates PowerPoint bullets from a Report Definition, a Data Page or a PageList.

|Data Source|Syntax|Description
|--|--|--|
|Report Definition|`{{$Bullet Data-PegaPowerPointTest.PegaPowerPointTabeTest.Name}}`|Generates bullets from the `Name` property of `PegaPowerPointTabeTest` Report Definition. 
|Report Definition with Parameters|`{{$Bullets Data-PegaPowerPointTest.PegaPowerPointTabeWithParamsTest.Name Quarter=.Quarter}}`|Generates bullets from the `Name` property of `PegaPowerPointTabeWithParamsTest` Report Definition with parameters Quarter equals .Quarter property.
|Data Page|`{{$Bullets D_ListPegaPowerPointBulletsTest.pxResults.Name}}`|Generates bullets from the `Name` property of `D_ListPegaPowerPointBulletsTest` Data Page.
|Data Page with Parameters|`{{$Bullets D_ListPegaPowerPointBulletsTest.pxResults.Name ID=.ID Quarter=.Quarter}}`|Generates bullets from the `Name` property of `D_ListPegaPowerPointBulletsTest` Data Page with parameters Quarter equals .Quarter property
|PageList|`{{$Bullets pyWorkPage.Clients.Name}}`|Generates bullets from the `Name` property of `Clients` PageList.

### PieChart Action

$PieChart action generates PowerPoint pie chart from a Report Definition, 

|Data Source|Syntax|Description
|--|--|--|
|Report Definition|`{{$PieChart Data-PegaPowerPointTest.PegaPowerPointTabeTest}}`|Generates pie chart from the `PegaPowerPointTabeTest` Report Definition. 
Report Definition with Parameters|`{{$PieChart Data-PegaPowerPointTest.PegaPowerPointTabeWithParamsTest Quarter=.Quarter}}`|Generates pie chart from the `PegaPowerPointTabeWithParamsTest` Report Definition with parameters Quarter equals .Quarter property.

### BarChart Action

$BarChart action generates PowerPoint bar chart from a Report Definition, 

|Data Source|Syntax|Description
|--|--|--|
|Report Definition|`{{$BarChart Data-PegaPowerPointTest.PegaPowerPointMultipleSeriesWithParamsTest}}`|Generates bar chart from the `PegaPowerPointMultipleSeriesWithParamsTest` Report Definition. 
Report Definition with Parameters|`{{$BarChart Data-PegaPowerPointTest.PegaPowerPointMultipleSeriesWithParamsTest Quarter=.Quarter}}`|Generates bar chart from the `PegaPowerPointMultipleSeriesWithParamsTest` Report Definition with parameters Quarter equals .Quarter property.

### LineChart Action

$LineChart action generates PowerPoint line chart from a Report Definition, 

|Data Source|Syntax|Description
|--|--|--|
|Report Definition|`{{$LineChart Data-PegaPowerPointTest.PegaPowerPointMultipleSeriesWithParamsTest}}`|Generates line chart from the `PegaPowerPointMultipleSeriesWithParamsTest` Report Definition. 
Report Definition with Parameters|`{{$LineChart Data-PegaPowerPointTest.PegaPowerPointMultipleSeriesWithParamsTest Quarter=.Quarter}}`|Generates line chart from the `PegaPowerPointMultipleSeriesWithParamsTest` Report Definition with parameters Quarter equals .Quarter property.

> BarChart and LineChart actions support both single series and multiple series charts.
> 
> Users can change the chart type by right clicking on the chart and selecting Change Chart Type > then selecting the desired chart type.

### SectionTable Action

$SectionTable action generates PowerPoint table from a table included in a section rule. Generated tables can span multiple pages.

|Data Source|Syntax|Description
|--|--|--|
|Section Rule|`{{$SectionTable Ppt-PptDemo-Work-Demo.TestTableInSection}}`|Generates a table from the HTML generated from a table in  `TestTableInSection` Section.

> Section rule can't contain more than one table.

### Image Action

$Image action generates a PowerPoint image from a binary file.

|Data Source|Syntax|Description
|--|--|--|
|Binary File Rule|`{{$Image RULE-FILE-BINARY.IMAGES.DEMOIMAGE.PNG}}`|Generates an image from IMAGES.DEMOIMAGE.PNG binary file.

## Pega Mail Merge Syntax

Pega Mail Merge is the expression language behind Pega PowerPoint Generator. Users  can include expressions in their PowerPoint template to reference clipboard properties or to perform actions. 

IDENTIFIER:				Any valid Pega property
property_reference:	Property | DataPage_Property | PageList Property
rule_reference: 			Class '.' IDENTIFIER+
reference:					property_reference | rule_reference
action:						'$'(Table|Bullets|PieChart|BarChart|LineChart|Image|SectionTable)
param:						IDENTIFIER'='reference
params:						param (' ' param)*
columns_expr:				Columns '=' '[' IDENTIFIER (', ' 	IDENTIFIER)* ']'
headers_expr:			Headers '=' '[' IDENTIFIER (', ' 	IDENTIFIER)* ']'
format_expr:					Format '=' '[' IDENTIFIER (', ' 	IDENTIFIER)* ']'
action_expr:				action ' ' reference
property_stmt:			'{{' rule_reference ('Format' '=' IDENTIFIER)? params? '}}'
action_stmt:				'{{' action_expr columns_expr? headers_expr? format_expr? params? '}}'

> Actions are not case sensitive

## Limitations

* Only one table action can be included in a slide. 
* A slide that includes a table action can't contain any other component.
* Table action can only be included in a title and content layout.
* Section rules referenced from a SectionTable action can't contain more than one table.
* Only one chart action can be included in a placeholder.
* Placeholders that contain a chart action can't contain any other component.
* Templates can't contain double curly braces: {{ or }}. To include double curly braces in  a template, users must add the escape character, e.g. \\{\\{ or \\}\\}. 

## Issues

* The generated chart title font is small. To fix that issue double click on the chart and select the desired chart style.
* Bar charts are shifted left. To fix that issue right click on the chart then select Change Chart Type > Columns > 2-D Column.
