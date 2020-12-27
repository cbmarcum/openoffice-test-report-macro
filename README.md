# openoffice-test-report-macro
An Apache OpenOffice macro written in Apache Groovy to import the results of OpenOffice verification tests into a Calc spreadsheet in a similar format as the JUnit HTML report.

## Prerequisites
Requires the [Groovy Scripting for OpenOffice](https://github.com/cbmarcum/openoffice-groovy) extension for OpenOffice which adds Groovy
 as a macro language and includes the [Groovy UNO Extension](https://github.com/cbmarcum/guno-extension) API.

## Usage
1. Create a new Groovy macro in OpenOffice.
2. Copy/Paste contents of ImportResultXmlMacro.groovy into it.
3. Run the macro and select a result.xml file to import.
