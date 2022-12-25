# OpenOffice Test Report Macros and Scripts
Apache OpenOffice macros written in Apache Groovy to import the results of OpenOffice verification tests into a Calc spreadsheet in a similar format as the JUnit HTML report.

There are two macros:
1. `ImportResultXmlMacro.groovy` to import an individual test result XML file into an OpenOffice Calc spreadsheet.
2. `BulkImportResultXmlMacro.groovy` to bulk import a set of 3 BVT and 3 FVT test results into an OpenOffice Calc spreadsheet all on their own sheet. These sets of results can be automated with the `compile-and-test.sh` script described below.

## Prerequisites
Requires the [Groovy Scripting for OpenOffice](https://github.com/cbmarcum/openoffice-groovy) extension for OpenOffice which adds Groovy
 as a macro language and includes the [Groovy UNO Extension](https://github.com/cbmarcum/guno-extension) API.

## Usage
### ImportResultXmlMacro
1. Create a new Groovy macro in OpenOffice.
2. Copy/Paste contents of ImportResultXmlMacro.groovy into it.
3. Run the macro and select a result.xml file to import.
### BulkImportResultXmlMacro
1. Create a new Groovy macro in OpenOffice.
2. Copy/Paste contents of BulkImportResultXmlMacro.groovy into it.
3. Run the macro and select the directory that contains the `bvt-*` and `fvt-*` result directories. By default this is `test/testspace`.

## Convenience Scripts
There are two scripts to automate testing. 
1. `copy-and-clean-build-LinuxX86-64.sh` for copying of the "installed" build of the office and SDK into a common directory for testing.
2. `compile-and-test.sh` to compile the automated tests and run the `BVT` and `FVT` test suites 3 times each and storing the results in `test/testspace` with directory names like `bvt-1`, `bvt-2`, etc.
   