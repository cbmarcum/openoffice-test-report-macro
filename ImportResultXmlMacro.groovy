/* ************************************************************************
 *
 * Copyright 2020 Code Builders, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *  http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 *********************************************************************** */

/*
 * ImportResultXmlMacro.groovy
 * An Apache OpenOffice macro written in Apache Groovy to import the results
 * of OpenOffice verification tests into a Calc spreadsheet in a similar
 * format as the JUnit HTML report.
 * Requires the openoffice-groovy [1] extension for OpenOffice which adds Groovy
 * as a macro language and includes the Groovy UNO Extension API [2].
 * [1] https://github.com/cbmarcum/openoffice-groovy
 * [2] https://github.com/cbmarcum/guno-extension
 */
import com.sun.star.frame.XModel
import com.sun.star.frame.XController
import org.openoffice.guno.UnoExtension

import com.sun.star.sheet.XSpreadsheet
import com.sun.star.sheet.XSpreadsheetDocument
import com.sun.star.sheet.XSpreadsheets
import com.sun.star.sheet.XSpreadsheetView
import com.sun.star.sheet.XViewFreezable

import com.sun.star.beans.XPropertySet
import com.sun.star.table.*

import groovy.xml.XmlParser

import javax.swing.filechooser.FileFilter
import javax.swing.JFileChooser

// final Integer BLUE = 0x004b8d
final Integer BLACK = 0x000000
final Integer WHITE = 0xffffff

// Custom Colors
final Integer BLUE = 0x004b8d
final Integer GREEN = 0x62bb46
final Integer TURQUOISE = 0x00a4d2
final Integer GOLD = 0xffcf22
final Integer ORANGE = 0xf79428
final Integer PURPLE = 0x6e298d
final Integer SKY_BLUE = 0x87e5ff
final Integer BLUE_GREEN = 0x00aa7e
final Integer LIGHT_GRAY = 0x959797
final Integer CRIMSON = 0xd31245
final Integer BROWN = 0x8a4b05
final Integer DARK_GRAY = 0x3f4040
final Integer LIGHTER_GRAY = 0xf2f2f2
final Integer LIGHT_YELLOW = 0xffff99


// a Groovy way use a Java Swing JFileChooser and filter for xml or directories
def openXmlDialog = new JFileChooser(
        dialogTitle: "Choose a result.xml file",
        fileSelectionMode: JFileChooser.FILES_ONLY,
        multiSelectionEnabled: false,
        fileFilter: [getDescription: { -> "*.xml" },
                     accept        : { file -> file ==~ /.*?\.xml/ || file.isDirectory() }] as FileFilter)

if (openXmlDialog.showOpenDialog() == JFileChooser.APPROVE_OPTION) {
    File xmlFile = openXmlDialog.getSelectedFile()
    println(xmlFile.getPath())

    // get the document model from the scripting context which is made available to all scripts
    XModel xModel = XSCRIPTCONTEXT.getDocument()
    XSpreadsheetDocument doc = xModel.guno(XSpreadsheetDocument.class)
    XSpreadsheets xSheets = doc.sheets
    XSpreadsheet sht = doc.getSheetByName("Sheet1")

    XController xController = xModel.currentController
    xSpreadsheetView = xController.guno(XSpreadsheetView.class)
    xFreeze = xController.guno(XViewFreezable.class)

    // create the custom styles
    XPropertySet colPs = doc.getCellStylePropertySet("ColHeading")
    colPs.putAt("IsCellBackgroundTransparent", false) // guno method to replace uno setPropertyValue method
    colPs["CellBackColor"] = LIGHTER_GRAY // which allows the shorter groovy subscript syntax
    colPs["CharWeight"] = com.sun.star.awt.FontWeight.BOLD
    colPs["IsTextWrapped"] = true
    colPs["VertJustify"] = CellVertJustify.BOTTOM
    colPs["HoriJustify"] = CellHoriJustify.LEFT

    XPropertySet rowPs = doc.getCellStylePropertySet("RowHeading")
    rowPs["HoriJustify"] = CellHoriJustify.LEFT
    rowPs["CellBackColor"] = LIGHTER_GRAY
    rowPs["CharWeight"] = com.sun.star.awt.FontWeight.BOLD

    XPropertySet oddRowPs = doc.getCellStylePropertySet("OddRow")
    oddRowPs["HoriJustify"] = CellHoriJustify.LEFT
    oddRowPs["IsCellBackgroundTransparent"] = false
    oddRowPs["CellBackColor"] = LIGHTER_GRAY
    oddRowPs["CharColor"] = DARK_GRAY

    XPropertySet evenRowPs = doc.getCellStylePropertySet("EvenRow")
    evenRowPs["HoriJustify"] = CellHoriJustify.LEFT
    evenRowPs["IsCellBackgroundTransparent"] = false
    evenRowPs["CellBackColor"] = WHITE
    evenRowPs["CharColor"] = DARK_GRAY

    XPropertySet redBgPs = doc.getCellStylePropertySet("RedBg")
    redBgPs["IsCellBackgroundTransparent"] = false
    redBgPs["CellBackColor"] = CRIMSON
    redBgPs["CharColor"] = WHITE


    /*
     Get properties from the xml report using Groovy's XmlParser and find method.
    The top level tag in the report is <testsuite> which holds properties for suite
    name, time, and test and error counts. We're counting them anyway.
    Test cases are the second level <testcase> tag which contains properties and
    the error stacktrace in the body which isn't used in this report.
    A second level <properties> tag contain many third level <property> tags
    with info like os, java version, etc.
     */

    groovy.util.Node testsuite = new XmlParser().parse(xmlFile.getPath())

    // bvt or fvt
    String suiteName = testsuite.@name

    // info.app.date = 2020-12-23
    def testDateProp = testsuite.properties.property.find { it.@name == 'info.app.date' }
    String testDate = testDateProp.@value

    // info.app.buildid = 450m1(Build:9900)
    // info.app.AllLanguages = (en-US)
    def buildIdProp = testsuite.properties.property.find { it.@name == 'info.app.buildid' }
    def buildLangsProp = testsuite.properties.property.find { it.@name == 'info.app.AllLanguages' }
    String buildId = buildIdProp.@value + " " + buildLangsProp.@value

    // info.app.Revision = b324a13c35 
    // full link = // https://github.com/apache/openoffice/commit/b324a13c35
    def buildRevProp = testsuite.properties.property.find { it.@name == 'info.app.Revision' }
    String buildRev = buildRevProp.@value
    String buildRevLink = "https://github.com/apache/openoffice/commit/${buildRev}"

    // Linux-3.10.0-1127.el7.x86_64-amd64
    // use info.os.name and info.os.version
    def osNameProp = testsuite.properties.property.find { it.@name == 'info.os.name' }
    def osVersionProp = testsuite.properties.property.find { it.@name == 'info.os.version' }
    String os = osNameProp.@value + " " + osVersionProp.@value

    // info.hostname
    def hostnameProp = testsuite.properties.property.find { it.@name == 'info.hostname' }
    String hostname = hostnameProp.@value

    // java.vendor + java.version
    def javaVendorProp = testsuite.properties.property.find { it.@name == 'java.vendor' }
    def javaVersionProp = testsuite.properties.property.find { it.@name == 'java.version' }
    String java = javaVendorProp.@value + " " + javaVersionProp.@value

    Integer testCount = 0
    Integer successCount = 0
    Integer failureCount = 0
    Integer errorCount = 0
    Integer ignoredCount = 0

    Integer rowOffset = 9 // make room for heading info

    sht.setFormulaOfCell(0, 0, "Information")
    sht.setFormulaOfCell(0, 1, "Test Date")
    sht.setFormulaOfCell(0, 2, "Build ID")
    sht.setFormulaOfCell(0, 3, "Revision")
    sht.setFormulaOfCell(0, 4, "OS")
    sht.setFormulaOfCell(0, 5, "Host Name")
    sht.setFormulaOfCell(0, 6, "Java")
    sht.setFormulaOfCell(0, 7, "Record")

    sht.setFormulaOfCell(1, 1, testDate)
    sht.setFormulaOfCell(1, 2, buildId)
    sht.setFormulaOfCell(1, 3, "=HYPERLINK(\"${buildRevLink}\";\"${buildRev}\")")
    sht.setFormulaOfCell(1, 4, os)
    sht.setFormulaOfCell(1, 5, hostname)
    sht.setFormulaOfCell(1, 6, java)

    sht.setFormulaOfCell(2, 0, "Summary")
    sht.setFormulaOfCell(2, 1, "Test Suite")
    sht.setFormulaOfCell(2, 2, "All")
    sht.setFormulaOfCell(2, 3, "Success")
    sht.setFormulaOfCell(2, 4, "Failure")
    sht.setFormulaOfCell(2, 5, "Error")
    sht.setFormulaOfCell(2, 6, "Ignored")

    // testcase heading
    sht.setFormulaOfCell(0, (rowOffset - 1), "Class Name")
    sht.setFormulaOfCell(1, (rowOffset - 1), "Method Name")
    sht.setFormulaOfCell(2, (rowOffset - 1), "Error")
    sht.setFormulaOfCell(3, (rowOffset - 1), "Failure")
    sht.setFormulaOfCell(4, (rowOffset - 1), "Ignored")

    /* In results.xml there can be failure with an error.
       If not failure, error, or ignored, report success
       by placing "No <result type>" in the cell */
    testsuite.testcase.eachWithIndex { testcase, i ->

        // keep track of errors to count success
        Boolean hasSuccess = true

        Integer row = i + rowOffset

        println "classname: ${testcase.@classname}, methodname: ${testcase.@methodname}"
        sht.setFormulaOfCell(0, row, testcase.@classname)
        sht.setFormulaOfCell(1, row, testcase.@methodname)
        if (testcase.error) {
            sht.setFormulaOfCell(2, row, testcase.error[0].@message)
            errorCount++
            hasSuccess = false
        } else {
            sht.setFormulaOfCell(2, row, "No Error")
        }
        if (testcase.failure) {
            sht.setFormulaOfCell(3, row, testcase.failure[0].@message)
            failureCount++
            hasSuccess = false
        } else {
            sht.setFormulaOfCell(3, row, "No Failure")
        }
        if (testcase.ignored) {
            sht.setFormulaOfCell(4, row, testcase.ignored[0].@message)
            ignoredCount++
            hasSuccess = false
        } else {
            sht.setFormulaOfCell(4, row, "No Ignored")
        }

        if (hasSuccess) {
            successCount++
        }
        testCount++

    }

    sht.setFormulaOfCell(3, 1, suiteName.toUpperCase())
    sht.setFormulaOfCell(3, 2, testCount.toString())
    sht.setFormulaOfCell(3, 3, successCount.toString())
    sht.setFormulaOfCell(3, 4, failureCount.toString())
    sht.setFormulaOfCell(3, 5, errorCount.toString())
    sht.setFormulaOfCell(3, 6, ignoredCount.toString())

    /*
     * Format everything
     */

    int colCount = 5

    // set the cell style of the header
    XCellRange xCellRange = sht.getCellRangeByPosition(0, (rowOffset - 1), (colCount - 1), (rowOffset - 1))
    XPropertySet xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["CellStyle"] = "ColHeading"
    xCellRange = null
    xCellRangePs = null

    // Stripe the content rows using our custom styles
    (1..testCount).each { r ->
        Integer row = r + rowOffset - 1
        xCellRange = sht.getCellRangeByPosition(0, row, (colCount - 1), row)
        xCellRangePs = xCellRange.guno(XPropertySet.class)
        if (row % 2) {
            // evenRowPs
            xCellRangePs["CellStyle"] =  "EvenRow"
        } else {
            // oddRowPs
            xCellRangePs["CellStyle"] =  "OddRow"
        }
        xCellRange = null
        xCellRangePs = null
    }

    // set 1st column width
    xCellRange = sht.getCellRangeByName("A1:E8")
    XColumnRowRange xColRowRange = xCellRange.guno(XColumnRowRange.class)

    XTableColumns xColumns = xColRowRange.columns

    Object aColumnObj = xColumns.getByIndex(0)
    XPropertySet aColPS = aColumnObj.guno(XPropertySet.class)
    aColPS["Width"] = 6000

    aColumnObj = xColumns.getByIndex(1)
    aColPS = aColumnObj.guno(XPropertySet.class)
    aColPS["Width"] = 6000

    aColumnObj = xColumns.getByIndex(2)
    aColPS = aColumnObj.guno(XPropertySet.class)
    aColPS["Width"] = 5000

    aColumnObj = xColumns.getByIndex(3)
    aColPS = aColumnObj.guno(XPropertySet.class)
    aColPS["Width"] = 5000

    aColumnObj = xColumns.getByIndex(4)
    aColPS = aColumnObj.guno(XPropertySet.class)
    aColPS["Width"] = 5000

    XTableRows xRows = xColRowRange.rows

    (0..7).each { row ->
        Object aRowObj = xRows.getByIndex(row)
        XPropertySet aRowPS = aRowObj.guno(XPropertySet.class)
        aRowPS["Height"] = 600
    }
    xRows = null

    // freeze using col, row
    xFreeze.freezeAtPosition(0, rowOffset)

    // format the header info
    xCellRange = sht.getCellRangeByName("A1:E1")
    xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["CharWeight"] = com.sun.star.awt.FontWeight.BOLD
    xCellRangePs["CellBackColor"] = TURQUOISE
    xCellRange = null
    xCellRangePs = null

    xCellRange = sht.getCellRangeByName("A8:E8")
    xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["CharWeight"] = com.sun.star.awt.FontWeight.BOLD
    xCellRangePs["CellBackColor"] = TURQUOISE
    xCellRange = null
    xCellRangePs = null

    xCellRange = sht.getCellRangeByName("A2:A7")
    xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["CellStyle"] = "RowHeading"
    xCellRange = null
    xCellRangePs = null

    xCellRange = sht.getCellRangeByName("D2:D7")
    xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["CharWeight"] = com.sun.star.awt.FontWeight.BOLD
    xCellRangePs["HoriJustify"] = CellHoriJustify.LEFT
    xCellRange = null
    xCellRangePs = null

    xCellRange = sht.getCellRangeByName("C2")
    xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["CellStyle"] = "RowHeading"
    xCellRange = null
    xCellRangePs = null

    xCellRange = sht.getCellRangeByName("C3")
    xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["CharWeight"] = com.sun.star.awt.FontWeight.BOLD
    xCellRangePs["CellBackColor"] = LIGHT_GRAY
    xCellRange = null
    xCellRangePs = null

    xCellRange = sht.getCellRangeByName("C4")
    xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["CharWeight"] = com.sun.star.awt.FontWeight.BOLD
    xCellRangePs["CellBackColor"] = GREEN
    xCellRange = null
    xCellRangePs = null

    xCellRange = sht.getCellRangeByName("C5")
    xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["CharWeight"] = com.sun.star.awt.FontWeight.BOLD
    xCellRangePs["CellBackColor"] = new Integer(0xFF9999)
    xCellRange = null
    xCellRangePs = null

    xCellRange = sht.getCellRangeByName("C6")
    xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["CharWeight"] = com.sun.star.awt.FontWeight.BOLD
    xCellRangePs["CellBackColor"] = CRIMSON
    xCellRange = null
    xCellRangePs = null

    xCellRange = sht.getCellRangeByName("C7")
    xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["CharWeight"] = com.sun.star.awt.FontWeight.BOLD
    xCellRangePs["CellBackColor"] = GOLD
    xCellRange = null
    xCellRangePs = null

    // format the Date
    xCellRange = sht.getCellRangeByName("B2")
    xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["NumberFormat"] = new Integer(84)
    xCellRangePs["HoriJustify"] = CellHoriJustify.LEFT
    xCellRange = null
    xCellRangePs = null

    // make it look like a hyperlink
    xCellRange = sht.getCellRangeByName("B4")
    xCellRangePs = xCellRange.guno(XPropertySet.class)
    xCellRangePs["CharColor"] = BLUE
    xCellRangePs["CharUnderline"] = com.sun.star.awt.FontUnderline.SINGLE
    xCellRangePs["CharUnderlineColor"] = BLUE
    xCellRange = null
    xCellRangePs = null

    println()


} else {
    println "FileChooser Cancelled"
}

// Groovy OpenOffice scripts should always return 0
return 0
