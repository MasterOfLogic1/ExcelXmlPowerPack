Imports System
Imports System.Data
Imports ExcelPowerPack.MOLExcelXml

Module Program
    Sub Main(args As String())
        'Console.WriteLine("Hello World!")

        'Dim cellValue As String = New ExcelWorkbook("C:\Automation\Test\Test.xlsx").ReadCellValue("Sheet1", "B1")
        ' Dim sheetNames As String() = New ExcelWorkbook("C:\Automation\Test\Test.xlsx").GetAllSheetNames()
        'Dim sheetName As String = New ExcelWorkbook("C:\Automation\Test\Test.xlsx").GetSheetByIndex(0)
        'Dim sheetIndex As String = New ExcelWorkbook("C:\Automation\Test\Test.xlsx").GetSheetIndexByName("Sheet1")
        'Dim LastUsedRow As String = New ExcelWorkbook("C:\Automation\Test\Test.xlsx").GetLastUsedRow("Sheet1")
        'Dim lastUsedColumnLetter As String = New ExcelWorkbook("C:\Automation\Test\Test.xlsx").GetLastUsedColumn("Sheet1").letter
        'Dim lastUsedColumnNumber As Integer = New ExcelWorkbook("C:\Automation\Test\Test.xlsx").GetLastUsedColumn("Sheet1").index
        'Dim usedRangeUper As String = New ExcelWorkbook("C:\Automation\Test\Test.xlsx").GetUsedRange("Sheet1").topLeft
        'Dim usedRangeLower As String = New ExcelWorkbook("C:\Automation\Test\Test.xlsx").GetUsedRange("Sheet1").bottomRight
        Dim e As New ExcelXmlPowerPack("C:\Automation\Ribon\GL Extract SNA_STANDARD - Reval.xlsx")
        Dim dt As DataTable = e.ReadSheetToDataTable("Sheet1")
        e.AddColorToRange("Sheet1", "A1", "B2", "FFFF0000")

    End Sub
End Module
