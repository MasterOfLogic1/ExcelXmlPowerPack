Imports ExcelXmlPowerPack.ExcelXmlMain

Module Program
    Sub Main(args As String())
        'Console.WriteLine("Hello World!")

        Dim cellValue As String = New ExcelXmlAction("C:\Automation\Test\Test.xlsx").ReadCellValue("Sheet1", "B1")
        Dim sheetNames As String() = New ExcelXmlAction("C:\Automation\Ribon\Robot.xlsx").GetAllSheetNames()
        Dim sheetName As String = New ExcelXmlAction("C:\Automation\Ribon\Robot.xlsx").GetSheetByIndex(12)
        Dim sheetIndex As String = New ExcelXmlAction("C:\Automation\Ribon\Robot.xlsx").GetSheetIndexByName("Sheet11")
        'Dim LastUsedRow As String = New ExcelXmlAction("C:\Automation\Ribon\Robot.xlsx").GetLastUsedRow("Sheet1")
        'Dim lastUsedColumnLetter As String = New ExcelXmlAction("C:\Automation\Ribon\Robot.xlsx").GetLastUsedColumn("Sheet1")(0).ToString()
        'Dim lastUsedColumnNumber As Integer = CInt(New ExcelXmlAction("C:\Automation\Ribon\Robot.xlsx").GetLastUsedColumn("Sheet1")(1))
        'Dim usedRangeUper As String = New ExcelXmlAction("C:\Automation\Ribon\Robot.xlsx").GetUsedRange("Sheet1")(0)
        'Dim usedRangeLower As String = New ExcelXmlAction("C:\Automation\Ribon\Robot.xlsx").GetUsedRange("Sheet1")(1)
        Dim e As New ExcelXmlAction("C:\Automation\Test\Test.xlsx")
        'e.RenameSheet("ROCKY", "Nasty")
        'e.DeleteSheet("Sheet1") 'delete a sheet
        'e.HideSheet("nasty")
        'e.DeleteRange("Robo", "A1", "B5")

        'ExcelCreator.CreateBlankExcel("C:\Automation\Test\you.xlsx")
        Dim dt As DataTable = e.ReadSheetToDataTable("Sheet1", Nothing, True)
        ExcelXmlAction.WriteDataTableToSheet("C:\Automation\Test\RIY.xlsx", "Sheet2", dt)
        'e.ReadSheetToDataTable("Sheet1")
        'e.AddColorToRange("Sheet1", "A1", "B2", "FFFF0000")

    End Sub
End Module
