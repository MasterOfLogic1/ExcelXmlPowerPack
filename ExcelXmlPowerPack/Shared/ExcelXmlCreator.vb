Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Public Class ExcelXmlCreator
    '...............................This is responsible for creating a blank excel only...................................................
    Public Shared Sub CreateBlankExcel(filePath As String)
        ' Create a new spreadsheet document
        Using spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Create(filePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)
            ' Add a WorkbookPart to the document
            Dim workbookPart As WorkbookPart = spreadsheetDocument.AddWorkbookPart()
            workbookPart.Workbook = New Workbook()

            ' Add a WorksheetPart to the WorkbookPart
            Dim worksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
            worksheetPart.Worksheet = New Worksheet(New SheetData())

            ' Add Sheets to the Workbook
            Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(New Sheets())

            ' Append a new worksheet and associate it with the workbook
            Dim sheet As Sheet = New Sheet() With {
                .Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                .SheetId = 1,
                .Name = "Sheet1"
            }
            sheets.Append(sheet)

            ' Save the workbook
            workbookPart.Workbook.Save()
            spreadsheetDocument.Dispose()
        End Using
    End Sub
End Class