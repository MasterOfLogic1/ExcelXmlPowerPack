Imports System.Data
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports ExcelPowerPack.MOLExcelXmlSuppourt

Namespace MOLExcelXml

    Public Class ExcelXmlPowerPack
        Private _filePath As String

        Public Sub New(filePath As String)
            Me._filePath = filePath
        End Sub

        'function the read cell value
        Public Function ReadCellValue(sheetName As String, cellReference As String) As String
            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet).FirstOrDefault(Function(s) s.Name = sheetName)
                If sheet IsNot Nothing Then
                    Dim worksheetPart As WorksheetPart = DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
                    Dim cell As Cell = worksheetPart.Worksheet.Descendants(Of Cell).FirstOrDefault(Function(c) c.CellReference = cellReference)
                    If cell IsNot Nothing AndAlso cell.CellValue IsNot Nothing Then
                        Dim sharedStringTablePart As SharedStringTablePart = workbookPart.GetPartsOfType(Of SharedStringTablePart)().FirstOrDefault()
                        If sharedStringTablePart IsNot Nothing AndAlso cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString Then
                            Dim sharedStringItem As SharedStringItem = sharedStringTablePart.SharedStringTable.ElementAt(Integer.Parse(cell.CellValue.InnerText))
                            If sharedStringItem IsNot Nothing Then
                                Return sharedStringItem.InnerText
                            End If
                        Else
                            Return cell.CellValue.InnerText
                        End If
                    End If
                End If
            End Using
            Return Nothing
        End Function


        Public Function GetAllSheetNames() As String()
            Dim sheetNames As New List(Of String)()
            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                For Each sheet As Sheet In workbookPart.Workbook.Descendants(Of Sheet)()
                    sheetNames.Add(sheet.Name)
                Next
            End Using
            Return sheetNames.ToArray()
        End Function

        Public Function GetSheetByIndex(index As Integer) As String
            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheets As IEnumerable(Of Sheet) = workbookPart.Workbook.Descendants(Of Sheet)()
                Dim sheetIndex As Integer = 0

                For Each sheet As Sheet In sheets
                    If sheetIndex = index Then
                        Return sheet.Name
                    End If
                    sheetIndex += 1
                Next
            End Using

            Return Nothing ' Return Nothing if index is out of range
        End Function


        Public Function GetSheetIndexByName(sheetName As String) As Integer?
            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheets As IEnumerable(Of Sheet) = workbookPart.Workbook.Descendants(Of Sheet)()
                Dim sheetIndex As Integer = 0

                For Each sheet As Sheet In sheets
                    If sheet.Name = sheetName Then
                        Return sheetIndex
                    End If
                    sheetIndex += 1
                Next
            End Using

            Throw New SystemException("Sheet not found")
        End Function

        Public Function GetLastUsedRow(sheetName As String) As Integer
            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)

                If sheet IsNot Nothing Then
                    Dim worksheetPart As WorksheetPart = DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
                    Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData)().First()

                    Dim lastRow As Row = sheetData.Elements(Of Row)().LastOrDefault()
                    If lastRow IsNot Nothing Then
                        Return If(lastRow.RowIndex IsNot Nothing, CInt(lastRow.RowIndex.Value), 0)
                    End If
                End If
            End Using

            Return 0 ' Return 0 if the sheet is not found or if it's empty
        End Function


        Public Function GetLastUsedColumn(sheetName As String) As String() '(letter As String, index As Integer)
            Dim lastColumnIndex As Integer = 0
            Dim lastColumnLetter As String = ""

            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)

                If sheet IsNot Nothing Then
                    Dim worksheetPart As WorksheetPart = DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
                    Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData)().First()

                    For Each row As Row In sheetData.Elements(Of Row)()
                        For Each cell As Cell In row.Elements(Of Cell)()
                            Dim columnName As String = ExcelXmlSupport.GetColumnName(cell.CellReference.Value)
                            Dim columnIndex As Integer = ExcelXmlSupport.CellReferenceToIndex(cell.CellReference.Value)

                            If columnIndex > lastColumnIndex Then
                                lastColumnIndex = columnIndex
                                lastColumnLetter = columnName
                            End If
                        Next
                    Next
                End If
            End Using

            Return {lastColumnLetter, lastColumnIndex}
        End Function

        Public Function GetUsedRange(sheetName As String) As String() '(topLeft As String, bottomRight As String)
            Dim topLeft As String = Nothing
            Dim bottomRight As String = Nothing

            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)

                If sheet IsNot Nothing Then
                    Dim worksheetPart As WorksheetPart = DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
                    Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData)().First()

                    Dim firstRow As Row = sheetData.Elements(Of Row)().FirstOrDefault()
                    Dim lastRow As Row = sheetData.Elements(Of Row)().LastOrDefault()

                    If firstRow IsNot Nothing AndAlso lastRow IsNot Nothing Then
                        Dim firstCell As Cell = firstRow.Elements(Of Cell)().FirstOrDefault()
                        Dim lastCell As Cell = Nothing

                        For Each row As Row In sheetData.Elements(Of Row)()
                            Dim lastCellInRow As Cell = row.Elements(Of Cell)().LastOrDefault()
                            If lastCellInRow IsNot Nothing Then
                                lastCell = lastCellInRow
                            End If
                        Next

                        If firstCell IsNot Nothing AndAlso lastCell IsNot Nothing Then
                            topLeft = firstCell.CellReference
                            bottomRight = lastCell.CellReference
                        End If
                    End If
                End If
            End Using

            Return {topLeft, bottomRight}
        End Function

        Public Sub DeleteSheet(sheetName As String)
            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheets As Sheets = workbookPart.Workbook.Sheets
                Dim sheet As Sheet = sheets.Elements(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)

                If sheet IsNot Nothing Then
                    Dim worksheetPart As WorksheetPart = CType(workbookPart.GetPartById(sheet.Id.Value), WorksheetPart)
                    ' Remove the sheet reference from the workbook
                    sheet.Remove()
                    ' Delete the worksheet part
                    workbookPart.DeletePart(worksheetPart)
                    ' Save the workbook
                    workbookPart.Workbook.Save()
                End If
            End Using
        End Sub

        ' Function to add a new sheet
        Public Sub AddSheet(sheetName As String)
            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim worksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
                worksheetPart.Worksheet = New Worksheet(New SheetData())

                ' Get the Sheets collection
                Dim sheets As Sheets = workbookPart.Workbook.GetFirstChild(Of Sheets)()
                Dim relationshipId As String = workbookPart.GetIdOfPart(worksheetPart)

                ' Generate a unique sheet ID
                Dim sheetId As UInt32 = 1
                If sheets.Elements(Of Sheet)().Any() Then
                    sheetId = sheets.Elements(Of Sheet)().Max(Function(s) s.SheetId.Value) + 1
                End If

                ' Append the new sheet to the Sheets collection
                Dim sheet As New Sheet() With {
                    .Id = relationshipId,
                    .SheetId = sheetId,
                    .Name = sheetName
                }
                sheets.Append(sheet)

                ' Save the workbook
                workbookPart.Workbook.Save()
            End Using
        End Sub

        ' Function to rename a sheet
        Public Sub RenameSheet(oldSheetName As String, newSheetName As String)
            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name = oldSheetName)

                If sheet IsNot Nothing Then
                    sheet.Name = newSheetName
                    ' Save the workbook
                    workbookPart.Workbook.Save()
                End If
            End Using
        End Sub

        ' Function to hide a sheet
        Public Sub HideSheet(sheetName As String)
            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)

                If sheet IsNot Nothing Then
                    sheet.State = SheetStateValues.Hidden
                    ' Save the workbook
                    workbookPart.Workbook.Save()
                End If
            End Using
        End Sub

        ' Function to unhide a sheet
        Public Sub UnhideSheet(sheetName As String)
            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)

                If sheet IsNot Nothing Then
                    sheet.State = SheetStateValues.Visible
                    ' Save the workbook
                    workbookPart.Workbook.Save()
                End If
            End Using
        End Sub

        ' Function to add color to a specific range
        Public Sub AddColorToRange(sheetName As String, startCell As String, endCell As String, colorHex As String)
            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)

                If sheet IsNot Nothing Then
                    Dim worksheetPart As WorksheetPart = DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
                    Dim worksheet As Worksheet = worksheetPart.Worksheet
                    Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()

                    ' Create a new fill pattern
                    Dim fills As Fills = workbookPart.WorkbookStylesPart.Stylesheet.Fills
                    Dim fillPattern As Fill = New Fill(New PatternFill(New ForegroundColor() With {.Rgb = HexBinaryValue.FromString(colorHex)}) With {.PatternType = PatternValues.Solid})
                    fills.Append(fillPattern)
                    fills.Count = CInt(fills.Count.ToString()) + 1
                    workbookPart.WorkbookStylesPart.Stylesheet.Save()

                    ' Create a new cell format
                    Dim cellFormats As CellFormats = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats
                    Dim cellFormat As CellFormat = New CellFormat() With {
                        .FillId = CInt(fills.Count.ToString()) - 1,
                        .BorderId = 0,
                        .FontId = 0,
                        .NumberFormatId = 0,
                        .ApplyFill = True
                    }
                    cellFormats.Append(cellFormat)
                    cellFormats.Count = CInt(cellFormats.Count.ToString()) + 1
                    workbookPart.WorkbookStylesPart.Stylesheet.Save()

                    ' Get the start and end cell references
                    Dim startColumn As String = ExcelXmlSupport.GetColumnName(startCell)
                    Dim startRow As Integer = ExcelXmlSupport.GetRowIndex(startCell)
                    Dim endColumn As String = ExcelXmlSupport.GetColumnName(endCell)
                    Dim endRow As Integer = ExcelXmlSupport.GetRowIndex(endCell)

                    ' Apply the new cell format to each cell in the range
                    For Each row As Row In sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex.Value >= startRow And r.RowIndex.Value <= endRow)
                        For Each cell As Cell In row.Elements(Of Cell)().Where(Function(c) ExcelXmlSupport.IsCellInRange(c.CellReference.Value, startColumn, endColumn))
                            cell.StyleIndex = CInt(cellFormats.Count.ToString()) - 1
                        Next
                    Next

                    worksheet.Save()
                End If
            End Using
        End Sub


        ' Function to delete a range of cells
        Public Sub DeleteRange(sheetName As String, startCell As String, endCell As String)
            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)

                If sheet IsNot Nothing Then
                    Dim worksheetPart As WorksheetPart = DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
                    Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()

                    ' Get the start and end cell references
                    Dim startColumn As String = ExcelXmlSupport.GetColumnName(startCell)
                    Dim startRow As Integer = ExcelXmlSupport.GetRowIndex(startCell)
                    Dim endColumn As String = ExcelXmlSupport.GetColumnName(endCell)
                    Dim endRow As Integer = ExcelXmlSupport.GetRowIndex(endCell)

                    ' Remove cells in the specified range
                    For Each row As Row In sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex.Value >= startRow And r.RowIndex.Value <= endRow).ToList()
                        For Each cell As Cell In row.Elements(Of Cell)().Where(Function(c) ExcelXmlSupport.IsCellInRange(c.CellReference.Value, startColumn, endColumn)).ToList()
                            cell.Remove()
                        Next
                        ' Remove row if it becomes empty
                        If Not row.Elements(Of Cell)().Any() Then
                            row.Remove()
                        End If
                    Next

                    worksheetPart.Worksheet.Save()
                End If
            End Using
        End Sub

        ' Function to read a sheet into a DataTable
        ' Function to read a sheet into a DataTable
        Public Function ReadSheetToDataTable(sheetName As String) As DataTable
            Dim dataTable As New DataTable()

            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                Dim workbookPart As WorkbookPart = document.WorkbookPart
                Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)

                If sheet IsNot Nothing Then
                    Dim worksheetPart As WorksheetPart = DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
                    Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()

                    ' Add columns to DataTable
                    Dim firstRow As Row = sheetData.Elements(Of Row)().FirstOrDefault()
                    If firstRow IsNot Nothing Then
                        For Each cell As Cell In firstRow.Elements(Of Cell)()
                            dataTable.Columns.Add(ExcelXmlSupport.GetCellValue(document, cell), GetType(String))
                        Next
                    End If

                    ' Add rows to DataTable
                    For Each row As Row In sheetData.Elements(Of Row)().Skip(1) ' Skip the first row for headers
                        Dim dataRow As DataRow = dataTable.NewRow()
                        Dim cellIndex As Integer = 0

                        For Each cell As Cell In row.Elements(Of Cell)()
                            Dim cellValue As String = ExcelXmlSupport.GetCellValue(document, cell)
                            If cellIndex < dataTable.Columns.Count Then
                                Dim columnName As String = dataTable.Columns(cellIndex).ColumnName
                                dataRow(columnName) = cellValue
                                cellIndex += 1
                            End If
                        Next

                        dataTable.Rows.Add(dataRow)
                    Next
                End If
            End Using

            Return dataTable
        End Function

    End Class
End Namespace


