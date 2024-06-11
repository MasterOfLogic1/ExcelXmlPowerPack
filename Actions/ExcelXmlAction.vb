Imports System.Data
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Drawing.Charts
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports ExcelXmlPowerPack.ExcelXmlHelper

Namespace ExcelXmlMain

    Public Class ExcelXmlAction
        'private variable holding the file path
        Private _filePath As String

        'the constructor intializes this class when called
        Public Sub New(filePath As String)
            Me._filePath = filePath
        End Sub

        'This returns  the cell value of the workbook specified
        Public Function ReadCellValue(sheetName As String, cellReference As String) As String
            Dim errorMessage As String = String.Empty
            Dim cellValue As String = Nothing
            Try
                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    If sheet IsNot Nothing Then
                        Dim worksheetPart As WorksheetPart = DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
                        Dim cell As Cell = worksheetPart.Worksheet.Descendants(Of Cell).FirstOrDefault(Function(c) c.CellReference = cellReference)
                        If cell IsNot Nothing AndAlso cell.CellValue IsNot Nothing Then
                            Dim sharedStringTablePart As SharedStringTablePart = workbookPart.GetPartsOfType(Of SharedStringTablePart)().FirstOrDefault()
                            If sharedStringTablePart IsNot Nothing AndAlso cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString Then
                                Dim sharedStringItem As SharedStringItem = sharedStringTablePart.SharedStringTable.ElementAt(Integer.Parse(cell.CellValue.InnerText))
                                If sharedStringItem IsNot Nothing Then
                                    cellValue = sharedStringItem.InnerText
                                    Return cellValue
                                End If
                            Else
                                cellValue = cell.CellValue.InnerText
                                Return cellValue
                            End If
                        End If
                    Else
                        Throw New SystemException(String.Format("Sheet [{0}] does not exist", sheetName))
                    End If
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while reading cell value : " + errorMessage)
            End If

            Return cellValue
        End Function

        'This returns all the sheetnames in an excel
        Public Function GetAllSheetNames() As String()
            Dim sheetNames As New List(Of String)()
            Dim errorMessage As String = String.Empty
            Try
                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    For Each sheet As Sheet In workbookPart.Workbook.Descendants(Of Sheet)()
                        sheetNames.Add(sheet.Name)
                    Next
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while getting all sheet names : " + errorMessage)
            End If
            Return sheetNames.ToArray()
        End Function


        'This returns the name of a sheet when given an index
        Public Function GetSheetByIndex(index As Integer) As String
            Dim errorMessage As String = String.Empty
            Try
                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim sheets As IEnumerable(Of Sheet) = workbookPart.Workbook.Descendants(Of Sheet)()
                    Dim sheetIndex As Integer = 0

                    For Each sheet As Sheet In sheets
                        If sheetIndex = index Then
                            'found sheet return and exit here
                            Return sheet.Name
                        End If
                        sheetIndex += 1
                    Next
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while getting sheet by index : " + errorMessage)
            End If

            Throw New SystemException("No sheet found at index " + index.ToString()) 'Throw exception when sheet index is out of range
        End Function

        'This returns the index of a sheet when a sheet name is provided
        Public Function GetSheetIndexByName(sheetName As String) As Integer?
            Dim errorMessage As String = String.Empty
            Try
                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim sheets As IEnumerable(Of Sheet) = workbookPart.Workbook.Descendants(Of Sheet)()
                    Dim sheetIndex As Integer = 0
                    For Each sheet As Sheet In sheets
                        If sheet.Name.ToString().Equals(sheetName, StringComparison.OrdinalIgnoreCase) Then
                            'found sheet and return index
                            Return sheetIndex
                        End If
                        sheetIndex += 1
                    Next
                    document.Dispose()
                End Using

            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while getting sheet by index : " + errorMessage)
            End If

            Throw New SystemException("No sheet as " + sheetName.ToString() + " found") 'Throw exception when sheet name is not found
        End Function

        'This returns the Last used row when a sheet name is provided
        Public Function GetLastUsedRow(sheetName As String) As Integer
            Dim errorMessage As String = String.Empty
            Try
                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))

                    If sheet IsNot Nothing Then
                        Dim worksheetPart As WorksheetPart = DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
                        Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData)().First()

                        Dim lastRow As Row = sheetData.Elements(Of Row)().LastOrDefault()
                        If lastRow IsNot Nothing Then
                            Return If(lastRow.RowIndex IsNot Nothing, CInt(lastRow.RowIndex.Value), 0)
                        End If
                    Else
                        Throw New SystemException("No sheet as " + sheetName.ToString() + " found")
                    End If
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while getting last used row : " + errorMessage)
            End If

            Return 0 ' Return 0 for anything else
        End Function

        'This returns the last used column letter and index as an array of object i.e [A,9]
        Public Function GetLastUsedColumn(sheetName As String) As Object() '(letter As String, index As Integer)
            Dim lastColumnIndex As Integer = 0
            Dim lastColumnLetter As String = ""
            Dim errorMessage As String = String.Empty
            Try

                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))

                    If sheet IsNot Nothing Then
                        Dim worksheetPart As WorksheetPart = DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
                        Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData)().First()

                        For Each row As Row In sheetData.Elements(Of Row)()
                            For Each cell As Cell In row.Elements(Of Cell)()
                                Dim columnName As String = ExcelXmlHelperActions.GetColumnName(cell.CellReference.Value)
                                Dim columnIndex As Integer = ExcelXmlHelperActions.CellReferenceToIndex(cell.CellReference.Value)

                                If columnIndex > lastColumnIndex Then
                                    lastColumnIndex = columnIndex
                                    lastColumnLetter = columnName
                                End If
                            Next
                        Next
                    Else
                        Throw New SystemException("No sheet as " + sheetName.ToString() + " found")
                    End If
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while getting last used column : " + errorMessage)
            End If

            Return {lastColumnLetter, lastColumnIndex}
        End Function


        'This returns the Used Range of a sheet as an array of object [A1,C5]
        Public Function GetUsedRange(sheetName As String) As Object() '(topLeft As String, bottomRight As String)
            Dim topLeft As String = Nothing
            Dim bottomRight As String = Nothing
            Dim errorMessage As String = String.Empty
            Try

                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))

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
                    Else
                        Throw New SystemException("No sheet as " + sheetName.ToString() + " found")
                    End If
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while getting used range : " + errorMessage)
            End If

            Return {topLeft, bottomRight}
        End Function



        'This does not return anything but delets a target sheet
        Public Sub DeleteSheet(sheetName As String)
            Dim errorMessage As String = String.Empty
            Try
                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim sheets As Sheets = workbookPart.Workbook.Sheets
                    Dim sheet As Sheet = sheets.Elements(Of Sheet)().FirstOrDefault(Function(s) s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    Dim isLastSheet As Boolean = If(sheets IsNot Nothing AndAlso sheets.Count() <= 1, True, False)
                    If Not isLastSheet Then
                        If sheet IsNot Nothing Then
                            Dim worksheetPart As WorksheetPart = CType(workbookPart.GetPartById(sheet.Id.Value), WorksheetPart)
                            ' Remove the sheet reference from the workbook
                            sheet.Remove()
                            ' Delete the worksheet part
                            workbookPart.DeletePart(worksheetPart)
                            ' Save the workbook
                            workbookPart.Workbook.Save()
                        Else
                            Throw New SystemException("No sheet as " + sheetName.ToString() + " found")
                        End If
                    Else
                        Throw New SystemException("to avoid corrupt excel, sheet [" + sheetName.ToString() + "] would not be deleted because it is the only sheet in this workbook.")
                    End If
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while trying to delete sheet [" + sheetName + "] : " + errorMessage)
            End If

        End Sub


        'This does not return anything but adds a new sheet
        Public Sub AddSheet(sheetName As String)
            Dim errorMessage As String = String.Empty
            Try
                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim worksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
                    worksheetPart.Worksheet = New Worksheet(New SheetData())

                    ' Get the Sheets collection
                    Dim sheets As Sheets = workbookPart.Workbook.GetFirstChild(Of Sheets)()
                    'check if sheet already exists
                    Dim existingSheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    Dim sheetAlreadyAdded As Boolean = If(existingSheet IsNot Nothing AndAlso existingSheet.Name.ToString().ToUpper().Equals(sheetName.ToUpper()), True, False)
                    If Not sheetAlreadyAdded Then
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
                    Else
                        Throw New SystemException("Sheet with name [" + sheetName + "] already exist ")
                    End If
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while trying to add sheet : " + errorMessage)
            End If

        End Sub



        ' This does not return anything but renames an old sheetname to a desired new sheetname
        Public Sub RenameSheet(oldSheetName As String, newSheetName As String)
            Dim errorMessage As String = String.Empty
            Try
                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name.Value.Equals(oldSheetName, StringComparison.OrdinalIgnoreCase))

                    If sheet IsNot Nothing Then
                        'check if sheet already exists
                        Dim existingSheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet).FirstOrDefault(Function(s) s.Name = newSheetName)
                        Dim sheetAlreadyAdded As Boolean = If(existingSheet IsNot Nothing AndAlso existingSheet.Name.ToString().ToUpper().Equals(newSheetName.ToUpper()), True, False)
                        If Not sheetAlreadyAdded Then
                            sheet.Name = newSheetName
                            ' Save the workbook
                            workbookPart.Workbook.Save()
                        Else
                            Throw New SystemException("Sheet with new name [" + newSheetName + "] already exist ")
                        End If
                    Else
                        Throw New SystemException("Old Sheet with name [" + oldSheetName + "] not found ")
                    End If
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while trying to rename sheet : " + errorMessage)
            End If
        End Sub



        ' This does not return anything but hides a sheet
        Public Sub HideSheet(sheetName As String)
            Dim errorMessage As String = String.Empty
            Try
                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))

                    If sheet IsNot Nothing Then
                        sheet.State = SheetStateValues.Hidden
                        ' Save the workbook
                        workbookPart.Workbook.Save()
                    Else
                        Throw New SystemException("No sheet as " + sheetName.ToString() + " found")
                    End If
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while trying to hide sheet : " + errorMessage)
            End If
        End Sub

        ' This does not return anything but unhides a sheet
        Public Sub UnhideSheet(sheetName As String)
            Dim errorMessage As String = String.Empty
            Try
                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))

                    If sheet IsNot Nothing Then
                        sheet.State = SheetStateValues.Visible
                        ' Save the workbook
                        workbookPart.Workbook.Save()
                    Else
                        Throw New SystemException("No sheet as " + sheetName.ToString() + " found")
                    End If
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while trying to unhide sheet : " + errorMessage)
            End If
        End Sub



        ' This does not return anything but adds color to a specific range in a sheet
        Public Sub AddColorToRange(sheetName As String, startCell As String, endCell As String, colorHex As String)
            Dim errorMessage As String = String.Empty
            Try
                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))

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
                        Dim startColumn As String = ExcelXmlHelperActions.GetColumnName(startCell)
                        Dim startRow As Integer = ExcelXmlHelperActions.GetRowIndex(startCell)
                        Dim endColumn As String = ExcelXmlHelperActions.GetColumnName(endCell)
                        Dim endRow As Integer = ExcelXmlHelperActions.GetRowIndex(endCell)

                        ' Apply the new cell format to each cell in the range
                        For Each row As Row In sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex.Value >= startRow And r.RowIndex.Value <= endRow)
                            For Each cell As Cell In row.Elements(Of Cell)().Where(Function(c) ExcelXmlHelperActions.IsCellInRange(c.CellReference.Value, startColumn, endColumn))
                                cell.StyleIndex = CInt(cellFormats.Count.ToString()) - 1
                            Next
                        Next

                        worksheet.Save()
                    Else
                        Throw New SystemException("No sheet as " + sheetName.ToString() + " found")
                    End If
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while trying to add colour to range : " + errorMessage)
            End If
        End Sub


        'This does not return anything but deletes a range of cells
        Public Sub DeleteRange(sheetName As String, startCell As String, endCell As String)
            Dim errorMessage As String = String.Empty
            Try
                Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, True)
                    Dim workbookPart As WorkbookPart = document.WorkbookPart
                    Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))

                    If sheet IsNot Nothing Then
                        Dim worksheetPart As WorksheetPart = DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
                        Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()

                        ' Get the start and end cell references
                        Dim startColumn As String = ExcelXmlHelperActions.GetColumnName(startCell)
                        Dim startRow As Integer = ExcelXmlHelperActions.GetRowIndex(startCell)
                        Dim endColumn As String = ExcelXmlHelperActions.GetColumnName(endCell)
                        Dim endRow As Integer = ExcelXmlHelperActions.GetRowIndex(endCell)

                        ' Remove cells in the specified range
                        For Each row As Row In sheetData.Elements(Of Row)().Where(Function(r) r.RowIndex.Value >= startRow And r.RowIndex.Value <= endRow).ToList()
                            For Each cell As Cell In row.Elements(Of Cell)().Where(Function(c) ExcelXmlHelperActions.IsCellInRange(c.CellReference.Value, startColumn, endColumn)).ToList()
                                cell.Remove()
                            Next
                            ' Remove row if it becomes empty
                            If Not row.Elements(Of Cell)().Any() Then
                                row.Remove()
                            End If
                        Next

                        worksheetPart.Worksheet.Save()

                    Else
                        Throw New SystemException("No sheet as " + sheetName.ToString() + " found")
                    End If
                    document.Dispose()
                End Using
            Catch ex As Exception
                errorMessage = ex.Message
            End Try

            If Not String.IsNullOrEmpty(errorMessage) Then
                Throw New SystemException("Error encountered while trying to delete a sheet : " + errorMessage)
            End If
        End Sub


        ' Function to read a sheet into a DataTable
        Public Function ReadSheetToDataTable(sheetName As String, Optional cellRange As String = Nothing, Optional hasHeader As Boolean = False) As System.Data.DataTable
            Dim dataTable As New System.Data.DataTable()

            Using document As SpreadsheetDocument = SpreadsheetDocument.Open(_filePath, False)
                Dim worksheetPart As WorksheetPart = ExcelXmlHelperActions.GetWorksheetPart(document, sheetName)
                If worksheetPart Is Nothing Then
                    Throw New Exception("Sheet not found.")
                End If

                Dim sheetData As SheetData = worksheetPart.Worksheet.Elements(Of SheetData)().First()

                Dim startCell, endCell As String
                If Not String.IsNullOrEmpty(cellRange) Then
                    ExcelXmlHelperActions.CheckRangeValidity(cellRange)
                    Dim cells = cellRange.Trim().Replace(" ", String.Empty).Split(":"c)
                    startCell = cells(0)
                    endCell = cells(1)
                Else
                    startCell = "A1"
                    endCell = ExcelXmlHelperActions.GetLastUsedColumn(document, sheetName)(0) & ExcelXmlHelperActions.GetLastRow(document, sheetName)
                End If

                Dim startCol As Integer = ExcelXmlHelperActions.ColumnLetterToNumber(ExcelXmlHelperActions.ExtractLetters(startCell))
                Dim endCol As Integer = ExcelXmlHelperActions.ColumnLetterToNumber(ExcelXmlHelperActions.ExtractLetters(endCell))
                Dim startRow As Integer = Integer.Parse(ExcelXmlHelperActions.ExtractNumbers(startCell))
                Dim endRow As Integer = Integer.Parse(ExcelXmlHelperActions.ExtractNumbers(endCell))
                Dim rows As IEnumerable(Of Row) = sheetData.Elements(Of Row)()

                If Not hasHeader Then
                    ' Initialize columns with generic headers
                    For colNumber As Integer = startCol To endCol
                        dataTable.Columns.Add("Column" & colNumber.ToString(), GetType(String))
                    Next
                Else
                    'headear indicated then take first row as header
                    Dim firstRow As Row = sheetData.Elements(Of Row)().FirstOrDefault()
                    rows = rows.Skip(1)
                    For colNumber As Integer = startCol To endCol
                        Dim cn As Integer = colNumber
                        Dim cell = firstRow.Elements(Of Cell)().FirstOrDefault(Function(c) ExcelXmlHelperActions.ColumnLetterToNumber(ExcelXmlHelperActions.ExtractLetters(c.CellReference)) = cn)
                        Dim colName As String = Nothing
                        If cell IsNot Nothing Then
                            colName = ExcelXmlHelperActions.GetCellValue(cell, document.WorkbookPart)
                        Else
                            colName = "Column" + cn.ToString()
                        End If
                        dataTable.Columns.Add(colName, GetType(String))
                    Next

                End If

                ' Loop through rows
                For Each row As Row In rows
                    If CInt(row.RowIndex.ToString) < startRow OrElse CInt(row.RowIndex.ToString) > endRow Then Continue For

                    Dim rowOutput As New List(Of Object)
                    For colNumber As Integer = startCol To endCol
                        Dim cn As Integer = colNumber
                        Dim cell = row.Elements(Of Cell)().FirstOrDefault(Function(c) ExcelXmlHelperActions.ColumnLetterToNumber(ExcelXmlHelperActions.ExtractLetters(c.CellReference)) = cn)
                        If cell IsNot Nothing Then
                            rowOutput.Add(ExcelXmlHelperActions.GetCellValue(cell, document.WorkbookPart))

                        Else
                            rowOutput.Add(Nothing)
                        End If
                    Next
                    dataTable.Rows.Add(rowOutput.ToArray())
                Next
                document.Dispose()
            End Using

            Return dataTable
        End Function


        'static or shared Function to write a sheet to a table
        Public Shared Sub WriteDataTableToSheet(filePath As String, sheetName As String, dataTable As System.Data.DataTable)
            ' Open or create the spreadsheet document
            Dim spreadsheetDocument As SpreadsheetDocument
            If System.IO.File.Exists(filePath) Then
                spreadsheetDocument = SpreadsheetDocument.Open(filePath, True)
            Else
                spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook)
                ' Add a WorkbookPart to the document
                Dim wbprt As WorkbookPart = spreadsheetDocument.AddWorkbookPart()
                wbprt.Workbook = New Workbook()
                wbprt.Workbook.AppendChild(New Sheets())
            End If

            ' Get the WorkbookPart
            Dim workbookPart As WorkbookPart = spreadsheetDocument.WorkbookPart
            Dim sheets As Sheets = workbookPart.Workbook.Sheets

            ' Check if the sheet already exists
            Dim sheet As Sheet = sheets.Elements(Of Sheet)().FirstOrDefault(Function(s) s.Name = sheetName)
            Dim worksheetPart As WorksheetPart

            If sheet Is Nothing Then
                ' Create a new worksheet part
                worksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
                worksheetPart.Worksheet = New Worksheet(New SheetData())

                ' Create a new sheet
                sheet = New Sheet() With {
                        .Id = workbookPart.GetIdOfPart(worksheetPart),
                        .SheetId = sheets.Count() + 1,
                        .Name = sheetName
                    }
                sheets.Append(sheet)
            Else
                ' Get the existing worksheet part
                worksheetPart = DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
            End If

            ' Get the SheetData from the worksheet
            Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()

            ' Write the DataTable to the worksheet
            ' Write the column headers
            Dim headerRow As Row = New Row()
            For Each column As DataColumn In dataTable.Columns
                Dim cell As Cell = New Cell() With {
                        .CellValue = New CellValue(column.ColumnName),
                        .DataType = CellValues.String
                    }
                headerRow.Append(cell)
            Next
            sheetData.Append(headerRow)

            ' Write the rows
            For Each dataRow As DataRow In dataTable.Rows
                Dim newRow As Row = New Row()
                For Each column As DataColumn In dataTable.Columns
                    Dim cellValue As String = If(dataRow(column) IsNot DBNull.Value, dataRow(column).ToString(), String.Empty)
                    Dim cell As Cell = New Cell() With {
                            .CellValue = New CellValue(cellValue),
                            .DataType = CellValues.String
                        }
                    newRow.Append(cell)
                Next
                sheetData.Append(newRow)
            Next

            ' Save the worksheet
            worksheetPart.Worksheet.Save()

            ' Save the workbook
            workbookPart.Workbook.Save()

            ' Close the document
            spreadsheetDocument.Dispose()
        End Sub


    End Class


End Namespace


