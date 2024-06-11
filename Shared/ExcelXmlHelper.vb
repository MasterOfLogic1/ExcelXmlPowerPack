Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Wordprocessing
Imports System.Linq
Imports System.Text.RegularExpressions

Namespace ExcelXmlHelper
    Public Class ExcelXmlHelperActions

        '----------------------------------------------------------------------------All Public Helper Function---------------------------------------------------------------------------------------------------------
        '------------------------------------------These functions are not directly required or used by excel users but help the ExcelXmlPowerPack achieve most of its functions----------------------------------------

        'Helper function to get column name from cell reference
        Public Shared Function GetColumnName(cellReference As String) As String
            Dim regex As New Regex("[A-Za-z]+")
            Dim match As Match = regex.Match(cellReference)
            Return match.Value
        End Function

        'Helper function to get column index from cell refernce
        Public Shared Function CellReferenceToIndex(cellReference As String) As Integer
            Dim columnLetters As String = GetColumnName(cellReference)
            Dim columnNumber As Integer = 0

            For Each c As Char In columnLetters
                columnNumber = (columnNumber * 26) + (Asc(Char.ToUpper(c)) - Asc("A")) + 1
            Next

            Return columnNumber
        End Function

        'helper function to get row index from cell refernce
        Public Shared Function GetRowIndex(cellReference As String) As Integer
            Dim regex As New Regex("\d+")
            Dim match As Match = regex.Match(cellReference)
            Return Integer.Parse(match.Value)
        End Function

        'helper function check if cell refernce is within range
        Public Shared Function IsCellInRange(cellReference As String, startColumn As String, endColumn As String) As Boolean
            Dim columnName As String = GetColumnName(cellReference)
            Return String.Compare(columnName, startColumn) >= 0 And String.Compare(columnName, endColumn) <= 0
        End Function

        'Helper function to get last used column

        Public Shared Function GetLastUsedColumn(document As SpreadsheetDocument, sheetName As String) As Object() '(letter As String, index As Integer)
            Dim lastColumnIndex As Integer = 0
            Dim lastColumnLetter As String = ""
            Dim errorMessage As String = String.Empty
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

            Return {lastColumnLetter, lastColumnIndex}
        End Function

        'Helper Function to get last used row
        Public Shared Function GetLastRow(document As SpreadsheetDocument, sheetName As String) As String
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
            Return 0 ' Return 0 for anything else
        End Function


        Public Shared Function IsDateFormat(numberFormatId As UInt32Value) As Boolean
            Dim dateFormatIds As UInt32() = {14, 15, 16, 17, 22, 45, 46, 47}
            Return dateFormatIds.Contains(numberFormatId.Value)
        End Function


        'Helper Function to extract letters from a cell reference
        Public Shared Function ExtractLetters(cell As String) As String
            Dim letters As String = String.Empty
            For Each ch As Char In cell
                If Char.IsLetter(ch) Then
                    letters &= ch
                Else
                    Exit For
                End If
            Next
            Return letters
        End Function

        'Helper Function to extract numbers from a cell reference
        Public Shared Function ExtractNumbers(cell As String) As String
            Dim numbers As String = String.Empty
            For Each ch As Char In cell
                If Char.IsDigit(ch) Then
                    numbers &= ch
                End If
            Next
            Return numbers
        End Function

        'Helper Function to convert column letter to number
        Public Shared Function ColumnLetterToNumber(columnLetter As String) As Integer
            Dim sum As Integer = 0
            For i As Integer = 0 To columnLetter.Length - 1
                sum *= 26
                sum += (Asc(columnLetter(i)) - Asc("A"c)) + 1
            Next
            Return sum
        End Function

        ' Function to convert column number to letter
        Public Shared Function ColumnNumberToLetter(columnNumber As Integer) As String
            Dim columnLetter As String = String.Empty
            While columnNumber > 0
                columnNumber -= 1
                columnLetter = Chr((columnNumber Mod 26) + Asc("A"c)) & columnLetter
                columnNumber \= 26
            End While
            Return columnLetter
        End Function

        Public Shared Function GetWorksheetPart(document As SpreadsheetDocument, sheetName As String) As WorksheetPart
            Dim workbookPart As WorkbookPart = document.WorkbookPart
            Dim sheet As Sheet = workbookPart.Workbook.Descendants(Of Sheet)().FirstOrDefault(Function(s) s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))

            If sheet IsNot Nothing Then
                Return DirectCast(workbookPart.GetPartById(sheet.Id), WorksheetPart)
            End If

            Return Nothing
        End Function

        ' Helper function to get cell value
        Public Shared Function GetCellValue(cell As Cell, workbookPart As WorkbookPart) As String
            If cell.CellValue Is Nothing Then
                Return String.Empty
            End If

            Dim value As String = cell.CellValue.InnerText

            If cell.DataType IsNot Nothing Then
                Select Case cell.DataType.Value
                    Case CellValues.SharedString
                        Dim sharedStringTable = workbookPart.GetPartsOfType(Of SharedStringTablePart).FirstOrDefault()
                        If sharedStringTable IsNot Nothing Then
                            value = sharedStringTable.SharedStringTable.ChildElements(Integer.Parse(value)).InnerText
                        End If
                    Case CellValues.Boolean
                        value = If(value = "1", "TRUE", "FALSE")
                        ' Add cases for other cell data types as needed
                    Case CellValues.Date
                        If Double.TryParse(value, Nothing) Then
                            value = DateTime.FromOADate(Double.Parse(value)).ToString("dd/MM/yyyy")
                        End If
                    Case Else
                        ' For other types, use the cell value as is
                End Select
            ElseIf cell.StyleIndex IsNot Nothing Then
                ' Check for date format in the style
                Dim stylesPart As WorkbookStylesPart = workbookPart.WorkbookStylesPart
                Dim cellFormat As CellFormat = DirectCast(stylesPart.Stylesheet.CellFormats.ChildElements(CInt(cell.StyleIndex.Value)), CellFormat)
                If cellFormat.NumberFormatId IsNot Nothing AndAlso IsDateFormat(cellFormat.NumberFormatId) Then
                    'If Double.TryParse(value, value) Then
                    value = DateTime.FromOADate(Double.Parse(value)).ToString("dd/MM/yyyy")
                    'End If
                End If
            End If

            Return value
        End Function

        Public Shared Sub CheckRangeValidity(cellRange As String)
            Try
                ' Trim and split the input range
                Dim rangeParts As String() = cellRange.Trim().Replace(" ", String.Empty).Split(":"c)

                ' If the range does not have exactly two parts, it's invalid
                If rangeParts.Length <> 2 Then
                    Throw New ArgumentException("Invalid range format.")
                End If

                ' Validate each part of the range
                ValidateCell(rangeParts(0))
                ValidateCell(rangeParts(1))

                ' Ensure that the first cell is top-left and the second is bottom-right
                If Not IsRangeValid(rangeParts(0), rangeParts(1)) Then
                    Throw New ArgumentException("Invalid range: start cell is not top-left of end cell.")
                End If

                Console.WriteLine("The range is valid.")
            Catch ex As Exception
                Throw New SystemException(ex.Message)
            End Try
        End Sub


        '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


        '------------------------------------------------------------------------------All Private Helper Function -----------------------------------------------------------------------------------------------------------------

        Private Shared Sub ValidateCell(cell As String)
            ' Ensure the cell reference has at least one letter and one number
            If String.IsNullOrEmpty(cell) OrElse Not cell.Any(AddressOf Char.IsLetter) OrElse Not cell.Any(AddressOf Char.IsDigit) Then
                Throw New ArgumentException($"Invalid cell reference: {cell}")
            End If

            ' Extract column letters and row numbers
            Dim letters As String = ExcelXmlHelperActions.ExtractLetters(cell)
            Dim numbers As String = ExcelXmlHelperActions.ExtractNumbers(cell)

            ' Validate column letters and row numbers
            If String.IsNullOrEmpty(letters) OrElse String.IsNullOrEmpty(numbers) Then
                Throw New ArgumentException($"Invalid cell reference: {cell}")
            End If

            ' Convert column letters to a number and check if it is within Excel's column limits
            Dim colNumber As Integer = ExcelXmlHelperActions.ColumnLetterToNumber(letters)
            If colNumber < 1 OrElse colNumber > 16384 Then
                Throw New ArgumentException($"Invalid column in cell reference: {cell}")
            End If

            ' Convert row numbers to an integer and check if it is within Excel's row limits
            Dim rowNumber As Integer = Integer.Parse(numbers)
            If rowNumber < 1 OrElse rowNumber > 1048576 Then
                Throw New ArgumentException($"Invalid row in cell reference: {cell}")
            End If
        End Sub

        Private Shared Function IsRangeValid(startCell As String, endCell As String) As Boolean
            ' Extract letters and numbers from both cells
            Dim startCol As Integer = ExcelXmlHelperActions.ColumnLetterToNumber(ExcelXmlHelperActions.ExtractLetters(startCell))
            Dim startRow As Integer = Integer.Parse(ExcelXmlHelperActions.ExtractNumbers(startCell))
            Dim endCol As Integer = ExcelXmlHelperActions.ColumnLetterToNumber(ExcelXmlHelperActions.ExtractLetters(endCell))
            Dim endRow As Integer = Integer.Parse(ExcelXmlHelperActions.ExtractNumbers(endCell))

            ' Check if the start cell is top-left of the end cell
            Return startCol <= endCol AndAlso startRow <= endRow
        End Function

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


    End Class
End Namespace

