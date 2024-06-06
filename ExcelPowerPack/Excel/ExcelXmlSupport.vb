Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports System.Linq
Imports System.Text.RegularExpressions

Namespace MOLExcelXmlSuppourt
    Public Class ExcelXmlSupport

        Public Shared Function GetColumnName(cellReference As String) As String
            Dim regex As New Regex("[A-Za-z]+")
            Dim match As Match = regex.Match(cellReference)
            Return match.Value
        End Function

        Public Shared Function CellReferenceToIndex(cellReference As String) As Integer
            Dim columnLetters As String = GetColumnName(cellReference)
            Dim columnNumber As Integer = 0

            For Each c As Char In columnLetters
                columnNumber = (columnNumber * 26) + (Asc(Char.ToUpper(c)) - Asc("A")) + 1
            Next

            Return columnNumber
        End Function


        Public Shared Function GetRowIndex(cellReference As String) As Integer
            Dim regex As New Regex("\d+")
            Dim match As Match = regex.Match(cellReference)
            Return Integer.Parse(match.Value)
        End Function

        Public Shared Function IsCellInRange(cellReference As String, startColumn As String, endColumn As String) As Boolean
            Dim columnName As String = GetColumnName(cellReference)
            Return String.Compare(columnName, startColumn) >= 0 And String.Compare(columnName, endColumn) <= 0
        End Function

        Public Shared Function GetCellValue(document As SpreadsheetDocument, cell As Cell) As String
            If cell.CellValue Is Nothing Then
                Return String.Empty
            End If

            Dim value As String = cell.CellValue.InnerText

            If cell.DataType IsNot Nothing Then
                Select Case cell.DataType.Value
                    Case CellValues.SharedString
                        Dim stringTablePart As SharedStringTablePart = document.WorkbookPart.SharedStringTablePart
                        value = stringTablePart.SharedStringTable.ChildElements(Integer.Parse(value)).InnerText
                    Case CellValues.Boolean
                        value = If(value = "1", "TRUE", "FALSE")
                    Case CellValues.Date
                        If Double.TryParse(value, Nothing) Then
                            value = DateTime.FromOADate(Double.Parse(value)).ToString("MM/dd/yyyy")
                        End If
                    Case Else
                        ' For other types, use the cell value as is
                End Select
            ElseIf cell.StyleIndex IsNot Nothing Then
                ' Check for date format in the style
                Dim stylesPart As WorkbookStylesPart = document.WorkbookPart.WorkbookStylesPart
                Dim cellFormat As CellFormat = DirectCast(stylesPart.Stylesheet.CellFormats.ChildElements(CInt(cell.StyleIndex.Value)), CellFormat)
                If cellFormat.NumberFormatId IsNot Nothing AndAlso IsDateFormat(cellFormat.NumberFormatId) Then
                    'If Double.TryParse(value, value) Then
                    value = DateTime.FromOADate(Double.Parse(value)).ToString("MM/dd/yyyy")
                    'End If
                End If
            End If

            Return value
        End Function

        Public Shared Function IsDateFormat(numberFormatId As UInt32Value) As Boolean
            Dim dateFormatIds As UInt32() = {14, 15, 16, 17, 22, 45, 46, 47}
            Return dateFormatIds.Contains(numberFormatId.Value)
        End Function


    End Class
End Namespace

