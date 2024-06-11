# ExcelXmlPowerPack README

## Overview
The `ExcelXmlPowerPack` library provides a set of functions to facilitate reading, writing, and manipulating Excel files using the Open XML SDK. It allows users to perform a variety of operations on Excel workbooks, such as reading cell values, adding or deleting sheets, and modifying cell formats. This DLL is designed to be integrated with Blue Prism to automate Excel-related tasks.

## Features
- Read cell values from specified sheets.
- Retrieve all sheet names in a workbook.
- Get sheet name by index and index by sheet name.
- Identify the last used row and column in a sheet.
- Get the used range of a sheet.
- Add, delete, hide, and unhide sheets.
- Apply color to a range of cells.
- Read a sheet into a DataTable and write a DataTable to a sheet.

## Installation
1. Add the Open XML SDK to your project:
   ```
   Install-Package DocumentFormat.OpenXml
   ```

2. Reference the `ExcelXmlPowerPack` DLL in your Blue Prism project:
   - Open Blue Prism and navigate to the "Objects" section.
   - Create a new Business Object or edit an existing one.
   - In the "Initialize" action, add a reference to the `ExcelXmlPowerPack.dll` by selecting "Imports" and browsing to the location of the DLL file.

3. Include the `ExcelXmlPowerPack` namespace in your Blue Prism code stages:
   ```vb
   Imports ExcelXmlPowerPack
   ```

## Usage

### Initialization
Create an instance of the `ExcelXmlAction` class by providing the file path of the Excel workbook.
```vb
Dim excelActions As New ExcelXmlMain.ExcelXmlAction("[Path to your Excel file]")
```

### Reading Cell Values
Read the value of a specific cell from a specified sheet.
```vb
Dim cellValue As String = excelActions.ReadCellValue("Sheet1", "A1")
```

### Retrieve All Sheet Names
Get all the sheet names in the workbook.
```vb
Dim sheetNames() As String = excelActions.GetAllSheetNames()
```

### Get Sheet Name by Index
Retrieve the sheet name by its index.
```vb
Dim sheetName As String = excelActions.GetSheetByIndex(0)
```

### Get Sheet Index by Name
Retrieve the index of a sheet by its name.
```vb
Dim sheetIndex As Integer? = excelActions.GetSheetIndexByName("Sheet1")
```

### Get Last Used Row and Column
Get the last used row in a sheet.
```vb
Dim lastRow As Integer = excelActions.GetLastUsedRow("Sheet1")
```

Get the last used column in a sheet.
```vb
Dim lastColumn As Object() = excelActions.GetLastUsedColumn("Sheet1")
```

### Get Used Range
Retrieve the used range of a sheet.
```vb
Dim usedRange As Object() = excelActions.GetUsedRange("Sheet1")
```

### Add, Delete, Hide, and Unhide Sheets
Add a new sheet to the workbook.
```vb
excelActions.AddSheet("NewSheet")
```

Delete an existing sheet.
```vb
excelActions.DeleteSheet("SheetToDelete")
```

Hide a sheet.
```vb
excelActions.HideSheet("SheetToHide")
```

Unhide a sheet.
```vb
excelActions.UnhideSheet("SheetToUnhide")
```

### Apply Color to a Range
Apply a color to a range of cells in a sheet.
```vb
excelActions.AddColorToRange("Sheet1", "A1", "B2", "FFFF00")
```

### Read and Write DataTable
Read a sheet into a DataTable.
```vb
Dim dataTable As DataTable = excelActions.ReadSheetToDataTable("Sheet1", "A1:C10", True)
```

Write a DataTable to a sheet.
```vb
ExcelXmlMain.ExcelXmlAction.WriteDataTableToSheet("[Path to your Excel file]", "SheetName", dataTable)
```

## Exception Handling
The library includes comprehensive exception handling for various operations. If an error occurs, the methods throw `SystemException` with a detailed error message.

## License
This library is licensed under the MIT License. Feel free to modify and distribute it as needed.

## Contributing
Contributions are welcome! Please fork the repository and submit pull requests for any enhancements or bug fixes.

## Contact
For any questions or support, please contact the author at [masteroflogic.mol@gmail.com].

---
This README provides a comprehensive guide to using the `ExcelXmlPowerPack` library, covering all the main functionalities and providing examples for each.
