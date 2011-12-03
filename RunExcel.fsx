// Add references for Excel-related things.
#r "office.dll"
#r "Microsoft.Office.Interop.Excel"

open System
open System.Collections
open System.Text.RegularExpressions
open Microsoft.Office.Interop.Excel

// Start Excel.
let excel = ApplicationClass(Visible = true)

// Open a workbook.
let workbookDir = "C:\\Iouri\\ExcelFS\\" // Update this as necessary.
excel.Workbooks.Open(workbookDir + "Temperatures 2011.08.09.14.58.xls")

// Get a reference to the workbook, then get a reference to a
// work sheet in that workbook as well as some named ranges in the
// worksheet.
let workbook = excel.Workbooks.Item("Temperatures 2011.08.09.14.58.xls")
let temperatureSheet = workbook.Sheets.["Temperatures"] :?> Worksheet
let calculationsSheet = workbook.Sheets.["Calculations"] :?> Worksheet
let datetimeColumn = temperatureSheet.Range("Temperatures_DateTime")
let temperatureColumns = temperatureSheet.Range("Temperatures_Data")

// Recalculate a worksheet.
calculationsSheet.Activate()
calculationsSheet.Calculate()
temperatureSheet.Activate()

// Run a macro.  The Run() function can accept multiple arguments.  The
// first argument is the name of the macro to run, any additional
// arguments become parameters passed to that macro.  In this case
// the macro has no additional parameters.
excel.Run("UpdateLastRunDate")

// Read some cell values, specifically the Yahoo! Weather location codes.
// The cell values are extracted from each Range object using Value2 property,
// and will need to be typecast into specific types.
// For a multi-cell range, Value2 contains a 2-d array of Objects.
// Note that we are assuming that some cells may be empty.  For empty
// cells, Value2 property will be null.
let maxCityCodes = 10
let cityCodeRow = 1
let cityCodes = [|for column in 0 .. (maxCityCodes - 1) do
                   let cell = temperatureColumns.Offset(cityCodeRow, column)
                   match cell.Value2 with
                   | :? string as code -> yield code
                   | _ -> ()|]

// Insert a new row.
let newRow = 3
datetimeColumn.Offset(newRow).EntireRow.Insert()

// Write data into the new row.
datetimeColumn.Offset(newRow).Value2 <- 
    calculationsSheet.Range("Calculations_Now").Value2
let temperatures = [| 18.0; 15.0 |]
for column in 0 .. temperatures.Length - 1 do
    temperatureColumns.Offset(newRow, column).Value2 <- temperatures.[column]

// Save the worksheet under a new name.
// Remember to specify the full path.
let textDate = 
    calculationsSheet.Range("Calculations_NowText").Value2 :?> string
let savedFileName = workbookDir + "Temperatures " + textDate + ".xls"
workbook.SaveAs(savedFileName)

// Exit Excel.  Need to take extra cleanup step to release the COM object,
// associated with the Excel process, or else Excel may quietly stick around.
excel.Quit()
System.Runtime.InteropServices.Marshal.ReleaseComObject excel