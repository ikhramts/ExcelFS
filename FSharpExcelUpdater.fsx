#r "office.dll"
#r "Microsoft.Office.Interop.Excel"
#r "System.Xml.Linq"

open System
open System.Collections
open System.Diagnostics
open System.IO
open System.Net.Mail
open System.Reflection
open System.Text.RegularExpressions
open System.Xml.Linq
open Microsoft.Office.Interop.Excel

//------------------------------------------------------------------------------------------
//                          Settings
// Worksheet settings.
let worksheetDir = "C:\\Iouri\\ExcelFS\\"
let worksheetBaseName = "Temperatures "
let worksheetDateTimePattern = @"\d\d\d\d.\d\d.\d\d.\d\d.\d\d"
//                                  year  month date  hr  min

// Email Settings.
// Change "useEmail" below to 'true' and update the email settings to send emails.
let useEmail = false
let fromAddress = new MailAddress("you@gmail.com", "Your Name")
let password = "xxxxxxx"

let onSuccessEmailTo = new MailAddressCollection()
onSuccessEmailTo.Add(new MailAddress("you@gmail.com", "Your Name"))
onSuccessEmailTo.Add(new MailAddress("someone.elses@email.com", "Someone's Name"))

let onFailEmailTo  = new MailAddressCollection()
onFailEmailTo.Add(new MailAddress("your@email.com", "Your Name"))

let smtp = new SmtpClient("smtp.gmail.com", 587)
smtp.Credentials <- new System.Net.NetworkCredential("you@gmail.com", password)
smtp.EnableSsl <- true

let mutable excel:ApplicationClass = null 

//------------------------------------------------------------------------------------------
//                          Useful Functions

// A shortcut for easier work with Linq to XML.
let xname name = XName.Get(name)

// Get a temperature at specific location from Yahoo Weather RSS feed.
let getTemperatureAt (whereOnEarthId:string) : double = 
    let url = "http://weather.yahooapis.com/forecastrss?p=" + 
              whereOnEarthId + "&u=c"
    let elementName = xname "{http://xml.weather.yahoo.com/ns/rss/1.0}condition"

    // Load the XML document from Yahoo and pick the first occuring
    // element yweather:condition.  
    let element = XDocument.Load(url).Descendants(elementName)
                    |> Seq.pick (fun xelement -> Some xelement)
    let temperature = Double.Parse(element.Attribute(xname "temp").Value)
    temperature

let quitExcel () :unit=
    printfn "Quitting"
    excel.Quit()
    System.Runtime.InteropServices.Marshal.ReleaseComObject excel |> ignore

let sendEmail (toAddresses: MailAddressCollection) 
              (subject:string) 
              (body:string) 
              (ccAddresses: MailAddressCollection) 
              (attachment:Attachment) :unit = 
    let message = new MailMessage()
    message.From <- fromAddress
    for toAddress in toAddresses do 
        message.To.Add toAddress
    message.Subject <- subject
    message.Body <- body
    
    match ccAddresses with
    | null -> ()
    | _ -> for cc in ccAddresses do 
            message.CC.Add(cc)
    
    match attachment with
    | null -> ()
    | _ -> message.Attachments.Add(attachment)
    
    match useEmail with
    | false -> ()
    | true -> smtp.Send(message)

let failWithError (message:string) = 
    quitExcel()

    // Notify of failure by email.
    let subject = "ERROR: Temperatures Updater "
    sendEmail onFailEmailTo subject message null null
    
    failwith message

//------------------------------------------------------------------------
//                        Start Script.

// Start Excel.  If you want to run things in the background, leave it invisible.
// I find however that keeping Excel visible helps with debugging.
excel <- ApplicationClass(Visible = true)

// Open the the temperatures workbook with the latest date.
let worksheetNamePattern = worksheetBaseName + worksheetDateTimePattern + ".xls"
let workbookFullName = 
    Directory.GetFiles(worksheetDir)
    |> Array.filter (fun name -> Regex.IsMatch(name, worksheetNamePattern))
    |> Array.sort
    |> Array.rev
    |> Array.pick (fun name -> Some name)

excel.Workbooks.Open(workbookFullName)

// Get references to various useful objects in Excel.
let workbookName = workbookFullName.Substring(workbookFullName.LastIndexOf('\\') + 1)
let workbook = excel.Workbooks.Item(workbookName)
let temperatureSheet = workbook.Sheets.[ "Temperatures"] :?> Worksheet
let calculationsSheet = workbook.Sheets.[ "Calculations"] :?> Worksheet
let datetimeColumn = temperatureSheet.Range("Temperatures_DateTime")
let temperatureColumns = temperatureSheet.Range("Temperatures_Data")

// Find the Yahoo WOEID codes for the cities whose temperature we want to get.
let woeidRow = 1
let numWoeids = 10

let maxCityCodes = 10
let cityCodeRow = 1
let cityCodes = [|for column in 0 .. (maxCityCodes - 1) do
                   let cell = temperatureColumns.Offset(cityCodeRow, column)
                   match cell.Value2 with
                   | :? string as code -> yield code
                   | _ -> ()|]

// Find the temperatures.  Assume that there are no gaps betwen the begining and the end
// of the temperature columns.
let temperatures = cityCodes
                   |> Array.map getTemperatureAt

// Write the tempearatures in a new row in the table.
let newRow = 3
datetimeColumn.Offset(newRow).EntireRow.Insert()
calculationsSheet.Activate()
calculationsSheet.Calculate()

datetimeColumn.Offset(newRow).Value2 <- calculationsSheet.Range("Calculations_Now").Value2
temperatureSheet.Activate()

for column in 0 .. temperatures.Length - 1 do
    temperatureColumns.Offset(newRow, column).Value2 <- temperatures.[column]

// Save the workbook with a new name.
let dateTimeText = calculationsSheet.Range("Calculations_NowText").Value2 :?> string
let savedFileName = worksheetDir + worksheetBaseName + dateTimeText + ".xls"
workbook.SaveAs(savedFileName)

quitExcel()

// Send success email.
printfn "Sending Email"
let subject = "Temperature Update for " + dateTimeText
let body = "The updated temperature data is attached."
let attachment = new Attachment(savedFileName)

sendEmail onSuccessEmailTo subject body null attachment

