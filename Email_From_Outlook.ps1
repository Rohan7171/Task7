# Excel file path
$excelFilePath = "C:\OneDrive_Razorleaf Corporation_Rohan Dalvi\OneDrive - Razorleaf Corporation\Documents\Employee Details.xlsx"

 

 

 

# Connection string for Excel
$connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$excelFilePath;Extended Properties='Excel 12.0;HDR=YES;'"

 

 

 

# Create a connection object
$connection = New-Object System.Data.OleDb.OleDbConnection($connectionString)

 

 

 

# Open the connection
$connection.Open()

 

 

 

# SQL query to select data from the Excel sheet
$query = "SELECT * FROM [Sheet1$]"

 

 

 

# Create a command object
$command = New-Object System.Data.OleDb.OleDbCommand($query, $connection)

 

 

 

# Execute the command and retrieve the data
$dataAdapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
$dataTable = New-Object System.Data.DataTable
$dataAdapter.Fill($dataTable)

 

 

 

# Close the connection
$connection.Close()

 

 

 

# Specify the column indices to retrieve data from
$columnIndex1 = 1  # First column
$columnIndex2 = 2  # Second column
$columnIndex3 = 4  # Second column

 

 

 

# Iterate over each row in the data t
foreach ($row in $dataTable.Rows) {
    # Retrieve the data from the specific column
    $columnData1 = $row[$columnIndex2]
    $columnData2 = $row[$columnIndex3]
    $columnData3 = $row[$columnIndex1]


 

 

 

    # Create Outlook item and send email for each recipient
    $OL = New-Object -ComObject outlook.application
    Start-Sleep 5

 

 

 

    # Create Item
    $mItem = $OL.CreateItem("olMailItem")
    $mItem.To = "$columnData1"
    $mItem.Subject = "Razorleaf"
    $mItem.Body = "Hii $columnData3,Please check the below attachments"
    # Attach a file
   $attachmentPath = "$columnData2"
Write-Host "$attachmentPath"
   $attachment = $mItem.Attachments.Add($attachmentPath)

 

 

 

    # Send the email
    $mItem.Send()
}