<# Script to add linebreaks to delimited text data. Assumes first value for each row is 9 digit numeric. #>

<# Necessary info about data. #>
$numColumns = 12
$digitsInColumn1 = 9 # 9?
$delimiter = "^"
$columnToMerge = 13

<# Filenames. #>
$dataFile = "C:\LabContingency\send_orders_to_nt_server.txt"
$outFile = "C:\LabContingency\out.txt"

<# Get contents of unformatted text file, then trim to get rid of trailing spaces. #>
$rawData = [IO.File]::ReadAllText($dataFile)
$rawData = $rawData.Trim()

<# StringBuilder object to hold strings. More efficient than string concatenation. #>
$contentString = [System.Text.StringBuilder]""
$contentString.Append($rawData) | Out-Null
$formattedString = [System.Text.StringBuilder]""

<# Set regex pattern.
Each row starts with a number that has $digitsInColumn1 digits, so we can assume that 
a row ends right before a number with $digitsInColumn1 digits appears in the substring. 
We can use that pattern to find a regex match and get the index of the row's end. #>
$regexColNum = $numColumns - 1 # Can't do math within a regex pattern
# $regexPattern = "([^$delimiter]*\$delimiter){$regexColNum}(?=[\d{$digitsInColumn1}])"
$regexPattern = "([^^]*\^){11}.*?(?=\d{9}\^)"

<# Loop through string, deleting matches from original after adding to formatted version. #>
while ($contentString -match $regexPattern) {
    $formattedString.AppendLine($Matches[0]) | Out-Null # $Matches is a hash table generated from -matches
    <# Delete matched string from original. #>
    $contentString.Remove(0, $Matches[0].Length) | Out-Null
}

# Remove trailing newline.
$formattedString.Remove($formattedString.Length - 2, 2) | Out-Null

# Remove nth column from header by removing delimiter and combining with (n-1)th column's title.
$regexColNum = $columnToMerge - 1
$regexPattern = "([^$delimiter]*\$delimiter){$regexColNum}"
$formattedString -match $regexPattern | Out-Null
$formattedString.Remove($Matches[0].Length - 1, 1) | Out-Null

<# Create file for formatted text. Delete old file first if it already exists. #>
if (Test-Path($outFile)) {
    Remove-Item $outFile | Out-Null
}
New-Item $outFile | Out-Null # Piping to Out-Null hides output that would normally be written to shell.

<# Add formatted string to new file. #>
Add-Content $outFile $formattedString.ToString()

### Now that data is formatted, we can create labels. ###

# Load .NET class 
Add-Type -assembly System.Windows.Forms

# Create the window (form) that will hold elements
$MainWindow = New-Object System.Windows.Forms.Form

# Set parameters for window
$MainWindow.Text ='Contingency Label Printer'
$MainWindow.Width = 1200
$MainWindow.Height = 800
$MainWindow.FormBorderStyle = "Fixed3D"
$MainWindow.AutoSize = $true  # Form enlarges if elements are out-of-bounds

# Create label
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Select order and click button to print label."

# Set label location to (x, y) pixels
$Label.Location  = New-Object System.Drawing.Point(50,15)
$Label.AutoSize = $true

# Add label to form
$MainWindow.Controls.Add($Label)

# Create button, add to form
$Button = New-Object System.Windows.Forms.Button
$Button.Text = "Print Row"
$Button.Location = New-Object System.Drawing.Point(360, 10)
$Button.AutoSize = $true
$MainWindow.Controls.Add($Button)

# Get data from file
$Data = Import-Csv -Path C:\LabContingency\out.txt -Delimiter '^'

# DataGridView
$OrderTable = New-Object System.Windows.Forms.DataGridView
$OrderTable.Location = New-Object System.Drawing.Point(50,50)
$OrderTable.Height = 700
$OrderTable.Dock = "Bottom"
$OrderTable.ReadOnly = $true
$OrderTable.AllowUserToAddRows = $false
$OrderTable.AllowDrop = $false
$OrderTable.AllowUserToDeleteRows = $false
$OrderTable.AllowUserToOrderColumns = $false
$OrderTable.AllowUserToResizeColumns = $false
$OrderTable.AllowUserToResizeRows = $false
$OrderTable.ColumnHeadersHeightSizeMode = "DisableResizing"
$OrderTable.RowHeadersVisible = $false
$OrderTable.SelectionMode = "FullRowSelect"
$OrderTable.AutoSizeColumnsMode = "Fill"
$OrderTable.Columns.Add("SSN", "SSN") | Out-Null
$OrderTable.Columns.Add("ORDNUM", "ORDER #") | Out-Null
$OrderTable.Columns.Add("PAT", "PATIENT") | Out-Null
$OrderTable.Columns.Add("TEST", "TEST") | Out-Null
$OrderTable.Columns.Add("SAMP", "SAMPLE") | Out-Null
$OrderTable.Columns.Add("SPEC", "SPECIMEN") | Out-Null
$OrderTable.Columns.Add("HOWCOLL", "METHOD") | Out-Null
$OrderTable.Columns.Add("PROV", "PROVIDER") | Out-Null
$OrderTable.Columns.Add("STATUS", "STATUS") | Out-Null
$OrderTable.Columns.Add("URG", "URGENCY") | Out-Null
$OrderTable.Columns.Add("ROUTE", "ROUTE") | Out-Null
$OrderTable.Columns.Add("PHYLOCPORDLOC", "LOCATION") | Out-Null
$OrderTable.Columns.Add("ACC", "ACCESSION") | Out-Null

$OrderTable.AutoResizeColumns()
# Add rows from data to DataGridView object
foreach($row in $Data) {
        $OrderTable.Rows.Add($row.SSN, $row.ORDNUM, $row.PAT, $row.TEST, $row.SAMP, $row.SPEC, $row.HOWCOLL, $row.PROV, $row.STATUS, $row.URG, $row.ROUTE, $row.PHYLOCPORDLOC) | Out-Null
}

# Function to create an accession number for order
# Using system time, likely unique enough? Maybe better to use hostname or location...
# Last digit of year, Month, Day, minute, second, tenth of second
function GenerateAccession() {
        $AsscString = Get-Date -Format "yyMMddHHmmssf"
        return $AsscString.Substring(1)
}

# Function to append order to log file.
function UpdateLog([string]$SSNVal, [string]$OrderVal, [string]$NameVal, [string]$AsscVal, [string]$TestVal) {
        $OutputString = $SSNVal + '^' + $OrderVal + '^' + $NameVal + '^' + $AsscVal + '^' + $TestVal
        Out-File -FilePath C:\LabContingency\Log.txt -InputObject $OutputString -Append
}

# Insert data into IPL print job before sending to printer
function SendPrintJobIPL([string]$SSNVal, [string]$OrderVal, [string]$NameVal, [string]$AsscVal, [string]$TestVal) {

        # TODO: Let user change printer name
        $PrinterName = "Generic / Text Only"

        # Generate IPL for label format/data.  Temp format is created and sent to printer each time label is printed.
        # TODO:  Add format documentation comments for each IPL command
        $PrintJobIPL = "<STX><ESC>C<ETX>`n"
        $PrintJobIPL += "<STX><ESC>P<ETX>`n"
        $PrintJobIPL += "<STX>E*;F*;<ETX>`n"
        $PrintJobIPL += "<STX>H1;f0;o20,51;c30;b0;h1;w1;d3,SSN:<ETX>`n"
        $PrintJobIPL += "<STX>H2;f0;o20,79;c30;b0;h1;w1;d3,ORD:<ETX>`n"
        $PrintJobIPL += "<STX>H3;f0;o11,24;c30;b0;h1;w1;d3,NAME:<ETX>`n"
        $PrintJobIPL += "<STX>H4;f3;o185,116;c30;b0;h1;w1;d3,CONTINGENCY ASSC:<ETX>`n"
        $PrintJobIPL += "<STX>H5;f3;o40,116;c30;b0;h1;w1;d3,TST:<ETX>`n"
        $PrintJobIPL += "<STX>H6;f0;o60,51;c30;b0;h1;w1;d3," + $SSNVal + "<ETX>`n"
        $PrintJobIPL += "<STX>H7;f0;o60,79;c30;b0;h1;w1;d3," + $OrderVal + "<ETX>`n"
        $PrintJobIPL += "<STX>H8;f0;o60,24;c30;b0;h1;w1;d3," + $NameVal + "<ETX>`n"
        $PrintJobIPL += "<STX>H9;f3;o185,286;c30;b0;h1;w1;d3," + $AsscVal + "<ETX>`n"
        $PrintJobIPL += "<STX>H10;f3;o40,156;c30;b0;h1;w1;d3," + $TestVal + "<ETX>`n"
        $PrintJobIPL += "<STX>B11;f3;o151,111;c6,0,0,3;w3;h102;d3," + $AsscVal + "<ETX>`n"
        $PrintJobIPL += "<STX>R;<ETX>`n"
        $PrintJobIPL += "<STX><ESC>E*<CAN><ETX>`n"
        $PrintJobIPL += "<STX><ETB><ETX>"

        # Write-Host $PrintJobIPL

        Out-Printer -Name $PrinterName -InputObject $PrintJobIPL
}

# Function to print row when button is clicked
$Button.Add_MouseClick({Button_Click})
function Button_Click() {
        # Get values from selected row
        $row_index = $OrderTable.CurrentRow.Index
        $ssn = $OrderTable.Rows[$row_index].Cells[0].Value
        $name = $OrderTable.Rows[$row_index].Cells[2].Value
        $order = $OrderTable.Rows[$row_index].Cells[1].Value
        $test = $OrderTable.Rows[$row_index].Cells[3].Value

        # Create and store accession number
        $OrderTable.Rows[$row_index].Cells[12].Value = GenerateAccession
        $assc = $OrderTable.Rows[$row_index].Cells[12].Value

        UpdateLog -SSNVal $ssn -NameVal $name -OrderVal $order -TestVal $test -AsscVal $assc
        SendPrintJobIPL -SSNVal $ssn -NameVal $name -OrderVal $order -TestVal $test -AsscVal $assc
}

$MainWindow.Controls.Add($OrderTable)


# Display form
$MainWindow.ShowDialog()
