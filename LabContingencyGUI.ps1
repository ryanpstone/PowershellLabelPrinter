<# Script to add linebreaks to delimited text data. Assumes first value for each row is 9 digit numeric. #>

<# Necessary info about data. #>
$numColumns = 12
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
$MainWindow.Height = 700
$MainWindow.FormBorderStyle = "Fixed3D"
$MainWindow.AutoSize = $true  # Form enlarges if elements are out-of-bounds

# Create label
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Select order and click button to print label."

# Set label location to (x, y) pixels
$Label.Location  = New-Object System.Drawing.Point(50,515)
$Label.AutoSize = $true

# Add label to form
$MainWindow.Controls.Add($Label)

# Label for new order data grid
$NewOrderLabel = New-Object System.Windows.Forms.Label
$NewOrderLabel.Text = "Enter data for new order and click button to print"
$NewOrderLabel.Location  = New-Object System.Drawing.Point(50,585)
$NewOrderLabel.AutoSize = $true
$MainWindow.Controls.Add($NewOrderLabel)

# Create button, add to form
$Button = New-Object System.Windows.Forms.Button
$Button.Text = "Print Row"
$Button.Location = New-Object System.Drawing.Point(333, 510)
$Button.AutoSize = $true
$MainWindow.Controls.Add($Button)

# Button for new order data grid
$NewOrderButton = New-Object System.Windows.Forms.Button
$NewOrderButton.Text = "Print New Order"
$NewOrderButton.Location = New-Object System.Drawing.Point(333, 580)
$NewOrderButton.AutoSize = $true
$MainWindow.Controls.Add($NewOrderButton)

# Button to clear new order data grid
$ClearNewOrderButton = New-Object System.Windows.Forms.Button
$ClearNewOrderButton.Text = "Clear New Order"
$ClearNewOrderButton.Location = New-Object System.Drawing.Point(500, 580)
$ClearNewOrderButton.AutoSize = $true
$MainWindow.Controls.Add($ClearNewOrderButton)

# Get data from file
$Data = Import-Csv -Path C:\LabContingency\out.txt -Delimiter '^'

$NewOrderTable = New-Object System.Windows.Forms.DataGridView
$NewOrderTable.Location = New-Object System.Drawing.Point(50,50)
$NewOrderTable.Height = 50
$NewOrderTable.Dock = "Bottom"
$NewOrderTable.AllowUserToAddRows = $false
$NewOrderTable.AllowDrop = $false
$NewOrderTable.AllowUserToDeleteRows = $false
$NewOrderTable.AllowUserToOrderColumns = $false
$NewOrderTable.AllowUserToResizeColumns = $false
$NewOrderTable.AllowUserToResizeRows = $false
$NewOrderTable.ColumnHeadersHeightSizeMode = "DisableResizing"
$NewOrderTable.RowHeadersVisible = $true
# $NewOrderTable.SelectionMode = "FullRowSelect"
$NewOrderTable.AutoSizeColumnsMode = "Fill"
$NewOrderTable.Columns.Add("SSN", "SSN") | Out-Null
$NewOrderTable.Columns.Add("ORDNUM", "ORDER #") | Out-Null
$NewOrderTable.Columns.Add("PAT", "PATIENT") | Out-Null
$NewOrderTable.Columns.Add("TEST", "TEST") | Out-Null
$NewOrderTable.Columns.Add("SAMP", "SAMPLE") | Out-Null
$NewOrderTable.Columns.Add("SPEC", "SPECIMEN") | Out-Null
$NewOrderTable.Columns.Add("HOWCOLL", "METHOD") | Out-Null
$NewOrderTable.Columns.Add("PROV", "PROVIDER") | Out-Null
$NewOrderTable.Columns.Add("STATUS", "STATUS") | Out-Null
$NewOrderTable.Columns.Add("URG", "URGENCY") | Out-Null
$NewOrderTable.Columns.Add("ROUTE", "ROUTE") | Out-Null
$NewOrderTable.Columns.Add("PHYLOCPORDLOC", "LOCATION") | Out-Null
$NewOrderTable.Columns.Add("ACC", "ACCESSION") | Out-Null
$NewOrderTable.Columns[12].ReadOnly = $true  # User can't edit accession
$NewOrderTable.Rows.Add() | Out-Null

# DataGridView for lab orders
$OrderTable = New-Object System.Windows.Forms.DataGridView
$OrderTable.Location = New-Object System.Drawing.Point(50,50)
$OrderTable.Height = 500
$OrderTable.Dock = "Top"
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
# Using system time, likely unique enough? Maybe better to use hostname or location
# Last digit of year, Month, Day, minute, second, tenth of second
function GenerateAccession([string]$Assc) {
        if ($Assc -eq "") {
                $Clinic = "A"

                $Date = Get-Date

                $Year = Get-Date -Date $Date -Format "yy"
                $Year = $Year.Substring(1)

                $DayOfYear = $Date.DayOfYear
                $Hour = Get-Date -Date $Date -Format "HH"
                $Minute = Get-Date -Date $Date -Format "mm"
                $Second = Get-Date -Date $Date -Format "ss"
                $SecondsInYear = [int]$DayOfYear * 24 * 60 * 60 + [int]$Hour * 60 * 60 + [int]$Minute * 60 + [int]$Second
                $Hex = '{0:X}' -f $SecondsInYear

                $AsscString = $Clinic + $Year + $Hex
                return $AsscString
        }
        else {
                return $Assc
        }
}

# Function to append order to log file.
function UpdateLog([string]$SSNVal, [string]$OrderVal, [string]$NameVal, [string]$TestVal, [string]$SampVal, [string]$SpecVal, [string]$HowCollVal, [string]$ProvVal, [string]$StatVal, [string]$UrgVal, [string]$RouteVal, [string]$LocVal, [string]$AsscVal) {
        $OutputString = $SSNVal + '^' + $OrderVal + '^' + $NameVal + '^' + $TestVal + '^' + $SampVal + '^' + $SpecVal + '^' + $HowCollVal + '^' + $ProvVal + '^' + $StatVal + '^' + $UrgVal + '^' + $RouteVal + '^' + $LocVal + '^' + $AsscVal
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

        # Uncomment to view print job data in console
        # Write-Host $PrintJobIPL

        Out-Printer -Name $PrinterName -InputObject $PrintJobIPL
}

# Functions to print row when button is clicked
$Button.Add_MouseClick({Button_Click})
function Button_Click() {
        # Get values from selected row
        $row_index = $OrderTable.CurrentRow.Index
        $ssn = $OrderTable.Rows[$row_index].Cells[0].Value
        $order = $OrderTable.Rows[$row_index].Cells[1].Value
        $name = $OrderTable.Rows[$row_index].Cells[2].Value
        $test = $OrderTable.Rows[$row_index].Cells[3].Value
        $samp = $OrderTable.Rows[$row_index].Cells[4].Value
        $spec = $OrderTable.Rows[$row_index].Cells[5].Value
        $howcoll = $OrderTable.Rows[$row_index].Cells[6].Value
        $prov = $OrderTable.Rows[$row_index].Cells[7].Value
        $status = $OrderTable.Rows[$row_index].Cells[8].Value
        $urg = $OrderTable.Rows[$row_index].Cells[9].Value
        $route = $OrderTable.Rows[$row_index].Cells[10].Value
        $location = $OrderTable.Rows[$row_index].Cells[11].Value

        # Create and store accession number
        $assc = $OrderTable.Rows[$row_index].Cells[12].Value
        $OrderTable.Rows[$row_index].Cells[12].Value = GenerateAccession -Assc $assc
        $assc = $OrderTable.Rows[$row_index].Cells[12].Value

        UpdateLog -SSNVal $ssn -NameVal $name -OrderVal $order -TestVal $test -SampVal $samp -SpecVal $spec -HowCollVal $howcoll -ProvVal $prov -StatVal $status -UrgVal $urg -RouteVal $route -LocVal $location -AsscVal $assc
        SendPrintJobIPL -SSNVal $ssn -NameVal $name -OrderVal $order -TestVal $test -AsscVal $assc
}

$NewOrderButton.Add_MouseClick({NewOrder_Button_Click})
function NewOrder_Button_Click() {
        $ssn = $NewOrderTable.Rows[0].Cells[0].Value
        $order = $NewOrderTable.Rows[0].Cells[1].Value
        $name = $NewOrderTable.Rows[0].Cells[2].Value
        $test = $NewOrderTable.Rows[0].Cells[3].Value
        $samp = $NewOrderTable.Rows[0].Cells[4].Value
        $spec = $NewOrderTable.Rows[0].Cells[5].Value
        $howcoll = $NewOrderTable.Rows[0].Cells[6].Value
        $prov = $NewOrderTable.Rows[0].Cells[7].Value
        $status = $NewOrderTable.Rows[0].Cells[8].Value
        $urg = $NewOrderTable.Rows[0].Cells[9].Value
        $route = $NewOrderTable.Rows[0].Cells[10].Value
        $location = $NewOrderTable.Rows[0].Cells[11].Value

        # Create and store accession number
        $assc = $NewOrderTable.Rows[0].Cells[12].Value
        $NewOrderTable.Rows[0].Cells[12].Value = GenerateAccession -Assc $assc
        $assc = $NewOrderTable.Rows[0].Cells[12].Value

        UpdateLog -SSNVal $ssn -NameVal $name -OrderVal $order -TestVal $test -SampVal $samp -SpecVal $spec -HowCollVal $howcoll -ProvVal $prov -StatVal $status -UrgVal $urg -RouteVal $route -LocVal $location -AsscVal $assc
        SendPrintJobIPL -SSNVal $ssn -NameVal $name -OrderVal $order -TestVal $test -AsscVal $assc
}

$ClearNewOrderButton.Add_MouseClick({ClearNewOrder_Button_Click})
function ClearNewOrder_Button_Click() {
        $NewOrderTable.Rows.Clear()
        $NewOrderTable.Rows.Add() | Out-Null
}

$MainWindow.Controls.Add($OrderTable)
$MainWindow.Controls.Add($NewOrderTable)

# Display form
$MainWindow.ShowDialog()
