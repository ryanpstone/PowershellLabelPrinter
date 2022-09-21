# Load .NET class 
Add-Type -assembly System.Windows.Forms

# Create the window (form) that will hold elements
$MainWindow = New-Object System.Windows.Forms.Form

# Set parameters for window
$MainWindow.Text ='Contingency Label Printer'
$MainWindow.Width = 800
$MainWindow.Height = 600
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

# Get header from file
$Data = Import-Csv -Path .\data.txt -Delimiter '^'

# DataGridView
$OrderTable = New-Object System.Windows.Forms.DataGridView
$OrderTable.Location = New-Object System.Drawing.Point(50,50)
$OrderTable.Height = 500
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
$OrderTable.Columns.Add("columnId1", "Column 1")
$OrderTable.Columns.Add("columnId2", "Column 2")
$OrderTable.Columns.Add("columnId3", "Column 3")
$OrderTable.Columns.Add("columnId4", "Column 4")
$OrderTable.Columns.Add("columnId5", "Column 5")

$OrderTable.AutoResizeColumns()
# Add rows from data to DataGridView object
foreach($row in $Data) {
        $OrderTable.Rows.Add($row.columnId1, $row.columnId2, $row.columnId3, $row.columnId4, $row.columnId5)
}

# Function to create an accession number for order
# Using system time, likely unique enough? Maybe better to use hostname or location...
# Last digit of year, Month, Day, minute, second, tenth of second
function GenerateAccession() {
        $AsscString = Get-Date -Format "yyMMddHHmmssf"
        return $AsscString.Substring(1)
}

# Function to append order to log file.
# TODO: Create file if it doesn't exist.
function UpdateLog([string]$SSNVal, [string]$OrderVal, [string]$NameVal, [string]$AsscVal, [string]$TestVal) {
        $OutputString = $SSNVal + '^' + $OrderVal + '^' + $NameVal + '^' + $AsscVal + '^' + $TestVal
        Out-File -FilePath .\Log.txt -InputObject $OutputString -Append
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
        $name = $OrderTable.Rows[$row_index].Cells[1].Value
        $order = $OrderTable.Rows[$row_index].Cells[2].Value
        $test = $OrderTable.Rows[$row_index].Cells[3].Value

        # Create and store accession number
        $OrderTable.Rows[$row_index].Cells[4].Value = GenerateAccession
        $assc = $OrderTable.Rows[$row_index].Cells[4].Value

        UpdateLog -SSNVal $ssn -NameVal $name -OrderVal $order -TestVal $test -AsscVal $assc
        SendPrintJobIPL -SSNVal $ssn -NameVal $name -OrderVal $order -TestVal $test -AsscVal $assc
}

$MainWindow.Controls.Add($OrderTable)


# Display form
$MainWindow.ShowDialog()
