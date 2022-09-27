# Load .NET class 
Add-Type -assembly System.Windows.Forms

# Get data
$Data = Import-Csv -Path C:\LabContingency\Log.txt -Delimiter '^'

# Create the window (form) that will hold elements
$MainWindow = New-Object System.Windows.Forms.Form

# Set parameters for window
$MainWindow.Text ='Accessions'
$MainWindow.Width = 1200
$MainWindow.Height = 700
$MainWindow.FormBorderStyle = "Fixed3D"
$MainWindow.AutoSize = $true  # Form enlarges if elements are out-of-bounds

# DataGridView for lab orders
$OrderTable = New-Object System.Windows.Forms.DataGridView
$OrderTable.Location = New-Object System.Drawing.Point(0,0)
$OrderTable.Height = 700
$OrderTable.Dock = "Fill"
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
        $OrderTable.Rows.Add($row.SSN, $row.ORDNUM, $row.PAT, $row.TEST, $row.SAMP, $row.SPEC, $row.HOWCOLL, $row.PROV, $row.STATUS, $row.URG, $row.ROUTE, $row.PHYLOCPORDLOC, $row.ACC) | Out-Null
}

$MainWindow.Controls.Add($OrderTable)
$MainWindow.ShowDialog()