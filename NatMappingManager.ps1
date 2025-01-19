Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll", ExactSpelling = true)]
    public static extern IntPtr GetConsoleWindow();
}
"@

# Minimize the PowerShell console
$consolePtr = [Win32]::GetConsoleWindow()
[Win32]::ShowWindow($consolePtr, 0)  # 0 = Hide the window

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "NAT Static Mapping Manager"
$form.Size = New-Object System.Drawing.Size(800, 400)

# DataGridView to display NAT mappings
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Size = New-Object System.Drawing.Size(770, 250)
$dataGridView.Location = New-Object System.Drawing.Point(10, 10)
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.AllowUserToAddRows = $false
$dataGridView.ReadOnly = $true
$form.Controls.Add($dataGridView)

# Buttons
$addButton = New-Object System.Windows.Forms.Button
$addButton.Text = "Add"
$addButton.Location = New-Object System.Drawing.Point(10, 270)
$form.Controls.Add($addButton)

$editButton = New-Object System.Windows.Forms.Button
$editButton.Text = "Edit"
$editButton.Location = New-Object System.Drawing.Point(90, 270)
$form.Controls.Add($editButton)

$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "Delete"
$deleteButton.Location = New-Object System.Drawing.Point(170, 270)
$form.Controls.Add($deleteButton)

# Custom input dialog function
function Show-InputDialog {
    param (
        [string]$title,
        [hashtable]$defaults = @{}
    )

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = $title
    $dialog.Size = New-Object System.Drawing.Size(300, 350)
    $dialog.StartPosition = "CenterParent"

    $controls = @{}
    $y = 10

    foreach ($field in @("NatName", "ExternalIPAddress", "ExternalPort", "InternalIPAddress", "InternalPort", "Protocol")) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $field
        $label.Location = New-Object System.Drawing.Point(10, $y)
        $label.Size = New-Object System.Drawing.Size(120, 20)
        $dialog.Controls.Add($label)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Text = $defaults[$field] -as [string]
        $textBox.Location = New-Object System.Drawing.Point(140, $y)
        $textBox.Size = New-Object System.Drawing.Size(120, 20)
        $dialog.Controls.Add($textBox)

        $controls[$field] = $textBox
        $y += 30
    }

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(60, $y)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dialog.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(140, $y)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialog.Controls.Add($cancelButton)

    $dialog.AcceptButton = $okButton
    $dialog.CancelButton = $cancelButton

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return @(
            $controls["NatName"].Text,
            $controls["ExternalIPAddress"].Text,
            $controls["ExternalPort"].Text,
            $controls["InternalIPAddress"].Text,
            $controls["InternalPort"].Text,
            $controls["Protocol"].Text
        )
    } else {
        return $null
    }
}

# Load NAT Mappings
function Load-NatMappings {
    $dataGridView.DataSource = $null # Clear the data source
    $mappings = Get-NetNatStaticMapping | ForEach-Object {
        [PSCustomObject]@{
            ID                = $_.StaticMappingID
            NatName           = $_.NatName
            Protocol          = $_.Protocol
            ExternalIPAddress = $_.ExternalIPAddress
            ExternalPort      = $_.ExternalPort
            InternalIPAddress = $_.InternalIPAddress
            InternalPort      = $_.InternalPort
            Active            = $_.Active
        }
    }

    # Convert data to a DataTable for binding
    $dataTable = New-Object System.Data.DataTable
    if ($mappings.Count -gt 0) {
        $mappings[0].PSObject.Properties.Name | ForEach-Object { $dataTable.Columns.Add($_) }
        $mappings | ForEach-Object {
            $row = $dataTable.NewRow()
            $_.PSObject.Properties | ForEach-Object { $row.($_.Name) = $_.Value }
            $dataTable.Rows.Add($row)
        }
    }
    $dataGridView.DataSource = $dataTable
}

# Add Mapping
$addButton.Add_Click({
    $result = Show-InputDialog -title "Add New NAT Mapping"
    if ($result) {
        $natName, $externalIP, $externalPort, $internalIP, $internalPort, $protocol = $result
        try {
            Add-NetNatStaticMapping -NatName $natName -ExternalIPAddress $externalIP -ExternalPort $externalPort -InternalIPAddress $internalIP -InternalPort $internalPort -Protocol $protocol -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show("New mapping added successfully.", "Info")
            Load-NatMappings
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error adding mapping: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

$editButton.Add_Click({
    $selectedRow = $dataGridView.CurrentRow
    if ($selectedRow) {
        $defaults = @{
            NatName           = $selectedRow.Cells["NatName"].Value
            ExternalIPAddress = $selectedRow.Cells["ExternalIPAddress"].Value
            ExternalPort      = $selectedRow.Cells["ExternalPort"].Value
            InternalIPAddress = $selectedRow.Cells["InternalIPAddress"].Value
            InternalPort      = $selectedRow.Cells["InternalPort"].Value
            Protocol          = $selectedRow.Cells["Protocol"].Value
        }

        $result = Show-InputDialog -title "Edit NAT Mapping" -defaults $defaults
        if ($result) {
            $natName, $externalIP, $externalPort, $internalIP, $internalPort, $protocol = $result
            $id = $selectedRow.Cells["ID"].Value
            try {
                # Remove the old mapping with -Confirm:$false
                Remove-NetNatStaticMapping -StaticMappingID $id -Confirm:$false -ErrorAction Stop

                # Add the updated mapping
                Add-NetNatStaticMapping -NatName $natName -ExternalIPAddress $externalIP -ExternalPort $externalPort -InternalIPAddress $internalIP -InternalPort $internalPort -Protocol $protocol -ErrorAction Stop

                [System.Windows.Forms.MessageBox]::Show("Mapping updated successfully.", "Info")
                Load-NatMappings
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error editing mapping: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No mapping selected.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$deleteButton.Add_Click({
    $selectedRow = $dataGridView.CurrentRow
    if ($selectedRow) {
        $id = $selectedRow.Cells["ID"].Value
        if ($id -ne $null) {
            try {
                # Suppress confirmation with -Confirm:$false
                Remove-NetNatStaticMapping -StaticMappingID $id -Confirm:$false -ErrorAction Stop

                [System.Windows.Forms.MessageBox]::Show("Mapping deleted successfully.", "Info")
                Load-NatMappings
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error deleting mapping: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No mapping selected.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})



# Load the initial data
Load-NatMappings

# Show the form
[void]$form.ShowDialog()
