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
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.Size = New-Object System.Drawing.Size(820, 450)

# Optional: Change the form's icon to match PowerShell
# or any other .ico path you have
# $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell.exe).Path)

# Use a nicer font
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

# Create a group box to hold the DataGridView
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = "Existing NAT Mappings"
$groupBox.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$groupBox.Location = New-Object System.Drawing.Point(10, 10)
$groupBox.Size = New-Object System.Drawing.Size(790, 290)
$form.Controls.Add($groupBox)

# DataGridView to display NAT mappings
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Size = New-Object System.Drawing.Size(760, 240)
$dataGridView.Location = New-Object System.Drawing.Point(15, 25)
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.AllowUserToAddRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.MultiSelect = $false

# DataGridView styling
$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
$dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$dataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::WhiteSmoke
$dataGridView.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::LightSteelBlue
$dataGridView.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black

$groupBox.Controls.Add($dataGridView)

# Buttons
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(10, 310)
$buttonPanel.Size = New-Object System.Drawing.Size(790, 50)
$form.Controls.Add($buttonPanel)

$addButton = New-Object System.Windows.Forms.Button
$addButton.Text = "Add"
$addButton.Size = New-Object System.Drawing.Size(75, 30)
$addButton.Location = New-Object System.Drawing.Point(0, 10)

$editButton = New-Object System.Windows.Forms.Button
$editButton.Text = "Edit"
$editButton.Size = New-Object System.Drawing.Size(75, 30)
$editButton.Location = New-Object System.Drawing.Point(90, 10)

$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "Delete"
$deleteButton.Size = New-Object System.Drawing.Size(75, 30)
$deleteButton.Location = New-Object System.Drawing.Point(180, 10)

$buttonPanel.Controls.Add($addButton)
$buttonPanel.Controls.Add($editButton)
$buttonPanel.Controls.Add($deleteButton)

# ToolTips for buttons
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($addButton, "Add a new NAT mapping")
$toolTip.SetToolTip($editButton, "Edit the selected NAT mapping")
$toolTip.SetToolTip($deleteButton, "Delete the selected NAT mapping")

# Custom input dialog function
function Show-InputDialog {
    param (
        [string]$title,
        [hashtable]$defaults = @{ }
    )

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = $title
    $dialog.Size = New-Object System.Drawing.Size(300, 350)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $controls = @{ }
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
            [System.Windows.Forms.MessageBox]::Show("Error adding mapping: $($_.Exception.Message)", "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Edit Mapping
$editButton.Add_Click({
    $selectedRow = $dataGridView.CurrentRow
    if ($selectedRow) {
        $defaults = @{}
        $defaults["NatName"]           = $selectedRow.Cells["NatName"].Value
        $defaults["ExternalIPAddress"] = $selectedRow.Cells["ExternalIPAddress"].Value
        $defaults["ExternalPort"]      = $selectedRow.Cells["ExternalPort"].Value
        $defaults["InternalIPAddress"] = $selectedRow.Cells["InternalIPAddress"].Value
        $defaults["InternalPort"]      = $selectedRow.Cells["InternalPort"].Value
        $defaults["Protocol"]          = $selectedRow.Cells["Protocol"].Value

        $result = Show-InputDialog -title "Edit NAT Mapping" -defaults $defaults
        if ($result) {
            $natName, $externalIP, $externalPort, $internalIP, $internalPort, $protocol = $result
            $id = $selectedRow.Cells["ID"].Value
            try {
                Remove-NetNatStaticMapping -StaticMappingID $id -Confirm:$false -ErrorAction Stop
                Add-NetNatStaticMapping -NatName $natName -ExternalIPAddress $externalIP -ExternalPort $externalPort `
                    -InternalIPAddress $internalIP -InternalPort $internalPort -Protocol $protocol -ErrorAction Stop

                [System.Windows.Forms.MessageBox]::Show("Mapping updated successfully.", "Info")
                Load-NatMappings
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error editing mapping: $($_.Exception.Message)", "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("No mapping selected.", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Delete Mapping
$deleteButton.Add_Click({
    $selectedRow = $dataGridView.CurrentRow
    if ($selectedRow) {
        $id = $selectedRow.Cells["ID"].Value
        if ($id -ne $null) {
            try {
                Remove-NetNatStaticMapping -StaticMappingID $id -Confirm:$false -ErrorAction Stop
                [System.Windows.Forms.MessageBox]::Show("Mapping deleted successfully.", "Info")
                Load-NatMappings
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error deleting mapping: $($_.Exception.Message)", "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("No mapping selected.", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Load the initial data
Load-NatMappings

# Show the form
[void]$form.ShowDialog()
