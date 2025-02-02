# =========================
# Full Integrated NAT Manager GUI (Updated)
# =========================

# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import Win32 functions to hide the console
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

# Minimize (hide) the PowerShell console
$consolePtr = [Win32]::GetConsoleWindow()
[Win32]::ShowWindow($consolePtr, 0)  # 0 = Hide

# =========================
# MAIN FORM (Static Mappings)
# =========================

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "NAT Static Mapping Manager"
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.Size = New-Object System.Drawing.Size(840, 500)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

# GroupBox for Static Mappings DataGridView
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = "Existing NAT Mappings"
$groupBox.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$groupBox.Location = New-Object System.Drawing.Point(10, 10)
$groupBox.Size = New-Object System.Drawing.Size(800, 290)
$form.Controls.Add($groupBox)

# DataGridView for Static Mappings
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Size = New-Object System.Drawing.Size(760, 240)
$dataGridView.Location = New-Object System.Drawing.Point(15, 25)
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.AllowUserToAddRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.MultiSelect = $false

# DataGridView styling for Static Mappings
$dataGridView.EnableHeadersVisualStyles = $false
$dataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(225,225,225)
$dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
$dataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$dataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::WhiteSmoke
$dataGridView.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::LightSteelBlue
$dataGridView.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black

$groupBox.Controls.Add($dataGridView)

# Panel for buttons on the main form
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(10, 310)
$buttonPanel.Size = New-Object System.Drawing.Size(800, 50)
$form.Controls.Add($buttonPanel)

# Static mapping buttons
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

# Button for NAT Networks management
$natNetworksButton = New-Object System.Windows.Forms.Button
$natNetworksButton.Text = "NAT Network"
$natNetworksButton.Size = New-Object System.Drawing.Size(100, 30)
$natNetworksButton.Location = New-Object System.Drawing.Point(270, 10)

# Close button for main form
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Size = New-Object System.Drawing.Size(75, 30)
$closeButton.Location = New-Object System.Drawing.Point(380, 10)
$closeButton.Add_Click({ $form.Close() })

$buttonPanel.Controls.Add($addButton)
$buttonPanel.Controls.Add($editButton)
$buttonPanel.Controls.Add($deleteButton)
$buttonPanel.Controls.Add($natNetworksButton)
$buttonPanel.Controls.Add($closeButton)

# ToolTips for buttons on the main form
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($addButton, "Add a new NAT mapping")
$toolTip.SetToolTip($editButton, "Edit the selected NAT mapping")
$toolTip.SetToolTip($deleteButton, "Delete the selected NAT mapping")
$toolTip.SetToolTip($natNetworksButton, "Manage NAT Networks")
$toolTip.SetToolTip($closeButton, "Close this window")

# --- Custom Input Dialog for Static Mappings ---
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

# --- Load Static NAT Mappings ---
function Load-NatMappings {
    $dataGridView.DataSource = $null  # Clear the data source
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
    # Convert to DataTable for binding
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

# --- Static Mapping Button Click Handlers ---

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

# Edit Mapping (remove and re-add)
$editButton.Add_Click({
    $selectedRow = $dataGridView.CurrentRow
    if ($selectedRow) {
        $defaults = @{
            "NatName"           = $selectedRow.Cells["NatName"].Value
            "ExternalIPAddress" = $selectedRow.Cells["ExternalIPAddress"].Value
            "ExternalPort"      = $selectedRow.Cells["ExternalPort"].Value
            "InternalIPAddress" = $selectedRow.Cells["InternalIPAddress"].Value
            "InternalPort"      = $selectedRow.Cells["InternalPort"].Value
            "Protocol"          = $selectedRow.Cells["Protocol"].Value
        }
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
    } else {
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
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error deleting mapping: $($_.Exception.Message)", "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No mapping selected.", "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Load initial static mappings
Load-NatMappings

# =========================
# NAT NETWORKS MANAGEMENT FORM
# =========================

function Show-NatNetworksForm {
    # Create a new form for NAT networks management
    $natForm = New-Object System.Windows.Forms.Form
    $natForm.Text = "NAT Networks Manager"
    $natForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $natForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $natForm.MaximizeBox = $false
    $natForm.MinimizeBox = $false
    $natForm.Size = New-Object System.Drawing.Size(840, 500)
    $natForm.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    
    # GroupBox for NAT Networks DataGridView
    $natGroupBox = New-Object System.Windows.Forms.GroupBox
    $natGroupBox.Text = "NAT Networks"
    $natGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $natGroupBox.Location = New-Object System.Drawing.Point(10, 10)
    $natGroupBox.Size = New-Object System.Drawing.Size(800, 350)
    $natForm.Controls.Add($natGroupBox)
    
    # DataGridView for NAT Networks
    $natDataGrid = New-Object System.Windows.Forms.DataGridView
    $natDataGrid.Size = New-Object System.Drawing.Size(760, 310)
    $natDataGrid.Location = New-Object System.Drawing.Point(15, 25)
    $natDataGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $natDataGrid.AllowUserToAddRows = $false
    $natDataGrid.ReadOnly = $true
    $natDataGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $natDataGrid.MultiSelect = $false
    # Styling for NAT Networks grid
    $natDataGrid.EnableHeadersVisualStyles = $false
    $natDataGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(225,225,225)
    $natDataGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
    $natDataGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $natDataGrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::WhiteSmoke
    $natDataGrid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::LightSteelBlue
    $natDataGrid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
    $natGroupBox.Controls.Add($natDataGrid)
    
    # Panel for NAT Networks buttons
    $natButtonPanel = New-Object System.Windows.Forms.Panel
    $natButtonPanel.Location = New-Object System.Drawing.Point(10, 370)
    $natButtonPanel.Size = New-Object System.Drawing.Size(800, 50)
    $natForm.Controls.Add($natButtonPanel)
    
    $natAddButton = New-Object System.Windows.Forms.Button
    $natAddButton.Text = "Add"
    $natAddButton.Size = New-Object System.Drawing.Size(75, 30)
    $natAddButton.Location = New-Object System.Drawing.Point(0, 10)
    
    $natEditButton = New-Object System.Windows.Forms.Button
    $natEditButton.Text = "Edit"
    $natEditButton.Size = New-Object System.Drawing.Size(75, 30)
    $natEditButton.Location = New-Object System.Drawing.Point(90, 10)
    
    $natDeleteButton = New-Object System.Windows.Forms.Button
    $natDeleteButton.Text = "Delete"
    $natDeleteButton.Size = New-Object System.Drawing.Size(75, 30)
    $natDeleteButton.Location = New-Object System.Drawing.Point(180, 10)
    
    $natRefreshButton = New-Object System.Windows.Forms.Button
    $natRefreshButton.Text = "Refresh"
    $natRefreshButton.Size = New-Object System.Drawing.Size(75, 30)
    $natRefreshButton.Location = New-Object System.Drawing.Point(270, 10)
    
    # Close button for NAT networks form
    $natCloseButton = New-Object System.Windows.Forms.Button
    $natCloseButton.Text = "Close"
    $natCloseButton.Size = New-Object System.Drawing.Size(75, 30)
    $natCloseButton.Location = New-Object System.Drawing.Point(360, 10)
    $natCloseButton.Add_Click({ $natForm.Close() })
    
    $natButtonPanel.Controls.Add($natAddButton)
    $natButtonPanel.Controls.Add($natEditButton)
    $natButtonPanel.Controls.Add($natDeleteButton)
    $natButtonPanel.Controls.Add($natRefreshButton)
    $natButtonPanel.Controls.Add($natCloseButton)
    
    # ToolTips for NAT Networks buttons
    $natToolTip = New-Object System.Windows.Forms.ToolTip
    $natToolTip.SetToolTip($natAddButton, "Add a new NAT network")
    $natToolTip.SetToolTip($natEditButton, "Edit the selected NAT network")
    $natToolTip.SetToolTip($natDeleteButton, "Delete the selected NAT network")
    $natToolTip.SetToolTip($natRefreshButton, "Refresh the NAT networks list")
    $natToolTip.SetToolTip($natCloseButton, "Close this window")
    
    # --- Custom Input Dialog for NAT Networks (Only 3 fields) ---
    function Show-NatNetworkInputDialog {
        param (
            [string]$title,
            [hashtable]$defaults = @{ }
        )
        $dialog = New-Object System.Windows.Forms.Form
        $dialog.Text = $title
        $dialog.Size = New-Object System.Drawing.Size(350, 250)
        $dialog.StartPosition = "CenterParent"
        $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $dialog.MaximizeBox = $false
        $dialog.MinimizeBox = $false
        $dialog.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
        $controls = @{}
        $y = 10
        # Only these three fields will be shown:
        $fields = @(
            "Name",
            "InternalIPInterfaceAddressPrefix",
            "ExternalIPInterfaceAddressPrefix"
        )
    
        foreach ($field in $fields) {
            $label = New-Object System.Windows.Forms.Label
            $label.Text = $field
            $label.Location = New-Object System.Drawing.Point(10, $y)
            $label.Size = New-Object System.Drawing.Size(160, 20)
            $dialog.Controls.Add($label)
    
            $textBox = New-Object System.Windows.Forms.TextBox
            $textBox.Text = $defaults[$field] -as [string]
            $textBox.Location = New-Object System.Drawing.Point(180, $y)
            $textBox.Size = New-Object System.Drawing.Size(120, 20)
            $dialog.Controls.Add($textBox)
    
            $controls[$field] = $textBox
            $y += 30
        }
    
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(80, $y)
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Controls.Add($okButton)
    
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(180, $y)
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.Controls.Add($cancelButton)
    
        $dialog.AcceptButton = $okButton
        $dialog.CancelButton = $cancelButton
    
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $result = @{}
            foreach ($field in $fields) {
                $result[$field] = $controls[$field].Text
            }
            return $result
        }
        else {
            return $null
        }
    }
    
    # --- Load NAT Networks ---
    function Load-NatNetworks {
        $natDataGrid.DataSource = $null  # Clear grid
        # Force Get-NetNat output into an array
        $natList = @(Get-NetNat | ForEach-Object {
            [PSCustomObject]@{
                Name                             = $_.Name
                ExternalIPInterfaceAddressPrefix = $_.ExternalIPInterfaceAddressPrefix
                InternalIPInterfaceAddressPrefix = $_.InternalIPInterfaceAddressPrefix
                IcmpQueryTimeout                 = $_.IcmpQueryTimeout
                TcpEstablishedConnectionTimeout  = $_.TcpEstablishedConnectionTimeout
                TcpTransientConnectionTimeout    = $_.TcpTransientConnectionTimeout
                TcpFilteringBehavior             = $_.TcpFilteringBehavior
                UdpFilteringBehavior             = $_.UdpFilteringBehavior
                UdpIdleSessionTimeout            = $_.UdpIdleSessionTimeout
                UdpInboundRefresh                = $_.UdpInboundRefresh
                Store                            = $_.Store
                Active                           = $_.Active
            }
        })
    
        # Debug output
        if ($natList.Count -eq 0) {
            Write-Host "No NAT networks found."
        }
        else {
            Write-Host "Found $($natList.Count) NAT network(s)."
        }
    
        $dt = New-Object System.Data.DataTable
        if ($natList.Count -gt 0) {
            $natList[0].PSObject.Properties.Name | ForEach-Object { $dt.Columns.Add($_) }
            foreach ($item in $natList) {
                $row = $dt.NewRow()
                $item.PSObject.Properties | ForEach-Object { $row.($_.Name) = $_.Value }
                $dt.Rows.Add($row)
            }
        }
        $natDataGrid.DataSource = $dt
    }
    
    # --- NAT Network Button Handlers ---
    
    # Add NAT Network (using New-NetNat) with only 3 fields
    $natAddButton.Add_Click({
        # Default values as example
        $defaults = @{
            "Name" = "EXAMPLE"
            "InternalIPInterfaceAddressPrefix" = "192.168.0.0/24"
            "ExternalIPInterfaceAddressPrefix" = ""
        }
        $input = Show-NatNetworkInputDialog -title "Add New NAT Network" -defaults $defaults
        if ($input) {
            try {
                $params = @{
                    Name = $input["Name"]
                    InternalIPInterfaceAddressPrefix = $input["InternalIPInterfaceAddressPrefix"]
                }
                if ($input["ExternalIPInterfaceAddressPrefix"]) {
                    $params["ExternalIPInterfaceAddressPrefix"] = $input["ExternalIPInterfaceAddressPrefix"]
                }
                New-NetNat @params -ErrorAction Stop
                [System.Windows.Forms.MessageBox]::Show("New NAT network added successfully.", "Info")
                Load-NatNetworks
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error adding NAT network: $($_.Exception.Message)", "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    
    # Edit NAT Network (remove and re-create) with only 3 fields
    $natEditButton.Add_Click({
        $selectedRow = $natDataGrid.CurrentRow
        if ($selectedRow) {
            $defaults = @{
                "Name" = $selectedRow.Cells["Name"].Value
                "InternalIPInterfaceAddressPrefix" = $selectedRow.Cells["InternalIPInterfaceAddressPrefix"].Value
                "ExternalIPInterfaceAddressPrefix" = $selectedRow.Cells["ExternalIPInterfaceAddressPrefix"].Value
            }
            $input = Show-NatNetworkInputDialog -title "Edit NAT Network" -defaults $defaults
            if ($input) {
                try {
                    $netName = $selectedRow.Cells["Name"].Value
                    Remove-NetNat -Name $netName -Confirm:$false -ErrorAction Stop
                    $params = @{
                        Name = $input["Name"]
                        InternalIPInterfaceAddressPrefix = $input["InternalIPInterfaceAddressPrefix"]
                    }
                    if ($input["ExternalIPInterfaceAddressPrefix"]) {
                        $params["ExternalIPInterfaceAddressPrefix"] = $input["ExternalIPInterfaceAddressPrefix"]
                    }
                    New-NetNat @params -ErrorAction Stop
                    [System.Windows.Forms.MessageBox]::Show("NAT network updated successfully.", "Info")
                    Load-NatNetworks
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error editing NAT network: $($_.Exception.Message)", "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("No NAT network selected.", "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    
    # Delete NAT Network
    $natDeleteButton.Add_Click({
        $selectedRow = $natDataGrid.CurrentRow
        if ($selectedRow) {
            $netName = $selectedRow.Cells["Name"].Value
            if ($netName) {
                try {
                    Remove-NetNat -Name $netName -Confirm:$false -ErrorAction Stop
                    [System.Windows.Forms.MessageBox]::Show("NAT network deleted successfully.", "Info")
                    Load-NatNetworks
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error deleting NAT network: $($_.Exception.Message)", "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("No NAT network selected.", "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    
    # Refresh NAT Networks
    $natRefreshButton.Add_Click({
        Load-NatNetworks
    })
    
    # Load initial NAT networks data and show the NAT Networks form
    Load-NatNetworks
    [void]$natForm.ShowDialog()
}

# --- NAT Networks Button Handler on the Main Form ---
$natNetworksButton.Add_Click({
    Show-NatNetworksForm
})

# Show the main form
[void]$form.ShowDialog()
