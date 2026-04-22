# =========================
# NAT Manager GUI for Windows Systems
# =========================

# Prerequisites validated below:
# - PowerShell 5.1 or later
# - NetNat module
# - Administrator privileges

# Load necessary assemblies for Windows Forms before any MessageBox usage
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Check for PowerShell version 5.1 or later
if ($PSVersionTable.PSVersion -lt [Version]"5.1") {
    [System.Windows.Forms.MessageBox]::Show(
        "This script requires PowerShell version 5.1 or later. You are running version $($PSVersionTable.PSVersion).",
        "Requirement Not Met",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
}

# Check if the required module 'NetNat' is available
if (-not (Get-Module -ListAvailable -Name NetNat)) {
    [System.Windows.Forms.MessageBox]::Show(
        "The required module 'NetNat' is not available. Please install the module before running this script.",
        "Requirement Not Met",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
}

# Check if the script is running with administrator privileges
$currentUser = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentUser.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    [System.Windows.Forms.MessageBox]::Show(
        "This script must be run as an Administrator.",
        "Administrator Privileges Required",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
}

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
# SHARED HELPERS
# =========================

function Show-AppMessage {
    param (
        [string]$message,
        [string]$title = "NAT Manager",
        [System.Windows.Forms.MessageBoxIcon]$icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )

    [System.Windows.Forms.MessageBox]::Show(
        $message,
        $title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $icon
    ) | Out-Null
}

function Confirm-AppAction {
    param (
        [string]$message,
        [string]$title = "Confirm"
    )

    return ([System.Windows.Forms.MessageBox]::Show(
        $message,
        $title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    ) -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Test-ValidIPAddress {
    param ([string]$value)

    $ipAddress = $null
    return [System.Net.IPAddress]::TryParse($value, [ref]$ipAddress)
}

function Test-ValidPort {
    param ([string]$value)

    $port = 0
    return ([int]::TryParse($value, [ref]$port) -and $port -ge 1 -and $port -le 65535)
}

function Test-ValidCidrPrefix {
    param (
        [string]$value,
        [switch]$AllowEmpty
    )

    if ([string]::IsNullOrWhiteSpace($value)) {
        return $AllowEmpty.IsPresent
    }

    $parts = $value.Trim().Split("/")
    if ($parts.Count -ne 2) {
        return $false
    }

    $prefixLength = 0
    if (-not (Test-ValidIPAddress $parts[0])) {
        return $false
    }
    if (-not [int]::TryParse($parts[1], [ref]$prefixLength)) {
        return $false
    }

    $isIPv6 = $parts[0].Contains(":")
    if ($isIPv6) {
        return ($prefixLength -ge 0 -and $prefixLength -le 128)
    }

    return ($prefixLength -ge 0 -and $prefixLength -le 32)
}

function ConvertTo-DataTable {
    param ([object[]]$items)

    $dataTable = New-Object System.Data.DataTable
    if ($items.Count -eq 0) {
        return ,$dataTable
    }

    $items[0].PSObject.Properties.Name | ForEach-Object { [void]$dataTable.Columns.Add($_) }
    foreach ($item in $items) {
        $row = $dataTable.NewRow()
        $item.PSObject.Properties | ForEach-Object { $row.($_.Name) = $_.Value }
        [void]$dataTable.Rows.Add($row)
    }

    return ,$dataTable
}

function Set-GridStyle {
    param ([System.Windows.Forms.DataGridView]$grid)

    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.AllowUserToAddRows = $false
    $grid.ReadOnly = $true
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $false
    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(225,225,225)
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
    $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::WhiteSmoke
    $grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::LightSteelBlue
    $grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
}

function Test-StaticMappingInput {
    param (
        [string]$natName,
        [string]$externalIP,
        [string]$externalPort,
        [string]$internalIP,
        [string]$internalPort,
        [string]$protocol,
        [string]$ignoreMappingId
    )

    $errors = New-Object System.Collections.Generic.List[string]

    if ([string]::IsNullOrWhiteSpace($natName)) {
        $errors.Add("NAT Name is required.")
    }
    elseif (-not (Get-NetNat -Name $natName -ErrorAction SilentlyContinue)) {
        $errors.Add("NAT Name '$natName' does not exist.")
    }

    if (-not (Test-ValidIPAddress $externalIP)) {
        $errors.Add("External IP Address is invalid.")
    }
    if (-not (Test-ValidIPAddress $internalIP)) {
        $errors.Add("Internal IP Address is invalid.")
    }
    if (-not (Test-ValidPort $externalPort)) {
        $errors.Add("External Port must be between 1 and 65535.")
    }
    if (-not (Test-ValidPort $internalPort)) {
        $errors.Add("Internal Port must be between 1 and 65535.")
    }
    if ($protocol -notin @("TCP", "UDP")) {
        $errors.Add("Protocol must be TCP or UDP.")
    }

    if ($errors.Count -eq 0) {
        $existingMappings = @(Get-NetNatStaticMapping -ErrorAction SilentlyContinue | Where-Object {
            $_.NatName -eq $natName -and
            "$($_.Protocol)".ToUpperInvariant() -eq $protocol -and
            "$($_.ExternalIPAddress)" -eq $externalIP -and
            [int]$_.ExternalPort -eq [int]$externalPort -and
            "$($_.StaticMappingID)" -ne "$ignoreMappingId"
        })

        if ($existingMappings.Count -gt 0) {
            $errors.Add("A mapping for $natName $protocol $externalIP`:$externalPort already exists.")
        }
    }

    return $errors
}

function Test-NatNetworkInput {
    param (
        [string]$name,
        [string]$internalPrefix,
        [string]$externalPrefix,
        [string]$originalName
    )

    $errors = New-Object System.Collections.Generic.List[string]

    if ([string]::IsNullOrWhiteSpace($name)) {
        $errors.Add("Name is required.")
    }
    elseif ($name -ne $originalName -and (Get-NetNat -Name $name -ErrorAction SilentlyContinue)) {
        $errors.Add("A NAT network named '$name' already exists.")
    }

    if (-not (Test-ValidCidrPrefix $internalPrefix)) {
        $errors.Add("Internal IP Interface Address Prefix must be a valid CIDR prefix.")
    }
    if (-not (Test-ValidCidrPrefix $externalPrefix -AllowEmpty)) {
        $errors.Add("External IP Interface Address Prefix must be empty or a valid CIDR prefix.")
    }

    return $errors
}

# =========================
# MAIN FORM (Static Mappings)
# =========================

$form = New-Object System.Windows.Forms.Form
$form.Text = "NAT Static Mapping Manager"
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.Size = New-Object System.Drawing.Size(840, 520)  # Slightly taller to fit help text
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

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
Set-GridStyle $dataGridView
$groupBox.Controls.Add($dataGridView)

# Panel for main form buttons
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(10, 310)
$buttonPanel.Size = New-Object System.Drawing.Size(800, 50)
$form.Controls.Add($buttonPanel)

# Buttons: Add/Edit/Delete/Close
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

# NAT Networks button
$natNetworksButton = New-Object System.Windows.Forms.Button
$natNetworksButton.Text = "NAT Network"
$natNetworksButton.Size = New-Object System.Drawing.Size(100, 30)
$natNetworksButton.Location = New-Object System.Drawing.Point(270, 10)

# Refresh button
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh"
$refreshButton.Size = New-Object System.Drawing.Size(75, 30)
$refreshButton.Location = New-Object System.Drawing.Point(380, 10)

# Close button
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "Close"
$closeButton.Size = New-Object System.Drawing.Size(75, 30)
$closeButton.Location = New-Object System.Drawing.Point(470, 10)
$closeButton.Add_Click({ $form.Close() })

$buttonPanel.Controls.Add($addButton)
$buttonPanel.Controls.Add($editButton)
$buttonPanel.Controls.Add($deleteButton)
$buttonPanel.Controls.Add($natNetworksButton)
$buttonPanel.Controls.Add($refreshButton)
$buttonPanel.Controls.Add($closeButton)

# ToolTips
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($addButton, "Add a new NAT mapping")
$toolTip.SetToolTip($editButton, "Edit the selected NAT mapping")
$toolTip.SetToolTip($deleteButton, "Delete the selected NAT mapping")
$toolTip.SetToolTip($natNetworksButton, "Manage NAT Networks")
$toolTip.SetToolTip($refreshButton, "Refresh NAT mappings")
$toolTip.SetToolTip($closeButton, "Close this window")

# --- Add a label to display info about NAT fields, at bottom ---
$infoLabel = New-Object System.Windows.Forms.Label
$infoLabel.AutoSize = $false
$infoLabel.Width = 800
$infoLabel.Height = 90
$infoLabel.Location = New-Object System.Drawing.Point(10, 370)
$infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$infoLabel.Text = @"
NAT Name: The name of your NAT Network. (Use the 'NAT Network' button above to manage networks.)
External IP Address: By default 0.0.0.0, which covers all external IPs.
External Port: The external port used for NAT.
Internal Port: The internal port used for NAT.
Protocol: TCP or UDP.
"@
$form.Controls.Add($infoLabel)

# --- Input Dialog for Static Mappings ---
function Show-InputDialog {
    param (
        [string]$title,
        [hashtable]$defaults = @{ },
        [string]$ignoreMappingId = $null
    )
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = $title
    $dialog.Size = New-Object System.Drawing.Size(340, 350)
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

        if ($field -eq "NatName") {
            $comboBox = New-Object System.Windows.Forms.ComboBox
            $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            $comboBox.Location = New-Object System.Drawing.Point(140, $y)
            $comboBox.Size = New-Object System.Drawing.Size(160, 20)
            @(Get-NetNat -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name) | ForEach-Object {
                [void]$comboBox.Items.Add($_)
            }
            if ($defaults[$field] -and $comboBox.Items.Contains($defaults[$field])) {
                $comboBox.SelectedItem = $defaults[$field]
            }
            elseif ($comboBox.Items.Count -gt 0) {
                $comboBox.SelectedIndex = 0
            }
            $dialog.Controls.Add($comboBox)
            $controls[$field] = $comboBox
        }
        elseif ($field -eq "Protocol") {
            $comboBox = New-Object System.Windows.Forms.ComboBox
            $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            $comboBox.Location = New-Object System.Drawing.Point(140, $y)
            $comboBox.Size = New-Object System.Drawing.Size(160, 20)
            [void]$comboBox.Items.Add("TCP")
            [void]$comboBox.Items.Add("UDP")
            $protocolDefault = if ($defaults[$field]) { "$($defaults[$field])".ToUpperInvariant() } else { "TCP" }
            if ($comboBox.Items.Contains($protocolDefault)) {
                $comboBox.SelectedItem = $protocolDefault
            }
            else {
                $comboBox.SelectedIndex = 0
            }
            $dialog.Controls.Add($comboBox)
            $controls[$field] = $comboBox
        }
        else {
            $textBox = New-Object System.Windows.Forms.TextBox
            $textBox.Text = $defaults[$field] -as [string]
            $textBox.Location = New-Object System.Drawing.Point(140, $y)
            $textBox.Size = New-Object System.Drawing.Size(160, 20)
            $dialog.Controls.Add($textBox)
            $controls[$field] = $textBox
        }
        $y += 30
    }

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(80, $y)
    $dialog.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(170, $y)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialog.Controls.Add($cancelButton)

    $dialog.AcceptButton = $okButton
    $dialog.CancelButton = $cancelButton
    $dialog.Tag = $null

    $okButton.Add_Click({
        $natName = $controls["NatName"].Text.Trim()
        $externalIP = $controls["ExternalIPAddress"].Text.Trim()
        $externalPort = $controls["ExternalPort"].Text.Trim()
        $internalIP = $controls["InternalIPAddress"].Text.Trim()
        $internalPort = $controls["InternalPort"].Text.Trim()
        $protocol = $controls["Protocol"].Text.Trim().ToUpperInvariant()

        $errors = Test-StaticMappingInput `
            -natName $natName `
            -externalIP $externalIP `
            -externalPort $externalPort `
            -internalIP $internalIP `
            -internalPort $internalPort `
            -protocol $protocol `
            -ignoreMappingId $ignoreMappingId

        if ($errors.Count -gt 0) {
            Show-AppMessage ($errors -join [Environment]::NewLine) "Invalid NAT Mapping" ([System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $dialog.Tag = @(
            $natName,
            $externalIP,
            [int]$externalPort,
            $internalIP,
            [int]$internalPort,
            $protocol
        )
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dialog.Close()
    })

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.Tag
    }
    else {
        return $null
    }
}

# --- Load Static NAT Mappings ---
function Load-NatMappings {
    $dataGridView.DataSource = $null
    try {
        $mappings = @(Get-NetNatStaticMapping -ErrorAction Stop | ForEach-Object {
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
        })
        $dataGridView.DataSource = (ConvertTo-DataTable $mappings)
    }
    catch {
        Show-AppMessage "Error loading NAT mappings.`n`n$($_.Exception.Message)" "Error" ([System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        Update-MappingButtons
    }
}

function Update-MappingButtons {
    $hasRows = ($dataGridView.Rows.Count -gt 0)
    $hasSelection = ($hasRows -and $null -ne $dataGridView.CurrentRow)
    $editButton.Enabled = $hasSelection
    $deleteButton.Enabled = $hasSelection
}

# --- Static Mapping Button Click Handlers ---
$addButton.Add_Click({
    $result = Show-InputDialog -title "Add New NAT Mapping"
    if ($result) {
        $natName, $externalIP, $externalPort, $internalIP, $internalPort, $protocol = $result
        try {
            Add-NetNatStaticMapping -NatName $natName -ExternalIPAddress $externalIP -ExternalPort $externalPort `
                -InternalIPAddress $internalIP -InternalPort $internalPort -Protocol $protocol -ErrorAction Stop
            Show-AppMessage "New mapping added successfully.`n`n$natName $protocol $externalIP`:$externalPort -> $internalIP`:$internalPort" "Info"
            Load-NatMappings
        } catch {
            Show-AppMessage "Error adding mapping for $natName $protocol $externalIP`:$externalPort -> $internalIP`:$internalPort`n`n$($_.Exception.Message)" "Error" ([System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

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
        $id = $selectedRow.Cells["ID"].Value
        $result = Show-InputDialog -title "Edit NAT Mapping" -defaults $defaults -ignoreMappingId $id
        if ($result) {
            $natName, $externalIP, $externalPort, $internalIP, $internalPort, $protocol = $result
            try {
                Remove-NetNatStaticMapping -StaticMappingID $id -Confirm:$false -ErrorAction Stop
                try {
                    Add-NetNatStaticMapping -NatName $natName -ExternalIPAddress $externalIP -ExternalPort $externalPort `
                        -InternalIPAddress $internalIP -InternalPort $internalPort -Protocol $protocol -ErrorAction Stop
                }
                catch {
                    $addError = $_.Exception.Message
                    try {
                        Add-NetNatStaticMapping `
                            -NatName $defaults["NatName"] `
                            -ExternalIPAddress $defaults["ExternalIPAddress"] `
                            -ExternalPort ([int]$defaults["ExternalPort"]) `
                            -InternalIPAddress $defaults["InternalIPAddress"] `
                            -InternalPort ([int]$defaults["InternalPort"]) `
                            -Protocol "$($defaults["Protocol"])" `
                            -ErrorAction Stop
                        Show-AppMessage "Could not apply the edited mapping, so the original mapping was restored.`n`nTarget: $natName $protocol $externalIP`:$externalPort -> $internalIP`:$internalPort`n`nError: $addError" "Edit Failed - Restored" ([System.Windows.Forms.MessageBoxIcon]::Warning)
                    }
                    catch {
                        Show-AppMessage "Could not apply the edited mapping, and the rollback also failed.`n`nTarget: $natName $protocol $externalIP`:$externalPort -> $internalIP`:$internalPort`n`nEdit error: $addError`nRollback error: $($_.Exception.Message)" "Edit Failed" ([System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                    Load-NatMappings
                    return
                }
                Show-AppMessage "Mapping updated successfully.`n`n$natName $protocol $externalIP`:$externalPort -> $internalIP`:$internalPort" "Info"
                Load-NatMappings
            } catch {
                Show-AppMessage "Error editing mapping ID $id.`n`n$($_.Exception.Message)" "Error" ([System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } else {
        Show-AppMessage "No mapping selected." "Error" ([System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$deleteButton.Add_Click({
    $selectedRow = $dataGridView.CurrentRow
    if ($selectedRow) {
        $id = $selectedRow.Cells["ID"].Value
        $natName = $selectedRow.Cells["NatName"].Value
        $externalIP = $selectedRow.Cells["ExternalIPAddress"].Value
        $externalPort = $selectedRow.Cells["ExternalPort"].Value
        $internalIP = $selectedRow.Cells["InternalIPAddress"].Value
        $internalPort = $selectedRow.Cells["InternalPort"].Value
        $protocol = $selectedRow.Cells["Protocol"].Value

        if (-not (Confirm-AppAction "Delete mapping $natName $protocol $externalIP`:$externalPort -> $internalIP`:$internalPort?" "Delete NAT Mapping")) {
            return
        }

        try {
            Remove-NetNatStaticMapping -StaticMappingID $id -Confirm:$false -ErrorAction Stop
            Show-AppMessage "Mapping deleted successfully.`n`n$natName $protocol $externalIP`:$externalPort -> $internalIP`:$internalPort" "Info"
            Load-NatMappings
        } catch {
            Show-AppMessage "Error deleting mapping $natName $protocol $externalIP`:$externalPort -> $internalIP`:$internalPort`n`n$($_.Exception.Message)" "Error" ([System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        Show-AppMessage "No mapping selected." "Error" ([System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$refreshButton.Add_Click({
    Load-NatMappings
})

$dataGridView.Add_SelectionChanged({
    Update-MappingButtons
})

$dataGridView.Add_DataBindingComplete({
    Update-MappingButtons
})

# Load initial static mappings
Load-NatMappings

# =========================
# NAT NETWORKS MANAGEMENT FORM
# =========================

function Show-NatNetworksForm {

    # Check if WinNAT service is installed. If not, show error and return.
    # If installed but not running, ask the user whether to start it.
    try {
        $winnatService = Get-Service -Name 'WinNAT' -ErrorAction Stop
        if ($winnatService.Status -ne 'Running') {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "WinNAT Service not running. Do you want to start it to proceed?",
                "WinNAT Service",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Service -Name 'WinNAT'
                Start-Sleep -Seconds 2
            }
            else {
                return  # user chose No => do not open form
            }
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "The WinNAT service is not installed or cannot be accessed.
The NAT Networks Manager cannot continue without WinNAT.",
            "WinNAT Service Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # Create NAT networks form
    $natForm = New-Object System.Windows.Forms.Form
    $natForm.Text = "NAT Networks Manager"
    $natForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $natForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $natForm.MaximizeBox = $false
    $natForm.MinimizeBox = $false
    $natForm.Size = New-Object System.Drawing.Size(840, 500)
    $natForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # GroupBox for NAT networks
    $natGroupBox = New-Object System.Windows.Forms.GroupBox
    $natGroupBox.Text = "NAT Networks"
    $natGroupBox.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $natGroupBox.Location = New-Object System.Drawing.Point(10, 10)
    $natGroupBox.Size = New-Object System.Drawing.Size(800, 350)
    $natForm.Controls.Add($natGroupBox)
    
    # DataGridView for NAT networks
    $natDataGrid = New-Object System.Windows.Forms.DataGridView
    $natDataGrid.Size = New-Object System.Drawing.Size(760, 310)
    $natDataGrid.Location = New-Object System.Drawing.Point(15, 25)
    Set-GridStyle $natDataGrid
    $natGroupBox.Controls.Add($natDataGrid)
    
    # Panel for NAT Networks form buttons
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
    
    $natToolTip = New-Object System.Windows.Forms.ToolTip
    $natToolTip.SetToolTip($natAddButton, "Add a new NAT network")
    $natToolTip.SetToolTip($natEditButton, "Edit the selected NAT network")
    $natToolTip.SetToolTip($natDeleteButton, "Delete the selected NAT network")
    $natToolTip.SetToolTip($natRefreshButton, "Refresh the NAT networks list")
    $natToolTip.SetToolTip($natCloseButton, "Close this window")

    # --- Custom Input Dialog for NAT Networks (Wider) ---
    function Show-NatNetworkInputDialog {
        param (
            [string]$title,
            [hashtable]$defaults = @{ },
            [string]$originalName = $null
        )
        $dialog = New-Object System.Windows.Forms.Form
        $dialog.Text = $title
        $dialog.Size = New-Object System.Drawing.Size(420, 250)
        $dialog.StartPosition = "CenterParent"
        $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $dialog.MaximizeBox = $false
        $dialog.MinimizeBox = $false
        $dialog.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
        $controls = @{}
        $y = 10
        $fields = @(
            "Name",
            "InternalIPInterfaceAddressPrefix",
            "ExternalIPInterfaceAddressPrefix"
        )
    
        foreach ($field in $fields) {
            $label = New-Object System.Windows.Forms.Label
            $label.Text = $field
            $label.Location = New-Object System.Drawing.Point(10, $y)
            $label.Size = New-Object System.Drawing.Size(200, 20)
            $dialog.Controls.Add($label)
    
            $textBox = New-Object System.Windows.Forms.TextBox
            $textBox.Text = $defaults[$field] -as [string]
            $textBox.Location = New-Object System.Drawing.Point(220, $y)
            $textBox.Size = New-Object System.Drawing.Size(180, 20)
            $dialog.Controls.Add($textBox)
    
            $controls[$field] = $textBox
            $y += 30
        }
    
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(100, $y)
        $dialog.Controls.Add($okButton)
    
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(200, $y)
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.Controls.Add($cancelButton)
    
        $dialog.AcceptButton = $okButton
        $dialog.CancelButton = $cancelButton
        $dialog.Tag = $null

        $okButton.Add_Click({
            $name = $controls["Name"].Text.Trim()
            $internalPrefix = $controls["InternalIPInterfaceAddressPrefix"].Text.Trim()
            $externalPrefix = $controls["ExternalIPInterfaceAddressPrefix"].Text.Trim()

            $errors = Test-NatNetworkInput `
                -name $name `
                -internalPrefix $internalPrefix `
                -externalPrefix $externalPrefix `
                -originalName $originalName

            if ($errors.Count -gt 0) {
                Show-AppMessage ($errors -join [Environment]::NewLine) "Invalid NAT Network" ([System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            $dialog.Tag = @($name, $internalPrefix, $externalPrefix)
            $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dialog.Close()
        })
    
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.Tag
        }
        else {
            return $null
        }
    }

    # --- Load NAT Networks ---
    function Load-NatNetworks {
        $natDataGrid.DataSource = $null
        try {
            $natList = @(Get-NetNat -ErrorAction Stop | ForEach-Object {
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
            $natDataGrid.DataSource = (ConvertTo-DataTable $natList)
        }
        catch {
            Show-AppMessage "Error loading NAT networks.`n`n$($_.Exception.Message)" "Error" ([System.Windows.Forms.MessageBoxIcon]::Error)
        }
        finally {
            Update-NatNetworkButtons
        }
    }

    function Update-NatNetworkButtons {
        $hasRows = ($natDataGrid.Rows.Count -gt 0)
        $hasSelection = ($hasRows -and $null -ne $natDataGrid.CurrentRow)
        $natEditButton.Enabled = $hasSelection
        $natDeleteButton.Enabled = $hasSelection
    }
    
    # Add NAT Network
    $natAddButton.Add_Click({
        # ICS/HNS check
        $icsService = Get-Service -Name SharedAccess -ErrorAction SilentlyContinue
        $hnsService = Get-Service -Name hns -ErrorAction SilentlyContinue

        $servicesRunning = @()
        if (($icsService -and $icsService.Status -eq 'Running') -or
            ($hnsService -and $hnsService.Status -eq 'Running')) {
            if ($icsService -and $icsService.Status -eq 'Running') {
                $servicesRunning += "Internet Connection Sharing (SharedAccess)"
            }
            if ($hnsService -and $hnsService.Status -eq 'Running') {
                $servicesRunning += "Host Network Service (hns)"
            }
        }

        if ($servicesRunning.Count -gt 0) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "The following services are running and may manage or conflict with NAT configuration:`n`n$($servicesRunning -join "`n")`n`nNo services will be stopped or disabled. Continue anyway?",
                "ICS/HNS Detected",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
                return
            }
        }

        # Show the Add NAT Network dialog
        $defaults = @{
            "Name"                             = "EXAMPLE"
            "InternalIPInterfaceAddressPrefix" = "192.168.0.0/24"
            "ExternalIPInterfaceAddressPrefix" = ""
        }
        $input = Show-NatNetworkInputDialog -title "Add New NAT Network" -defaults $defaults
        if ($input) {
            try {
                $params = @{
                    Name                             = $input[0]
                    InternalIPInterfaceAddressPrefix = $input[1]
                }
                if ($input[2]) {
                    $params["ExternalIPInterfaceAddressPrefix"] = $input[2]
                }
                New-NetNat @params -ErrorAction Stop
                Show-AppMessage "New NAT network added successfully.`n`nName: $($input[0])`nInternal Prefix: $($input[1])" "Info"
                Load-NatNetworks
            } catch {
                Show-AppMessage "Error adding NAT network '$($input[0])'.`n`n$($_.Exception.Message)" "Error" ([System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    
    # Edit NAT Network
    $natEditButton.Add_Click({
        $selectedRow = $natDataGrid.CurrentRow
        if ($selectedRow) {
            $defaults = @{
                "Name" = $selectedRow.Cells["Name"].Value
                "InternalIPInterfaceAddressPrefix" = $selectedRow.Cells["InternalIPInterfaceAddressPrefix"].Value
                "ExternalIPInterfaceAddressPrefix" = $selectedRow.Cells["ExternalIPInterfaceAddressPrefix"].Value
            }
            $input = Show-NatNetworkInputDialog -title "Edit NAT Network" -defaults $defaults -originalName $selectedRow.Cells["Name"].Value
            if ($input) {
                $netName = $selectedRow.Cells["Name"].Value
                if (-not (Confirm-AppAction "Editing NAT network '$netName' requires recreating it. Existing static mappings for this NAT network may be removed by Windows during this operation.`n`nContinue?" "Edit NAT Network")) {
                    return
                }

                try {
                    Remove-NetNat -Name $netName -Confirm:$false -ErrorAction Stop

                    $params = @{
                        Name = $input[0]
                        InternalIPInterfaceAddressPrefix = $input[1]
                    }
                    if ($input[2]) {
                        $params["ExternalIPInterfaceAddressPrefix"] = $input[2]
                    }
                    try {
                        New-NetNat @params -ErrorAction Stop
                    }
                    catch {
                        $newError = $_.Exception.Message
                        try {
                            $rollbackParams = @{
                                Name = $defaults["Name"]
                                InternalIPInterfaceAddressPrefix = $defaults["InternalIPInterfaceAddressPrefix"]
                            }
                            if ($defaults["ExternalIPInterfaceAddressPrefix"]) {
                                $rollbackParams["ExternalIPInterfaceAddressPrefix"] = $defaults["ExternalIPInterfaceAddressPrefix"]
                            }
                            New-NetNat @rollbackParams -ErrorAction Stop
                            Show-AppMessage "Could not apply the edited NAT network, so the original NAT network was restored.`n`nTarget: $($input[0])`nError: $newError" "Edit Failed - Restored" ([System.Windows.Forms.MessageBoxIcon]::Warning)
                        }
                        catch {
                            Show-AppMessage "Could not apply the edited NAT network, and the rollback also failed.`n`nTarget: $($input[0])`nEdit error: $newError`nRollback error: $($_.Exception.Message)" "Edit Failed" ([System.Windows.Forms.MessageBoxIcon]::Error)
                        }
                        Load-NatNetworks
                        return
                    }

                    Show-AppMessage "NAT network updated successfully.`n`nName: $($input[0])`nInternal Prefix: $($input[1])" "Info"
                    Load-NatNetworks
                } catch {
                    Show-AppMessage "Error editing NAT network '$netName'.`n`n$($_.Exception.Message)" "Error" ([System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        } else {
            Show-AppMessage "No NAT network selected." "Error" ([System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    
    # Delete NAT Network
    $natDeleteButton.Add_Click({
        $selectedRow = $natDataGrid.CurrentRow
        if ($selectedRow) {
            $netName = $selectedRow.Cells["Name"].Value
            if ($netName) {
                if (-not (Confirm-AppAction "Delete NAT network '$netName'?`n`nStatic mappings attached to this NAT network may be removed by Windows." "Delete NAT Network")) {
                    return
                }

                try {
                    Remove-NetNat -Name $netName -Confirm:$false -ErrorAction Stop
                    Show-AppMessage "NAT network deleted successfully.`n`nName: $netName" "Info"
                    Load-NatNetworks
                } catch {
                    Show-AppMessage "Error deleting NAT network '$netName'.`n`n$($_.Exception.Message)" "Error" ([System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        } else {
            Show-AppMessage "No NAT network selected." "Error" ([System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # Refresh NAT Networks
    $natRefreshButton.Add_Click({
        Load-NatNetworks
    })

    $natDataGrid.Add_SelectionChanged({
        Update-NatNetworkButtons
    })

    $natDataGrid.Add_DataBindingComplete({
        Update-NatNetworkButtons
    })

    # Load initial NAT networks data and show the NAT Networks form
    Load-NatNetworks
    [void]$natForm.ShowDialog()
}

# NAT Networks Button on the Main Form
$natNetworksButton.Add_Click({
    Show-NatNetworksForm
    Load-NatMappings
})

# Show the main form
[void]$form.ShowDialog()
