# ===============================================================
# CROWDSTRIKE TOOL - HARDENED PRODUCTION VERSION
# ===============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -----------------------------
# SAFE POWERSHELL PATH
# -----------------------------
$PSExe = (Get-Command powershell.exe).Source

# -----------------------------
# AUTO ELEVATION (SAFE)
# -----------------------------
$identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)

if (-not $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Start-Process -FilePath ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -Verb RunAs
    exit
}

# -----------------------------
# TRUSTED HOSTS
# -----------------------------
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force -ErrorAction SilentlyContinue

# -----------------------------
# ENABLE WINRM FUNCTION
# -----------------------------
function Enable-WinRM {
    param(
        [string]$IP,
        [System.Management.Automation.PSCredential]$cred
    )
    
    try {
        $session = New-PSSession -ComputerName $IP -Credential $cred -ErrorAction Stop
        Invoke-Command -Session $session -ScriptBlock {
            Set-Service -Name WinRM -StartupType Automatic -ErrorAction SilentlyContinue
            Start-Service -Name WinRM -ErrorAction SilentlyContinue
            winrm quickconfig -quiet 2>$null
            Set-Item WSMan:\localhost\Service\AllowUnencrypted -Value $true -Force -ErrorAction SilentlyContinue
            Set-Item WSMan:\localhost\Service\Auth\Basic -Value $true -Force -ErrorAction SilentlyContinue
            Set-Item WSMan:\localhost\Client\AllowUnencrypted -Value $true -Force -ErrorAction SilentlyContinue
            netsh advfirewall firewall add rule name="WinRM-HTTP" dir=in action=allow protocol=tcp localport=5985 -ErrorAction SilentlyContinue
            netsh advfirewall firewall add rule name="WinRM-HTTPS" dir=in action=allow protocol=tcp localport=5986 -ErrorAction SilentlyContinue
        } -ErrorAction SilentlyContinue
        Remove-PSSession $session
        return $true
    } catch {
        return $false
    }
}

# -----------------------------
# TEST AND ENABLE WINRM FUNCTION
# -----------------------------
function Test-AndEnable-WinRM {
    param(
        [string]$IP,
        [System.Management.Automation.PSCredential]$cred
    )
    
    try {
        $result = Invoke-Command -ComputerName $IP -Credential $cred -ScriptBlock { $true } -ErrorAction Stop
        return $true
    } catch {
        try {
            $wmiSession = New-PSSession -ComputerName $IP -Credential $cred -AllowClobber -ErrorAction SilentlyContinue
            if ($wmiSession) {
                Invoke-Command -Session $wmiSession -ScriptBlock {
                    param($psexe)
                    Start-Process $psexe -ArgumentList "-ExecutionPolicy Bypass -Command winrm quickconfig -quiet" -Wait -WindowStyle Hidden
                    Set-Service -Name WinRM -StartupType Automatic
                    Start-Service -Name WinRM
                    netsh advfirewall firewall add rule name="WinRM-HTTP-In" dir=in action=allow protocol=tcp localport=5985 | Out-Null
                } -ArgumentList $PSExe -ErrorAction SilentlyContinue
                Remove-PSSession $wmiSession
                return $true
            }
        } catch {
            try {
                $wmi = [WMIClass]"\\$IP\root\cimv2:Win32_Process"
                $wmi.Create("cmd /c winrm quickconfig -q") | Out-Null
                $wmi.Create("cmd /c netsh firewall add portopening TCP 5985 'WinRM HTTP'") | Out-Null
                $wmi.Create("cmd /c sc config WinRM start= auto") | Out-Null
                $wmi.Create("cmd /c net start WinRM") | Out-Null
                return $true
            } catch {
                return $false
            }
        }
    }
    return $false
}

# -----------------------------
# UI
# -----------------------------
$form = New-Object Windows.Forms.Form
$form.Text = "CrowdStrike Deployment Tool"
$form.Size = "520,300"
$form.StartPosition = "CenterScreen"

$btnCSV = New-Object Windows.Forms.Button
$btnCSV.Text = "Select CSV"
$btnCSV.Location = "50,30"
$form.Controls.Add($btnCSV)

$btnFolder = New-Object Windows.Forms.Button
$btnFolder.Text = "Select CS Folder"
$btnFolder.Location = "200,30"
$form.Controls.Add($btnFolder)

$btnStart = New-Object Windows.Forms.Button
$btnStart.Text = "Start Deployment"
$btnStart.Location = "150,80"
$form.Controls.Add($btnStart)

$progress = New-Object Windows.Forms.ProgressBar
$progress.Location = "50,130"
$progress.Size = "400,20"
$form.Controls.Add($progress)

$status = New-Object Windows.Forms.Label
$status.Location = "50,170"
$status.Size = "400,40"
$form.Controls.Add($status)

$global:CsvPath = $null
$global:CSFolder = $null

# -----------------------------
# FILE SELECT
# -----------------------------
$btnCSV.Add_Click({
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Filter = "csv files (*.csv)|*.csv"
    if ($dlg.ShowDialog()) {
        $global:CsvPath = $dlg.FileName
        $status.Text = "CSV Selected: $(Split-Path $global:CsvPath -Leaf)"
    }
})

$btnFolder.Add_Click({
    $dlg = New-Object Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog()) {
        $global:CSFolder = $dlg.SelectedPath
        $status.Text = "Folder Selected: $(Split-Path $global:CSFolder -Leaf)"
    }
})

# -----------------------------
# MAIN EXECUTION
# -----------------------------
$btnStart.Add_Click({

    if (-not $global:CsvPath -or -not $global:CSFolder) {
        [Windows.Forms.MessageBox]::Show("Select CSV and Folder")
        return
    }

    if (-not (Test-Path "$global:CSFolder\WindowsSensor.exe")) {
        [Windows.Forms.MessageBox]::Show("WindowsSensor.exe not found in selected folder")
        return
    }

    $Machines = Import-Csv $global:CsvPath
    $progress.Maximum = $Machines.Count

    $OutputDir = "C:\CS_Report_$(Get-Date -Format yyyyMMdd_HHmmss)"
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

    foreach ($m in $Machines) {

        $IP = $m.IPAddress
        $User = $m.Username
        $Pass = $m.Password

        $status.Text = "Processing $IP"
        $form.Refresh()

        try {
            $sec = ConvertTo-SecureString $Pass -AsPlainText -Force
            $cred = New-Object pscredential($User,$sec)

            $connected = $false
            for ($i=1; $i -le 3; $i++) {
                try {
                    Invoke-Command -ComputerName $IP -Credential $cred -ScriptBlock { hostname } -ErrorAction Stop
                    $connected = $true
                    break
                } catch {
                    if ($i -eq 1) {
                        $status.Text = "Enabling WinRM on $IP..."
                        $form.Refresh()
                        Test-AndEnable-WinRM -IP $IP -cred $cred | Out-Null
                    }
                    Start-Sleep 3
                }
            }

            if (-not $connected) {
                "$IP FAILED - CONNECTION" | Out-File "$OutputDir\result.txt" -Append
                continue
            }

            Invoke-Command -ComputerName $IP -Credential $cred -ScriptBlock {
                New-Item -ItemType Directory -Path C:\TempCS -Force | Out-Null
            }

            $session = New-PSSession -ComputerName $IP -Credential $cred -ErrorAction SilentlyContinue

            if ($session) {
                $csFolder = $using:CSFolder
                if (Test-Path "$csFolder\WindowsSensor.exe") {
                    Copy-Item "$csFolder\WindowsSensor.exe" `
                        -Destination "C:\TempCS\WindowsSensor.exe" `
                        -ToSession $session
                }
                Remove-PSSession $session
            }

            Invoke-Command -ComputerName $IP -Credential $cred -ScriptBlock {
                Start-Process "C:\TempCS\WindowsSensor.exe" -ArgumentList "/install /quiet" -Wait
            }

            "$IP SUCCESS" | Out-File "$OutputDir\result.txt" -Append

        } catch {
            "$IP FAILED - $($_.Exception.Message)" | Out-File "$OutputDir\result.txt" -Append
        }

        $progress.Value++
    }

    $status.Text = "Completed"
    Start-Process $OutputDir
})

$form.ShowDialog()
