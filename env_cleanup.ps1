<#
    Improved PowerShell Script for Cleaning Duplicate and Invalid Environment Variables

    Features:
      - Backs up current User and System environment variables (to a timestamped backup file)
      - Supports a dry-run mode (via -WhatIfChanges switch) to preview changes without making them
      - Uses verbose logging and robust error handling
      - Checks for administrative privileges when modifying system variables
      - Cleans semicolon-separated lists (like PATH) by:
          • Removing duplicate entries while preserving the first occurrence
          • Trimming extraneous spaces and semicolons
          • Removing entries that are empty or point to non-existent directories
      - For non-list variables, if the value “looks like” a directory but does not exist, it is removed
      - Checks for size errors (e.g. if the cleaned PATH exceeds Windows’ limit of 32767 characters)
      - Compares environment variable names between User and System scopes for additional insight
#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [switch]$WhatIfChanges  # Dry-run mode: preview changes without updating
)

# Backup function – save current environment variables to a backup file
function Backup-EnvironmentVariables {
    param(
        [string]$Scope  # "User" or "System"
    )
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupFile = "$env:USERPROFILE\EnvBackup_$Scope`_$timestamp.txt"
    try {
        if ($Scope -eq "User") {
            $regPath = "HKCU:\Environment"
        }
        elseif ($Scope -eq "System") {
            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
        }
        else {
            throw "Invalid scope provided: $Scope"
        }
        # Export properties (excluding PS metadata)
        $envData = Get-ItemProperty -Path $regPath | Select-Object * -ExcludeProperty PSPath, PSParentPath, PSChildName, PSDrive, PSProvider
        $envData | Out-File -FilePath $backupFile -Encoding UTF8
        Write-Verbose "Backup of $Scope environment variables saved to $backupFile"
    }
    catch {
        Write-Warning "Failed to backup $Scope environment variables: $_"
    }
}

# Function to clean a single environment variable value.
function Clean-EnvVarValue {
    param(
        [string]$varValue
    )
    if (-not $varValue) { return $null }

    # If value contains semicolons, treat it as a list (e.g., PATH)
    if ($varValue -match ';') {
        $entries = $varValue -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        $uniqueEntries = @()
        foreach ($entry in $entries) {
            if ($uniqueEntries -notcontains $entry) {
                $uniqueEntries += $entry
            }
        }
        # Only keep entries that exist (using Test-Path) to avoid invalid paths
        $validEntries = $uniqueEntries | Where-Object { Test-Path $_ }
        if ($validEntries.Count -gt 0) {
            return $validEntries -join ';'
        }
        else {
            return $null
        }
    }
    else {
        # For single-value entries: trim and if it “looks like” a path, verify its existence.
        $trimValue = $varValue.Trim()
        if ($trimValue -eq "") { return $null }
        if ($trimValue -match '[\\:]') {
            if (-not (Test-Path $trimValue)) { return $null }
        }
        return $trimValue
    }
}

# Function to process environment variables for a given registry path and scope.
function Process-EnvVars {
    param(
        [string]$regPath,
        [string]$Scope  # "User" or "System"
    )
    Write-Output "Processing $Scope environment variables from $regPath..."
    try {
        $envProps = Get-Item -Path $regPath |
            Get-ItemProperty |
            Get-Member -MemberType NoteProperty |
            Select-Object -ExpandProperty Name |
            Where-Object { $_ -notmatch '^PS' }
    }
    catch {
        Write-Error "Failed to access registry path $regPath. Ensure you have the required permissions."
        return
    }

    foreach ($varName in $envProps) {
        try {
            $oldValue = (Get-ItemProperty -Path $regPath -Name $varName).$varName
            $newValue = Clean-EnvVarValue -varValue $oldValue

            # For PATH variable, trim extraneous semicolons and check size constraints.
            if ($varName -ieq "Path" -and $newValue) {
                $newValue = $newValue.Trim(';')
                $maxLength = 32767
                if ($newValue.Length -gt $maxLength) {
                    Write-Warning "The cleaned PATH for $Scope variable '$varName' is $($newValue.Length) characters long, exceeding the maximum allowed ($maxLength). Skipping update. Please manually trim non-critical entries."
                    continue
                }
            }

            if ($newValue -ne $oldValue) {
                if ($WhatIfChanges -or $PSCmdlet.ShouldProcess("$Scope variable '$varName'", "Update")) {
                    if ($newValue) {
                        if (-not $WhatIfChanges) {
                            Set-ItemProperty -Path $regPath -Name $varName -Value $newValue -ErrorAction Stop
                        }
                        Write-Output "Updated $Scope variable '$varName':"
                        Write-Output "   Old: $oldValue"
                        Write-Output "   New: $newValue"
                    }
                    else {
                        if ($WhatIfChanges -or $PSCmdlet.ShouldProcess("$Scope variable '$varName'", "Remove")) {
                            if (-not $WhatIfChanges) {
                                Remove-ItemProperty -Path $regPath -Name $varName -ErrorAction Stop
                            }
                            Write-Output "Removed $Scope variable '$varName' (empty or invalid after cleaning)."
                        }
                    }
                }
            }
        }
        catch {
            Write-Warning "Failed to process variable '$varName': $_"
        }
    }
}

# Function to check if running as administrator (required for System changes).
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Define registry paths.
$userEnvRegPath   = "HKCU:\Environment"
$systemEnvRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"

# Backup environment variables for both scopes.
Backup-EnvironmentVariables -Scope "User"
Backup-EnvironmentVariables -Scope "System"

# Process User environment variables.
Process-EnvVars -regPath $userEnvRegPath -Scope "User"

# Process System environment variables.
if (-not (Test-Administrator)) {
    Write-Warning "Not running as administrator. Skipping processing of System environment variables."
}
else {
    Process-EnvVars -regPath $systemEnvRegPath -Scope "System"
}

# Compare common variable names between User and System scopes.
try {
    $userVars   = Get-Item -Path $userEnvRegPath | Get-ItemProperty |
                  Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    $systemVars = Get-Item -Path $systemEnvRegPath | Get-ItemProperty |
                  Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    $commonVars = Compare-Object -ReferenceObject $userVars -DifferenceObject $systemVars -IncludeEqual -ExcludeDifferent |
                  Select-Object -ExpandProperty InputObject
    if ($commonVars) {
        Write-Output "The following environment variables exist in both User and System scopes:"
        $commonVars | ForEach-Object { Write-Output "  $_" }
    }
    else {
        Write-Output "No common environment variables found between User and System scopes."
    }
}
catch {
    Write-Warning "Failed to compare common environment variables: $_"
}

# Refresh environment variables so that changes are propagated.
function Refresh-Environment {
    $signature = @"
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SendMessageTimeout(IntPtr hWnd, int Msg, UIntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
"@
    try {
        Add-Type -MemberDefinition $signature -Name NativeMethods -Namespace Win32 -ErrorAction Stop
        [UIntPtr]$result = [UIntPtr]::Zero
        [Win32.NativeMethods]::SendMessageTimeout([IntPtr]::Zero, 0x001A, [UIntPtr]::Zero, "Environment", 0, 1000, [ref]$result) | Out-Null
        Write-Output "Environment variables refreshed."
    }
    catch {
        Write-Warning "Failed to refresh environment variables: $_"
    }
}
Refresh-Environment

Write-Output "✔️✔️✔️ Complete ✔️✔️✔️"
