<#
.SYNOPSIS
    Anonymizes sensitive data in RVTools VMware infrastructure exports.

.DESCRIPTION
    This script replaces identifying information (VM names, hostnames, IPs, etc.)
    with generic anonymized values while maintaining data relationships across sheets.
    
    Requires the ImportExcel PowerShell module:
    Install-Module ImportExcel -Scope CurrentUser

.PARAMETER Path
    Path to the RVTools Excel file (.xlsx) to anonymize.

.PARAMETER OutputPath
    Optional output path. Defaults to <filename>_anonymized.xlsx in the same directory.

.PARAMETER ExportMappings
    If specified, exports the original-to-anonymized mappings to a separate file.

.EXAMPLE
    .\Anonymize-RVTools.ps1 -Path "C:\exports\rvtools.xlsx"
    
.EXAMPLE
    .\Anonymize-RVTools.ps1 -Path "rvtools.xlsx" -ExportMappings

.EXAMPLE
    .\Anonymize-RVTools.ps1 -Path "rvtools.xlsx" -OutputPath "anonymized_output.xlsx"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$Path,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportMappings
)

# Check for ImportExcel module
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Error: ImportExcel module is required." -ForegroundColor Red
    Write-Host "Install it with: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

Import-Module ImportExcel -ErrorAction Stop

# Sensitive columns mapped to their data type
$SensitiveColumns = @{
    'VM'                 = 'vm'
    'Name'               = 'name'
    'Host'               = 'host'
    'Cluster'            = 'cluster'
    'Datacenter'         = 'datacenter'
    'DNS Name'           = 'dns'
    'Domain'             = 'domain'
    'DNS Search Order'   = 'domain'
    'DNS Servers'        = 'ip'
    'Folder'             = 'folder'
    'vApp'               = 'folder'
    'Resource pool'      = 'folder'
    'Path'               = 'path'
    'Network'            = 'network'
    'Portgroup'          = 'network'
    'VI SDK Server'      = 'ip'
    'Primary IP Address' = 'ip'
    'Annotation'         = 'annotation'
    'Notes'              = 'annotation'
}

# Prefixes for anonymized values
$Prefixes = @{
    'vm'         = 'VM-'
    'host'       = 'HOST-'
    'cluster'    = 'CLUSTER-'
    'datacenter' = 'DC-'
    'datastore'  = 'DS-'
    'network'    = 'NET-'
    'folder'     = 'FOLDER-'
    'domain'     = 'domain'
    'ip'         = '10.0.'
    'name'       = 'NAME-'
    'annotation' = ''
}

# Global mappings and counters
$script:Mappings = @{}
$script:Counters = @{}

function Get-AnonymizedValue {
    param(
        [string]$Original,
        [string]$DataType
    )
    
    if ([string]::IsNullOrWhiteSpace($Original)) {
        return ''
    }
    
    # Initialize mapping dictionary for this type if needed
    if (-not $script:Mappings.ContainsKey($DataType)) {
        $script:Mappings[$DataType] = @{}
        $script:Counters[$DataType] = 0
    }
    
    # Check if already mapped
    if ($script:Mappings[$DataType].ContainsKey($Original)) {
        return $script:Mappings[$DataType][$Original]
    }
    
    # Generate new anonymized value
    $script:Counters[$DataType]++
    $counter = $script:Counters[$DataType]
    $prefix = $Prefixes[$DataType]
    
    switch ($DataType) {
        'ip' {
            $newValue = "10.0.$([math]::Floor($counter / 256)).$($counter % 256)"
        }
        'domain' {
            $newValue = "domain$counter.local"
        }
        'annotation' {
            $newValue = ''
        }
        default {
            $newValue = "{0}{1:D4}" -f $prefix, $counter
        }
    }
    
    $script:Mappings[$DataType][$Original] = $newValue
    return $newValue
}

function Get-AnonymizedDnsName {
    param([string]$DnsName)
    
    if ([string]::IsNullOrWhiteSpace($DnsName) -or -not $DnsName.Contains('.')) {
        return Get-AnonymizedValue -Original $DnsName -DataType 'vm'
    }
    
    $parts = $DnsName.Split('.', 2)
    $hostname = $parts[0]
    $domain = if ($parts.Length -gt 1) { $parts[1] } else { '' }
    
    $anonHost = Get-AnonymizedValue -Original $hostname -DataType 'vm'
    $anonDomain = if ($domain) { Get-AnonymizedValue -Original $domain -DataType 'domain' } else { '' }
    
    if ($anonDomain) {
        return "$anonHost.$anonDomain"
    }
    return $anonHost
}

function Get-AnonymizedPath {
    param([string]$PathValue)
    
    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        return ''
    }
    
    # Handle [DATASTORE] VM_FOLDER/VM_FILE.ext paths
    if ($PathValue -match '^\[([^\]]+)\]\s*(.*)$') {
        $dsName = $Matches[1]
        $restOfPath = $Matches[2].Trim()
        $anonDs = Get-AnonymizedValue -Original $dsName -DataType 'datastore'
        
        # Check if there's a folder/file structure
        if ($restOfPath -match '^([^/]+)/(.*)$') {
            $vmFolder = $Matches[1]
            $filename = $Matches[2]
            
            # Anonymize the VM folder name
            $anonVmFolder = Get-AnonymizedValue -Original $vmFolder -DataType 'vm'
            
            # Replace original VM name in filename with anonymized version
            $anonFilename = if ($filename -and $vmFolder) {
                $filename -replace [regex]::Escape($vmFolder), $anonVmFolder
            } else {
                $filename
            }
            
            return "[$anonDs] $anonVmFolder/$anonFilename"
        }
        else {
            return "[$anonDs] $restOfPath"
        }
    }
    
    return Get-AnonymizedValue -Original $PathValue -DataType 'folder'
}

function ConvertTo-AnonymizedValue {
    param(
        [object]$Value,
        [string]$ColumnName,
        [string]$SheetName
    )
    
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace($Value.ToString())) {
        return $Value
    }
    
    $strValue = $Value.ToString()
    $dataType = $SensitiveColumns[$ColumnName]
    
    if (-not $dataType) {
        return $Value
    }
    
    switch ($ColumnName) {
        'DNS Name' {
            return Get-AnonymizedDnsName -DnsName $strValue
        }
        'Host' {
            if ($strValue.Contains('.')) {
                return Get-AnonymizedDnsName -DnsName $strValue
            }
            return Get-AnonymizedValue -Original $strValue -DataType 'host'
        }
        'Path' {
            return Get-AnonymizedPath -PathValue $strValue
        }
        'Name' {
            # "Name" varies by sheet
            switch ($SheetName) {
                'vDatastore' { return Get-AnonymizedValue -Original $strValue -DataType 'datastore' }
                'vCluster' { return Get-AnonymizedValue -Original $strValue -DataType 'cluster' }
                default { return Get-AnonymizedValue -Original $strValue -DataType 'name' }
            }
        }
        default {
            return Get-AnonymizedValue -Original $strValue -DataType $dataType
        }
    }
}

function Export-Mappings {
    param([string]$OutputFile)
    
    $mappingData = @{}
    
    foreach ($dataType in $script:Mappings.Keys) {
        $rows = @()
        foreach ($original in $script:Mappings[$dataType].Keys) {
            $rows += [PSCustomObject]@{
                'Original Value'   = $original
                'Anonymized Value' = $script:Mappings[$dataType][$original]
            }
        }
        if ($rows.Count -gt 0) {
            $mappingData[$dataType] = $rows
        }
    }
    
    # Export each mapping type to a separate worksheet
    $first = $true
    foreach ($dataType in $mappingData.Keys) {
        $sheetName = "${dataType}_mappings"
        if ($sheetName.Length -gt 31) {
            $sheetName = $sheetName.Substring(0, 31)
        }
        
        if ($first) {
            $mappingData[$dataType] | Export-Excel -Path $OutputFile -WorksheetName $sheetName -AutoSize
            $first = $false
        }
        else {
            $mappingData[$dataType] | Export-Excel -Path $OutputFile -WorksheetName $sheetName -AutoSize -Append
        }
    }
    
    Write-Host "Mappings exported to: $OutputFile" -ForegroundColor Green
}

# Main script execution
$ErrorActionPreference = 'Stop'

# Resolve full path
$inputFile = Resolve-Path $Path
$inputDir = Split-Path $inputFile -Parent
$inputName = [System.IO.Path]::GetFileNameWithoutExtension($inputFile)
$inputExt = [System.IO.Path]::GetExtension($inputFile)

# Set output path
if (-not $OutputPath) {
    $OutputPath = Join-Path $inputDir "${inputName}_anonymized${inputExt}"
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "RVTools Anonymizer (PowerShell)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Input:  $inputFile"
Write-Host "Output: $OutputPath`n"

# Get all sheet names
Write-Host "Loading workbook..." -ForegroundColor Yellow
$sheetNames = Get-ExcelSheetInfo -Path $inputFile | Select-Object -ExpandProperty Name

$totalSheets = $sheetNames.Count
$currentSheet = 0

# Process each sheet
foreach ($sheetName in $sheetNames) {
    $currentSheet++
    Write-Host "Processing sheet $currentSheet/$totalSheets`: $sheetName" -ForegroundColor Gray
    
    # Import sheet data
    $data = Import-Excel -Path $inputFile -WorksheetName $sheetName
    
    if ($null -eq $data -or $data.Count -eq 0) {
        continue
    }
    
    # Get column names that need anonymization
    $columns = $data[0].PSObject.Properties.Name
    $sensitiveColumnsInSheet = $columns | Where-Object { $SensitiveColumns.ContainsKey($_) }
    
    if ($sensitiveColumnsInSheet.Count -eq 0) {
        # No sensitive columns, export as-is
        if ($currentSheet -eq 1) {
            $data | Export-Excel -Path $OutputPath -WorksheetName $sheetName -AutoSize
        }
        else {
            $data | Export-Excel -Path $OutputPath -WorksheetName $sheetName -AutoSize -Append
        }
        continue
    }
    
    # Anonymize sensitive columns
    foreach ($row in $data) {
        foreach ($colName in $sensitiveColumnsInSheet) {
            $originalValue = $row.$colName
            if ($null -ne $originalValue) {
                $row.$colName = ConvertTo-AnonymizedValue -Value $originalValue -ColumnName $colName -SheetName $sheetName
            }
        }
    }
    
    # Export anonymized data
    if ($currentSheet -eq 1) {
        $data | Export-Excel -Path $OutputPath -WorksheetName $sheetName -AutoSize
    }
    else {
        $data | Export-Excel -Path $OutputPath -WorksheetName $sheetName -AutoSize -Append
    }
}

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "Anonymization Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "Output file: $OutputPath"

# Export mappings if requested
if ($ExportMappings) {
    $mappingsFile = Join-Path $inputDir "anonymization_mappings.xlsx"
    Export-Mappings -OutputFile $mappingsFile
}

# Summary
Write-Host "`nSummary:" -ForegroundColor Cyan
foreach ($dataType in $script:Counters.Keys | Sort-Object) {
    Write-Host "  $($dataType): $($script:Counters[$dataType]) values anonymized"
}

Write-Host "`nDone!" -ForegroundColor Green
