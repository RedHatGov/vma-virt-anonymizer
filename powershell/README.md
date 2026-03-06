# RVTools Anonymizer - PowerShell Version

A PowerShell script to anonymize sensitive data in RVTools exports. Works on Windows without needing Excel installed.

## Requirements

- Windows PowerShell 5.1+ or PowerShell Core 7+
- ImportExcel module (no Excel installation required)

## Quick Start

### 1. Install ImportExcel Module (One-Time Setup)

```powershell
# No admin required - installs to user scope
Install-Module ImportExcel -Scope CurrentUser
```

If prompted about an untrusted repository, type `Y` to confirm.

### 2. Run the Anonymizer

```powershell
# Basic usage - creates <filename>_anonymized.xlsx
.\Anonymize-RVTools.ps1 -Path "C:\exports\rvtools_export.xlsx"

# With mapping export (shows what was changed)
.\Anonymize-RVTools.ps1 -Path "rvtools_export.xlsx" -ExportMappings

# Specify custom output path
.\Anonymize-RVTools.ps1 -Path "rvtools_export.xlsx" -OutputPath "clean_export.xlsx"
```

## Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Path` | Yes | Path to the RVTools Excel file to anonymize |
| `-OutputPath` | No | Custom output path (default: `<input>_anonymized.xlsx`) |
| `-ExportMappings` | No | Export original-to-anonymized mappings to separate file |

## What Gets Anonymized

| Data Type | Example Original | Anonymized As |
|-----------|------------------|---------------|
| VM Names | `PRODDB01` | `VM-0001` |
| DNS Names | `server.corp.local` | `VM-0001.domain1.local` |
| Hosts | `esxi01.corp.local` | `HOST-0001.domain1.local` |
| Clusters | `Production-Cluster` | `CLUSTER-0001` |
| Datacenters | `Chicago-DC` | `DC-0001` |
| Datastores | `SAN_LUN_01` | `DS-0001` |
| Networks | `VLAN_100_Prod` | `NET-0001` |
| Paths | `[SAN_LUN_01] VM/VM.vmx` | `[DS-0001] VM-0001/VM-0001.vmx` |
| IPs | `192.168.1.100` | `10.0.0.1` |
| Folders | `/Production/Apps` | `FOLDER-0001` |

## Output Files

After running, you'll have:

1. **Anonymized Export** (`*_anonymized.xlsx`) - Safe to share externally
2. **Mappings File** (`anonymization_mappings.xlsx`) - Reference showing original values (keep secure!)

## Troubleshooting

### "Running scripts is disabled on this system"

PowerShell's execution policy may prevent running scripts. To fix:

```powershell
# Allow scripts for current user only
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### "ImportExcel module not found"

Make sure to install the module first:

```powershell
Install-Module ImportExcel -Scope CurrentUser -Force
```

### Script runs slowly

Large RVTools files (10,000+ VMs) may take several minutes. The script shows progress for each sheet being processed.

## Example Session

```
PS C:\exports> .\Anonymize-RVTools.ps1 -Path rvtools_export.xlsx -ExportMappings

========================================
RVTools Anonymizer (PowerShell)
========================================
Input:  C:\exports\rvtools_export.xlsx
Output: C:\exports\rvtools_export_anonymized.xlsx

Loading workbook...
Processing sheet 1/27: vInfo
Processing sheet 2/27: vCPU
Processing sheet 3/27: vMemory
...
Processing sheet 27/27: vMetaData

========================================
Anonymization Complete!
========================================
Output file: C:\exports\rvtools_export_anonymized.xlsx
Mappings exported to: C:\exports\anonymization_mappings.xlsx

Summary:
  cluster: 15 values anonymized
  datacenter: 8 values anonymized
  datastore: 120 values anonymized
  domain: 3 values anonymized
  folder: 45 values anonymized
  host: 42 values anonymized
  ip: 12 values anonymized
  network: 67 values anonymized
  vm: 1250 values anonymized

Done!
```
