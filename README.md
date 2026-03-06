# vma-virt-anonymizer

A tool to anonymize sensitive data in [RVTools](https://www.dell.com/en-us/shop/vmware/sl/rvtools) VMware infrastructure exports while preserving data structure and integrity for safe sharing with vendors and partners.

## Overview

When sharing RVTools exports externally, you may need to remove identifying information such as:

- VM names and DNS names
- Host names (FQDNs)
- Cluster and datacenter names
- Datastore and network names
- IP addresses and domain names
- Folder paths and resource pools

This tool creates **consistent anonymized mappings** across all 27 RVTools sheets, ensuring that relationships between entities are preserved for analysis.

## Features

- **Consistent Mapping**: Same original value always maps to same anonymized value across all sheets
- **Preserves Structure**: Keeps headers, column order, data types (numbers, dates, booleans)
- **Configurable**: Choose which data types to anonymize
- **Non-Destructive**: Creates a new file, never modifies the original
- **Mapping Export**: Optionally save the original-to-anonymized mappings for your reference
- **Two Options**: Excel VBA macros (no installation) or Python script

## Sample Output

| Data Type | Original | Anonymized |
|-----------|----------|------------|
| VM Name | `PRODDB01` | `VM-0001` |
| DNS Name | `proddb01.corp.example.com` | `VM-0001.domain1.local` |
| Host | `esxi01.corp.example.com` | `HOST-0001.domain1.local` |
| Cluster | `Production-Cluster` | `CLUSTER-0001` |
| Datacenter | `Chicago-DC` | `DC-0001` |
| Datastore | `SAN_LUN_01` | `DS-0001` |
| Network | `VLAN_100_Prod` | `NET-0001` |
| IP Address | `192.168.1.100` | `10.0.0.1` |

## Quick Start

### Option 1: Excel VBA (Windows - No Installation Required)

Best for customers who cannot install software.

1. Download `AnonymizerModule.bas` and `QuickLauncher.bas`
2. Open Excel, create a new macro-enabled workbook (`.xlsm`)
3. Press **Alt+F11** to open VBA Editor
4. Right-click the project → **Import File** → Import both `.bas` files
5. Press **Alt+F8** → Run `RunAnonymizer`
6. Follow the prompts to select your RVTools file

See [Excel Setup Guide](docs/excel-setup.md) for detailed instructions.

### Option 2: Python Script

Best for automation and batch processing.

```bash
# Install dependencies
pip install openpyxl

# Analyze a file for sensitive data
python validate_anonymization.py --analyze rvtools_export.xlsx

# Anonymize a file
python validate_anonymization.py --anonymize rvtools_export.xlsx --export-mappings
```

## Requirements

### For Excel VBA Version
- Microsoft Excel (Windows)
- Macros enabled
- No admin privileges or software installation required

### For Python Version
- Python 3.7+
- openpyxl library

## What Gets Anonymized

### Sensitive Columns (Anonymized)

| Column Name | Found In Sheets | Anonymized As |
|-------------|-----------------|---------------|
| VM | vInfo, vCPU, vMemory, vDisk, etc. | `VM-0001` |
| Host | vHost, vInfo, vNetwork, etc. | `HOST-0001` or FQDN |
| Cluster | vCluster, vInfo, vHost, etc. | `CLUSTER-0001` |
| Datacenter | Most sheets | `DC-0001` |
| DNS Name | vInfo | `VM-0001.domain1.local` |
| Name | vDatastore, vCluster | `DS-0001`, `CLUSTER-0001` |
| Network/Portgroup | vNetwork | `NET-0001` |
| Path | vDisk, vInfo | `[DS-0001] ...` |
| Folder/vApp | vInfo, vUSB | `FOLDER-0001` |
| Domain | vHost | `domain1.local` |
| IP addresses | vSwitch, vPort, etc. | `10.0.x.x` |

### Preserved Data (Not Anonymized)

- Power state, template status, config status
- CPU, memory, disk metrics and statistics
- VMware Tools versions and status
- Dates and timestamps
- Boolean flags (HA, DRS settings, etc.)
- All numeric values

## Project Structure

```
vma-virt-anonymizer/
├── README.md                    # This file
├── LICENSE                      # Apache 2.0 License
├── AnonymizerModule.bas         # Main VBA anonymization logic
├── QuickLauncher.bas            # Simple VBA launcher macros
├── AnonymizerForm.frm           # Optional VBA UserForm
├── validate_anonymization.py    # Python script for testing/automation
└── docs/
    └── excel-setup.md           # Detailed Excel setup instructions
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.

## Related Projects

- [vma-virt-analysis](https://github.com/RedHatGov/vma-virt-analysis) - VMware infrastructure analysis tools
- [RVTools](https://www.dell.com/en-us/shop/vmware/sl/rvtools) - VMware infrastructure inventory tool by Dell

## Support

For issues or feature requests, please open a GitHub issue.
