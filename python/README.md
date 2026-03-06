# RVTools Anonymizer - Python Version

A Python script for anonymizing, validating, and analyzing RVTools exports. Best for automation, batch processing, and CI/CD pipelines.

## Requirements

- Python 3.7+
- openpyxl library

## Installation

```bash
pip install openpyxl
```

## Usage

### Analyze a File

Identify sensitive data before anonymizing:

```bash
python validate_anonymization.py --analyze rvtools_export.xlsx
```

### Anonymize a File

```bash
# Basic anonymization
python validate_anonymization.py --anonymize rvtools_export.xlsx

# With custom output path
python validate_anonymization.py --anonymize rvtools_export.xlsx -o anonymized.xlsx

# With mapping export
python validate_anonymization.py --anonymize rvtools_export.xlsx --export-mappings
```

### Validate Anonymization

Check that anonymization was performed correctly:

```bash
python validate_anonymization.py --validate original.xlsx anonymized.xlsx
```

## Command Line Options

```
usage: validate_anonymization.py [-h] [--analyze FILE] [--anonymize FILE]
                                  [--validate ORIGINAL ANONYMIZED]
                                  [--output FILE] [--export-mappings]

Options:
  --analyze FILE          Analyze an RVTools file for sensitive data
  --anonymize FILE        Anonymize an RVTools file (creates _anonymized copy)
  --validate ORIG ANON    Validate anonymization was performed correctly
  --output, -o FILE       Custom output path for anonymized file
  --export-mappings       Export value mappings to separate file
```

## What Gets Anonymized

| Column | Data Type | Anonymized As |
|--------|-----------|---------------|
| VM | vm | `VM-0001` |
| Host | host/dns | `HOST-0001` or FQDN |
| Cluster | cluster | `CLUSTER-0001` |
| Datacenter | datacenter | `DC-0001` |
| DNS Name | dns | `VM-0001.domain1.local` |
| Path | path | `[DS-0001] VM-0001/...` |
| Network/Portgroup | network | `NET-0001` |
| Folder/vApp | folder | `FOLDER-0001` |
| Domain | domain | `domain1.local` |
| IP addresses | ip | `10.0.x.x` |
| Annotation/Notes | annotation | (cleared) |

## Example Output

```
$ python validate_anonymization.py --anonymize data/rvtools.xlsx --export-mappings

Loading workbook: data/rvtools.xlsx
Processing sheet 1/27: vInfo
Processing sheet 2/27: vCPU
...
Processing sheet 27/27: vMetaData
Saving anonymized workbook: data/rvtools_anonymized.xlsx

✓ Anonymization complete: data/rvtools_anonymized.xlsx
Mappings exported to: data/anonymization_mappings.xlsx
```

## Programmatic Usage

```python
from validate_anonymization import RVToolsAnonymizer

# Create anonymizer instance
anonymizer = RVToolsAnonymizer()

# Anonymize a workbook
output_path = anonymizer.anonymize_workbook('input.xlsx', 'output.xlsx')

# Export mappings
anonymizer.export_mappings('mappings.xlsx')

# Access mappings programmatically
vm_mappings = anonymizer.mappings['vm']
print(f"Anonymized {len(vm_mappings)} VMs")
```
