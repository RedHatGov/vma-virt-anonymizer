# Excel VBA Setup Guide

This guide walks you through setting up the RVTools Anonymizer in Microsoft Excel.

## Prerequisites

- Microsoft Excel (Windows version)
- RVTools export file (.xlsx)
- No admin privileges or software installation required

## Setup Steps

### Step 1: Download the VBA Files

Download these files from the repository:
- `AnonymizerModule.bas` - Main anonymization logic
- `QuickLauncher.bas` - Simple launcher macros

### Step 2: Create the Anonymizer Workbook

1. Open Microsoft Excel
2. Create a new blank workbook
3. Save it as **`RVTools_Anonymizer.xlsm`**
   - Click **File → Save As**
   - Choose location (e.g., Desktop)
   - Select **"Excel Macro-Enabled Workbook (*.xlsm)"** as file type
   - Click **Save**

### Step 3: Open the VBA Editor

Press **Alt + F11** to open the Visual Basic Editor.

Alternatively:
1. Click the **Developer** tab in the ribbon
2. Click **Visual Basic**

> **Note**: If you don't see the Developer tab:
> 1. Click **File → Options**
> 2. Click **Customize Ribbon**
> 3. Check the box next to **Developer**
> 4. Click **OK**

### Step 4: Import the VBA Modules

1. In the VBA Editor, look at the **Project Explorer** panel (usually on the left)
2. Find **VBAProject (RVTools_Anonymizer.xlsm)**
3. Right-click on it
4. Select **Import File...**
5. Navigate to where you saved the `.bas` files
6. Select **AnonymizerModule.bas** and click **Open**
7. Repeat steps 3-6 for **QuickLauncher.bas**

You should now see both modules listed under **Modules** in your project.

### Step 5: Save and Close the VBA Editor

1. Press **Ctrl + S** to save
2. Press **Alt + Q** to close the VBA Editor and return to Excel

### Step 6: Enable Macros (if prompted)

If Excel asks about macros:
1. Click **Enable Content** or **Enable Macros**
2. If you see a security warning bar, click **Enable Content**

## Using the Anonymizer

### Method 1: Run from Macro Dialog

1. Press **Alt + F8** to open the Macro dialog
2. Select **RunAnonymizer** from the list
3. Click **Run**

### Method 2: Add a Button (Optional)

1. Go to **Developer → Insert → Button (Form Control)**
2. Draw a button on your worksheet
3. In the "Assign Macro" dialog, select **RunAnonymizer**
4. Click **OK**
5. Right-click the button to edit its text (e.g., "Anonymize RVTools File")

## Anonymization Process

When you run the anonymizer:

1. **Choose anonymization level**:
   - **Yes (Full)**: Anonymizes everything - VMs, hosts, clusters, datacenters, datastores, networks, folders, domains, IPs
   - **No (Minimal)**: Only anonymizes VM names and DNS names

2. **Select your RVTools file**:
   - A file browser opens
   - Navigate to your RVTools export (.xlsx file)
   - Select it and click **Open**

3. **Confirm**:
   - Review the confirmation message
   - Click **Yes** to proceed

4. **Wait for processing**:
   - Progress shows in the Excel status bar
   - Large files (10,000+ VMs) may take several minutes

5. **Export mappings** (optional):
   - When prompted, choose whether to export the value mappings
   - This creates a separate file showing original → anonymized values
   - **Keep this file secure** - it contains the original values!

## Output Files

After anonymization, you'll have:

| File | Description |
|------|-------------|
| `original_file_anonymized.xlsx` | The anonymized RVTools export |
| `anonymization_mappings.xlsx` | (Optional) Mapping of original to anonymized values |

The original file is **never modified**.

## Available Macros

| Macro | Description |
|-------|-------------|
| `RunAnonymizer` | Interactive mode with prompts |
| `RunAnonymizer_FullAnonymization` | Full anonymization, no prompts |
| `RunAnonymizer_VMsAndHostsOnly` | Minimal anonymization, no prompts |

## Troubleshooting

### "Macros are disabled"

1. Click **File → Options → Trust Center**
2. Click **Trust Center Settings**
3. Click **Macro Settings**
4. Select **"Disable all macros with notification"** (recommended) or **"Enable all macros"**
5. Click **OK** and restart Excel

### "Compile error" when running

Ensure both `.bas` files were imported correctly:
1. Press **Alt + F11** to open VBA Editor
2. Check that both `AnonymizerModule` and `QuickLauncher` appear under **Modules**
3. If missing, re-import the files

### Large files take too long

- Files with 10,000+ VMs may take 5-10 minutes
- Progress is shown in the Excel status bar (bottom of window)
- Do not close Excel while processing

### "Subscript out of range" error

This usually means a sheet structure differs from expected:
1. Verify the file is a genuine RVTools export
2. Ensure the file hasn't been manually modified
3. Try with a fresh RVTools export

### "File not found" after selecting file

- Ensure the file path doesn't contain special characters
- Try copying the file to a simpler path (e.g., `C:\Temp\`)

## Security Notes

- The anonymizer runs entirely locally - no data is sent anywhere
- Original files are never modified
- Keep the mappings file secure if you export it
- Consider deleting the mappings file after use if not needed
