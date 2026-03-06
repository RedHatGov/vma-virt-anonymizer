VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AnonymizerForm 
   Caption         =   "RVTools Anonymizer"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "AnonymizerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AnonymizerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    ' Set default checkbox values
    chkVMs.Value = True
    chkHosts.Value = True
    chkClusters.Value = True
    chkDatacenters.Value = True
    chkDatastores.Value = True
    chkNetworks.Value = True
    chkFolders.Value = True
    chkDomains.Value = True
    chkIPs.Value = True
    chkStripDNS.Value = True
    
    ' Update status
    lblStatus.Caption = "Ready. Select an RVTools export file to anonymize."
End Sub

Private Sub btnBrowse_Click()
    ' Open file dialog to select RVTools export
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select RVTools Export File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            txtFilePath.Text = .SelectedItems(1)
            lblStatus.Caption = "File selected. Click 'Anonymize' to process."
        End If
    End With
End Sub

Private Sub btnAnonymize_Click()
    ' Validate file selection
    If Trim(txtFilePath.Text) = "" Then
        MsgBox "Please select an RVTools export file first.", vbExclamation, "No File Selected"
        Exit Sub
    End If
    
    If Dir(txtFilePath.Text) = "" Then
        MsgBox "The selected file does not exist.", vbExclamation, "File Not Found"
        Exit Sub
    End If
    
    ' Transfer settings to module
    AnonymizerModule.AnonymizeVMs = chkVMs.Value
    AnonymizerModule.AnonymizeHosts = chkHosts.Value
    AnonymizerModule.AnonymizeClusters = chkClusters.Value
    AnonymizerModule.AnonymizeDatacenters = chkDatacenters.Value
    AnonymizerModule.AnonymizeDatastores = chkDatastores.Value
    AnonymizerModule.AnonymizeNetworks = chkNetworks.Value
    AnonymizerModule.AnonymizeFolders = chkFolders.Value
    AnonymizerModule.AnonymizeDomains = chkDomains.Value
    AnonymizerModule.AnonymizeIPs = chkIPs.Value
    AnonymizerModule.StripDNSSuffix = chkStripDNS.Value
    
    ' Disable UI during processing
    EnableControls False
    lblStatus.Caption = "Processing... Please wait."
    DoEvents
    
    ' Open the source workbook and anonymize
    On Error GoTo ErrorHandler
    
    Dim sourceWB As Workbook
    Set sourceWB = Workbooks.Open(txtFilePath.Text, ReadOnly:=True)
    
    AnonymizerModule.AnonymizeWorkbook sourceWB
    
    sourceWB.Close SaveChanges:=False
    
    lblStatus.Caption = "Anonymization complete!"
    MsgBox "Anonymization complete!" & vbCrLf & vbCrLf & _
           "The anonymized file has been saved with '_anonymized' suffix.", _
           vbInformation, "Complete"
    
    EnableControls True
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Error: " & Err.Description
    MsgBox "An error occurred: " & vbCrLf & Err.Description, vbCritical, "Error"
    EnableControls True
    
    On Error Resume Next
    If Not sourceWB Is Nothing Then sourceWB.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub btnClose_Click()
    AnonymizerModule.ResetStatusBar
    Unload Me
End Sub

Private Sub EnableControls(enabled As Boolean)
    btnBrowse.enabled = enabled
    btnAnonymize.enabled = enabled
    chkVMs.enabled = enabled
    chkHosts.enabled = enabled
    chkClusters.enabled = enabled
    chkDatacenters.enabled = enabled
    chkDatastores.enabled = enabled
    chkNetworks.enabled = enabled
    chkFolders.enabled = enabled
    chkDomains.enabled = enabled
    chkIPs.enabled = enabled
    chkStripDNS.enabled = enabled
End Sub

Private Sub chkSelectAll_Click()
    Dim selectAll As Boolean
    selectAll = chkSelectAll.Value
    
    chkVMs.Value = selectAll
    chkHosts.Value = selectAll
    chkClusters.Value = selectAll
    chkDatacenters.Value = selectAll
    chkDatastores.Value = selectAll
    chkNetworks.Value = selectAll
    chkFolders.Value = selectAll
    chkDomains.Value = selectAll
    chkIPs.Value = selectAll
End Sub
