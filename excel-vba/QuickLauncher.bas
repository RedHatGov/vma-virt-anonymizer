Attribute VB_Name = "QuickLauncher"
Option Explicit

' ============================================================================
' RVTools Anonymizer - Quick Launcher (Cross-Platform: Windows & Mac)
' ============================================================================
' This module provides a simple way to run the anonymizer using dialogs.
' Works on both Windows and Mac versions of Excel.
' ============================================================================

Public Sub RunAnonymizer()
    ' Main entry point - run this macro to start the anonymizer
    
    ' Initialize the anonymizer
    AnonymizerModule.InitializeAnonymizer
    
    ' Show configuration dialog
    If Not ConfigureAnonymizer() Then
        MsgBox "Anonymization cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If
    
    ' Select the file to anonymize
    Dim filePath As String
    filePath = SelectRVToolsFile()
    
    If filePath = "" Then
        MsgBox "No file selected. Anonymization cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If
    
    ' Confirm before processing
    Dim msg As String
    msg = "Ready to anonymize:" & vbCrLf & vbCrLf & _
          filePath & vbCrLf & vbCrLf & _
          "A new file will be created with '_anonymized' suffix." & vbCrLf & _
          "The original file will NOT be modified." & vbCrLf & vbCrLf & _
          "Continue?"
    
    If MsgBox(msg, vbYesNo + vbQuestion, "Confirm Anonymization") <> vbYes Then
        MsgBox "Anonymization cancelled.", vbInformation, "Cancelled"
        Exit Sub
    End If
    
    ' Open and process the file
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Opening file..."
    
    Dim sourceWB As Workbook
    Set sourceWB = Workbooks.Open(filePath, ReadOnly:=True)
    
    AnonymizerModule.AnonymizeWorkbook sourceWB
    
    sourceWB.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "An error occurred:" & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "File path: " & filePath, vbCritical, "Error"
    
    On Error Resume Next
    If Not sourceWB Is Nothing Then sourceWB.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function SelectRVToolsFile() As String
    ' Open file dialog to select RVTools export (cross-platform)
    
    #If Mac Then
        ' Mac version - use AppleScript
        Dim macPath As String
        Dim script As String
        
        script = "set theFile to choose file of type {""xlsx"", ""xls"", ""com.microsoft.excel.xls"", ""org.openxmlformats.spreadsheetml.sheet""} with prompt ""Select RVTools Export File""" & vbCrLf & _
                 "return POSIX path of theFile"
        
        On Error Resume Next
        macPath = MacScript(script)
        On Error GoTo 0
        
        If macPath <> "" Then
            SelectRVToolsFile = macPath
        Else
            SelectRVToolsFile = ""
        End If
    #Else
        ' Windows version - use FileDialog
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        
        With fd
            .Title = "Select RVTools Export File"
            .Filters.Clear
            .Filters.Add "Excel Files", "*.xlsx; *.xls"
            .AllowMultiSelect = False
            
            If .Show = -1 Then
                SelectRVToolsFile = .SelectedItems(1)
            Else
                SelectRVToolsFile = ""
            End If
        End With
    #End If
End Function

Private Function ConfigureAnonymizer() As Boolean
    ' Simple configuration using message boxes
    
    Dim msg As String
    Dim response As VbMsgBoxResult
    
    ' Ask about anonymization level
    msg = "Choose anonymization level:" & vbCrLf & vbCrLf & _
          "YES = Full anonymization (recommended)" & vbCrLf & _
          "       - VM names, hosts, clusters, datacenters" & vbCrLf & _
          "       - Datastores, networks, folders" & vbCrLf & _
          "       - Domains, IP addresses" & vbCrLf & vbCrLf & _
          "NO = Minimal anonymization" & vbCrLf & _
          "       - Only VM names and DNS names" & vbCrLf & vbCrLf & _
          "CANCEL = Abort"
    
    response = MsgBox(msg, vbYesNoCancel + vbQuestion, "Configure Anonymization")
    
    Select Case response
        Case vbYes
            ' Full anonymization - all options enabled (already set by InitializeAnonymizer)
            ConfigureAnonymizer = True
            
        Case vbNo
            ' Minimal anonymization
            AnonymizerModule.AnonymizeVMs = True
            AnonymizerModule.AnonymizeHosts = False
            AnonymizerModule.AnonymizeClusters = False
            AnonymizerModule.AnonymizeDatacenters = False
            AnonymizerModule.AnonymizeDatastores = False
            AnonymizerModule.AnonymizeNetworks = False
            AnonymizerModule.AnonymizeFolders = False
            AnonymizerModule.AnonymizeDomains = True
            AnonymizerModule.AnonymizeIPs = False
            AnonymizerModule.StripDNSSuffix = True
            ConfigureAnonymizer = True
            
        Case vbCancel
            ConfigureAnonymizer = False
    End Select
End Function

' ============================================================================
' Alternative: Run with specific settings
' ============================================================================

Public Sub RunAnonymizer_FullAnonymization()
    ' Run anonymizer with all options enabled
    AnonymizerModule.InitializeAnonymizer
    
    ' All options are enabled by default
    RunAnonymizerOnSelectedFile
End Sub

Public Sub RunAnonymizer_VMsAndHostsOnly()
    ' Run anonymizer for only VMs and hosts
    AnonymizerModule.InitializeAnonymizer
    
    AnonymizerModule.AnonymizeVMs = True
    AnonymizerModule.AnonymizeHosts = True
    AnonymizerModule.AnonymizeClusters = False
    AnonymizerModule.AnonymizeDatacenters = False
    AnonymizerModule.AnonymizeDatastores = False
    AnonymizerModule.AnonymizeNetworks = False
    AnonymizerModule.AnonymizeFolders = False
    AnonymizerModule.AnonymizeDomains = True
    AnonymizerModule.AnonymizeIPs = False
    AnonymizerModule.StripDNSSuffix = True
    
    RunAnonymizerOnSelectedFile
End Sub

Private Sub RunAnonymizerOnSelectedFile()
    ' Helper to run anonymizer on a selected file
    
    Dim filePath As String
    filePath = SelectRVToolsFile()
    
    If filePath = "" Then
        MsgBox "No file selected.", vbInformation, "Cancelled"
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim sourceWB As Workbook
    Set sourceWB = Workbooks.Open(filePath, ReadOnly:=True)
    
    AnonymizerModule.AnonymizeWorkbook sourceWB
    
    sourceWB.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
    
    On Error Resume Next
    If Not sourceWB Is Nothing Then sourceWB.Close SaveChanges:=False
    On Error GoTo 0
End Sub
