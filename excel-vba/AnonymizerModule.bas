Attribute VB_Name = "AnonymizerModule"
Option Explicit

' ============================================================================
' RVTools Anonymizer - Main Module (Cross-Platform: Windows & Mac)
' ============================================================================
' This module provides functionality to anonymize sensitive data in RVTools
' Excel exports while maintaining data integrity and cross-sheet consistency.
' Uses native VBA Collections for Mac compatibility.
' ============================================================================

' Collections to store value mappings (original -> anonymized)
Private vmMappings As Collection
Private vmKeys As Collection
Private hostMappings As Collection
Private hostKeys As Collection
Private clusterMappings As Collection
Private clusterKeys As Collection
Private datacenterMappings As Collection
Private datacenterKeys As Collection
Private datastoreMappings As Collection
Private datastoreKeys As Collection
Private networkMappings As Collection
Private networkKeys As Collection
Private folderMappings As Collection
Private folderKeys As Collection
Private domainMappings As Collection
Private domainKeys As Collection
Private ipMappings As Collection
Private ipKeys As Collection

' Counters for generating sequential IDs
Private vmCounter As Long
Private hostCounter As Long
Private clusterCounter As Long
Private datacenterCounter As Long
Private datastoreCounter As Long
Private networkCounter As Long
Private folderCounter As Long
Private domainCounter As Long
Private ipCounter As Long

' Configuration flags
Public AnonymizeVMs As Boolean
Public AnonymizeHosts As Boolean
Public AnonymizeClusters As Boolean
Public AnonymizeDatacenters As Boolean
Public AnonymizeDatastores As Boolean
Public AnonymizeNetworks As Boolean
Public AnonymizeFolders As Boolean
Public AnonymizeDomains As Boolean
Public AnonymizeIPs As Boolean
Public StripDNSSuffix As Boolean

' Export collections for mapping output
Private allMappingTypes As Collection

Public Sub InitializeAnonymizer()
    ' Create collection objects for mappings
    Set vmMappings = New Collection
    Set vmKeys = New Collection
    Set hostMappings = New Collection
    Set hostKeys = New Collection
    Set clusterMappings = New Collection
    Set clusterKeys = New Collection
    Set datacenterMappings = New Collection
    Set datacenterKeys = New Collection
    Set datastoreMappings = New Collection
    Set datastoreKeys = New Collection
    Set networkMappings = New Collection
    Set networkKeys = New Collection
    Set folderMappings = New Collection
    Set folderKeys = New Collection
    Set domainMappings = New Collection
    Set domainKeys = New Collection
    Set ipMappings = New Collection
    Set ipKeys = New Collection
    
    ' Reset counters
    vmCounter = 0
    hostCounter = 0
    clusterCounter = 0
    datacenterCounter = 0
    datastoreCounter = 0
    networkCounter = 0
    folderCounter = 0
    domainCounter = 0
    ipCounter = 0
    
    ' Set default configuration - all enabled
    AnonymizeVMs = True
    AnonymizeHosts = True
    AnonymizeClusters = True
    AnonymizeDatacenters = True
    AnonymizeDatastores = True
    AnonymizeNetworks = True
    AnonymizeFolders = True
    AnonymizeDomains = True
    AnonymizeIPs = True
    StripDNSSuffix = True
End Sub

Private Function CollectionContains(col As Collection, key As String) As Boolean
    ' Check if a collection contains a key (Mac-compatible)
    On Error GoTo NotFound
    Dim temp As Variant
    temp = col(key)
    CollectionContains = True
    Exit Function
NotFound:
    CollectionContains = False
End Function

Private Function GetFromCollection(col As Collection, key As String) As String
    ' Get value from collection by key
    On Error GoTo NotFound
    GetFromCollection = col(key)
    Exit Function
NotFound:
    GetFromCollection = ""
End Function

Public Function GetAnonymizedValue(originalValue As Variant, dataType As String) As String
    ' Returns an anonymized value for the given original value
    ' Maintains consistency - same original value always returns same anonymized value
    
    Dim strValue As String
    Dim mappings As Collection
    Dim keys As Collection
    Dim counter As Long
    Dim prefix As String
    Dim newValue As String
    
    If IsEmpty(originalValue) Or IsNull(originalValue) Then
        GetAnonymizedValue = ""
        Exit Function
    End If
    
    strValue = CStr(originalValue)
    If Trim(strValue) = "" Then
        GetAnonymizedValue = ""
        Exit Function
    End If
    
    ' Select appropriate mapping collection and settings
    Select Case LCase(dataType)
        Case "vm"
            Set mappings = vmMappings
            Set keys = vmKeys
            counter = vmCounter
            prefix = "VM-"
        Case "host"
            Set mappings = hostMappings
            Set keys = hostKeys
            counter = hostCounter
            prefix = "HOST-"
        Case "cluster"
            Set mappings = clusterMappings
            Set keys = clusterKeys
            counter = clusterCounter
            prefix = "CLUSTER-"
        Case "datacenter"
            Set mappings = datacenterMappings
            Set keys = datacenterKeys
            counter = datacenterCounter
            prefix = "DC-"
        Case "datastore"
            Set mappings = datastoreMappings
            Set keys = datastoreKeys
            counter = datastoreCounter
            prefix = "DS-"
        Case "network"
            Set mappings = networkMappings
            Set keys = networkKeys
            counter = networkCounter
            prefix = "NET-"
        Case "folder"
            Set mappings = folderMappings
            Set keys = folderKeys
            counter = folderCounter
            prefix = "FOLDER-"
        Case "domain"
            Set mappings = domainMappings
            Set keys = domainKeys
            counter = domainCounter
            prefix = "domain"
        Case "ip"
            Set mappings = ipMappings
            Set keys = ipKeys
            counter = ipCounter
            prefix = "10.0."
        Case Else
            GetAnonymizedValue = strValue
            Exit Function
    End Select
    
    ' Check if already mapped
    If CollectionContains(mappings, strValue) Then
        GetAnonymizedValue = GetFromCollection(mappings, strValue)
        Exit Function
    End If
    
    ' Generate new anonymized value
    counter = counter + 1
    
    If prefix = "10.0." Then
        ' Special handling for IP addresses
        newValue = prefix & (counter \ 256) & "." & (counter Mod 256)
    ElseIf prefix = "domain" Then
        ' Special handling for domains
        newValue = prefix & counter & ".local"
    Else
        newValue = prefix & Format(counter, "0000")
    End If
    
    ' Store mapping
    mappings.Add newValue, strValue
    keys.Add strValue, CStr(keys.Count + 1)
    
    ' Update counter in module-level variable
    Select Case LCase(dataType)
        Case "vm": vmCounter = counter
        Case "host": hostCounter = counter
        Case "cluster": clusterCounter = counter
        Case "datacenter": datacenterCounter = counter
        Case "datastore": datastoreCounter = counter
        Case "network": networkCounter = counter
        Case "folder": folderCounter = counter
        Case "domain": domainCounter = counter
        Case "ip": ipCounter = counter
    End Select
    
    GetAnonymizedValue = newValue
End Function

Public Function AnonymizeDNSName(dnsName As String) As String
    ' Anonymize a DNS name by replacing the hostname and domain
    ' e.g., "server01.corp.example.com" -> "VM-0001.domain1.local"
    
    If Trim(dnsName) = "" Then
        AnonymizeDNSName = ""
        Exit Function
    End If
    
    Dim parts() As String
    parts = Split(dnsName, ".")
    
    If UBound(parts) < 1 Then
        ' No domain, just anonymize as VM name
        AnonymizeDNSName = GetAnonymizedValue(dnsName, "vm")
        Exit Function
    End If
    
    ' Get the hostname (first part)
    Dim hostPart As String
    hostPart = parts(0)
    
    ' Get the domain (remaining parts)
    Dim domainPart As String
    Dim i As Long
    For i = 1 To UBound(parts)
        If i > 1 Then domainPart = domainPart & "."
        domainPart = domainPart & parts(i)
    Next i
    
    ' Anonymize both parts
    Dim anonHost As String
    Dim anonDomain As String
    
    anonHost = GetAnonymizedValue(hostPart, "vm")
    
    If StripDNSSuffix Then
        anonDomain = GetAnonymizedValue(domainPart, "domain")
        AnonymizeDNSName = anonHost & "." & anonDomain
    Else
        AnonymizeDNSName = anonHost
    End If
End Function

Public Function AnonymizePath(pathValue As String) As String
    ' Anonymize paths like "[DATASTORE] VM_NAME/VM_NAME.vmx"
    ' Also anonymizes VM folder names and filenames in the path
    
    If Trim(pathValue) = "" Then
        AnonymizePath = ""
        Exit Function
    End If
    
    Dim result As String
    result = pathValue
    
    ' Handle datastore paths: [DATASTORE_NAME] VM_FOLDER/VM_FILE.ext
    If Left(pathValue, 1) = "[" Then
        Dim bracketEnd As Long
        bracketEnd = InStr(pathValue, "]")
        If bracketEnd > 2 Then
            ' Anonymize datastore name
            Dim dsName As String
            dsName = Mid(pathValue, 2, bracketEnd - 2)
            Dim anonDS As String
            anonDS = GetAnonymizedValue(dsName, "datastore")
            
            ' Get the rest of the path after the datastore
            Dim restOfPath As String
            restOfPath = Mid(pathValue, bracketEnd + 1)
            
            ' Remove leading space if present
            If Left(restOfPath, 1) = " " Then
                restOfPath = Mid(restOfPath, 2)
            End If
            
            ' Extract VM folder name (first path component before /)
            Dim slashPos As Long
            slashPos = InStr(restOfPath, "/")
            
            If slashPos > 1 Then
                Dim vmFolder As String
                vmFolder = Left(restOfPath, slashPos - 1)
                
                ' Anonymize the VM folder name
                Dim anonVMFolder As String
                anonVMFolder = GetAnonymizedValue(vmFolder, "vm")
                
                ' Get the filename part
                Dim fileName As String
                fileName = Mid(restOfPath, slashPos + 1)
                
                ' Replace original VM name in filename with anonymized version
                ' Handle patterns like VM_NAME.vmx, VM_NAME_1.vmdk, VM_NAME-flat.vmdk
                Dim anonFileName As String
                anonFileName = ReplaceVMNameInFileName(fileName, vmFolder, anonVMFolder)
                
                result = "[" & anonDS & "] " & anonVMFolder & "/" & anonFileName
            Else
                ' No folder, just a filename
                result = "[" & anonDS & "] " & restOfPath
            End If
        End If
    Else
        ' Non-datastore path (like resource pool paths)
        result = GetAnonymizedValue(pathValue, "folder")
    End If
    
    AnonymizePath = result
End Function

Private Function ReplaceVMNameInFileName(fileName As String, originalVM As String, anonVM As String) As String
    ' Replace VM name occurrences in filename while preserving suffixes
    ' e.g., "MYVM_1.vmdk" -> "VM-0001_1.vmdk"
    
    Dim result As String
    result = fileName
    
    ' Simple replacement - replace all occurrences of original VM name
    Dim pos As Long
    pos = InStr(1, result, originalVM, vbTextCompare)
    
    Do While pos > 0
        result = Left(result, pos - 1) & anonVM & Mid(result, pos + Len(originalVM))
        pos = InStr(pos + Len(anonVM), result, originalVM, vbTextCompare)
    Loop
    
    ReplaceVMNameInFileName = result
End Function

Public Sub AnonymizeWorkbook(sourceWB As Workbook, Optional targetPath As String = "")
    ' Main entry point - anonymize all sheets in the workbook
    
    Dim ws As Worksheet
    Dim targetWB As Workbook
    Dim newPath As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Create a copy of the workbook
    If targetPath = "" Then
        ' Handle file extension properly to avoid double-suffix
        Dim basePath As String
        Dim ext As String
        
        If Right(LCase(sourceWB.FullName), 5) = ".xlsx" Then
            basePath = Left(sourceWB.FullName, Len(sourceWB.FullName) - 5)
            ext = ".xlsx"
        ElseIf Right(LCase(sourceWB.FullName), 4) = ".xls" Then
            basePath = Left(sourceWB.FullName, Len(sourceWB.FullName) - 4)
            ext = ".xls"
        Else
            basePath = sourceWB.FullName
            ext = ""
        End If
        
        newPath = basePath & "_anonymized" & ext
    Else
        newPath = targetPath
    End If
    
    sourceWB.SaveCopyAs newPath
    Set targetWB = Workbooks.Open(newPath)
    
    ' Process each sheet
    Dim totalSheets As Long
    totalSheets = targetWB.Worksheets.Count
    Dim sheetNum As Long
    
    For sheetNum = 1 To totalSheets
        Set ws = targetWB.Worksheets(sheetNum)
        UpdateStatus "Processing sheet " & sheetNum & " of " & totalSheets & ": " & ws.Name
        AnonymizeSheet ws
    Next sheetNum
    
    ' Save the anonymized workbook
    targetWB.Save
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    UpdateStatus "Anonymization complete! Saved to: " & newPath
    
    ' Optionally export the mapping
    If MsgBox("Export value mappings to a separate file?", vbYesNo + vbQuestion, "Export Mappings") = vbYes Then
        ExportMappings targetWB.Path
    End If
End Sub

Private Sub AnonymizeSheet(ws As Worksheet)
    ' Anonymize a single worksheet
    
    If ws.UsedRange.Rows.Count < 2 Then Exit Sub
    
    ' Find header row and identify columns to process
    Dim headers As Collection
    Set headers = New Collection
    
    Dim lastCol As Long
    lastCol = ws.UsedRange.Columns.Count
    
    Dim col As Long
    Dim headerValue As String
    
    ' Map column positions to header names
    For col = 1 To lastCol
        headerValue = Trim(CStr(ws.Cells(1, col).Value))
        If headerValue <> "" Then
            On Error Resume Next
            headers.Add headerValue, CStr(col)
            On Error GoTo 0
        End If
    Next col
    
    ' Process each row (skip header)
    Dim lastRow As Long
    lastRow = ws.UsedRange.Rows.Count
    
    Dim row As Long
    Dim cellValue As Variant
    Dim newValue As String
    
    For row = 2 To lastRow
        For col = 1 To lastCol
            headerValue = ""
            On Error Resume Next
            headerValue = headers(CStr(col))
            On Error GoTo 0
            
            If headerValue <> "" Then
                cellValue = ws.Cells(row, col).Value
                
                If Not IsEmpty(cellValue) And Not IsNull(cellValue) Then
                    newValue = ProcessCellValue(CStr(cellValue), headerValue)
                    If newValue <> CStr(cellValue) Then
                        ws.Cells(row, col).Value = newValue
                    End If
                End If
            End If
        Next col
    Next row
End Sub

Private Function ProcessCellValue(cellValue As String, columnName As String) As String
    ' Process a cell value based on its column type
    
    ProcessCellValue = cellValue
    If Trim(cellValue) = "" Then Exit Function
    
    Select Case columnName
        Case "VM"
            If AnonymizeVMs Then ProcessCellValue = GetAnonymizedValue(cellValue, "vm")
            
        Case "Host"
            If AnonymizeHosts Then
                ' Host can be FQDN or short name
                If InStr(cellValue, ".") > 0 Then
                    ProcessCellValue = AnonymizeDNSName(cellValue)
                Else
                    ProcessCellValue = GetAnonymizedValue(cellValue, "host")
                End If
            End If
            
        Case "Cluster"
            If AnonymizeClusters Then ProcessCellValue = GetAnonymizedValue(cellValue, "cluster")
            
        Case "Datacenter"
            If AnonymizeDatacenters Then ProcessCellValue = GetAnonymizedValue(cellValue, "datacenter")
            
        Case "DNS Name"
            If AnonymizeVMs Or AnonymizeDomains Then ProcessCellValue = AnonymizeDNSName(cellValue)
            
        Case "Name"
            ' "Name" column varies by sheet - could be datastore, cluster, etc.
            If AnonymizeDatastores Then ProcessCellValue = GetAnonymizedValue(cellValue, "datastore")
            
        Case "Network", "Portgroup"
            If AnonymizeNetworks Then ProcessCellValue = GetAnonymizedValue(cellValue, "network")
            
        Case "Folder", "vApp", "Resource pool"
            If AnonymizeFolders Then ProcessCellValue = GetAnonymizedValue(cellValue, "folder")
            
        Case "Path"
            If AnonymizeDatastores Or AnonymizeVMs Then ProcessCellValue = AnonymizePath(cellValue)
            
        Case "Domain", "DNS Search Order"
            If AnonymizeDomains Then ProcessCellValue = GetAnonymizedValue(cellValue, "domain")
            
        Case "DNS Servers", "VI SDK Server"
            If AnonymizeIPs Then ProcessCellValue = AnonymizeIPAddress(cellValue)
            
        Case "Primary IP Address"
            If AnonymizeIPs Then ProcessCellValue = AnonymizeIPAddress(cellValue)
            
        Case "Annotation", "Notes"
            ' Clear annotations/notes as they may contain sensitive info
            ProcessCellValue = ""
            
    End Select
End Function

Private Function AnonymizeIPAddress(ipValue As String) As String
    ' Anonymize IP addresses
    
    If Trim(ipValue) = "" Then
        AnonymizeIPAddress = ""
        Exit Function
    End If
    
    ' Check if it looks like an IP address
    If IsIPAddress(ipValue) Then
        AnonymizeIPAddress = GetAnonymizedValue(ipValue, "ip")
    Else
        AnonymizeIPAddress = ipValue
    End If
End Function

Private Function IsIPAddress(value As String) As Boolean
    ' Simple check if value looks like an IP address
    Dim parts() As String
    parts = Split(value, ".")
    
    If UBound(parts) <> 3 Then
        IsIPAddress = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To 3
        If Not IsNumeric(parts(i)) Then
            IsIPAddress = False
            Exit Function
        End If
        If CLng(parts(i)) < 0 Or CLng(parts(i)) > 255 Then
            IsIPAddress = False
            Exit Function
        End If
    Next i
    
    IsIPAddress = True
End Function

Public Sub ExportMappings(targetFolder As String)
    ' Export all mappings to a new workbook for reference
    
    Dim mapWB As Workbook
    Set mapWB = Workbooks.Add
    
    ' Export each mapping type to a separate sheet
    ExportMappingToSheet mapWB, vmMappings, vmKeys, "VM_Mappings"
    ExportMappingToSheet mapWB, hostMappings, hostKeys, "Host_Mappings"
    ExportMappingToSheet mapWB, clusterMappings, clusterKeys, "Cluster_Mappings"
    ExportMappingToSheet mapWB, datacenterMappings, datacenterKeys, "Datacenter_Mappings"
    ExportMappingToSheet mapWB, datastoreMappings, datastoreKeys, "Datastore_Mappings"
    ExportMappingToSheet mapWB, networkMappings, networkKeys, "Network_Mappings"
    ExportMappingToSheet mapWB, folderMappings, folderKeys, "Folder_Mappings"
    ExportMappingToSheet mapWB, domainMappings, domainKeys, "Domain_Mappings"
    ExportMappingToSheet mapWB, ipMappings, ipKeys, "IP_Mappings"
    
    ' Delete the default Sheet1 if it exists and is empty
    On Error Resume Next
    Application.DisplayAlerts = False
    If mapWB.Sheets.Count > 1 Then
        mapWB.Sheets("Sheet1").Delete
    End If
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Save the mappings workbook
    Dim mapPath As String
    mapPath = targetFolder & Application.PathSeparator & "anonymization_mappings.xlsx"
    mapWB.SaveAs mapPath, xlOpenXMLWorkbook
    mapWB.Close
    
    MsgBox "Mappings exported to: " & mapPath, vbInformation, "Mappings Exported"
End Sub

Private Sub ExportMappingToSheet(wb As Workbook, mappings As Collection, keys As Collection, sheetName As String)
    ' Export a single mapping collection to a worksheet
    
    If keys.Count = 0 Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add
    ws.Name = sheetName
    
    ' Headers
    ws.Cells(1, 1).Value = "Original Value"
    ws.Cells(1, 2).Value = "Anonymized Value"
    ws.Range("A1:B1").Font.Bold = True
    
    ' Data
    Dim i As Long
    Dim originalKey As String
    For i = 1 To keys.Count
        originalKey = keys(i)
        ws.Cells(i + 1, 1).Value = originalKey
        ws.Cells(i + 1, 2).Value = GetFromCollection(mappings, originalKey)
    Next i
    
    ' Autofit columns
    ws.Columns("A:B").AutoFit
End Sub

Private Sub UpdateStatus(message As String)
    ' Update status bar
    Application.StatusBar = message
    DoEvents
End Sub

Public Sub ResetStatusBar()
    Application.StatusBar = False
End Sub

' ============================================================================
' Quick Launch - Run this to start the anonymizer
' ============================================================================
Public Sub LaunchAnonymizer()
    ' Initialize and show the configuration form
    InitializeAnonymizer
    ' For cross-platform, use QuickLauncher instead of form
    QuickLauncher.RunAnonymizer
End Sub
