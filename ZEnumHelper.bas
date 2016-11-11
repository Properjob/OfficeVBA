Attribute VB_Name = "ZEnumHelper"
Public RegistryKeys As SimpleEnum
Public References As SimpleEnum
Dim RegistryKeysLoaded As Boolean

Public Function RegistryKeysItem(IndexVal) As String
    RegistryKeysItem = RegistryKeys.Values.Item(IndexVal)
End Function

Public Function ReferencesItem(IndexVal) As String
    ReferencesItem = References.Values.Item(IndexVal)
End Function

Sub Load()
    Set RegistryKeys = New SimpleEnum
    Set References = New SimpleEnum
    '
    RegistryKeysLoad
    ReferencesLoad
    End Sub

Sub RegistryKeysLoad()
    Set RegistryKeys = RegistryKeys.Init()
    '
    RegistryKeys.Values.Add "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Word\Options\AutoScanCompFiles", "AutoScanCompFiles"
    RegistryKeys.Values.Add "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Word\Options\CompFilesPath", "CompFilesPath"
    RegistryKeys.Values.Add "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Word\Security\AccessVBOM", "AccessVBProj"
End Sub

Sub ReferencesLoad()
    References.Values.Add "C:\PROGRA~2\COMMON~1\MICROS~1\VBA\VBA7\VBE7.DLL", "VBA"
    References.Values.Add "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE", "Excel"
    References.Values.Add "C:\Windows\SysWOW64\stdole2.tlb", "stdole"
    References.Values.Add "C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\MSO.DLL", "Office"
    References.Values.Add "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB", "VBIDE"
    References.Values.Add Application.Path & "\MSOUTL.OLB", "Outlook"
    References.Values.Add Application.Path & "\MSWORD.OLB", "Word"
End Sub

