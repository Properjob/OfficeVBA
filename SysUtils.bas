Attribute VB_Name = "SysUtils"
Function checkRegKey(RegKey As String, Value As String) As Boolean
Dim shellObj As Object
Dim Val As String

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set shellObj = CreateObject("WScript.Shell")
  'try to read the registry key
  Val = shellObj.RegRead(RegKey)
  If Value <> Val Then GoTo ErrorHandler
  'key was found
  checkRegKey = True
  Exit Function
  
ErrorHandler:
  'key was not found
  checkRegKey = False
End Function

Function getRegKeyVal(RegKey As String) As Variant
Dim shellObj As Object
Dim Val As String

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set shellObj = CreateObject("WScript.Shell")
  'try to read the registry key
  Val = shellObj.RegRead(RegKey)
  'key was found
  getRegKeyVal = Val
  Exit Function
  
ErrorHandler:
  'key was not found
  getRegKeyVal = ""
End Function

Sub setRegKey(RegKey As String, Value As String)
    Dim shellObj As Object
    On Error GoTo ErrorHandler
    Set shellObj = CreateObject("WScript.Shell")
    '
    shellObj.RegWrite RegKey, Value
ErrorHandler:
End Sub

Sub setRegKeyDWord(RegKey As String, Value As Integer)
    Dim shellObj As Object
    On Error GoTo ErrorHandler
    Set shellObj = CreateObject("WScript.Shell")
    '
    shellObj.RegWrite RegKey, Value, "REG_DWORD"
ErrorHandler:
End Sub

Sub CheckReferences(RequiredRefs As String)
    Dim VBAEditor As VBIDE.VBE
    Dim vbProj As VBIDE.VBProject
    Dim chkRef As VBIDE.Reference
    Dim Refs() As String
    Dim Pce As Integer

    Set VBAEditor = Application.VBE
    Set vbProj = VBAEditor.ActiveVBProject
    For Each chkRef In vbProj.References
        If Strings.Replace(RequiredRefs, ",", "") = "" Then Exit For
        If Contains(chkRef.Name, RequiredRefs, ",") Then
            RequiredRefs = Strings.Replace(RequiredRefs, chkRef.Name, "")   ' remove from list
        End If
    Next
    
    If Strings.Replace(RequiredRefs, ",", "") = "" Then
        Exit Sub
    Else
        Refs = Strings.Split(RequiredRefs, ",")
        ZEnumHelper.Load
        For Pce = LBound(Refs) To UBound(Refs)
            AddReference CStr(Refs(Pce)), vbProj
        Next Pce
    End If
End Sub

Sub AddReference(Ref As String, vbProj As VBIDE.VBProject)
    Dim Loc As String
    
    If (Ref = "") Or (IsNull(vbProj)) Then Exit Sub
    Loc = ZEnumHelper.References.Values(Ref)
    vbProj.References.AddFromFile (Loc)
End Sub

