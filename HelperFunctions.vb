Private Sub rangeToDoc()
    On Error Resume Next
    Dim wObj As Word.Application
    Dim Doc As Word.Document
    ' Get existing instance of Word if it exists.
    Set wObj = GetObject(, "Word.Application")
    If Err <> 0 Then
        ' If GetObject fails, then use CreateObject instead.
        Set wObj = CreateObject("Word.Application")
    End If
    '
    ' Show application
    wObj.Visible = True
    wObj.Activate
    wObj.WindowState = wdWindowStateMinimize
    '
    ' Add a new document.
    ' For loop of rows
    Set Doc = wObj.Documents.Add("C:\template")
    '
    ' Insert Data into documents
     
    '
    ' Show Document
    Doc.Activate
    ' Next
    '
End Sub

Sub sendWindowMessage(windowName As String, message As Long, wParam As Long, lParam As Long)
'
' Sends Window Message to Task that matches windowName
Dim taskLoop As Task
For Each taskLoop In Tasks
	If InStr(taskLoop.Name, windowName) > 0 Then
	taskLoop.SendWindowMessage message, wParam, lParam
	Exit For
	End If
Next taskLoop
End Sub

Sub initJSON()
    Dim JSON As String
    Dim scriptControl As Object
    Dim o, i, num
    Dim Attachment
    Dim AttachCnt As Integer

    Set scriptControl = CreateObject("scriptcontrol")
    scriptControl.Language = "JScript"

    'JSON = JSON Text
    '
    scriptControl.Eval "var obj=(" & JSON & ")" 'evaluate the json response
    'add some accessor functions
    scriptControl.AddCode "function GetAttachCnt(){return obj.Attachment.length;}"	' Standard Function
    scriptControl.AddCode "function GetAttachment(i){return obj.Attachment[i];}"	' Return Array[]
    '
    scriptControl.Run("GetAttachCnt")
End Sub