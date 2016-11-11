VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim pValues As New Collection
Dim pCount As Integer

Function Init(Optional File As String) As SimpleEnum
    If File <> "" Then
        
    End If
    Set Init = Me
End Function

Public Property Get Values() As Collection
    Set Values = pValues
End Property

Public Property Get Count() As Integer
    Count = pValues.Count
End Property

Private Sub Class_Initialize()
    Set pValues = New Collection
End Sub

