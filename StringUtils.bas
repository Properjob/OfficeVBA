Attribute VB_Name = "StringUtils"
Function StringUtils() As String
    StringUtils = "1.2" ' Version
    
    ' 1.0 Initial Work, Added Piece, Extract, TextFromCell
    ' 1.1 Fix to Extract
    ' 1.2 Fix to TextFromCell
End Function

Function Piece(Str As String, Delim As String, StartPiece As Integer, Optional EndPiece As Integer) As String
    Dim RetStr As String
    Dim RetArr() As String
    Dim LBnd As Integer
    Dim UBnd As Integer
    '
    RetStr = ""
    '
    ' Check that delim exists in string
    If Not InStr(1, Str, Delim) > 0 Then
        RetStr = Str
        GoTo Finish
    End If
    '
    ' Split into array
    RetArr = Strings.Split(Str, Delim)
    LBnd = LBound(RetArr) + 1
    UBnd = UBound(RetArr) + 1
    '
    ' Check StartPiece is Lower than the last piece
    If StartPiece > UBnd Then GoTo Finish
    If StartPiece = -1 Then StartPiece = UBnd
    '
    If Not EndPiece = 0 Then
        If EndPiece = -1 Then EndPiece = UBnd
        If EndPiece > UBnd Then GoTo Finish
        Else: EndPiece = StartPiece
    End If
    '
    ' Give pieces requested
    For Pce = StartPiece - 1 To EndPiece - 1
        RetStr = RetStr & RetArr(Pce)
        If Not Pce = EndPiece - 1 Then RetStr = RetStr & Delim
    Next Pce
    
Finish:
    Piece = RetStr
End Function

Public Function Extract(Str As String, Optional StartPos As Integer = 1, Optional EndPos As Integer = -1)
    Dim StrLen As Integer
    '
    StrLen = Len(Str)
    '
    If EndPos = -1 Then EndPos = StrLen
    If StartPos = -1 Then StartPos = StrLen
    Extract = Strings.Mid$(Str, StartPos, EndPos - StartPos + 1)
End Function

Public Function TextFromCell(pCell As Cell) As String
    Dim cellText As String
    '
    cellText = pCell.Range.Text
    TextFromCell = cellText
    '
    If Extract(cellText, Len(cellText) - 1) = (Chr(13) & Chr(7)) Then
        TextFromCell = Mid$(cellText, 1, Len(cellText) - 2)
    End If
End Function

Public Function Contains(Str As String, StrIn As String, Delim As String) As Boolean
    Dim SplitArr() As String
    Dim RetStr As Boolean
    Dim Pce As Integer
    '
    RetStr = False
    '
    SplitArr = Strings.Split(StrIn, Delim)
    
    For Pce = LBound(SplitArr) To UBound(SplitArr)
        If SplitArr(Pce) = Str Then
            RetStr = True
            Exit For
        End If
    Next Pce
    
    Contains = RetStr
End Function
