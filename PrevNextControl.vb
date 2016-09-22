Dim cCount As Integer
Dim cRibbonUI As IRibbonUI

Sub doNextThing()
    cCount = cCount + 1
    Debug.Print cCount
End Sub

Sub doPrevThing()
    cCount = cCount - 1
    Debug.Print cCount
End Sub

'Callback for btnNextThing onAction
Sub NextThing(control As IRibbonControl)
    RefreshRibbon
    doNextThing
End Sub

'Callback for btnNextThing getEnabled
Sub GetEnabledNextThing(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (cCount < 5)
End Sub

'Callback for btnPrevThing onAction
Sub PrevThing(control As IRibbonControl)
    RefreshRibbon
    doPrevThing
End Sub

'Callback for btnPrevThing getEnabled
Sub GetEnabledPrevThing(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (cCount > 0)
End Sub

'Callback for customUI.onLoad
Sub RibbonLoaded(ribbon As IRibbonUI)
    Set cRibbonUI = ribbon
End Sub

Sub RefreshRibbon()
    cRibbonUI.InvalidateControl "btnPrevThing"
    cRibbonUI.InvalidateControl "btnNextThing"
    cRibbonUI.InvalidateControl "lblCurrThing"
End Sub

'Callback for lblCurrThing getLabel
Sub getCurrentThing(control As IRibbonControl, ByRef returnedVal)
    returnedVal = cCount
End Sub