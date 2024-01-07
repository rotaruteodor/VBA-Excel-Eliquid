Private Sub Workbook_Open()
    Application.Calculation = xlAutomatic
    Call LiquidCalculatorSheet.initializeGlobalVariables
    LiquidCalculatorSheet.Activate
End Sub

Private Sub Workbook_SheetActivate(ByVal ClickedSheet As Object)
    If StrComp(ClickedSheet.Name, IngredientsInfoSheet.Name) <> 0 Then
        IngredientsInfoSheet.Visible = False
    End If

    If StrComp(ClickedSheet.Name, LiquidHistorySheet.Name) = 0 Then
        Call LiquidHistorySheet.updateSteepTimeOfAllLiquids
    End If
End Sub
