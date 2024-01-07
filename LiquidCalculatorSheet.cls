Option Explicit

Private baseVgPercentageCell As Range
Private basePgPercentageCell As Range
Private nicVgPercentageCell As Range
Private nicPgPercentageCell As Range
Private nicMgCell As Range
Private finalAmountCell As Range
Private finalNicStrengthCell As Range
Private desiredVgPercentageCell As Range
Private desiredPgPercentageCell As Range

Private Const FLAVOR_INPUT_ID As String = "flavorInput"
Private Const VG_RESULT_ID As String = "vgResult"
Private Const PG_RESULT_ID As String = "pgResult"
Private Const BASE_RESULT_ID As String = "baseResult"
Private Const NIC_BASE_RESULT_ID As String = "nicBaseResult"
Private Const FLAVOR_RESULT_ID As String = "flavorResult"
Private Const RECIPE_TOTALS_ID As String = "recipeTotalsResult"
Private Const VG_PERCENTAGE_RESULT_ID As String = "vgPercentageResult"
Private Const PG_PERCENTAGE_RESULT_ID As String = "pgPercentageResult"
Private Const TOTAL_FLAVOR_PERCENTAGE_RESULT As String = "totalFlavorPercentageResult"

Private Const FLAVOR_COLOR As Long = 8696052
Private Const BACKGROUND_COLOR As Long = 4210752
Private Const BASE_COLOR As Long = 10086143
Private Const LIQUID_HISTORY_ROW_COLOR As Long = 1117248

Sub initializeGlobalVariables()
    Set baseVgPercentageCell = LiquidCalculatorSheet.Range("D3")
    Set basePgPercentageCell = LiquidCalculatorSheet.Range("D4")
    Set nicVgPercentageCell = LiquidCalculatorSheet.Range("D5")
    Set nicPgPercentageCell = LiquidCalculatorSheet.Range("D6")
    Set nicMgCell = LiquidCalculatorSheet.Range("D7")
    Set finalAmountCell = LiquidCalculatorSheet.Range("C11")
    Set finalNicStrengthCell = LiquidCalculatorSheet.Range("C12")
    Set desiredVgPercentageCell = LiquidCalculatorSheet.Range("E11")
    Set desiredPgPercentageCell = LiquidCalculatorSheet.Range("E12")
End Sub

Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo AllErr
    Application.EnableEvents = False
        
    If Not Intersect(Target, baseVgPercentageCell) Is Nothing Then
        If baseVgPercentageCell.Value >= 0 And baseVgPercentageCell.Value <= 1 Then
            basePgPercentageCell.Value = 1 - baseVgPercentageCell.Value
        Else
            baseVgPercentageCell.Value = 1 - basePgPercentageCell.Value
            MsgBox "Must be a value between 0% and 100%", vbCritical
        End If
    End If
    
    If Not Intersect(Target, basePgPercentageCell) Is Nothing Then
        If basePgPercentageCell.Value >= 0 And basePgPercentageCell.Value <= 1 Then
            baseVgPercentageCell.Value = 1 - basePgPercentageCell.Value
        Else
            basePgPercentageCell.Value = 1 - baseVgPercentageCell.Value
            MsgBox "Must be a value between 0% and 100%", vbCritical
        End If
    End If
    
    If Not Intersect(Target, nicVgPercentageCell) Is Nothing Then
        If nicVgPercentageCell.Value >= 0 And nicVgPercentageCell.Value <= 1 Then
            nicPgPercentageCell.Value = 1 - nicVgPercentageCell.Value
        Else
            nicVgPercentageCell.Value = 1 - nicPgPercentageCell.Value
            MsgBox "Must be a value between 0% and 100%", vbCritical
        End If
    End If
    
    If Not Intersect(Target, nicPgPercentageCell) Is Nothing Then
        If nicPgPercentageCell.Value >= 0 And nicPgPercentageCell.Value <= 1 Then
            nicVgPercentageCell.Value = 1 - nicPgPercentageCell.Value
        Else
            nicPgPercentageCell.Value = 1 - nicVgPercentageCell.Value
            MsgBox "Must be a value between 0% and 100%", vbCritical
        End If
    End If
    
    If Not Intersect(Target, nicMgCell) Is Nothing Then
        If nicMgCell.Value <= 0 Then
            MsgBox "Nicotine mg can't be negative", vbCritical
            nicMgCell.Value = 20
        End If
    End If
    
    If Not Intersect(Target, finalAmountCell) Is Nothing Then
        If finalAmountCell.Value < 0 Then
            MsgBox "Final amount can't be negative", vbCritical
            finalAmountCell.Value = 30
        End If
    End If
    
    If Not Intersect(Target, finalNicStrengthCell) Is Nothing Then
        If finalNicStrengthCell.Value < 0 Then
            MsgBox "Nicotine strength can't be negative", vbCritical
            finalNicStrengthCell.Value = 3
        End If
    End If
    
    If isSeparateVgPgCbChecked And Not IsEmpty(desiredVgPercentageCell.Value) Then
        Dim minimumPgInRecipe As Double: minimumPgInRecipe = getMinimumPgInFinalAmount
        Dim minimumVgInRecipe As Double: minimumVgInRecipe = getMinimumVgInFinalAmount
        
        If Not Intersect(Target, desiredVgPercentageCell) Is Nothing Then
            If desiredVgPercentageCell.Value >= minimumVgInRecipe And desiredVgPercentageCell.Value <= 1 - minimumPgInRecipe Then
                desiredPgPercentageCell.Value = 1 - desiredVgPercentageCell.Value
            Else
                desiredVgPercentageCell.Value = 1 - desiredPgPercentageCell.Value
                MsgBox "Must be a value between " & _
                minimumVgInRecipe * 100 & "% and " & _
                (1 - minimumPgInRecipe) * 100 & "%", _
                vbCritical
            End If
        ElseIf Not Intersect(Target, desiredPgPercentageCell) Is Nothing Then
            If desiredPgPercentageCell.Value >= minimumPgInRecipe And desiredPgPercentageCell.Value <= 1 - minimumVgInRecipe Then
                desiredVgPercentageCell.Value = 1 - desiredPgPercentageCell.Value
            Else
                desiredPgPercentageCell.Value = 1 - desiredVgPercentageCell.Value
                MsgBox "Must be a value between " & _
                minimumPgInRecipe * 100 & "% and " & _
                (1 - minimumVgInRecipe) * 100 & "%", _
                vbCritical
            End If
        ElseIf desiredPgPercentageCell.Value < minimumPgInRecipe Then
            desiredPgPercentageCell.Value = minimumPgInRecipe
            desiredVgPercentageCell.Value = 1 - desiredPgPercentageCell.Value
        ElseIf desiredVgPercentageCell.Value < minimumVgInRecipe Then
            desiredVgPercentageCell.Value = minimumVgInRecipe
            desiredPgPercentageCell.Value = 1 - desiredVgPercentageCell.Value
        End If
    End If
    
FinallyBlock:
    Application.EnableEvents = True
    Exit Sub
    
AllErr:
    If Err.Number = 13 Then ' Type mismatch error
        Debug.Print Err.description
    Else
        MsgBox Err.description, vbCritical
    End If
    Resume FinallyBlock
End Sub

Sub saveLiquid()
    On Error GoTo AllErr
    Dim currentLiquidRow As Integer: currentLiquidRow = LiquidHistorySheet.Range("C1048576").End(xlUp).Offset(1).Row
    Dim dateOfMakingDescriptionTag As String: dateOfMakingDescriptionTag = "Date of making: "
    Dim flavorsDescriptionTag As String: flavorsDescriptionTag = "Flavors:"
    Dim totalFlavorDescriptionTag As String: totalFlavorDescriptionTag = "TOTAL:"
    Dim vgDescriptionTag As String: vgDescriptionTag = "VG:"
    Dim currentLiquidIndexInAllHistory As Integer: currentLiquidIndexInAllHistory = IIf( _
    currentLiquidRow = 3, 1, LiquidHistorySheet.Range("A1").Value + 1)
    
    With LiquidHistorySheet.Range("A1")
        .Value = currentLiquidIndexInAllHistory
        .Font.Color = BACKGROUND_COLOR
    End With
    With LiquidHistorySheet.Range("C" & currentLiquidRow)
        .Font.Bold = True
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
        .Formula = "=CONCAT(""#"", ROW() - 2)"
    End With
    Dim liquidDescription As String: liquidDescription = getCurrentLiquidDescriptionForHistory
    With LiquidHistorySheet.Range("D" & currentLiquidRow)
        .Value = liquidDescription
        .WrapText = True
        .RowHeight = .RowHeight + 20
        .IndentLevel = 2
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignCenter
        .Characters(1, InStr(liquidDescription, flavorsDescriptionTag) + Len(flavorsDescriptionTag)).Font.Bold = True
        .Characters(InStr(liquidDescription, totalFlavorDescriptionTag), Len(totalFlavorDescriptionTag)).Font.Underline = True
        .Characters(InStr(liquidDescription, vgDescriptionTag), Len(liquidDescription)).Font.Bold = True
    End With
    With LiquidHistorySheet.Range("E" & currentLiquidRow)
        .Value = "0 days"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    With LiquidHistorySheet.Range("F" & currentLiquidRow)
        .Value = "No Rating Yet"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    With LiquidHistorySheet.Range("G" & currentLiquidRow)
        .Value = "No Notes Yet"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    With LiquidHistorySheet.Range("H" & currentLiquidRow)
        .Value = getCurrentLiquidDescriptionAsJsonForHistory
        .Font.Color = BACKGROUND_COLOR
    End With
    With LiquidHistorySheet.Range("B" & currentLiquidRow & ":G" & currentLiquidRow)
        .Interior.Color = LIQUID_HISTORY_ROW_COLOR
        .Font.Italic = True
        .Font.Color = vbWhite
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .Color = vbWhite
        End With
    End With
    
    Call createHistoryLiquidButtons(LiquidHistorySheet.Range("B" & currentLiquidRow), currentLiquidIndexInAllHistory)
    Call displayAutoDismissByTimerMsgBox("SUCCESS!", "Your E-liquid has been saved.", 1)

FinallyBlock:
    Application.EnableEvents = True
    Exit Sub
    
AllErr:
    MsgBox "Couldn't save E-liquid! " & vbNewLine & "ERROR: " & Err.description, vbCritical
    Resume FinallyBlock
End Sub

Sub createHistoryLiquidButtons(buttonRange As Range, btnIndex As Integer)
    Dim remixButton As Button: Set remixButton = _
    LiquidHistorySheet.Buttons.Add(buttonRange.Left, buttonRange.Top, buttonRange.Width, buttonRange.Height / 2)
    With remixButton
      .OnAction = "LiquidHistorySheet.remixLiquid"
      .Caption = "Mix again"
      .Name = "BtnRemixLiquid" & btnIndex
      With .Font
        .Bold = True
        .Color = vbBlack
        .Size = 20
      End With
    End With
    
    Dim startLiquidButton As Button: Set startLiquidButton = _
    LiquidHistorySheet.Buttons.Add(buttonRange.Left, buttonRange.Top + buttonRange.Height / 2, buttonRange.Width / 3, buttonRange.Height / 4)
    With startLiquidButton
      .OnAction = "LiquidHistorySheet.startLiquidButtonClick"
      .Caption = "Not Started"
      .Name = "BtnStartLiquid" & btnIndex
      .Font.Bold = True
      .Font.Color = vbRed
    End With
    
    Dim finishLiquidButton As Button: Set finishLiquidButton = _
    LiquidHistorySheet.Buttons.Add(buttonRange.Left + buttonRange.Width / 3, buttonRange.Top + buttonRange.Height / 2, buttonRange.Width / 3, buttonRange.Height / 4)
    With finishLiquidButton
      .OnAction = "LiquidHistorySheet.finishLiquidButtonClick"
      .Caption = "Not Finished"
      .Name = "BtnFinishLiquid" & btnIndex
      .Font.Bold = True
      .Font.Color = vbRed
    End With
    
    Dim seeStatsButton As Button: Set seeStatsButton = _
    LiquidHistorySheet.Buttons.Add(buttonRange.Left + 2 * buttonRange.Width / 3, buttonRange.Top + buttonRange.Height / 2, buttonRange.Width / 3, buttonRange.Height / 4)
    With seeStatsButton
      .OnAction = "LiquidHistorySheet.seeLiquidStats"
      .Caption = "See Stats"
      .Name = "BtnSeeLiquidStats" & btnIndex
      .Font.Bold = True
      .Font.Color = vbBlue
    End With
    
    Dim deleteButton As Button: Set deleteButton = _
    LiquidHistorySheet.Buttons.Add(buttonRange.Left, buttonRange.Top + 3 * buttonRange.Height / 4, buttonRange.Width, buttonRange.Height / 4)
    With deleteButton
      .OnAction = "LiquidHistorySheet.deleteLiquid"
      .Caption = "Delete"
      .Name = "BtnDeleteLiquid" & btnIndex
      .Font.Bold = True
      .Font.Color = vbRed
    End With
End Sub

Function getCurrentLiquidDescriptionForHistory() As String
    Dim liquidDescription As String
    Dim flavorsStringForDescription As String: flavorsStringForDescription = "Flavors: " & vbNewLine
    Dim rowNr As Integer
    For rowNr = LiquidCalculatorSheet.Range("G1").End(xlDown).Row To LiquidCalculatorSheet.Range("G1048576").End(xlUp).Row
        If StrComp(LiquidCalculatorSheet.Range("G" & rowNr).Value, FLAVOR_INPUT_ID) = 0 Then
            flavorsStringForDescription = flavorsStringForDescription & _
            Space(6) & "- " & LiquidCalculatorSheet.Range("B" & rowNr).Value & " @ " & LiquidCalculatorSheet.Range("C" & rowNr).Value * 100 & "%" & vbNewLine
        End If
    Next rowNr
    flavorsStringForDescription = flavorsStringForDescription & _
    Space(6) & "TOTAL: " & LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(TOTAL_FLAVOR_PERCENTAGE_RESULT)).Value * 100 & "%"
    
    liquidDescription = _
    "Amount: " & finalAmountCell.Value & "ml" & vbNewLine & _
    "Nicotine: " & finalNicStrengthCell.Value & "mg" & vbNewLine & _
    flavorsStringForDescription & vbNewLine & vbNewLine & _
    "VG: " & Format(LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(VG_PERCENTAGE_RESULT_ID)).Value * 100, "0") & "% | " & _
    "PG: " & Format(LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(PG_PERCENTAGE_RESULT_ID)).Value * 100, "0") & "%" & vbNewLine & _
    "Price: " & Format(LiquidCalculatorSheet.Range("F" & getRowNrOfComponent(RECIPE_TOTALS_ID)).Value, "0.0") & " $" & vbNewLine & _
    "Date of making: " & Format(Now(), "dd/mm/yyyy || hh:mm")
    
    getCurrentLiquidDescriptionForHistory = liquidDescription
End Function

Function getCurrentLiquidDescriptionAsJsonForHistory() As String
    Dim liquidDescriptionAsJson As Object: Set liquidDescriptionAsJson = JsonConverter.ParseJson("{}")
    Dim flavorDictionary As Dictionary
    Dim flavorsCollection As Collection: Set flavorsCollection = New Collection

    Dim rowNr As Integer
    For rowNr = LiquidCalculatorSheet.Range("G1").End(xlDown).Row To LiquidCalculatorSheet.Range("G1048576").End(xlUp).Row
        If StrComp(LiquidCalculatorSheet.Range("G" & rowNr).Value, FLAVOR_INPUT_ID) = 0 Then
            Set flavorDictionary = New Dictionary
            flavorDictionary.Add "flavorName", LiquidCalculatorSheet.Range("B" & rowNr).Value
            flavorDictionary.Add "percentage", LiquidCalculatorSheet.Range("C" & rowNr).Value
            flavorsCollection.Add flavorDictionary
        End If
    Next rowNr
    
    liquidDescriptionAsJson.Add "flavors", flavorsCollection
    liquidDescriptionAsJson("amount") = finalAmountCell.Value
    liquidDescriptionAsJson("nicotine") = finalNicStrengthCell.Value
    liquidDescriptionAsJson("finalVg") = LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(VG_PERCENTAGE_RESULT_ID)).Value
    liquidDescriptionAsJson("finalPg") = LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(PG_PERCENTAGE_RESULT_ID)).Value
    liquidDescriptionAsJson("dateOfMaking") = Format(Now(), "dd/mm/yyyy  hh:mm:ss")
    liquidDescriptionAsJson("price") = LiquidCalculatorSheet.Range("F" & getRowNrOfComponent(RECIPE_TOTALS_ID)).Value
    liquidDescriptionAsJson("startDate") = ""
    liquidDescriptionAsJson("finishDate") = ""
    
    getCurrentLiquidDescriptionAsJsonForHistory = JsonConverter.ConvertToJson(liquidDescriptionAsJson)
End Function

Sub addFlavor()
    Dim currentFlavorInputRowNr As Integer: currentFlavorInputRowNr = IIf( _
    getRowNrOfComponent(FLAVOR_INPUT_ID) = -1, finalNicStrengthCell.Row + 1, getRowNrOfComponent(FLAVOR_INPUT_ID, False) + 1)
    LiquidCalculatorSheet.Rows(currentFlavorInputRowNr).Insert
    With LiquidCalculatorSheet.Range("G" & currentFlavorInputRowNr)
        .Value = FLAVOR_INPUT_ID
    End With
    With LiquidCalculatorSheet.Range("B" & currentFlavorInputRowNr)
        .Font.Bold = True
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=OFFSET(" & FlavorsSheet.Name & "!$A$1,1,0,COUNTA(" & FlavorsSheet.Name & "!$A:$A)-1)"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = False
        End With
        .Value = "Flavor name"
    End With
    With LiquidCalculatorSheet.Range("C" & currentFlavorInputRowNr)
        .NumberFormat = "0.00%"
        .Value = 0
    End With
    With LiquidCalculatorSheet.Range("B" & currentFlavorInputRowNr & ":C" & currentFlavorInputRowNr)
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        .Interior.Color = FLAVOR_COLOR
    End With
    With LiquidCalculatorSheet.Range("D" & currentFlavorInputRowNr & ":E" & currentFlavorInputRowNr)
        .Interior.Color = BACKGROUND_COLOR
    End With
    
    Call createDeleteFlavorButton(LiquidCalculatorSheet.Range("D" & currentFlavorInputRowNr))
    
    Dim currentFlavorResultRowNr As Integer: currentFlavorResultRowNr = IIf( _
    getRowNrOfComponent(FLAVOR_RESULT_ID) = -1, getRowNrOfComponent(NIC_BASE_RESULT_ID) + 1, getRowNrOfComponent(FLAVOR_RESULT_ID, False) + 1)
    LiquidCalculatorSheet.Rows(currentFlavorResultRowNr).Insert
    Dim recipeTotalsRowNr As Integer: recipeTotalsRowNr = getRowNrOfComponent(RECIPE_TOTALS_ID)
    With LiquidCalculatorSheet.Range("B" & currentFlavorResultRowNr)
        .Font.Bold = True
        .Formula = "=B" & currentFlavorInputRowNr
    End With
    With LiquidCalculatorSheet.Range("C" & currentFlavorResultRowNr)
        .NumberFormat = "0.00" & Space(1) & Chr$(34) & " ml" & Chr$(34)
        .Formula = "=" & finalAmountCell.Address & "*C" & currentFlavorInputRowNr
    End With
    With LiquidCalculatorSheet.Range("D" & currentFlavorResultRowNr)
        .NumberFormat = "0.00" & Space(1) & Chr$(34) & " g" & Chr$(34)
        .Formula = "=C" & currentFlavorResultRowNr
    End With
    With LiquidCalculatorSheet.Range("E" & currentFlavorResultRowNr)
        .NumberFormat = "0.00 %"
        .Formula = "=C" & currentFlavorResultRowNr & "/C" & recipeTotalsRowNr
    End With
    With LiquidCalculatorSheet.Range("G" & currentFlavorResultRowNr)
        .Value = FLAVOR_RESULT_ID
    End With
    With LiquidCalculatorSheet.Range("F" & currentFlavorResultRowNr)
        .NumberFormat = "0.00" & Space(1) & Chr$(34) & " $" & Chr$(34)
        .Formula = getVlookupFormulaForIngredientsCost(currentFlavorResultRowNr)
    End With
    With LiquidCalculatorSheet.Range("B" & currentFlavorResultRowNr & ":F" & currentFlavorResultRowNr)
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        .Interior.Color = FLAVOR_COLOR
    End With
    
    Dim firstRowNrInResultTable As Integer
    If isSeparateVgPgCbChecked Then
        firstRowNrInResultTable = getRowNrOfComponent(VG_RESULT_ID)
        LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(PG_RESULT_ID)).Formula = _
        "=(" & finalAmountCell.Address & "*" & desiredPgPercentageCell.Address & ")" & _
        "-(SUM(C" & getRowNrOfComponent(FLAVOR_RESULT_ID) & ":C" & getRowNrOfComponent(FLAVOR_RESULT_ID, False) & ")" & _
        "+(" & nicPgPercentageCell.Address & "*C" & getRowNrOfComponent(NIC_BASE_RESULT_ID) & "))"
    Else
        firstRowNrInResultTable = getRowNrOfComponent(BASE_RESULT_ID)
        LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(BASE_RESULT_ID)).Formula = _
        "=" & finalAmountCell.Address & _
        "-SUM(C" & getRowNrOfComponent(NIC_BASE_RESULT_ID) & ":C" & currentFlavorResultRowNr & ")"
    End If
    
    LiquidCalculatorSheet.Range("C" & recipeTotalsRowNr).Formula = "=SUM(C" & firstRowNrInResultTable & ":C" & currentFlavorResultRowNr & ")"
    LiquidCalculatorSheet.Range("C" & recipeTotalsRowNr).Copy
    LiquidCalculatorSheet.Range("D" & recipeTotalsRowNr & ":F" & recipeTotalsRowNr).PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
    LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(TOTAL_FLAVOR_PERCENTAGE_RESULT)).Formula = _
    "=SUM(C" & getRowNrOfComponent(FLAVOR_INPUT_ID) & ":C" & currentFlavorInputRowNr & ")"
End Sub

Sub createDeleteFlavorButton(buttonRange As Range)
    Dim deleteButton As Button: Set deleteButton = LiquidCalculatorSheet.Buttons.Add( _
    buttonRange.Left, _
    buttonRange.Top, _
    buttonRange.Width / 3, _
    buttonRange.Height)
    With deleteButton
      .OnAction = "LiquidCalculatorSheet.deleteFlavor"
      .Caption = "X"
      .Name = "ButtonDeleteFlavorInput" & buttonRange.Row
      .Font.Bold = True
      .Font.Color = vbRed
    End With
End Sub

Sub deleteFlavor(Optional ByVal flavorInputRowNr As Integer = -1)
    If flavorInputRowNr = -1 Then
        flavorInputRowNr = LiquidCalculatorSheet.Shapes(Application.Caller).TopLeftCell.Row
    End If
    
    Dim flavorRowNumberResultTable As Integer: flavorRowNumberResultTable = IIf( _
    isSeparateVgPgCbChecked, flavorInputRowNr + 6 + getNrOfFlavors(FLAVOR_INPUT_ID), flavorInputRowNr + 5 + getNrOfFlavors(FLAVOR_INPUT_ID))
    LiquidCalculatorSheet.Rows(flavorInputRowNr).EntireRow.Delete
    LiquidCalculatorSheet.Rows(flavorRowNumberResultTable).EntireRow.Delete
    
    If getNrOfFlavors(FLAVOR_RESULT_ID) = 0 Then
        LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(TOTAL_FLAVOR_PERCENTAGE_RESULT)).Value = 0#
        If isSeparateVgPgCbChecked Then
            LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(PG_RESULT_ID)).Formula = _
            "=(" & finalAmountCell.Address & "*" & desiredPgPercentageCell.Address & ")" & _
            "-((" & nicPgPercentageCell.Address & "*C" & getRowNrOfComponent(NIC_BASE_RESULT_ID) & "))"
        End If
    End If
    
    Dim btn As Object
    For Each btn In LiquidCalculatorSheet.Buttons
        If InStr(btn.Name, "ButtonDeleteFlavorInput") <> 0 Then
            btn.Name = "ButtonDeleteFlavorInput" & btn.TopLeftCell.Row
        End If
    Next btn
End Sub

Sub checkBoxSeparateVgPgClick()
    Dim firstRowInResultTable As Integer
    Dim lastRowInResultTable As Integer
    Dim recipeTotalsRowNr As Integer
    
    If isSeparateVgPgCbChecked Then
        LiquidCalculatorSheet.Rows(baseVgPercentageCell.Row & ":" & basePgPercentageCell.Row).EntireRow.Hidden = True
        LiquidCalculatorSheet.Range("B" & getRowNrOfComponent(VG_PERCENTAGE_RESULT_ID) & ":C" & getRowNrOfComponent(PG_PERCENTAGE_RESULT_ID)).Copy
        LiquidCalculatorSheet.Range("D11").PasteSpecial Paste:=xlPasteValues
        LiquidCalculatorSheet.Range("D11").Value = "Final VG %"
        LiquidCalculatorSheet.Range("D12").Value = "Final PG %"
        LiquidCalculatorSheet.Range("D11").PasteSpecial Paste:=xlPasteFormats
        LiquidCalculatorSheet.Rows(getRowNrOfComponent(BASE_RESULT_ID)).EntireRow.Delete

        Dim vgResultRowNr As Integer: vgResultRowNr = getRowNrOfComponent(NIC_BASE_RESULT_ID)
        LiquidCalculatorSheet.Rows(vgResultRowNr).Insert
        With LiquidCalculatorSheet.Range("G" & vgResultRowNr)
            .Value = VG_RESULT_ID
        End With
        With LiquidCalculatorSheet.Range("B" & vgResultRowNr)
            .Value = "VG"
        End With
        With LiquidCalculatorSheet.Range("C" & vgResultRowNr)
            .NumberFormat = "0.00" & Space(1) & Chr$(34) & " ml" & Chr$(34)
            .Formula = _
            "=(" & finalAmountCell.Address & "*" & desiredVgPercentageCell.Address & ")" & _
            "-(" & nicVgPercentageCell.Address & "*C" & getRowNrOfComponent(NIC_BASE_RESULT_ID) & ")"
            .Font.Bold = False
        End With
        With LiquidCalculatorSheet.Range("D" & vgResultRowNr)
            .NumberFormat = "0.00" & Space(1) & Chr$(34) & " g" & Chr$(34)
            .Formula = "=C" & vgResultRowNr & "*'" & IngredientsInfoSheet.Name & "'!C3"
            .Font.Bold = False
        End With
        With LiquidCalculatorSheet.Range("E" & vgResultRowNr)
            .NumberFormat = "0.0 %"
            .Formula = "=C" & vgResultRowNr & "/C" & getRowNrOfComponent(RECIPE_TOTALS_ID)
            .Font.Bold = False
        End With
        With LiquidCalculatorSheet.Range("F" & vgResultRowNr)
            .NumberFormat = "0.00" & Space(1) & Chr$(34) & " $" & Chr$(34)
            .Formula = getVlookupFormulaForIngredientsCost(vgResultRowNr)
            .Font.Bold = False
        End With
        With LiquidCalculatorSheet.Range("B" & vgResultRowNr & ":F" & vgResultRowNr)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            .Interior.Color = BASE_COLOR
            .Font.Color = vbBlack
        End With

        Dim pgResultRowNr As Integer: pgResultRowNr = getRowNrOfComponent(NIC_BASE_RESULT_ID)
        LiquidCalculatorSheet.Rows(pgResultRowNr).Insert
        With LiquidCalculatorSheet.Range("G" & pgResultRowNr)
            .Value = PG_RESULT_ID
        End With
        With LiquidCalculatorSheet.Range("B" & pgResultRowNr)
            .Font.Bold = True
            .Value = "PG"
        End With
        Dim flavorsSumFormula As String: flavorsSumFormula = IIf( _
        getNrOfFlavors(FLAVOR_INPUT_ID) > 0, "SUM(C" & getRowNrOfComponent(FLAVOR_RESULT_ID) & ":C" & getRowNrOfComponent(FLAVOR_RESULT_ID, False) & ")", "")
        With LiquidCalculatorSheet.Range("C" & pgResultRowNr)
            .NumberFormat = "0.00" & Space(1) & Chr$(34) & " ml" & Chr$(34)
            .Formula = _
            "=(" & finalAmountCell.Address & "*" & desiredPgPercentageCell.Address & ")" & _
            "-(" & flavorsSumFormula & _
            "+(" & nicPgPercentageCell.Address & "*C" & getRowNrOfComponent(NIC_BASE_RESULT_ID) & "))"
        End With
        With LiquidCalculatorSheet.Range("D" & pgResultRowNr)
            .NumberFormat = "0.00" & Space(1) & Chr$(34) & " g" & Chr$(34)
            .Formula = "=C" & pgResultRowNr & "*'" & IngredientsInfoSheet.Name & "'!C4"
        End With
        With LiquidCalculatorSheet.Range("E" & pgResultRowNr)
            .NumberFormat = "0.0 %"
            .Formula = "=C" & pgResultRowNr & "/C" & getRowNrOfComponent(RECIPE_TOTALS_ID)
        End With
        With LiquidCalculatorSheet.Range("F" & pgResultRowNr)
            .NumberFormat = "0.00" & Space(1) & Chr$(34) & " $" & Chr$(34)
            .Formula = getVlookupFormulaForIngredientsCost(pgResultRowNr)
        End With
        With LiquidCalculatorSheet.Range("B" & pgResultRowNr & ":F" & pgResultRowNr)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            .Interior.Color = BASE_COLOR
        End With
        LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(VG_PERCENTAGE_RESULT_ID)).Formula = _
        "=(C" & vgResultRowNr & "+C" & getRowNrOfComponent(NIC_BASE_RESULT_ID) & _
        "*" & nicVgPercentageCell.Address & ")/" & finalAmountCell.Address
    Else
        LiquidCalculatorSheet.Rows(baseVgPercentageCell.Row & ":" & basePgPercentageCell.Row).EntireRow.Hidden = False
        With LiquidCalculatorSheet.Range("D11:" & desiredPgPercentageCell.Address)
            .Clear
            .Interior.Color = BACKGROUND_COLOR
        End With
        LiquidCalculatorSheet.Rows(getRowNrOfComponent(PG_RESULT_ID)).EntireRow.Delete
        Dim baseResultRowNr As Integer: baseResultRowNr = getRowNrOfComponent(VG_RESULT_ID)
        With LiquidCalculatorSheet.Range("G" & baseResultRowNr)
            .Value = BASE_RESULT_ID
        End With
        With LiquidCalculatorSheet.Range("B" & baseResultRowNr)
            .Font.Bold = True
            .Value = "Base"
        End With
        lastRowInResultTable = IIf( _
        getNrOfFlavors(FLAVOR_RESULT_ID) = 0, getRowNrOfComponent(NIC_BASE_RESULT_ID), getRowNrOfComponent(FLAVOR_RESULT_ID, False))
        With LiquidCalculatorSheet.Range("C" & baseResultRowNr)
            .NumberFormat = "0.00" & Space(1) & Chr$(34) & " ml" & Chr$(34)
            .Formula = "=" & finalAmountCell.Address & _
            "-SUM(C" & getRowNrOfComponent(NIC_BASE_RESULT_ID) & ":C" & lastRowInResultTable & ")"
        End With
        With LiquidCalculatorSheet.Range("D" & baseResultRowNr)
            .NumberFormat = "0.00" & Space(1) & Chr$(34) & " g" & Chr$(34)
            .Formula = "=C" & baseResultRowNr & _
            "*(" & baseVgPercentageCell.Address & "*'" & IngredientsInfoSheet.Name & "'!C3" & _
            "+" & basePgPercentageCell.Address & "*'" & IngredientsInfoSheet.Name & "'!C4)"
        End With
        With LiquidCalculatorSheet.Range("E" & baseResultRowNr)
            .NumberFormat = "0.0 %"
            .Formula = "=C" & baseResultRowNr & "/C" & getRowNrOfComponent(RECIPE_TOTALS_ID)
        End With
        With LiquidCalculatorSheet.Range("F" & baseResultRowNr)
            .NumberFormat = "0.00" & Space(1) & Chr$(34) & " $" & Chr$(34)
            .Formula = getVlookupFormulaForIngredientsCost(baseResultRowNr)
        End With
        LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(VG_PERCENTAGE_RESULT_ID)).Formula = _
        "=(C" & getRowNrOfComponent(BASE_RESULT_ID) & "*" & baseVgPercentageCell.Address & _
        "+C" & getRowNrOfComponent(NIC_BASE_RESULT_ID) & "*" & nicVgPercentageCell.Address & ")" & _
        "/" & finalAmountCell.Address
    End If
    
    recipeTotalsRowNr = getRowNrOfComponent(RECIPE_TOTALS_ID)
    firstRowInResultTable = IIf( _
    isSeparateVgPgCbChecked, getRowNrOfComponent(VG_RESULT_ID), getRowNrOfComponent(BASE_RESULT_ID))
    lastRowInResultTable = IIf( _
    getNrOfFlavors(FLAVOR_RESULT_ID) = 0, getRowNrOfComponent(NIC_BASE_RESULT_ID), getRowNrOfComponent(FLAVOR_RESULT_ID, False))
    LiquidCalculatorSheet.Range("C" & recipeTotalsRowNr).Formula = "=SUM(C" & firstRowInResultTable & ":C" & lastRowInResultTable & ")"
    LiquidCalculatorSheet.Range("C" & recipeTotalsRowNr).Copy
    LiquidCalculatorSheet.Range("D" & recipeTotalsRowNr & ":F" & recipeTotalsRowNr).PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
End Sub

Function isSeparateVgPgCbChecked() As Boolean
    isSeparateVgPgCbChecked = LiquidCalculatorSheet.CheckBoxes("CbSeparateVgPg").Value = 1
End Function

Sub editIngredientsInfoClick()
    ThisWorkbook.Worksheets(IngredientsInfoSheet.Name).Visible = True
    ThisWorkbook.Worksheets(IngredientsInfoSheet.Name).Activate
End Sub

Function getRowNrOfComponent(componentId As String, Optional ByVal findFirst As Boolean = True) As Integer
    Dim rowNr, startRowNr, endRowNr, forStep As Integer
    If findFirst Then
        startRowNr = LiquidCalculatorSheet.Range("G1").End(xlDown).Row
        endRowNr = LiquidCalculatorSheet.Range("G1048576").End(xlUp).Row
        forStep = 1
    Else
        startRowNr = LiquidCalculatorSheet.Range("G1048576").End(xlUp).Row
        endRowNr = LiquidCalculatorSheet.Range("G1").End(xlDown).Row
        forStep = -1
    End If
    
    For rowNr = startRowNr To endRowNr Step forStep
        If StrComp(LiquidCalculatorSheet.Range("G" & rowNr).Value, componentId) = 0 Then
            getRowNrOfComponent = rowNr
            Exit Function
        End If
    Next rowNr
    
    getRowNrOfComponent = -1
End Function

Function getNrOfFlavors(flavorTypeId As String) As Integer
    Dim rowNr, nrOfFlavors As Integer
    nrOfFlavors = 0
    
    For rowNr = LiquidCalculatorSheet.Range("G1").End(xlDown).Row To LiquidCalculatorSheet.Range("G1048576").End(xlUp).Row
        If StrComp(LiquidCalculatorSheet.Range("G" & rowNr).Value, flavorTypeId) = 0 Then
            nrOfFlavors = nrOfFlavors + 1
        End If
    Next rowNr
    
    getNrOfFlavors = nrOfFlavors
End Function

Function getMinimumPgInFinalAmount() As Double
    Dim flavorsPgPercentageInFinalAmount As Double: flavorsPgPercentageInFinalAmount = _
    LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(TOTAL_FLAVOR_PERCENTAGE_RESULT)).Value
    Dim nicBasePgPercentageInFinalAmount As Double: nicBasePgPercentageInFinalAmount = _
    (LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(NIC_BASE_RESULT_ID)).Value / finalAmountCell.Value) * nicPgPercentageCell.Value
    Dim minimumPgInRecipe As Double: minimumPgInRecipe = flavorsPgPercentageInFinalAmount + nicBasePgPercentageInFinalAmount
    
    getMinimumPgInFinalAmount = minimumPgInRecipe
End Function

Function getMinimumVgInFinalAmount() As Double
    Dim nicBaseVgPercentageInFinalAmount As Double: nicBaseVgPercentageInFinalAmount = _
    (LiquidCalculatorSheet.Range("C" & getRowNrOfComponent(NIC_BASE_RESULT_ID)).Value / finalAmountCell.Value) * nicVgPercentageCell.Value
    
    getMinimumVgInFinalAmount = nicBaseVgPercentageInFinalAmount
End Function

Function getVlookupFormulaForIngredientsCost(lookupValueRowNr As Integer) As String
    'Not using XLOOKUP just for compatibility concerns
    If StrComp(LiquidCalculatorSheet.Range("G" & lookupValueRowNr).Value, FLAVOR_RESULT_ID) = 0 Then
        getVlookupFormulaForIngredientsCost = "=IF(ISNA(" & _
        "VLOOKUP(B" & lookupValueRowNr & "," & FlavorsSheet.Name & "!A:A,1,0))," & _
        "0," & _
        "VLOOKUP(B" & lookupValueRowNr & "," & FlavorsSheet.Name & "!A:B,2,0)" & _
        "*C" & lookupValueRowNr & _
        ")"
        Exit Function
    Else
        getVlookupFormulaForIngredientsCost = "=IF(ISNA(" & _
        "VLOOKUP(B" & lookupValueRowNr & ",'" & IngredientsInfoSheet.Name & "'!$B$8:$B$11,1,0))," & _
        "0," & _
        "VLOOKUP(B" & lookupValueRowNr & ",'" & IngredientsInfoSheet.Name & "'!$B$8:$F$11,5,0)" & _
        "*C" & lookupValueRowNr & _
        ")"
        Exit Function
    End If
End Function

Sub displayAutoDismissByTimerMsgBox(title As String, message As String, secondsUntilDismiss As Integer)
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    Select Case InfoBox.Popup(message & " (this window will close shortly on its own)", _
    secondsUntilDismiss, title, 0)
        Case 1, -1
            Exit Sub
    End Select
End Sub

Sub deleteAllFlavors()
    Dim firstFlavorInputRow, idx As Integer
    firstFlavorInputRow = getRowNrOfComponent(FLAVOR_INPUT_ID)
    For idx = 1 To getNrOfFlavors(FLAVOR_INPUT_ID)
        Call deleteFlavor(firstFlavorInputRow)
    Next idx
End Sub
