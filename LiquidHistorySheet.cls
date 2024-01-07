Option Explicit

Private Const FLAVOR_INPUT_ID As String = "flavorInput"

Sub updateSteepTimeOfAllLiquids()
    Dim liquidDescriptionAsJson As Object
    Dim lastLiquidRow As Integer: lastLiquidRow = LiquidHistorySheet.Range("C1048576").End(xlUp).Row
    Dim rowNr As Integer
    
    For rowNr = 3 To lastLiquidRow
        Set liquidDescriptionAsJson = JsonConverter.ParseJson(LiquidHistorySheet.Range("H" & rowNr).Value)
        If Len(liquidDescriptionAsJson("finishDate")) = 0 Then
            Dim steepSecondsTotal, steepHours, steepDays As Long
            steepSecondsTotal = DateDiff("s", CDate(liquidDescriptionAsJson("dateOfMaking")), Now())
            steepDays = Application.WorksheetFunction.RoundDown(steepSecondsTotal / 86400, 0)
            steepHours = Format((steepSecondsTotal / 3600) - (24 * steepDays), 0)
            LiquidHistorySheet.Range("E" & rowNr).Value = IIf(steepDays > 0, steepDays & " days, ", "") & steepHours & "h"
        End If
    Next rowNr
End Sub

Sub remixLiquid()
    Dim currentLiquidRowNumber As Integer: currentLiquidRowNumber = LiquidHistorySheet.Shapes(Application.Caller).TopLeftCell.Row
    Dim liquidDescriptionAsJson As Object: Set liquidDescriptionAsJson = _
    JsonConverter.ParseJson(LiquidHistorySheet.Range("H" & currentLiquidRowNumber).Value)
    
    If Not LiquidCalculatorSheet.isSeparateVgPgCbChecked Then
        LiquidCalculatorSheet.CheckBoxes("CbSeparateVgPg").Value = 1
        Call LiquidCalculatorSheet.checkBoxSeparateVgPgClick
    End If
    
    Call LiquidCalculatorSheet.deleteAllFlavors
    LiquidCalculatorSheet.Range("C11").Value = liquidDescriptionAsJson("amount")
    LiquidCalculatorSheet.Range("C12").Value = liquidDescriptionAsJson("nicotine")
    
    Dim flavor As Object
    Dim currentFlavorInputRowNr As Integer
    For Each flavor In liquidDescriptionAsJson("flavors")
        Call LiquidCalculatorSheet.addFlavor
        currentFlavorInputRowNr = LiquidCalculatorSheet.getRowNrOfComponent(FLAVOR_INPUT_ID, False)
        LiquidCalculatorSheet.Range("B" & currentFlavorInputRowNr).Value = flavor("flavorName")
        LiquidCalculatorSheet.Range("C" & currentFlavorInputRowNr).Value = flavor("percentage")
    Next flavor
    
    LiquidCalculatorSheet.Range("E11").Value = liquidDescriptionAsJson("finalVg")
    LiquidCalculatorSheet.Range("E12").Value = liquidDescriptionAsJson("finalPg")
    LiquidCalculatorSheet.Activate
End Sub

Sub startLiquidButtonClick()
    Dim currentLiquidRowNumber As Integer: currentLiquidRowNumber = LiquidHistorySheet.Shapes(Application.Caller).TopLeftCell.Row
    Dim liquidDescriptionAsJson As Object: Set liquidDescriptionAsJson = _
    JsonConverter.ParseJson(LiquidHistorySheet.Range("H" & currentLiquidRowNumber).Value)
    Dim currentLiquidDescription As String: currentLiquidDescription = LiquidHistorySheet.Range("D" & currentLiquidRowNumber).Value
    Dim currentDateTime As Date: currentDateTime = Now
    
    Dim currentLiquidDescriptionBeforeStartTag As String
    If InStr(currentLiquidDescription, "Started at") = 0 Then
        currentLiquidDescriptionBeforeStartTag = currentLiquidDescription & vbCrLf
    Else
        currentLiquidDescriptionBeforeStartTag = Left(currentLiquidDescription, InStr(currentLiquidDescription, "Started at") - 1)
    End If

    LiquidHistorySheet.Range("D" & currentLiquidRowNumber).Value = _
    currentLiquidDescriptionBeforeStartTag & "Started at: " & Format(currentDateTime, "dd/mm/yyyy || hh:mm")
    LiquidHistorySheet.Buttons(Application.Caller).Font.Color = vbGreen
    LiquidHistorySheet.Buttons(Application.Caller).Caption = "Started"
    With LiquidHistorySheet.Rows(currentLiquidRowNumber)
        .AutoFit
        .RowHeight = .RowHeight + 20
    End With
    liquidDescriptionAsJson("startDate") = Format(currentDateTime, "dd/mm/yyyy  hh:mm:ss")
    liquidDescriptionAsJson("finishDate") = ""
    LiquidHistorySheet.Range("H" & currentLiquidRowNumber).Value = JsonConverter.ConvertToJson(liquidDescriptionAsJson)
    
    Dim finishButtonName As String: finishButtonName = "BtnFinishLiquid" & Right(Application.Caller, Len(Application.Caller) - Len("BtnStartLiquid"))
    LiquidHistorySheet.Buttons(finishButtonName).Font.Color = vbRed
    LiquidHistorySheet.Buttons(finishButtonName).Caption = "Not Finished"
End Sub

Sub finishLiquidButtonClick(Optional ByVal currentLiquidRowNumber As Integer = -1)
    If currentLiquidRowNumber = -1 Then
        currentLiquidRowNumber = LiquidHistorySheet.Shapes(Application.Caller).TopLeftCell.Row
    End If
    Dim liquidDescriptionAsJson As Object: Set liquidDescriptionAsJson = _
    JsonConverter.ParseJson(LiquidHistorySheet.Range("H" & currentLiquidRowNumber).Value)
    Dim currentLiquidDescription As String: currentLiquidDescription = LiquidHistorySheet.Range("D" & currentLiquidRowNumber).Value
    Dim currentLiquidDescriptionBeforeFinishTag As String
    Dim currentDateTime As Date: currentDateTime = Now
    
    If Len(liquidDescriptionAsJson("startDate")) = 0 Then
        MsgBox "You need to start your e-liquid first!", vbCritical
        Exit Sub
    End If
    If InStr(currentLiquidDescription, "Finished at") = 0 Then
        currentLiquidDescriptionBeforeFinishTag = currentLiquidDescription & vbCrLf
    Else
        currentLiquidDescriptionBeforeFinishTag = Left(currentLiquidDescription, InStr(currentLiquidDescription, "Finished at") - 1)
    End If
    
    LiquidHistorySheet.Range("D" & currentLiquidRowNumber).Value = _
    currentLiquidDescriptionBeforeFinishTag & "Finished at: " & Format(currentDateTime, "dd/mm/yyyy || hh:mm")
    Dim finishLiquidButtonName As String: finishLiquidButtonName = IIf(InStr(Application.Caller, "BtnSeeLiquidStats") = 0, _
    Application.Caller, _
    Replace(Application.Caller, "BtnSeeLiquidStats", "BtnFinishLiquid"))
    LiquidHistorySheet.Buttons(finishLiquidButtonName).Font.Color = vbGreen
    LiquidHistorySheet.Buttons(finishLiquidButtonName).Caption = "Finished"
    With LiquidHistorySheet.Rows(currentLiquidRowNumber)
        .AutoFit
        .RowHeight = .RowHeight + 20
    End With
    liquidDescriptionAsJson("finishDate") = Format(currentDateTime, "dd/mm/yyyy  hh:mm:ss")
    LiquidHistorySheet.Range("H" & currentLiquidRowNumber).Value = JsonConverter.ConvertToJson(liquidDescriptionAsJson)
End Sub

Sub seeLiquidStats()
    Dim currentLiquidRowNumber As Integer: currentLiquidRowNumber = LiquidHistorySheet.Shapes(Application.Caller).TopLeftCell.Row
    Dim liquidDescriptionAsJson As Object: Set liquidDescriptionAsJson = _
    JsonConverter.ParseJson(LiquidHistorySheet.Range("H" & currentLiquidRowNumber).Value)
    
    If Len(liquidDescriptionAsJson("startDate")) = 0 And Len(liquidDescriptionAsJson("finishDate")) = 0 Then
        MsgBox "You need to start, then finish your liquid!", vbCritical
        Exit Sub
    ElseIf Len(liquidDescriptionAsJson("finishDate")) = 0 Then
        Dim msgBoxFinishLiquid As Integer: msgBoxFinishLiquid = MsgBox("You need to finish your liquid! Do you want to set it as Finished now?", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Error!")
        If msgBoxFinishLiquid = vbYes Then
            Call finishLiquidButtonClick(currentLiquidRowNumber)
        Else
            Exit Sub
        End If
    End If
    
    Set liquidDescriptionAsJson = JsonConverter.ParseJson(LiquidHistorySheet.Range("H" & currentLiquidRowNumber).Value)
    Dim startDate, finishDate As Date
    Dim secondsLastedTotal, hoursLasted, daysLasted As Long
    startDate = CDate(liquidDescriptionAsJson("startDate"))
    finishDate = CDate(liquidDescriptionAsJson("finishDate"))
    secondsLastedTotal = DateDiff("s", startDate, finishDate)
    daysLasted = Application.WorksheetFunction.RoundDown(secondsLastedTotal / 86400, 0)
    hoursLasted = Format((secondsLastedTotal / 3600) - (24 * daysLasted), 0)
    
    Dim mlPerDay, mlPerHour, nicMgPerDay, nicMgPerHour, finalAmount, finalNicStrength As Double
    finalAmount = liquidDescriptionAsJson("amount")
    finalNicStrength = liquidDescriptionAsJson("nicotine")
    mlPerDay = Format(finalAmount / (secondsLastedTotal / 86400), "0.0")
    mlPerHour = Format(finalAmount / (secondsLastedTotal / 3600), "0.0")
    nicMgPerDay = Format((finalAmount * finalNicStrength) / (secondsLastedTotal / 86400), "0.0")
    nicMgPerHour = Format((finalAmount * finalNicStrength) / (secondsLastedTotal / 3600), "0.00")
    
    ' TODO4future Maybe more stats?
    MsgBox _
    "Lasted: " & IIf(daysLasted > 0, daysLasted & " days, ", "") & hoursLasted & "h" & vbNewLine & _
    "Ml per day: " & mlPerDay & " ml" & vbNewLine & _
    "Ml per hour: " & mlPerHour & " ml" & vbNewLine & _
    "Nicotine mg per day: " & nicMgPerDay & " mg" & vbNewLine & _
    "Nicotine mg per hour: " & nicMgPerHour & " mg", _
    vbOKOnly, "STATS"
End Sub

Sub deleteLiquid()
    Dim msgBoxDeleteConfirmation As Integer
    Dim currentLiquidRowNumber As Integer: currentLiquidRowNumber = LiquidHistorySheet.Shapes(Application.Caller).TopLeftCell.Row
    msgBoxDeleteConfirmation = MsgBox("Are you sure you want to delete liquid #" & currentLiquidRowNumber - 2 & " ?", _
    vbQuestion + vbYesNo + vbDefaultButton2, _
    "Careful!")
    If msgBoxDeleteConfirmation = vbYes Then
        LiquidHistorySheet.Rows(currentLiquidRowNumber).EntireRow.Delete
    End If
End Sub
