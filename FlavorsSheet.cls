Option Explicit

Private Const BACKGROUND_COLOR As Long = 4210752

Sub addFlavor()
    Dim newFlavorRow As Integer: newFlavorRow = FlavorsSheet.Range("C1048576").End(xlUp).Offset(1).Row
    Dim currentFlavorIndexInAllHistory As Integer: currentFlavorIndexInAllHistory = IIf( _
    newFlavorRow = 2, 1, FlavorsSheet.Range("H1").Value + 1)
    
    FlavorsSheet.Range("A" & newFlavorRow).Formula = "=CONCAT(C" & newFlavorRow & ","" - "",D" & newFlavorRow & ")"
    FlavorsSheet.Range("B" & newFlavorRow).Formula = "=F" & newFlavorRow & "/E" & newFlavorRow
    FlavorsSheet.Range("C" & newFlavorRow).Value = "New Flavor Name"
    FlavorsSheet.Range("D" & newFlavorRow).Value = "Company"
    FlavorsSheet.Range("E" & newFlavorRow).Value = 10
    FlavorsSheet.Range("F" & newFlavorRow).Value = 15
    With FlavorsSheet.Range("A" & newFlavorRow & ":F" & newFlavorRow)
        .Interior.Color = BACKGROUND_COLOR
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlHairline
            .Color = vbWhite
        End With
    End With
    
    Call createDeleteFlavorButton(FlavorsSheet.Range("G" & newFlavorRow), currentFlavorIndexInAllHistory)
    FlavorsSheet.Range("H1").Value = currentFlavorIndexInAllHistory
    FlavorsSheet.Range("C" & newFlavorRow).Select
    Application.SendKeys ("{F2}{HOME}+{END}")
End Sub

Sub createDeleteFlavorButton(buttonRange As Range, buttonIndex As Integer)
    Dim deleteButton As Button: Set deleteButton = FlavorsSheet.Buttons.Add( _
    buttonRange.Left, _
    buttonRange.Top, _
    buttonRange.Width / 2, _
    buttonRange.Height)
    With deleteButton
        .OnAction = "FlavorsSheet.deleteFlavor"
        .Caption = "X"
        .Name = "BtnDelFlvDb" & buttonIndex
        .Font.Bold = True
        .Font.Color = vbRed
    End With
End Sub

Sub deleteFlavor()
    Dim flavorRowNumber As Integer: flavorRowNumber = FlavorsSheet.Shapes(Application.Caller).TopLeftCell.Row
    Dim msgBoxDeleteConfirmation As Integer
    
    msgBoxDeleteConfirmation = MsgBox("Are you sure you want to delete " & FlavorsSheet.Range("A" & flavorRowNumber) & " ?", _
    vbQuestion + vbYesNo + vbDefaultButton2, _
    "Careful!")
    If msgBoxDeleteConfirmation = vbYes Then
        FlavorsSheet.Rows(flavorRowNumber).EntireRow.Delete
    End If
End Sub
