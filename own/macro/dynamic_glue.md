

```Sub DymamicGlue()
Dim Rect1 As Shape, Rect2 As Shape
Dim CL1 As Shape, CL2 As Shape, CL3 As Shape

Set Rect1 = ActivePage.DrawRectangle(1, 1, 2, 2) ' draw 1st rectangle
Set Rect2 = ActivePage.DrawRectangle(4, 1, 5, 2) ' draw 2nd rectangle
Set CL1 = Application.ActiveWindow.Page.Drop(ActiveDocument.Masters.ItemU("Dynamic connector"), 0#, 0#) ' add 1st connector

Dim vsoCell1 As Visio.Cell
Dim vsoCell2 As Visio.Cell
Set vsoCell1 = CL1.CellsU("BeginX")
Set vsoCell2 = Rect1.CellsSRC(1, 1, 0)
vsoCell1.GlueTo vsoCell2
Set vsoCell1 = CL1.CellsU("EndX")
Set vsoCell2 = Rect2.CellsSRC(1, 1, 0)
vsoCell1.GlueTo vsoCell2
CL1.Cells("LineColor").FormulaU = "2" ' для демонстрации красим 1 линию в красный цвет
    
Set CL2 = Application.ActiveWindow.Page.Drop(ActiveDocument.Masters.ItemU("Dynamic connector"), 0#, 0#)

Dim vsoCell3 As Visio.Cell
Dim vsoCell4 As Visio.Cell
Set vsoCell3 = CL2.CellsU("BeginX")
Set vsoCell4 = Rect1.CellsSRC(1, 1, 0)
vsoCell3.GlueTo vsoCell4
Set vsoCell3 = CL2.CellsU("EndX")
Set vsoCell4 = Rect2.CellsSRC(1, 1, 0)  ' PinX
vsoCell3.GlueTo vsoCell4
CL2.Cells("LineColor").FormulaU = "3"  ' для демонстрации красим 2 линию в зеленый цвет
    
Set CL3 = Application.ActiveWindow.Page.Drop(ActiveDocument.Masters.ItemU("Dynamic connector"), 0#, 0#)

Set vsoCell3 = CL3.CellsU("BeginX")
Set vsoCell4 = Rect1.CellsSRC(1, 1, 0)
vsoCell3.GlueTo vsoCell4
Set vsoCell3 = CL3.CellsU("EndX")
Set vsoCell4 = Rect2.CellsSRC(1, 1, 0)
vsoCell3.GlueTo vsoCell4
CL3.Cells("LineColor").FormulaU = "4"  ' для демонстрации красим 3 линию в синий цвет
    
ActiveWindow.DeselectAll
End Sub```
