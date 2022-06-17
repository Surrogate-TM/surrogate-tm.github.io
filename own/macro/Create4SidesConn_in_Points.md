# Create4SidesConn_in_Points
Код добавляет соединительные точки с каждой из 4 сторон фигуры

```
Sub Create4SidesConn_in_Points()
Dim main As Shape
Dim val As String
Dim n As Integer
Set main = ActiveWindow.Selection.Item(1)
main.AddSection visSectionConnectionPts
For x = 0 To 1
main.AddRow visSectionConnectionPts, visRowLast, visTagDefault
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctX).FormulaU = "width*" & x
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctY).FormulaForceU = "Height*1/2"
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctDirX).FormulaForceU = "0 mm"
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctDirY).FormulaForceU = "0 mm"
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctType).FormulaForceU = 0
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctAutoGen).FormulaForceU = "0 mm"
main.CellsSRC(visSectionConnectionPts, visRowLast, 6).FormulaForceU = ""
Next x
For y = 0 To 1
main.AddRow visSectionConnectionPts, visRowLast, visTagDefault
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctX).FormulaU = "width*0.5"
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctY).FormulaForceU = "Height*" & y
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctDirX).FormulaForceU = "0 mm"
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctDirY).FormulaForceU = "0 mm"
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctType).FormulaForceU = 0
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctAutoGen).FormulaForceU = "0 mm"
main.CellsSRC(visSectionConnectionPts, visRowLast, 6).FormulaForceU = ""
Next y
End Sub
```
