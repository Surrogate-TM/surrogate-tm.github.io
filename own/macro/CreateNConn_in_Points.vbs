dim va
set va = getobject(, "visio.application")
dim vw
Dim main ' As Shape
Dim val ' As String
Dim n ' As Integer
n = InputBox("ââåäèòå êîëè÷åñòâî òî÷åê")
Set main = va.ActiveWindow.Selection(1)
'main.AddSection visSectionConnectionPts
For x = 1 To n
For y = 0 To 1
main.AddRow visSectionConnectionPts, visRowLast, visTagDefault
valx = "Width*" & x & "/" & n + 1
valy = "Height*" & y
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctX).FormulaU = valx
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctY).FormulaForceU = valy
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctDirX).FormulaForceU = "0 mm"
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctDirY).FormulaForceU = "0 mm"
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctType).FormulaForceU = 0
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctAutoGen).FormulaForceU = "0 mm"
main.CellsSRC(visSectionConnectionPts, visRowLast, 6).FormulaForceU = ""
Next y
Next x
For y = 1 To n
For x = 0 To 1
main.AddRow visSectionConnectionPts, visRowLast, visTagDefault
valy = "Height*" & y & "/" & n + 1
valx = "Width*" & x
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctX).FormulaU = valx
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctY).FormulaForceU = valy
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctDirX).FormulaForceU = "0 mm"
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctDirY).FormulaForceU = "0 mm"
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctType).FormulaForceU = 0
main.CellsSRC(visSectionConnectionPts, visRowLast, visCnnctAutoGen).FormulaForceU = "0 mm"
main.CellsSRC(visSectionConnectionPts, visRowLast, 6).FormulaForceU = ""
Next x
Next y
set va = nothing

