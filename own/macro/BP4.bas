Attribute VB_Name = "BP4"
Option Base 1
Sub BP4_corrector(shpObj As Visio.Shape, pp As Integer)
Dim isSpec As Boolean
isSpec = False
Dim ma() As Integer
Dim r%, form$
r = ActiveDocument.Pages.Count
ReDim ma(r)
Dim pg As Page, sh As Shape, listing$, wn As Window, N%, pos As Shape, prim As Shape
listing = "": N = 0
For i = pp To ActiveDocument.Pages.Count
Set pg = ActiveDocument.Pages(i)
pg.Shapes("Рамка").Cells("fields.value").FormulaU = "0"
pg.Shapes("Рамка").Cells("fields.value").FormulaU = "=PAGENUMBER()-1"
If InStr(1, pg.Shapes("Рамка").Cells("prop.chapter").ResultStr(""), "-Спец") Then
isSpec = True
Else
End If
If pg.Shapes("Рамка").Cells("prop.cnum").FormulaU = 1 Then listing = listing & ";" & pg.Name: N = N + 1: ma(N) = pg.Shapes("Рамка").ID
Next
Set pg = ActiveDocument.Pages(pp)
Set wn = Application.ActiveWindow.Page.PageSheet.OpenSheetWindow
Application.ActiveWindow.Shape.Cells("user.store").FormulaU = Chr(34) & listing & Chr(34)
wn.Close
Set sh = shpObj
For i = 1 To N
Set prim = sh.Shapes("pos" & i).Shapes(3)
Set pos = prim.Parent
pos.Cells("prop.det.format").FormulaForceU = "GUARD(ThePage!User.store)"
pos.Cells("prop.det.value").FormulaForceU = "INDEX(" & i & " ,Prop.det.Format)"
form = "IF(0=0,SETF(GetRef(User.ch)," & Chr(34) & "=Pages[" & Chr(34) & "&Prop.det&" & Chr(34) & "]!sheet." & ma(i) & "!user.ch" & Chr(34) & ")+SETF(GetRef(User.de)," & Chr(34) & "=Pages[" & Chr(34) & "&Prop.det&" & Chr(34) & "]!sheet." & ma(i) & "!user.de" & Chr(34) & ")+SETF(GetRef(User.pn)," & Chr(34) & "=Pages[" & Chr(34) & "&Prop.det&" & Chr(34) & "]!sheet." & ma(i) & "!fields.value" & Chr(34) & "),33)"
pos.Cells("user.set").FormulaU = form
pos.CellsSRC(visSectionAction, 0, visActionAction).FormulaU = "GOTOPAGE(Prop.det)"
pos.CellsSRC(visSectionAction, 0, visActionMenu).FormulaU = """Перейти на ""&Prop.det"
Next
If isSpec Then
N = N - 1
Set prim = sh.Shapes("pos" & N).Shapes("prim" & N)
Set pos = prim.Parent
prim.Cells("user.row_6").FormulaU = ""
'Debug.Print prim.Name, "IF(Sheet." & pos.ID & "!User.N=Sheet." & sh.ID & "!Prop.N,SETF(getref(User.n),Sheet." & pos.ID + 4 & "!User.pn)),SETF(getref(User.n),Sheet." & pos.ID + 4 & "!User.pn))"
prim.Cells("user.row_6").FormulaU = "IF(Sheet." & pos.ID & "!User.N=Sheet." & sh.ID & "!Prop.N,SETF(getref(User.n),Sheet." & pos.ID + 4 & "!User.pn),SETF(getref(User.n),Sheet." & pos.ID + 4 & "!User.pn))"
Else
End If
sh.Cells("prop.n").Formula = N
MsgBox "TheEnd!"
End Sub
