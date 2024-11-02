<H1>DEMO</H1>

You can find my example document [there](https://github.com/Surrogate-TM/surrogate-tm.github.io/blob/master/forum/CorpTemplate_v.7.40.vst).    
For operations with tables use this ribbons tab.     
![My custom ribbons tab](https://i.imgur.com/wiuAxAT.png)
1. Import data (from Excel range to Visio table)
2. Export data (from Visio table to new Excel workbook)
3. Delete Visio table from document

<h1>Code</h1>

        Sub fill_table() ' fill the specification
        Dim DocCell As Cell
        Dim FTx As Integer
        Dim FTy As Integer
        pNumber = 1
        Set DocCell = ActiveDocument.DocumentSheet.Cells("user.coc")
        DocCell.FormulaU = 1
        Dim pec As Integer ' counter of the number of specification lines on the page
        pec = 1
        Dim Mast As Master
        Dim pg As Page
        spn = Check_Spec_PageName(True)
        ActiveWindow.Page = ActiveDocument.Pages(spn)
        Set pg = ActivePage
        Set Mast = ActiveDocument.Masters.Item("Спецификация")
        pg.Drop Mast, 6.889764, 8.661417
        Dim target As Shape ' target shape
        Dim main As Shape ' shape - main group
        Dim rw As Shape ' shape - row
        Dim rn As String ' name of the shape-row
        Set main = ActivePage.Shapes.Item("Спецификация")
        ActivePage.Shapes.ItemFromID(1).Cells("Prop.tnum").Formula = "=thedoc!user.coc"
          Dim SSS As Shapes ' subset of shapes of the main group
        Dim tn As String ' name of the target shape
        Set SSS = main.Shapes
        For FTy = 1 To rc
        rn = "row" & pec
        Set rw = SSS.Item(rn)
        For FTx = 1 To sc
        tn = pec & "." & FTx
        Set target = rw.Shapes.Item(tn)
        If FTx = 2 Or FTx = 9 Then target.CellsSRC(visSectionParagraph, 0, visHorzAlign).FormulaU = "0"
        'target.Text = arr(FTy, FTx)
        target.Text = ttl(FTy, FTx)
        Next FTx
        If main.CellsSRC(visSectionUser, 7, visUserValue) = 2 Then
        For xx = 1 To sc
            tn = pec & "." & xx
            Set target = rw.Shapes.Item(tn)
            target.Text = " "
            target.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = "0 mm"
            Next xx
            FTy = FTy - 1
            DeleteUnusedRows main, pec
            pec = 0
           pNumber = pNumber + 1
           DocCell.Formula = pNumber
           Dim aPage As Page
           Set aPage = AddNamedPage(spn & "." & pNumber)
           ActivePage.PageSheet.Cells("PageWidth").Formula = "420 MM"
          ActivePage.PageSheet.Cells("PageHeight").Formula = "297 MM"
          ActivePage.PageSheet.Cells("Paperkind").Formula = 8
          ActivePage.PageSheet.Cells("PrintPageOrientation").Formula = 2
          ActivePage.Shapes(1).Cells("Prop.cnum.value") = pNumber
          ActivePage.Shapes.ItemFromID(1).Cells("prop.chapter.value").Formula = """C-Specification of equipment, cable products and materials"""
          ActivePage.Shapes(1).Cells("fields.value").FormulaU = "=pagenumber()" & "-1"
          ActivePage.Drop Mast, 6.889764, 8.661417
        Set main = ActivePage.Shapes.Item("Спецификация")
        Set SSS = main.Shapes
          If pNumber = 1 Then
          ActivePage.Shapes(1).Cells("Prop.cnum.value") = 1
          ActivePage.Shapes(1).Cells("user.n.value") = 3
        ' ActivePage.Shapes(1).Cells("Prop.tnum.value").Formula = "=thedoc!user.coc"
          Else
          ActivePage.Shapes(1).Cells("Prop.cnum.value") = pNumber
          ActivePage.Shapes(1).Cells("user.n.value") = 6
          ActivePage.Shapes(1).Cells("Prop.tnum.value").Formula = "=thedoc!user.coc"
            End If
           Else
             End If
        If main.CellsSRC(visSectionUser, 7, visUserValue) = 1 And FTy <> rc Then ' it's here !!!
            DeleteUnusedRows main, pec + 1
            pec = 0
           pNumber = pNumber + 1
           DocCell.Formula = pNumber
           Set aPage = AddNamedPage(spn & "." & pNumber)
           ActivePage.PageSheet.Cells("PageWidth").Formula = "420 MM"
          ActivePage.PageSheet.Cells("PageHeight").Formula = "297 MM"
          ActivePage.PageSheet.Cells("Paperkind").Formula = 8
          ActivePage.PageSheet.Cells("PrintPageOrientation").Formula = 2
          ActivePage.Shapes(1).Cells("Prop.cnum.value") = pNumber
          ActivePage.Shapes.ItemFromID(1).Cells("prop.chapter.value").Formula = """C-Sp specification of equipment, cable products and materials"""
          ActivePage.Shapes(1).Cells("fields.value").FormulaU = "=pagenumber()" & "-1"
          ActivePage.Drop Mast, 6.889764, 8.661417
        Set main = ActivePage.Shapes.Item("Спецификация")
        Set SSS = main.Shapes
          If pNumber = 1 Then
        ' ActivePage.Shapes(1).Cells("Prop.tnum.value").Formula = "=thedoc!user.coc"
          ActivePage.Shapes(1).Cells("Prop.cnum.value") = 1
          ActivePage. Shapes(1).Cells("user.n.value") = 3
          Else
          ActivePage.Shapes(1).Cells("Prop.tnum.value").Formula = "=thedoc!user.coc"
          ActivePage.Shapes( 1).Cells("Prop.cnum.value") = pNumber
          ActivePage.Shapes(1).Cells("user.n.value") = 6
            End If
           Else
          End If
        pec = pec + 1
        If pec > 30 Then pec = 0
        Next FTy
        DeleteUnusedRows main, pec
        If pNumber = 6 Then MsgBox "Attention"
        'sp.Close
        'Set sp = Nothing
        'oExcel.Quit
        pNumber = 1
        End Sub
