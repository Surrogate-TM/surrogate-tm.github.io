        Sub Draw_Table()
        ' this code initially posted at https://visio.getbb.ru/viewtopic.php?f=15&t=233
        Dim Mast As Master
        Dim ololo As Shape
        Dim ooo As Window
        Dim x As Integer
        Dim y As Integer
        Dim r As Shape
        Dim cn As String
        Dim rn As String
        Const rh = 0.31496063 ' row height 8 mm in inches
        Dim xc(9) As Single
        xc(0) = 0
        xc(1) = 0.787401575 ' column width 20 mm in inches
        xc( 2) = 5.905511811
        xc(3) = 8.267716535
        xc(4) = 9.645669291
        xc(5) = 11.41732283
        xc(6) = 12.20472441
        xc(7) = 12.99212598
        xc(8) = 13.97637795
        xc(9) = 15.5511811
        n = 1
        Dim rect As Shape
        Dim rw As Shape
        Dim rs As Selection
        Dim piny As String
        piny = "=0 mm"
        Set ooo = ActiveWindow
        For y = 1 To 30 ' Loop to create 30 rows
            Set rs = Nothing
            Set rw = ActiveWindow.Shape.DrawRectangle(0, 0, 0, 0) ' draw a rectangle of zero width and height
            rnm = "pos" & y ' define the name of the current lines as pos + line number
            rw.Name = rnm ' assign the resulting name to the line
            ActiveWindow.Selection.ConvertToGroup ' convert the shape into a group
            rw.OpenDrawWindow.Activate ' enter this group
            Set ooo = ActiveWindow
            
            For x = 1 To 9 ' start filling the row with shape columns
                tn = y & "." & x ' name of the current rectangle shape, consists of the line number.line number
                tx = xc(x - 1) ' left coordinate of the rectangle
                bx = xc(x) ' right lower coordinate of the rectangle
                ty = 0 + rh * (y - 1 ) ' lower coordinate of the rectangle
                by = 0 + rh * (y) ' upper coordinate of the rectangle
                Set rect = ooo.Shape.DrawRectangle(tx, ty, bx, by) ' draw a rectangle with the previously obtained coordinates
                rect.Style = "None" ' I disable the
                rect.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY) styles .FormulaU = "Height*1" ' set the PinY binding of the rectangle to the top edge of the shape
                rect.CellsSRC(visSectionParagraph, 0, visSpaceLine).FormulaU = "8 mm" ' make the distance between the lines of text of the rectangle 8 mm
                rect.AddSection visSectionUser ' add the User section
                rect.AddRow visSectionUser, visRowLast, visTagDefault ' add a row to the User section
                rect.CellsSRC(visSectionUser, 0, visUserValue).FormulaU = "=user.row_1.prompt*INT(TEXTHEIGHT(TheText,Width)/8 mm)*8 mm" ' enter the value in the Value field
                rect.CellsSRC(visSectionObject, visRowText, visTxtBlkVerticalAlign).FormulaU = "=if(height=8 mm,1,0)" ' enter the value in the Value field 
                rect.CellsSRC(visSectionObject, visRowText, visTxtBlkTopMargin).FormulaU = "0 pt" ' set zero indentation at the top of the shape text 
                rect.CellsSRC(visSectionObject, visRowText, visTxtBlkBottomMargin).FormulaU = "0 pt" ' set zero indentation at the top of the shape text 
                rect.Name = tn ' assign the previously obtained name to the shape, consisting of the line number.line number 
                rect.Text = y & "." & x ' we enter into the shape text, which consists of the line number.number in the line 
                rect.CellsSRC(visSectionUser, 0, visUserPrompt).FormulaU = "if(strsame(shapetext(thetext),"" ""),0,1)" ' we fill in the value into the Prompt field 
            Next x 
            Dim tex As String 
            ooo.Close 
            Application.ActiveWindow.Selection.UpdateAlignmentBox ' we align the sizes of the group, in accordance with the sizes of the child shapes 
            hrow = "=guard(max(sheet." & n + 1 & "!user.row_1,sheet." & n + 2 & "!user.row_1,sheet." & n + 3 & "!user.row_1,sheet." & n + 4 & "!user.row_1,sheet." & n + 5 &     "!user.row_1,sheet." & n + 6 & "!user.row_1,sheet." & n + 7 & "!user.row_1,sheet." & n + 8 & "!user.row_1,sheet." & n + 9 & "!user.row_1))" ' calculate the maximum height of the child shape 
            If n > 10 Then piny = "=guard(sheet." & n - 10 & "!piny-sheet." & n - 10 & "!height)" 
            rw.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = hrow
            rw.CellsSRC (visSection Object, visRowXFormOut, visXFormLocPinY).FormulaU = "Height*1" 
            rw.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).FormulaU = "Width*0" 
            rw.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).FormulaU = piny 
            n = n + 10 
        Next y 
        Application.ActiveWindow.SelectAll ActiveWindow.Selection.Group 
        End Sub
