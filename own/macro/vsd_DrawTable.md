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
    Const rh = 0.31496063 ' высота строки 8 мм в дюймах
    Dim xc(9) As Single
    xc(0) = 0
    xc(1) = 0.787401575 ' ширина столбца 20 мм в дюймах
    xc(2) = 5.905511811
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
    For y = 1 To 30 ' Цикл по созданию 30 строк
    Set rs = Nothing
    Set rw = ActiveWindow.Shape.DrawRectangle(0, 0, 0, 0) ' рисуем прямоугольник нулевой ширины и высоты
    rnm = "pos" & y ' определяем имя текущей строки как pos + номер строки
    rw.Name = rnm ' присваиваем полученное имя строке
    ActiveWindow.Selection.ConvertToGroup ' преобразуем шейп в группу
    rw.OpenDrawWindow.Activate ' входим внутрь данной группы
    Set ooo = ActiveWindow
    
    For x = 1 To 9 ' начинаем заполнять строку шейпами-столбцами
    tn = y & "." & x '  имя текущего шейпа-прямоугольника, состоит из номера строки.номера в строке
    tx = xc(x - 1) ' левая координата прямоугольника
    bx = xc(x) ' правая нижняя координата прямоугольника
    ty = 0 + rh * (y - 1) ' нижняя координата прямоугольника
    by = 0 + rh * (y) ' верхняя координата прямоугольника
    Set rect = ooo.Shape.DrawRectangle(tx, ty, bx, by) ' рисуем прямоугольник с полученными ранее координатами
    rect.Style = "None" ' отключаю стили
    rect.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).FormulaU = "Height*1" ' ставим привязку PinY прямоугольника по верхнему краю шейпа
    rect.CellsSRC(visSectionParagraph, 0, visSpaceLine).FormulaU = "8 mm" ' делаем расстояние между строками текста прямоугольника 8 мм
    rect.AddSection visSectionUser ' добавляем секцию User
    rect.AddRow visSectionUser, visRowLast, visTagDefault ' добавляем строку в секцию User
    rect.CellsSRC(visSectionUser, 0, visUserValue).FormulaU = "=user.row_1.prompt*INT(TEXTHEIGHT(TheText,Width)/8 mm)*8 mm" ' забиваем значение в поле Value
    rect.CellsSRC(visSectionObject, visRowText, visTxtBlkVerticalAlign).FormulaU = "=if(height=8 mm,1,0)" ' забиваем значение в поле Value
    rect.CellsSRC(visSectionObject, visRowText, visTxtBlkTopMargin).FormulaU = "0 pt" ' ставим нулевой отступ сверху в тексте шейпа
    rect.CellsSRC(visSectionObject, visRowText, visTxtBlkBottomMargin).FormulaU = "0 pt" ' ставим нулевой отступ сверху в тексте шейпа
    rect.Name = tn ' присваиваем шейпу полученное ранее имя, состоит из номера строки.номера в строке
    rect.Text = y & "." & x ' вписываем в шейп текст, который состоит из номера строки.номера в строке
    rect.CellsSRC(visSectionUser, 0, visUserPrompt).FormulaU = "if(strsame(shapetext(thetext),"" ""),0,1)" ' забиваем значение в поле Prompt
    Next x
    Dim tex As String
    ooo.Close
    Application.ActiveWindow.Selection.UpdateAlignmentBox ' производим выравнивание размеров группы, в соответствии с размерами дочерних шейпов
    hrow = "=guard(max(sheet." & n + 1 & "!user.row_1,sheet." & n + 2 & "!user.row_1,sheet." & n + 3 & "!user.row_1,sheet." & n + 4 & "!user.row_1,sheet." & n + 5 & "!user.row_1,sheet." & n + 6 & "!user.row_1,sheet." & n + 7 & "!user.row_1,sheet." & n + 8 & "!user.row_1,sheet." & n + 9 & "!user.row_1))" ' вычисляем максимальную высоту дочернего шейпа
    If n > 10 Then piny = "=guard(sheet." & n - 10 & "!piny-sheet." & n - 10 & "!height)"
    rw.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = hrow
    rw.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinY).FormulaU = "Height*1"
    rw.CellsSRC(visSectionObject, visRowXFormOut, visXFormLocPinX).FormulaU = "Width*0"
    rw.CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).FormulaU = piny
    n = n + 10
    Next y
    Application.ActiveWindow.SelectAll
    ActiveWindow.Selection.Group
    End Sub
