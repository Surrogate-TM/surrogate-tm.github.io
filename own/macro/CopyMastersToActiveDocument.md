# CopyMastersToActiveDocument
Код размещается в трафарете, предназначен для копирования мастер-шейпов из трафарета в document stencil (локальный набор элементов) активного документа
````
Sub CopyMastersToActiveDocument()
Dim mst As Master
For Each mst In Me.Masters
ActiveDocument.Masters.Drop mst, 0#, 0#
Next
End Sub
````
