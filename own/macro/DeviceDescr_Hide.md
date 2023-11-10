This code can hide description for shape **Hardware** from my [stencil Rack](https://surrogate-tm.gitbook.io/my-stencils/racks/stencil-for-create-rack-diagram).    
![](https://i.ytimg.com/vi/o9Vr19PKPCM/hqdefault_313166.jpg?sqp=-oaymwEcCNACELwBSFXyq4qpAw4IARUAAIhCGAFwAcABBg==&rs=AOn4CLC8Gj3jLRqta_BqNNde-o_IhMmNKw)    

```Sub DeviceDescr_Hide()
Dim sl As Selection, sh As Shape, ssh As Shape
Set sl = ActiveWindow.Selection
For Each sh In sl
    Set ssh = sh.Shapes(1)
    ssh.Cells("HideText").Formula = "True"
Next
MsgBox "TheEnd!!!"
End Sub
```
