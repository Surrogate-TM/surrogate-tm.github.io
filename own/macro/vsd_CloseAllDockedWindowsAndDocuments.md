```
Sub vsd_CloseAllDockedWindowsAndDocuments()
Dim va As visio.Application, vd As visio.Document, j%
Set va = CreateObject("visio.application")
va.Visible = True
va.AlertResponse = 7
Dim myFile As String
Dim path As String
Dim ext() As Variant
Dim x%, n%
n = 0
path = "i:\Motiv\plan\" '
If Right(path, 1) <> "\" Then path = path + "\"
ext = Array("*.vsd") 
For i = 0 To UBound(ext)
myFile = Dir$(path & ext(i))
While myFile <> ""
myFile = Dir$()
Set vd = va.Documents.Open(path & myFile)
For j = va.Documents.Count To 1 Step -1
If Right(va.Documents(j).Name, 3) = "VSS" Then va.Documents(j).Close ' çàêðûâàåì íàáîðû-äîêóìåíòû
Next j
For j = va.Windows.Count To 1 Step -1
If Right(va.Windows(j).Caption, 4) <> "plan" Then va.Windows(j).Close ' çàêðûâàåì íàáîðû-îêíà
Next j
vd.Save
vd.Close
Set vd = Nothing
Wend
Next
va.AlertResponse = 0
va.Quit
Set va = Nothing
End Sub
```