VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formInsert 
   Caption         =   "Insert Fields"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2310
   OleObjectBlob   =   "formInsert.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmRun_Click()
    Dim shp As Shape, shp2 As Shape
    Dim i As Integer
    
    Select Case ActiveWindow.Selection.Count
    Case 0:
        MsgBox "You need to select a shape to perform the operation"
        Exit Sub
    Case 1:
        Set shp = ActiveWindow.Selection(1)
        ActiveWindow.DeselectAll
        ActiveWindow.Select shp, visSelect
        ActiveWindow.Selection.Copy
        On Error Resume Next
        shp.ConvertToGroup
        On Error GoTo 0
        If shp.OneD = False Then
        shp.OpenDrawWindow.Activate
        ActiveWindow.Shape.Paste
        ActiveWindow.SelectAll
        Set shp2 = ActiveWindow.Selection(ActiveWindow.Selection.Count)
            shp2.Cells("PinX").Formula = shp.NameU & "!width*0.5"
            shp2.Cells("PinY").Formula = shp.NameU & "!height*0.5"

        shp2.DeleteSection visSectionProp
        shp2.SendToBack
        ActiveWindow.Close
        shp.DeleteSection visSectionFirstComponent
        End If
        If shp.SectionExists(visSectionProp, False) Then
            listProps.Clear
            For i = 0 To shp.Section(visSectionProp).Count - 1
                listProps.AddItem shp.Section(visSectionProp).Row(i).Name
            Next i
        Else
            MsgBox "Your shape does not have shape data"
            Exit Sub
        End If
    Case Else:
        MsgBox "please select only one shape"
        Exit Sub
    End Select
    
End Sub
Private Sub cmInsert_Click()
    Dim shp As Shape, shp2 As Shape
    Dim i As Integer
    
    Select Case ActiveWindow.Selection.Count
    Case 0:
        MsgBox "You need to select a shape to perform the operation"
        Exit Sub
    Case 1:
        Set shp = ActiveWindow.Selection(1)
        If shp.SectionExists(visSectionProp, False) Then
            For i = 0 To listProps.ListCount - 1
                If listProps.Selected(i) Then
                    insertField shp, listProps.List(i)
                End If
            Next i
        Else
            MsgBox "Your shape does not have shape data"
            Exit Sub
        End If
    Case Else:
        MsgBox "please select only one shape"
        Exit Sub
    End Select

End Sub

Sub insertField(shp As Shape, field As String)
Dim shp2 As Shape
Dim shpChars As Visio.Characters

    If Not shp.SectionExists(visSectionControls, False) Then
        shp.AddSection visSectionControls
    End If
    shp.AddRow visSectionControls, visRowLast, visTagDefault

    Set shp2 = shp.DrawRectangle(0, 0, 1, 1)
    shp2.TextStyle = "Normal"
    shp2.LineStyle = "Text Only"
    shp2.FillStyle = "Text Only"
    Set shpChars = shp2.Characters
    shpChars.Begin = 0
    shpChars.End = 0
    shpChars.AddCustomFieldU "sheet." & shp.ID & "!prop." & field, visFmtNumGenNoUnits
    shp2.Cells("PinX").Formula = "sheet." & shp.ID & "!controls.row_" & shp.RowCount(visSectionControls)
    shp2.Cells("PinY").Formula = "sheet." & shp.ID & "!controls.row_" & shp.RowCount(visSectionControls) & ".Y"
    shp2.Cells("Width").FormulaU = "textwidth(thetext)"
    shp2.Cells("Height").FormulaU = "textheight(thetext, textwidth(thetext))"

End Sub
