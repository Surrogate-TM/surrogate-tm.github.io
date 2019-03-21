Attribute VB_Name = "m_SourceCode"

'--------------Ìîäóëü õðàíèòü ïðîöåäóðû äëÿ ýêñïîðòà êîäà VBA è èñõîäíèêà ôàéëà âî âíåøíèå ìîäóëè-------------

'------------------Íóæåí ÷òîáû áûëà âîçìîæíîñòü êîììèòèòü êîä ÷åðåç ÃèòÕàá------------------

Public Sub SaveSourceCode()



Dim targetPath As String

    

    targetPath = GetCodePath

    ExportVBA targetPath

    ExportDocState targetPath

    MsgBox "Èñõîäíûé êîä ýêñïîðòèðîâàí"



End Sub