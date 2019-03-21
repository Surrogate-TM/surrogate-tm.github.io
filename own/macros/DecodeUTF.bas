Attribute VB_Name = "DecodeUTF"
Sub tesss()
Dim dn As String, fff$
fff = "вЂ” РєР°С‚РµРіРѕСЂРёСЏ С„СѓРЅРєС†РёР№ С‚Р°Р±Р»РёС†С‹ СЃРІРѕР№СЃС‚РІ (ShapeSheet)"
dn = DecodeUTF8(fff)

'SaveTextToFile Cells(9, 4), dn, "utf-8noBOM"
Debug.Print dn
End Sub
Function DecodeUTF8(s)
' URL:
    Dim i, c, n, b1, b2, b3

    i = 1
    Do While i <= Len(s)
        c = Asc(Mid(s, i, 1))
        If (c And &HC0) = &HC0 Then
            n = 1
            Do While i + n <= Len(s)
                If (Asc(Mid(s, i + n, 1)) And &HC0) <> &H80 Then
                    Exit Do
                End If
                n = n + 1
            Loop
            If n = 2 And ((c And &HE0) = &HC0) Then
                b1 = Asc(Mid(s, i + 1, 1)) And &H3F
                b2 = c And &H1F
                c = b1 + b2 * &H40
            ElseIf n = 3 And ((c And &HF0) = &HE0) Then
                b1 = Asc(Mid(s, i + 2, 1)) And &H3F
                b2 = Asc(Mid(s, i + 1, 1)) And &H3F
                b3 = c And &HF
                c = b3 * &H1000 + b2 * &H40 + b1
            Else
                ' ?????? ?????? U+FFFF ??? ???????????? ??????????????????
                c = &HFFFD
            End If
            s = Left(s, i - 1) + ChrW(c) + Mid(s, i + n)
        ElseIf (c And &HC0) = &H80 Then
            ' ??????????? ???????????? ????
            s = Left(s, i - 1) + ChrW(&HFFFD) + Mid(s, i + 1)
        End If
        i = i + 1
    Loop
    DecodeUTF8 = s
End Function



Function SaveTextToFile(ByVal txt$, ByVal filename$, Optional ByVal encoding$ = "windows-1251") As Boolean
    ' функция сохраняет текст txt в кодировке Charset$ в файл filename$
    On Error Resume Next: Err.Clear
    Select Case encoding$
 
        Case "windows-1251", "", "ansi"
            Set fso = CreateObject("scripting.filesystemobject")
            Set ts = fso.CreateTextFile(filename, True)
            ts.Write txt: ts.Close
            Set ts = Nothing: Set fso = Nothing
 
        Case "utf-16", "utf-16LE"
            Set fso = CreateObject("scripting.filesystemobject")
            Set ts = fso.CreateTextFile(filename, True, True)
            ts.Write txt: ts.Close
            Set ts = Nothing: Set fso = Nothing
 
        Case "utf-8noBOM"
            With CreateObject("ADODB.Stream")
                .Type = 2: .Charset = "utf-8": .Open
                .WriteText txt$
 
                Set binaryStream = CreateObject("ADODB.Stream")
                binaryStream.Type = 1: binaryStream.Mode = 3: binaryStream.Open
                .Position = 3: .CopyTo binaryStream        'Skip BOM bytes
                .flush: .Close
                binaryStream.SaveToFile filename$, 2
                binaryStream.Close
            End With
 
        Case Else
            With CreateObject("ADODB.Stream")
                .Type = 2: .Charset = encoding$: .Open
                .WriteText txt$
                .SaveToFile filename$, 2        ' сохраняем файл в заданной кодировке
                .Close
            End With
    End Select
    SaveTextToFile = Err = 0: DoEvents
End Function

