# Код для добавления пользовательской ленты в MS Visio 2010+
```
Sub vsd_import_xml()
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("d:\dropbox\Visio_UI.xml", 1)
XML = file.ReadAll
file.Close
ActiveDocument.CustomUI = XML
ActiveDocument.Save
End Sub
```

# Образец xml-файла для создания ленты
```
<customUI onLoad="Ribbon_Load" xmlns="http://schemas.microsoft.com/office/2009/07/customui">
    <ribbon>
        <tabs>
            <tab id="tabMyAddin" label="My Addin">
                <group id="groupTools" label="Tools">
                    <button id="buttonSelectedColor" imageMso="SmartArtChangeColorsGallery" label="Selected Color" size="large" onAction="SelectedColor"/>
                    <toggleButton id="toggleSample" label="Toggle Sample" showImage="false" onAction="ToggleSample"/>
                    <checkBox id="checkBoxSample" label="Checkbox Sample" onAction="CheckedSample"/>
                    <comboBox id="comboBoxColors" label="Colors" showImage="false" onChange="ColorsChanged" getEnabled="ColorsEnabled">
                        <item id="__id1" label="Red"/>
                        <item id="__id2" label="Blue"/>
                        <item id="__id3" label="Green"/>
                    </comboBox>
                </group>
                <group id="groupHelp" label="Help">
                    <button id="buttonContents" imageMso="FunctionsLogicalInsertGallery" label="Contents" size="large" onAction="ShowHelp"/>
                    <button id="buttonAbout" imageMso="HappyFace" label="About" size="large" onAction="ShowAbout"/>
                </group>
            </tab>
        </tabs>
    </ribbon>
</customUI>
```
